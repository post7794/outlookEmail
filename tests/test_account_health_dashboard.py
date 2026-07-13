import importlib
import os
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from unittest.mock import patch


os.environ.setdefault('SECRET_KEY', 'test-secret-key')
os.environ.setdefault('ACCOUNT_HEALTH_WORKER_ENABLED', 'false')
if 'DATABASE_PATH' not in os.environ:
    root = tempfile.mkdtemp(prefix='outlookEmail-health-dashboard-tests-')
    os.environ['DATABASE_PATH'] = os.path.join(root, 'test.db')

web_outlook_app = importlib.import_module('web_outlook_app')
UTC = timezone.utc


class AccountHealthDashboardTests(unittest.TestCase):
    def setUp(self):
        self.app = web_outlook_app.app
        self.app.config['TESTING'] = True
        self.app.config['WTF_CSRF_ENABLED'] = False
        self.client = self.app.test_client()
        with self.client.session_transaction() as sess:
            sess['logged_in'] = True
        with self.app.app_context():
            web_outlook_app.init_db()
            db = web_outlook_app.get_db()
            db.execute('DELETE FROM account_health_runs')
            db.execute('DELETE FROM account_refresh_logs')
            db.execute('DELETE FROM accounts')
            db.commit()
        with web_outlook_app.account_health_manual_task_lock:
            web_outlook_app.account_health_manual_task.update({
                'running': False,
                'started_at': '',
                'finished_at': '',
                'summary': None,
                'error': '',
            })

    def _insert_account(self, email, health_status='healthy', error_code='', due=False,
                        auth_failures=0, transient_failures=0):
        now = datetime.now(UTC)
        next_check = now - timedelta(minutes=1) if due else now + timedelta(hours=2)
        with self.app.app_context():
            db = web_outlook_app.get_db()
            cursor = db.execute(
                '''
                INSERT INTO accounts (
                    email, password, client_id, refresh_token, status,
                    account_type, provider, health_enrolled_at,
                    next_health_check_at, health_status,
                    consecutive_auth_failures, transient_failure_count,
                    last_health_error_code, remark
                ) VALUES (?, ?, 'client-id', ?, 'active', 'outlook', 'outlook',
                          ?, ?, ?, ?, ?, ?, 'dashboard test')
                ''',
                (
                    email,
                    web_outlook_app.encrypt_data('password-secret'),
                    web_outlook_app.encrypt_data('refresh-secret'),
                    now.isoformat(),
                    next_check.isoformat(),
                    health_status,
                    auth_failures,
                    transient_failures,
                    error_code or None,
                ),
            )
            account_id = int(cursor.lastrowid)
            db.execute(
                '''
                INSERT INTO account_refresh_logs (
                    account_id, account_email, refresh_type, status, error_message
                ) VALUES (?, ?, 'health_scheduled', ?, ?)
                ''',
                (
                    account_id,
                    email,
                    'success' if health_status == 'healthy' else 'failed',
                    '' if health_status == 'healthy' else f'auth:{error_code} safe message',
                ),
            )
            db.commit()
            return account_id

    def test_dashboard_requires_login(self):
        client = self.app.test_client()
        response = client.get('/api/account-health/dashboard')
        self.assertIn(response.status_code, (302, 401))

    def test_dashboard_returns_summary_filters_and_no_credentials(self):
        self._insert_account('healthy@example.com', 'healthy')
        self._insert_account(
            'suspect@example.com', 'suspect', 'invalid_grant', due=True,
            auth_failures=2,
        )
        self._insert_account(
            'transient@example.com', 'transient', 'TimeoutError',
            transient_failures=1,
        )

        response = self.client.get('/api/account-health/dashboard?limit=100')
        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertTrue(payload['success'])
        self.assertEqual(response.headers.get('Cache-Control'), 'no-store')
        self.assertEqual(payload['summary']['total'], 3)
        self.assertEqual(payload['summary']['healthy'], 1)
        self.assertEqual(payload['summary']['suspect'], 1)
        self.assertEqual(payload['summary']['transient'], 1)
        self.assertEqual(payload['summary']['due_now'], 1)
        self.assertEqual(payload['status_counts']['healthy'], 1)
        self.assertEqual(payload['status_counts']['suspect'], 1)
        self.assertEqual(payload['error_counts'][0]['count'], 1)
        body = response.get_data(as_text=True)
        self.assertNotIn('password-secret', body)
        self.assertNotIn('refresh-secret', body)
        self.assertNotIn('client-id', body)

        filtered = self.client.get(
            '/api/account-health/dashboard?status=suspect&error_code=invalid_grant&q=suspect'
        ).get_json()
        self.assertEqual(len(filtered['accounts']), 1)
        self.assertEqual(filtered['accounts'][0]['health_status'], 'suspect')
        self.assertEqual(filtered['accounts'][0]['last_health_error_code'], 'invalid_grant')

    def test_dashboard_clamps_limit(self):
        for index in range(3):
            self._insert_account(f'user{index}@example.com')
        payload = self.client.get('/api/account-health/dashboard?limit=1').get_json()
        self.assertEqual(len(payload['accounts']), 1)

    def test_manual_run_is_async_and_rejects_duplicate(self):
        fake_thread = unittest.mock.Mock()
        with patch.object(web_outlook_app, 'account_health_worker_enabled', return_value=True), \
                patch.object(web_outlook_app.threading, 'Thread', return_value=fake_thread):
            response = self.client.post('/api/account-health/run', json={'scope': 'due'})
        self.assertEqual(response.status_code, 202)
        fake_thread.start.assert_called_once_with()

        with patch.object(web_outlook_app, 'account_health_worker_enabled', return_value=True):
            duplicate = self.client.post('/api/account-health/run', json={'scope': 'due'})
        self.assertEqual(duplicate.status_code, 409)
        self.assertEqual(duplicate.get_json()['error_code'], 'health_run_busy')

    def test_manual_run_rejects_unsupported_scope(self):
        response = self.client.post('/api/account-health/run', json={'scope': 'all'})
        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.get_json()['error_code'], 'unsupported_scope')

    def test_manual_run_rejects_when_worker_disabled(self):
        with patch.object(web_outlook_app, 'account_health_worker_enabled', return_value=False):
            response = self.client.post('/api/account-health/run', json={'scope': 'due'})
        self.assertEqual(response.status_code, 409)
        self.assertEqual(response.get_json()['error_code'], 'health_worker_disabled')

    def test_health_run_is_persisted_for_dashboard(self):
        with patch.object(web_outlook_app, 'account_health_worker_enabled', return_value=True):
            summary = web_outlook_app.run_due_health_checks(trigger_type='manual')
        self.assertEqual(summary['status'], 'success')
        self.assertEqual(summary['selected'], 0)
        with self.app.app_context():
            row = web_outlook_app.get_db().execute(
                'SELECT * FROM account_health_runs ORDER BY id DESC LIMIT 1'
            ).fetchone()
        self.assertEqual(row['trigger_type'], 'manual')
        self.assertEqual(row['status'], 'success')
        self.assertIsNotNone(row['finished_at'])

        payload = self.client.get('/api/account-health/dashboard').get_json()
        self.assertEqual(payload['worker']['last_status'], 'success')
        self.assertEqual(payload['worker']['last_run']['trigger_type'], 'manual')
        self.assertEqual(len(payload['recent_runs']), 1)


if __name__ == '__main__':
    unittest.main()
