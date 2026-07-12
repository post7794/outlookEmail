import importlib
import os
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from types import SimpleNamespace
from unittest.mock import patch


os.environ.setdefault('SECRET_KEY', 'test-secret-key')
os.environ.setdefault('ACCOUNT_HEALTH_WORKER_ENABLED', 'false')
if 'DATABASE_PATH' not in os.environ:
    root = tempfile.mkdtemp(prefix='outlookEmail-health-tests-')
    os.environ['DATABASE_PATH'] = os.path.join(root, 'test.db')

web_outlook_app = importlib.import_module('web_outlook_app')
UTC = timezone.utc


class AccountHealthWorkerTests(unittest.TestCase):
    def setUp(self):
        self.app = web_outlook_app.app
        self.app.config['TESTING'] = True
        with self.app.app_context():
            web_outlook_app.init_db()
            db = web_outlook_app.get_db()
            db.execute('DELETE FROM account_refresh_logs')
            db.execute('DELETE FROM accounts')
            db.commit()

    def _insert_account(self, *, status='active', account_type='outlook', due=True):
        now = datetime(2026, 7, 13, 0, 0, tzinfo=UTC)
        next_check = now - timedelta(minutes=1) if due else now + timedelta(hours=1)
        with self.app.app_context():
            db = web_outlook_app.get_db()
            cursor = db.execute(
                '''
                INSERT INTO accounts (
                    email, password, client_id, refresh_token, status, account_type,
                    provider, health_enrolled_at, next_health_check_at,
                    health_status, consecutive_auth_failures, transient_failure_count
                ) VALUES (?, '', 'client', ?, ?, ?, 'outlook', ?, ?, 'healthy', 0, 0)
                ''',
                (
                    f'user-{status}-{account_type}-{due}@outlook.com',
                    web_outlook_app.encrypt_data('refresh-token'),
                    status,
                    account_type,
                    now.isoformat(),
                    next_check.isoformat(),
                ),
            )
            db.commit()
            return int(cursor.lastrowid), now

    def _row(self, account_id):
        db = web_outlook_app.get_db()
        return db.execute('SELECT * FROM accounts WHERE id = ?', (account_id,)).fetchone()

    def test_due_query_excludes_future_inactive_and_imap(self):
        active_id, now = self._insert_account()
        self._insert_account(due=False)
        self._insert_account(status='inactive')
        self._insert_account(account_type='imap')
        with self.app.app_context():
            rows = web_outlook_app.load_due_health_accounts(now=now, limit=20)
            self.assertEqual([row['id'] for row in rows], [active_id])

    def test_success_rotates_token_and_resets_counters(self):
        account_id, now = self._insert_account()
        with self.app.app_context():
            db = web_outlook_app.get_db()
            db.execute(
                '''UPDATE accounts SET consecutive_auth_failures=2,
                   transient_failure_count=2, health_status='suspect' WHERE id=?''',
                (account_id,),
            )
            db.commit()
            row = self._row(account_id)
            with patch.object(web_outlook_app, 'probe_refresh_token', return_value={
                'success': True,
                'result_class': 'success',
                'error_code': '',
                'error_message': '',
                'rotated_refresh_token': 'rotated-token',
            }), patch.object(web_outlook_app, 'probe_imap_mailbox_access', return_value={
                'success': True,
                'result_class': 'success',
                'rotated_refresh_token': '',
            }):
                result = web_outlook_app.check_account_health(row, db=db, now=now)
            updated = self._row(account_id)
            self.assertEqual(result['state'], 'healthy')
            self.assertEqual(updated['health_status'], 'healthy')
            self.assertEqual(updated['consecutive_auth_failures'], 0)
            self.assertEqual(updated['transient_failure_count'], 0)
            self.assertEqual(web_outlook_app.decrypt_data(updated['refresh_token']), 'rotated-token')

    def test_real_imap_token_helper_returns_rotated_refresh_token(self):
        response = SimpleNamespace(
            status_code=200,
            json=lambda: {
                'access_token': 'imap-access-token',
                'refresh_token': 'imap-rotated-token',
            },
        )
        with patch.object(web_outlook_app, 'request_imap_token_response', return_value=response):
            result = web_outlook_app.get_access_token_imap_result('client', 'old-token')
        self.assertTrue(result['success'])
        self.assertEqual(result['rotated_refresh_token'], 'imap-rotated-token')

    def test_imap_invalid_scope_is_operational_not_auth(self):
        token_result = {
            'success': False,
            'status_code': 400,
            'oauth_error': 'invalid_scope',
            'oauth_error_description': 'The requested scope is invalid.',
            'error': {
                'code': 'IMAP_TOKEN_FAILED',
                'message': '获取访问令牌失败',
                'status': 400,
            },
        }
        with patch.object(web_outlook_app, 'get_access_token_imap_result', return_value=token_result):
            result = web_outlook_app.probe_imap_mailbox_access(
                'user@outlook.com', 'client', 'refresh-token'
            )
        self.assertFalse(result['success'])
        self.assertEqual(result['result_class'], 'operational')
        self.assertEqual(result['error_code'], 'invalid_scope')

    def test_transient_failure_never_increments_auth_count(self):
        account_id, now = self._insert_account()
        with self.app.app_context():
            db = web_outlook_app.get_db()
            db.execute(
                'UPDATE accounts SET consecutive_auth_failures=2 WHERE id=?',
                (account_id,),
            )
            db.commit()
            row = self._row(account_id)
            with patch.object(web_outlook_app, 'probe_refresh_token', return_value={
                'success': False,
                'result_class': 'transient',
                'error_code': 'http_429',
                'error_message': 'rate limited',
                'rotated_refresh_token': '',
            }):
                web_outlook_app.check_account_health(row, db=db, now=now)
            updated = self._row(account_id)
            self.assertEqual(updated['health_status'], 'transient')
            self.assertEqual(updated['consecutive_auth_failures'], 0)
            self.assertEqual(updated['transient_failure_count'], 1)

    def test_imap_preflight_failure_counts_as_auth_failure(self):
        account_id, now = self._insert_account()
        with self.app.app_context():
            db = web_outlook_app.get_db()
            row = self._row(account_id)
            with patch.object(web_outlook_app, 'probe_refresh_token', return_value={
                'success': True,
                'result_class': 'success',
                'rotated_refresh_token': '',
            }), patch.object(web_outlook_app, 'probe_imap_mailbox_access', return_value={
                'success': False,
                'result_class': 'auth',
                'error_code': 'IMAP_AUTH_FAILED',
                'error_message': 'authentication failed',
            }):
                web_outlook_app.check_account_health(row, db=db, now=now)
            updated = self._row(account_id)
            self.assertEqual(updated['health_status'], 'suspect')
            self.assertEqual(updated['consecutive_auth_failures'], 1)

    def test_third_auth_failure_quarantines_account(self):
        account_id, now = self._insert_account()
        failure = {
            'success': False,
            'result_class': 'auth',
            'error_code': 'invalid_grant',
            'error_message': 'refresh token revoked',
            'rotated_refresh_token': '',
        }
        with self.app.app_context():
            db = web_outlook_app.get_db()
            with patch.object(web_outlook_app, 'probe_refresh_token', return_value=failure):
                for offset in range(3):
                    row = self._row(account_id)
                    result = web_outlook_app.check_account_health(
                        row,
                        db=db,
                        now=now + timedelta(hours=2 * offset),
                    )
            updated = self._row(account_id)
            self.assertEqual(result['state'], 'quarantined')
            self.assertEqual(updated['status'], 'inactive')
            self.assertEqual(updated['health_status'], 'quarantined')
            self.assertEqual(updated['consecutive_auth_failures'], 3)
            self.assertIsNotNone(updated['health_delete_after_at'])

    def test_quarantined_account_can_recover(self):
        account_id, now = self._insert_account(status='inactive')
        with self.app.app_context():
            db = web_outlook_app.get_db()
            db.execute(
                '''
                UPDATE accounts
                SET health_status='quarantined', consecutive_auth_failures=3,
                    next_health_check_at=?, quarantined_at=?
                WHERE id=?
                ''',
                ((now - timedelta(minutes=1)).isoformat(), now.isoformat(), account_id),
            )
            db.commit()
            due = web_outlook_app.load_due_health_accounts(db=db, now=now)
            self.assertEqual([row['id'] for row in due], [account_id])
            with patch.object(web_outlook_app, 'probe_refresh_token', return_value={
                'success': True,
                'result_class': 'success',
                'rotated_refresh_token': '',
            }), patch.object(web_outlook_app, 'probe_imap_mailbox_access', return_value={
                'success': True,
                'result_class': 'success',
                'rotated_refresh_token': '',
            }):
                web_outlook_app.check_account_health(due[0], db=db, now=now)
            updated = self._row(account_id)
            self.assertEqual(updated['status'], 'active')
            self.assertEqual(updated['health_status'], 'healthy')
            self.assertEqual(updated['consecutive_auth_failures'], 0)


if __name__ == '__main__':
    unittest.main()
