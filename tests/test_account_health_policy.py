import unittest
from datetime import datetime, timedelta, timezone

from outlook_web.account_health import (
    classify_token_failure,
    combine_probe_failures,
    health_interval_for_age,
    next_success_check_at,
    transient_retry_delay,
)


UTC = timezone.utc


class AccountHealthPolicyTests(unittest.TestCase):
    def setUp(self):
        self.now = datetime(2026, 7, 13, 0, 0, tzinfo=UTC)

    def test_age_cohort_boundaries(self):
        cases = [
            (timedelta(hours=23, minutes=59), timedelta(hours=2)),
            (timedelta(hours=24), timedelta(hours=6)),
            (timedelta(hours=71, minutes=59), timedelta(hours=6)),
            (timedelta(hours=72), timedelta(hours=12)),
            (timedelta(days=6, hours=23), timedelta(hours=12)),
            (timedelta(days=7), timedelta(hours=24)),
        ]
        for age, expected in cases:
            with self.subTest(age=age):
                self.assertEqual(health_interval_for_age(self.now - age, self.now), expected)

    def test_success_next_check_supports_fixed_jitter(self):
        enrolled = self.now - timedelta(hours=1)
        self.assertEqual(
            next_success_check_at(enrolled, 'account-1', self.now, jitter_seconds=0),
            self.now + timedelta(hours=2),
        )

    def test_transient_backoff(self):
        self.assertEqual(transient_retry_delay(1), timedelta(minutes=5))
        self.assertEqual(transient_retry_delay(2), timedelta(minutes=30))
        self.assertEqual(transient_retry_delay(3), timedelta(hours=2))
        self.assertEqual(transient_retry_delay(20), timedelta(hours=2))

    def test_failure_classification(self):
        self.assertEqual(
            classify_token_failure(status_code=429, error='temporarily_unavailable')['result_class'],
            'transient',
        )
        self.assertEqual(
            classify_token_failure(status_code=400, error='invalid_grant')['result_class'],
            'auth',
        )
        self.assertEqual(
            classify_token_failure(status_code=400, error='invalid_client')['result_class'],
            'operational',
        )
        self.assertEqual(
            classify_token_failure(status_code=400, description='AADSTS50057: account disabled')['result_class'],
            'auth',
        )
        self.assertEqual(
            classify_token_failure(status_code=418, error='unknown')['result_class'],
            'transient',
        )

    def test_conflicting_auth_and_transient_is_transient(self):
        combined = combine_probe_failures([
            {'result_class': 'auth', 'error_code': 'invalid_grant'},
            {'result_class': 'transient', 'error_code': 'http_503'},
        ])
        self.assertEqual(combined['result_class'], 'transient')


if __name__ == '__main__':
    unittest.main()
