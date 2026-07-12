"""Pure account-health policy helpers.

The worker deliberately keeps policy decisions independent from Flask and
SQLite so boundary conditions and failure classification can be tested without
starting the web application.
"""

from __future__ import annotations

import hashlib
import re
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, Iterable, Optional


TRANSIENT_HTTP_STATUSES = {408, 425, 429}
AUTH_ERRORS = {
    "invalid_grant",
    "interaction_required",
    "consent_required",
    "account_selection_required",
}
OPERATIONAL_ERRORS = {
    "invalid_client",
    "unauthorized_client",
    "invalid_scope",
}


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def parse_timestamp(value: Any) -> Optional[datetime]:
    if isinstance(value, datetime):
        parsed = value
    else:
        text = str(value or "").strip()
        if not text:
            return None
        if text.endswith("Z"):
            text = text[:-1] + "+00:00"
        try:
            parsed = datetime.fromisoformat(text)
        except ValueError:
            return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(timezone.utc)


def format_timestamp(value: datetime) -> str:
    return value.astimezone(timezone.utc).replace(microsecond=0).isoformat()


def health_interval_for_age(enrolled_at: Any, now: Optional[datetime] = None) -> timedelta:
    current = parse_timestamp(now) or utc_now()
    enrolled = parse_timestamp(enrolled_at) or current
    age = max(timedelta(0), current - enrolled)
    if age < timedelta(hours=24):
        return timedelta(hours=2)
    if age < timedelta(hours=72):
        return timedelta(hours=6)
    if age < timedelta(days=7):
        return timedelta(hours=12)
    return timedelta(hours=24)


def deterministic_jitter_seconds(account_key: Any, maximum_seconds: int = 600) -> int:
    maximum = max(0, int(maximum_seconds or 0))
    if maximum == 0:
        return 0
    digest = hashlib.sha256(str(account_key or "").encode("utf-8")).digest()
    span = maximum * 2 + 1
    return int.from_bytes(digest[:4], "big") % span - maximum


def next_success_check_at(enrolled_at: Any, account_key: Any, now: Optional[datetime] = None,
                          jitter_seconds: Optional[int] = None) -> datetime:
    current = parse_timestamp(now) or utc_now()
    jitter = (
        deterministic_jitter_seconds(account_key)
        if jitter_seconds is None
        else int(jitter_seconds)
    )
    candidate = current + health_interval_for_age(enrolled_at, current) + timedelta(seconds=jitter)
    return max(candidate, current + timedelta(minutes=5))


def transient_retry_delay(failure_count: int) -> timedelta:
    count = max(1, int(failure_count or 1))
    if count == 1:
        return timedelta(minutes=5)
    if count == 2:
        return timedelta(minutes=30)
    return timedelta(hours=2)


def _aadsts_code(text: str) -> str:
    match = re.search(r"AADSTS(\d+)", text or "", re.IGNORECASE)
    return f"AADSTS{match.group(1)}" if match else ""


def classify_token_failure(*, status_code: int = 0, error: str = "",
                           description: str = "", exception: Any = None) -> Dict[str, str]:
    if exception is not None:
        name = type(exception).__name__
        lowered = f"{name} {exception}".lower()
        if any(marker in lowered for marker in (
            "timeout", "connection", "proxy", "dns", "temporarily unavailable"
        )):
            return {"result_class": "transient", "error_code": name}
        return {"result_class": "operational", "error_code": name}

    normalized_error = str(error or "").strip().lower()
    details = str(description or "").strip()
    lowered_details = details.lower()
    aadsts = _aadsts_code(details)

    if status_code in TRANSIENT_HTTP_STATUSES or status_code >= 500:
        return {"result_class": "transient", "error_code": normalized_error or f"http_{status_code}"}
    if normalized_error in {"temporarily_unavailable", "server_error"}:
        return {"result_class": "transient", "error_code": normalized_error}
    if normalized_error in OPERATIONAL_ERRORS:
        return {"result_class": "operational", "error_code": normalized_error}
    if normalized_error in AUTH_ERRORS:
        return {"result_class": "auth", "error_code": normalized_error}
    if aadsts in {
        "AADSTS50034",  # user not found
        "AADSTS50053",  # locked account
        "AADSTS50055",  # password expired
        "AADSTS50057",  # disabled account
        "AADSTS700082", # refresh token expired due to inactivity
        "AADSTS70000",
    }:
        return {"result_class": "auth", "error_code": aadsts}
    if any(marker in lowered_details for marker in (
        "refresh token has expired",
        "token is expired",
        "token has been revoked",
        "account is locked",
        "account is disabled",
        "authenticationfailed",
    )):
        return {"result_class": "auth", "error_code": aadsts or normalized_error or "auth_failed"}

    # Unknown failures must never cause destructive account lifecycle changes.
    return {"result_class": "transient", "error_code": aadsts or normalized_error or f"http_{status_code or 0}"}


def combine_probe_failures(results: Iterable[Dict[str, str]]) -> Dict[str, str]:
    rows = [dict(row or {}) for row in results]
    if not rows:
        return {"result_class": "transient", "error_code": "no_probe_result"}
    by_class = {str(row.get("result_class") or "transient") for row in rows}
    # A platform/configuration failure must not be blamed on the mailbox.
    if "operational" in by_class:
        chosen = next(row for row in rows if row.get("result_class") == "operational")
        return chosen
    # Any transient path makes a conflicting auth result inconclusive.
    if "transient" in by_class:
        chosen = next(row for row in rows if row.get("result_class") == "transient")
        return chosen
    return rows[0]

