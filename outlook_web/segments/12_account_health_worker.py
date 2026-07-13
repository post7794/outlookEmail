"""Per-account health scheduling and lifecycle isolation."""

from outlook_web.account_health import (
    classify_token_failure,
    format_timestamp,
    next_success_check_at,
    parse_timestamp,
    transient_retry_delay,
    utc_now,
)


account_health_run_lock = threading.Lock()
IMAP_AUTHENTICATED_NOT_CONNECTED_MESSAGE = 'User is authenticated but not connected.'
account_health_manual_task_lock = threading.Lock()
account_health_manual_task = {
    'running': False,
    'started_at': '',
    'finished_at': '',
    'summary': None,
    'error': '',
}


def probe_imap_mailbox_access(email_addr: str, client_id: str, refresh_token: str,
                              proxy_url: str = None, fallback_proxy_urls=None):
    """Verify that a token can actually authenticate and select the Inbox."""
    token_result = get_access_token_imap_result(
        client_id,
        refresh_token,
        proxy_url,
        fallback_proxy_urls,
    )
    if not token_result.get('success'):
        error = token_result.get('error') or {}
        classification = classify_token_failure(
            status_code=int(token_result.get('status_code') or error.get('status') or 0),
            error=str(token_result.get('oauth_error') or ''),
            description=str(
                token_result.get('oauth_error_description')
                or error.get('details')
                or error.get('message')
                or ''
            ),
        )
        return {
            'success': False,
            'result_class': classification['result_class'],
            'error_code': str(classification['error_code'] or error.get('code') or 'IMAP_TOKEN_FAILED'),
            'error_message': sanitize_error_details(str(error.get('message') or 'IMAP Token 获取失败')),
        }

    mail = None
    try:
        with proxy_socket_context(proxy_url):
            mail = imaplib.IMAP4_SSL(IMAP_SERVER_NEW, IMAP_PORT, timeout=IMAP_TIMEOUT)
        access_token = str(token_result.get('access_token') or '')
        auth_string = f"user={email_addr}\1auth=Bearer {access_token}\1\1".encode('utf-8')
        mail.authenticate('XOAUTH2', lambda _challenge: auth_string)
        status, _response = mail.select('INBOX', readonly=True)
        if str(status or '').upper() != 'OK':
            return {
                'success': False,
                'result_class': 'auth',
                'error_code': 'IMAP_SELECT_FAILED',
                'error_message': 'IMAP 无法打开 INBOX',
            }
        noop_status, _noop_response = mail.noop()
        if str(noop_status or '').upper() != 'OK':
            return {
                'success': False,
                'result_class': 'transient',
                'error_code': 'IMAP_NOOP_FAILED',
                'error_message': 'IMAP NOOP 失败',
            }
        return {
            'success': True,
            'result_class': 'success',
            'error_code': '',
            'error_message': '',
            'rotated_refresh_token': str(token_result.get('rotated_refresh_token') or '').strip(),
        }
    except imaplib.IMAP4.error as exc:
        error_message = str(exc).strip()
        if error_message == IMAP_AUTHENTICATED_NOT_CONNECTED_MESSAGE:
            return {
                'success': False,
                'result_class': 'transient',
                'error_code': 'IMAP_AUTHENTICATED_NOT_CONNECTED',
                'error_message': error_message,
            }
        return {
            'success': False,
            'result_class': 'auth',
            'error_code': 'IMAP_AUTH_FAILED',
            'error_message': sanitize_error_details(error_message),
        }
    except Exception as exc:
        return {
            'success': False,
            'result_class': 'transient',
            'error_code': type(exc).__name__,
            'error_message': sanitize_error_details(str(exc)),
        }
    finally:
        if mail:
            try:
                mail.logout()
            except Exception:
                pass


def probe_graph_mailbox_access(client_id: str, refresh_token: str,
                               proxy_url: str = None, fallback_proxy_urls=None):
    """Verify Graph can refresh the token and access the real Inbox resource."""
    token_result = get_access_token_graph_result(
        client_id,
        refresh_token,
        proxy_url,
        fallback_proxy_urls,
    )
    if not token_result.get('success'):
        error = token_result.get('error') or {}
        classification = classify_token_failure(
            status_code=int(token_result.get('status_code') or error.get('status') or 0),
            error=str(token_result.get('oauth_error') or ''),
            description=str(
                token_result.get('oauth_error_description')
                or error.get('details')
                or error.get('message')
                or ''
            ),
        )
        return {
            'success': False,
            'result_class': classification['result_class'],
            'error_code': str(
                classification['error_code'] or error.get('code') or 'GRAPH_TOKEN_FAILED'
            ),
            'error_message': sanitize_error_details(
                str(error.get('message') or 'Graph Token 获取失败')
            ),
        }

    try:
        response = get_with_proxy_fallback(
            'https://graph.microsoft.com/v1.0/me/mailFolders/inbox',
            headers={'Authorization': f"Bearer {token_result.get('access_token') or ''}"},
            params={'$select': 'id,displayName,totalItemCount,unreadItemCount'},
            timeout=HTTP_REQUEST_TIMEOUT,
            proxy_url=proxy_url,
            fallback_proxy_urls=fallback_proxy_urls,
        )
        status_code = int(getattr(response, 'status_code', 0) or 0)
        if 200 <= status_code < 300:
            return {
                'success': True,
                'result_class': 'success',
                'error_code': '',
                'error_message': '',
                'rotated_refresh_token': str(
                    token_result.get('rotated_refresh_token') or ''
                ).strip(),
            }

        details = get_response_details(response)
        graph_error = details.get('error') if isinstance(details, dict) else None
        if isinstance(graph_error, dict):
            graph_code = str(graph_error.get('code') or '')
            graph_message = str(graph_error.get('message') or '')
        else:
            graph_code = str(graph_error or '') if isinstance(details, dict) else ''
            graph_message = str(details or '')
        classification = classify_token_failure(
            status_code=status_code,
            error=graph_code,
            description=graph_message,
        )
        return {
            'success': False,
            'result_class': classification['result_class'],
            'error_code': str(
                classification['error_code'] or graph_code or f'GRAPH_INBOX_HTTP_{status_code}'
            ),
            'error_message': sanitize_error_details(graph_message or 'Graph Inbox 访问失败'),
        }
    except Exception as exc:
        classification = classify_token_failure(exception=exc)
        return {
            'success': False,
            'result_class': classification['result_class'],
            'error_code': str(classification['error_code'] or type(exc).__name__),
            'error_message': sanitize_error_details(str(exc)),
        }


def account_health_worker_enabled() -> bool:
    return str(os.getenv('ACCOUNT_HEALTH_WORKER_ENABLED', 'false')).strip().lower() in {
        '1', 'true', 'yes', 'on'
    }


def account_health_poll_seconds() -> int:
    try:
        return max(60, min(3600, int(os.getenv('ACCOUNT_HEALTH_POLL_SECONDS', '60') or 60)))
    except (TypeError, ValueError):
        return 60


def account_health_batch_size() -> int:
    try:
        return max(1, min(100, int(os.getenv('ACCOUNT_HEALTH_BATCH_SIZE', '20') or 20)))
    except (TypeError, ValueError):
        return 20


def account_health_account_delay_seconds() -> int:
    try:
        return max(0, min(300, int(
            os.getenv('ACCOUNT_HEALTH_ACCOUNT_DELAY_SECONDS', '5') or 5
        )))
    except (TypeError, ValueError):
        return 5


def enroll_account_health(account_id: int, db=None, now=None) -> None:
    database = db or get_db()
    current = parse_timestamp(now) or utc_now()
    next_check = next_success_check_at(current, account_id, current)
    database.execute(
        '''
        UPDATE accounts
        SET health_enrolled_at = ?,
            next_health_check_at = ?,
            health_status = 'healthy',
            consecutive_auth_failures = 0,
            transient_failure_count = 0,
            last_health_error_code = NULL,
            quarantined_at = NULL,
            health_delete_after_at = NULL,
            status = 'active',
            updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
        ''',
        (format_timestamp(current), format_timestamp(next_check), int(account_id)),
    )


def load_due_health_accounts(db=None, limit=None, now=None):
    database = db or get_db()
    current = format_timestamp(parse_timestamp(now) or utc_now())
    batch_limit = account_health_batch_size() if limit is None else max(1, int(limit))
    return database.execute(
        '''
        SELECT a.*
        FROM accounts a
        WHERE (
                a.status = 'active'
                OR (a.status = 'inactive' AND a.health_status = 'quarantined')
              )
          AND COALESCE(a.account_type, 'outlook') = 'outlook'
          AND COALESCE(a.client_id, '') != ''
          AND COALESCE(a.refresh_token, '') != ''
          AND (a.next_health_check_at IS NULL OR a.next_health_check_at <= ?)
        ORDER BY COALESCE(a.next_health_check_at, a.created_at) ASC, a.id ASC
        LIMIT ?
        ''',
        (current, batch_limit),
    ).fetchall()


def _health_proxy_config(account, db):
    group_id = account['group_id'] if 'group_id' in account.keys() else None
    if group_id:
        group = db.execute(
            '''
            SELECT proxy_url, fallback_proxy_url_1, fallback_proxy_url_2
            FROM groups WHERE id = ?
            ''',
            (group_id,),
        ).fetchone()
        if group:
            return (
                get_group_proxy_url(dict(group)),
                get_group_proxy_failover_urls(dict(group)),
            )
    account_dict = dict(account)
    proxy_config = get_account_proxy_config(account_dict)
    return (
        proxy_config.get('proxy_url', '') or '',
        get_account_proxy_failover_urls(account_dict),
    )


def _apply_health_result(account, result, db, now=None):
    current = parse_timestamp(now) or utc_now()
    current_text = format_timestamp(current)
    account_id = int(account['id'])
    email_addr = str(account['email'] or '')
    result_class = str(result.get('result_class') or 'transient')
    error_code = str(result.get('error_code') or '')[:96]
    error_message = str(result.get('error_message') or '')[:500]

    if result.get('success'):
        rotated = str(result.get('rotated_refresh_token') or '').strip()
        if rotated:
            cursor = db.execute(
                '''
                UPDATE accounts
                SET refresh_token = ?,
                    refresh_token_updated_at = CURRENT_TIMESTAMP,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = ? AND refresh_token = ?
                ''',
                (encrypt_data(rotated), account_id, account['refresh_token']),
            )
            if cursor.rowcount != 1:
                transient_count = int(account['transient_failure_count'] or 0) + 1
                db.execute(
                    '''
                    UPDATE accounts
                    SET health_status = 'transient',
                        consecutive_auth_failures = 0,
                        transient_failure_count = ?,
                        last_health_error_code = 'TOKEN_ROTATION_CONFLICT',
                        next_health_check_at = ?,
                        updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                    ''',
                    (
                        transient_count,
                        format_timestamp(current + transient_retry_delay(transient_count)),
                        account_id,
                    ),
                )
                log_refresh_result(
                    account_id,
                    email_addr,
                    'health_scheduled',
                    'failed',
                    'operational:TOKEN_ROTATION_CONFLICT',
                    db_conn=db,
                )
                return 'transient'
        enrolled_at = account['health_enrolled_at'] or current_text
        next_check = next_success_check_at(enrolled_at, account_id, current)
        db.execute(
            '''
            UPDATE accounts
            SET status = 'active',
                health_enrolled_at = COALESCE(health_enrolled_at, ?),
                next_health_check_at = ?,
                health_status = 'healthy',
                consecutive_auth_failures = 0,
                transient_failure_count = 0,
                last_health_error_code = NULL,
                quarantined_at = NULL,
                health_delete_after_at = NULL,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
            ''',
            (current_text, format_timestamp(next_check), account_id),
        )
        log_refresh_result(account_id, email_addr, 'health_scheduled', 'success', db_conn=db)
        return 'healthy'

    if result_class == 'auth':
        failures = int(account['consecutive_auth_failures'] or 0) + 1
        if failures >= 3:
            quarantine_retry_at = next_success_check_at(
                account['health_enrolled_at'] or current_text,
                account_id,
                current,
                jitter_seconds=0,
            )
            delete_after = current + timedelta(hours=max(
                24,
                int(os.getenv('ACCOUNT_HEALTH_DELETE_AFTER_HOURS', '168') or 168),
            ))
            db.execute(
                '''
                UPDATE accounts
                SET status = 'inactive',
                    health_status = 'quarantined',
                    consecutive_auth_failures = ?,
                    transient_failure_count = 0,
                    last_health_error_code = ?,
                    next_health_check_at = ?,
                    quarantined_at = COALESCE(quarantined_at, ?),
                    health_delete_after_at = COALESCE(health_delete_after_at, ?),
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
                ''',
                (
                    failures,
                    error_code,
                    format_timestamp(quarantine_retry_at),
                    current_text,
                    format_timestamp(delete_after),
                    account_id,
                ),
            )
            state = 'quarantined'
        else:
            db.execute(
                '''
                UPDATE accounts
                SET health_status = 'suspect',
                    consecutive_auth_failures = ?,
                    transient_failure_count = 0,
                    last_health_error_code = ?,
                    next_health_check_at = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
                ''',
                (
                    failures,
                    error_code,
                    format_timestamp(current + timedelta(hours=2)),
                    account_id,
                ),
            )
            state = 'suspect'
    else:
        transient_count = int(account['transient_failure_count'] or 0) + 1
        retry_at = current + transient_retry_delay(transient_count)
        db.execute(
            '''
            UPDATE accounts
            SET health_status = 'transient',
                consecutive_auth_failures = 0,
                transient_failure_count = ?,
                last_health_error_code = ?,
                next_health_check_at = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
            ''',
            (transient_count, error_code, format_timestamp(retry_at), account_id),
        )
        state = 'transient'

    log_refresh_result(
        account_id,
        email_addr,
        'health_scheduled',
        'failed',
        f'{result_class}:{error_code} {error_message}'.strip(),
        db_conn=db,
    )
    return state


def check_account_health(account, db=None, now=None):
    database = db or get_db()
    try:
        encrypted_token = account['refresh_token']
        refresh_token = decrypt_data(encrypted_token) if encrypted_token else ''
    except Exception as exc:
        result = {
            'success': False,
            'result_class': 'operational',
            'error_code': 'TOKEN_DECRYPT_FAILED',
            'error_message': sanitize_error_details(str(exc)),
        }
    else:
        proxy_url, fallback_urls = _health_proxy_config(account, database)
        result = probe_refresh_token(
            str(account['client_id'] or ''),
            refresh_token,
            proxy_url,
            fallback_urls,
        )
        if result.get('success'):
            token_for_mailbox = (
                str(result.get('rotated_refresh_token') or '').strip()
                or refresh_token
            )
            mailbox_result = probe_graph_mailbox_access(
                str(account['client_id'] or ''),
                token_for_mailbox,
                proxy_url,
                fallback_urls,
            )
            if not mailbox_result.get('success'):
                mailbox_result = probe_imap_mailbox_access(
                    str(account['email'] or ''),
                    str(account['client_id'] or ''),
                    token_for_mailbox,
                    proxy_url,
                    fallback_urls,
                )
            if mailbox_result.get('success'):
                result['rotated_refresh_token'] = (
                    str(mailbox_result.get('rotated_refresh_token') or '').strip()
                    or str(result.get('rotated_refresh_token') or '').strip()
                )
            else:
                result = mailbox_result

    state = _apply_health_result(account, result, database, now=now)
    database.commit()
    return {'account_id': int(account['id']), 'state': state, **result}


def purge_expired_health_quarantine(db=None, now=None) -> int:
    if str(os.getenv('ACCOUNT_HEALTH_AUTO_DELETE_ENABLED', 'false')).lower() not in {
        '1', 'true', 'yes', 'on'
    }:
        return 0
    database = db or get_db()
    current = format_timestamp(parse_timestamp(now) or utc_now())
    rows = database.execute(
        '''
        SELECT id FROM accounts
        WHERE status = 'inactive'
          AND health_status = 'quarantined'
          AND health_delete_after_at IS NOT NULL
          AND health_delete_after_at <= ?
        ORDER BY health_delete_after_at, id
        LIMIT 20
        ''',
        (current,),
    ).fetchall()
    deleted = 0
    for row in rows:
        if delete_account_by_id(int(row['id'])):
            deleted += 1
    return deleted


def run_due_health_checks(now=None, trigger_type='scheduled'):
    if not account_health_worker_enabled():
        return {'status': 'disabled', 'processed': 0}
    if not account_health_run_lock.acquire(blocking=False):
        return {'status': 'busy', 'processed': 0}
    token_lock_acquired = False
    try:
        token_lock_acquired = token_refresh_run_lock.acquire(blocking=False)
        if not token_lock_acquired:
            return {'status': 'token_refresh_busy', 'processed': 0}
        with app.app_context():
            db = get_db()
            run_cursor = db.execute(
                '''
                INSERT INTO account_health_runs (trigger_type, status, started_at)
                VALUES (?, 'running', ?)
                ''',
                (str(trigger_type or 'scheduled')[:32], format_timestamp(utc_now())),
            )
            run_id = int(run_cursor.lastrowid)
            db.commit()
            try:
                accounts = load_due_health_accounts(db=db, now=now)
                states = []
                exception_count = 0
                for index, account in enumerate(accounts):
                    try:
                        states.append(check_account_health(account, db=db, now=now)['state'])
                    except Exception as exc:
                        exception_count += 1
                        try:
                            db.rollback()
                        except Exception:
                            pass
                        error_code = f'WORKER_EXCEPTION_{type(exc).__name__}'[:96]
                        transient_count = int(account['transient_failure_count'] or 0) + 1
                        db.execute(
                            '''
                            UPDATE accounts
                            SET health_status = 'transient',
                                consecutive_auth_failures = 0,
                                transient_failure_count = ?,
                                last_health_error_code = ?,
                                next_health_check_at = ?,
                                updated_at = CURRENT_TIMESTAMP
                            WHERE id = ?
                            ''',
                            (
                                transient_count,
                                error_code,
                                format_timestamp(
                                    (parse_timestamp(now) or utc_now())
                                    + transient_retry_delay(transient_count)
                                ),
                                int(account['id']),
                            ),
                        )
                        db.execute(
                            '''
                            INSERT INTO account_refresh_logs (
                                account_id, account_email, refresh_type, status, error_message
                            ) VALUES (?, ?, 'health_scheduled', 'failed', ?)
                            ''',
                            (int(account['id']), str(account['email'] or ''), error_code),
                        )
                        db.commit()
                        safe_console_print(
                            f"[账号测活] account_id={account['id']} 执行异常：{type(exc).__name__}"
                        )
                    if index + 1 < len(accounts):
                        time.sleep(account_health_account_delay_seconds())
                deleted = purge_expired_health_quarantine(db=db, now=now)
                run_status = 'partial_failed' if exception_count else 'success'
                processed_count = len(accounts)
                summary = {
                    'status': run_status,
                    'run_id': run_id,
                    'selected': len(accounts),
                    'processed': processed_count,
                    'healthy': states.count('healthy'),
                    'suspect': states.count('suspect'),
                    'transient': states.count('transient'),
                    'quarantined': states.count('quarantined'),
                    'deleted': deleted,
                    'exceptions': exception_count,
                }
                db.execute(
                    '''
                    UPDATE account_health_runs
                    SET status = ?, finished_at = ?, selected_count = ?, processed_count = ?,
                        healthy_count = ?, suspect_count = ?, transient_count = ?,
                        quarantined_count = ?, deleted_count = ?, exception_count = ?
                    WHERE id = ?
                    ''',
                    (
                        run_status, format_timestamp(utc_now()), len(accounts), processed_count,
                        summary['healthy'], summary['suspect'], summary['transient'],
                        summary['quarantined'], deleted, exception_count, run_id,
                    ),
                )
                db.commit()
                safe_console_print(f"[账号测活] {summary}")
                return summary
            except Exception as exc:
                try:
                    db.rollback()
                    db.execute(
                        '''
                        UPDATE account_health_runs
                        SET status = 'failed', finished_at = ?, exception_count = 1,
                            error_code = ?
                        WHERE id = ?
                        ''',
                        (format_timestamp(utc_now()), type(exc).__name__[:96], run_id),
                    )
                    db.commit()
                except Exception:
                    pass
                raise
    finally:
        if token_lock_acquired:
            token_refresh_run_lock.release()
        account_health_run_lock.release()


def account_health_dashboard_limit(value) -> int:
    try:
        return max(1, min(500, int(value or 100)))
    except (TypeError, ValueError):
        return 100


def load_account_health_dashboard(status_filter: str = '', error_code: str = '',
                                  query: str = '', limit: int = 100):
    """Return a credential-free snapshot for the authenticated health dashboard."""
    db = get_db()
    now_text = format_timestamp(utc_now())
    base_where = """
        COALESCE(a.account_type, 'outlook') = 'outlook'
        AND COALESCE(a.client_id, '') != ''
        AND COALESCE(a.refresh_token, '') != ''
    """
    summary_row = db.execute(
        f'''
        SELECT
            COUNT(*) AS total,
            SUM(CASE WHEN a.status = 'active' THEN 1 ELSE 0 END) AS active,
            SUM(CASE WHEN COALESCE(a.health_status, 'pending') = 'healthy' THEN 1 ELSE 0 END) AS healthy,
            SUM(CASE WHEN COALESCE(a.health_status, 'pending') = 'pending' THEN 1 ELSE 0 END) AS pending,
            SUM(CASE WHEN a.health_status = 'suspect' THEN 1 ELSE 0 END) AS suspect,
            SUM(CASE WHEN a.health_status = 'transient' THEN 1 ELSE 0 END) AS transient,
            SUM(CASE WHEN a.health_status = 'quarantined' THEN 1 ELSE 0 END) AS quarantined,
            SUM(CASE WHEN a.next_health_check_at IS NULL OR a.next_health_check_at <= ? THEN 1 ELSE 0 END) AS due_now,
            SUM(CASE WHEN COALESCE(a.consecutive_auth_failures, 0) > 0 THEN 1 ELSE 0 END) AS auth_failures,
            SUM(CASE WHEN COALESCE(a.transient_failure_count, 0) > 0 THEN 1 ELSE 0 END) AS transient_failures
        FROM accounts a
        WHERE {base_where}
        ''',
        (now_text,),
    ).fetchone()
    summary = {
        key: int(summary_row[key] or 0)
        for key in (
            'total', 'active', 'healthy', 'pending', 'suspect', 'transient',
            'quarantined', 'due_now', 'auth_failures', 'transient_failures',
        )
    }

    status_counts = {
        str(row['health_status'] or 'pending'): int(row['count'] or 0)
        for row in db.execute(
            f'''
            SELECT COALESCE(NULLIF(a.health_status, ''), 'pending') AS health_status,
                   COUNT(*) AS count
            FROM accounts a
            WHERE {base_where}
            GROUP BY COALESCE(NULLIF(a.health_status, ''), 'pending')
            ORDER BY count DESC, health_status
            ''',
        ).fetchall()
    }
    error_counts = [
        {
            'error_code': str(row['error_code'] or 'UNKNOWN'),
            'count': int(row['count'] or 0),
        }
        for row in db.execute(
            f'''
            SELECT COALESCE(NULLIF(a.last_health_error_code, ''), 'UNKNOWN') AS error_code,
                   COUNT(*) AS count
            FROM accounts a
            WHERE {base_where}
              AND COALESCE(a.last_health_error_code, '') != ''
            GROUP BY COALESCE(NULLIF(a.last_health_error_code, ''), 'UNKNOWN')
            ORDER BY count DESC, error_code
            LIMIT 20
            ''',
        ).fetchall()
    ]
    recent_run_rows = db.execute(
        '''
        SELECT id, trigger_type, status, started_at, finished_at,
               selected_count, processed_count, healthy_count, suspect_count,
               transient_count, quarantined_count, deleted_count,
               exception_count, error_code
        FROM account_health_runs
        ORDER BY started_at DESC, id DESC
        LIMIT 10
        ''',
    ).fetchall()
    recent_runs = [dict(row) for row in recent_run_rows]
    last_run = recent_runs[0] if recent_runs else None

    clauses = [base_where]
    params = []
    normalized_status = str(status_filter or '').strip().lower()
    if normalized_status and normalized_status != 'all':
        clauses.append("COALESCE(NULLIF(a.health_status, ''), 'pending') = ?")
        params.append(normalized_status)
    normalized_error = str(error_code or '').strip()
    if normalized_error and normalized_error.lower() != 'all':
        clauses.append("COALESCE(a.last_health_error_code, '') = ?")
        params.append(normalized_error)
    normalized_query = str(query or '').strip().lower()
    if normalized_query:
        clauses.append("(LOWER(a.email) LIKE ? OR LOWER(COALESCE(a.remark, '')) LIKE ?)")
        pattern = f'%{normalized_query}%'
        params.extend([pattern, pattern])
    params.append(account_health_dashboard_limit(limit))

    rows = db.execute(
        f'''
        SELECT
            a.id, a.email, a.remark, a.status, a.health_status,
            a.health_enrolled_at, a.next_health_check_at,
            a.consecutive_auth_failures, a.transient_failure_count,
            a.last_health_error_code, a.quarantined_at,
            a.created_at, a.updated_at,
            l.status AS last_check_status,
            l.error_message AS last_check_error,
            l.created_at AS last_checked_at
        FROM accounts a
        LEFT JOIN account_refresh_logs l ON l.id = (
            SELECT l2.id
            FROM account_refresh_logs l2
            WHERE l2.account_id = a.id
              AND l2.refresh_type = 'health_scheduled'
            ORDER BY l2.created_at DESC, l2.id DESC
            LIMIT 1
        )
        WHERE {' AND '.join(f'({clause})' for clause in clauses)}
        ORDER BY
            CASE COALESCE(NULLIF(a.health_status, ''), 'pending')
                WHEN 'quarantined' THEN 0
                WHEN 'suspect' THEN 1
                WHEN 'transient' THEN 2
                WHEN 'pending' THEN 3
                ELSE 4
            END,
            COALESCE(a.next_health_check_at, a.created_at) ASC,
            a.id DESC
        LIMIT ?
        ''',
        tuple(params),
    ).fetchall()
    accounts = []
    for row in rows:
        accounts.append({
            'id': int(row['id']),
            'email': str(row['email'] or ''),
            'remark': str(row['remark'] or ''),
            'status': str(row['status'] or 'active'),
            'health_status': str(row['health_status'] or 'pending'),
            'health_enrolled_at': str(row['health_enrolled_at'] or ''),
            'next_health_check_at': str(row['next_health_check_at'] or ''),
            'consecutive_auth_failures': int(row['consecutive_auth_failures'] or 0),
            'transient_failure_count': int(row['transient_failure_count'] or 0),
            'last_health_error_code': str(row['last_health_error_code'] or ''),
            'quarantined_at': str(row['quarantined_at'] or ''),
            'last_check_status': str(row['last_check_status'] or ''),
            'last_check_error': sanitize_error_details(str(row['last_check_error'] or '')),
            'last_checked_at': str(row['last_checked_at'] or ''),
            'created_at': str(row['created_at'] or ''),
            'updated_at': str(row['updated_at'] or ''),
        })
    return {
        'summary': summary,
        'status_counts': status_counts,
        'error_counts': error_counts,
        'accounts': accounts,
        'last_run': last_run,
        'recent_runs': recent_runs,
    }


def _run_manual_account_health_task():
    with app.app_context():
        try:
            result = run_due_health_checks(trigger_type='manual')
            error = ''
        except Exception as exc:
            result = None
            error = sanitize_error_details(str(exc))
        with account_health_manual_task_lock:
            account_health_manual_task.update({
                'running': False,
                'finished_at': format_timestamp(utc_now()),
                'summary': result,
                'error': error,
            })


@app.route('/api/account-health/dashboard', methods=['GET'])
@login_required
def api_account_health_dashboard():
    snapshot = load_account_health_dashboard(
        request.args.get('status', ''),
        request.args.get('error_code', ''),
        request.args.get('q', ''),
        request.args.get('limit', 100),
    )
    with account_health_manual_task_lock:
        manual_task = dict(account_health_manual_task)
    last_run = snapshot.get('last_run') or {}
    scheduler = globals().get('scheduler_instance')
    scheduler_running = bool(getattr(scheduler, 'running', False))
    try:
        health_job = scheduler.get_job('account_health') if scheduler else None
    except Exception:
        health_job = None
    worker_enabled = account_health_worker_enabled()
    worker_busy = account_health_run_lock.locked()
    worker_registered = health_job is not None
    if worker_busy:
        worker_state = 'running'
    elif not worker_enabled:
        worker_state = 'stopped'
    elif scheduler_running and worker_registered:
        worker_state = 'idle'
    else:
        worker_state = 'error'
    response = jsonify({
        'success': True,
        'generated_at': format_timestamp(utc_now()),
        **snapshot,
        'worker': {
            'enabled': worker_enabled,
            'state': worker_state,
            'scheduler_running': scheduler_running,
            'registered': worker_registered,
            'next_run_at': str(getattr(health_job, 'next_run_time', '') or ''),
            'poll_seconds': account_health_poll_seconds(),
            'batch_size': account_health_batch_size(),
            'account_delay_seconds': account_health_account_delay_seconds(),
            'busy': worker_busy,
            'last_run_at': last_run.get('finished_at') or last_run.get('started_at') or '',
            'last_processed': int(last_run.get('processed_count') or 0),
            'last_status': last_run.get('status') or '',
            'last_run': last_run or None,
            'manual_task': manual_task,
        },
    })
    response.headers['Cache-Control'] = 'no-store'
    return response


@app.route('/api/account-health/run', methods=['POST'])
@login_required
def api_run_account_health_now():
    data = request.get_json(silent=True) or {}
    scope = str(data.get('scope') or 'due').strip().lower()
    if scope != 'due':
        return jsonify({
            'success': False,
            'error': '当前仅支持扫描到期账号',
            'error_code': 'unsupported_scope',
        }), 400
    if not account_health_worker_enabled():
        return jsonify({
            'success': False,
            'error': '账号测活功能当前已停用',
            'error_code': 'health_worker_disabled',
        }), 409
    with account_health_manual_task_lock:
        if account_health_manual_task.get('running'):
            return jsonify({
                'success': False,
                'error': '测活任务正在运行',
                'error_code': 'health_run_busy',
            }), 409
        account_health_manual_task.update({
            'running': True,
            'started_at': format_timestamp(utc_now()),
            'finished_at': '',
            'summary': None,
            'error': '',
        })
    threading.Thread(
        target=_run_manual_account_health_task,
        name='account-health-manual-run',
        daemon=True,
    ).start()
    return jsonify({'success': True, 'started': True, 'scope': 'due'}), 202


def scheduled_health_task():
    return run_due_health_checks()


def add_account_health_job(scheduler, interval_trigger_cls, app_tzinfo) -> bool:
    if not account_health_worker_enabled():
        safe_console_print('✓ 账号分层测活 Worker 已禁用')
        return False
    scheduler.add_job(
        func=scheduled_health_task,
        trigger=interval_trigger_cls(seconds=account_health_poll_seconds(), timezone=app_tzinfo),
        id='account_health',
        name='账号分层测活',
        replace_existing=True,
        max_instances=1,
        coalesce=True,
    )
    safe_console_print(
        f"✓ 账号分层测活已启动：每 {account_health_poll_seconds()} 秒扫描到期账号"
    )
    return True
