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


def run_due_health_checks(now=None):
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
            accounts = load_due_health_accounts(db=db, now=now)
            states = []
            for index, account in enumerate(accounts):
                try:
                    states.append(check_account_health(account, db=db, now=now)['state'])
                except Exception as exc:
                    try:
                        db.rollback()
                    except Exception:
                        pass
                    safe_console_print(
                        f"[账号测活] account_id={account['id']} 执行异常：{type(exc).__name__}"
                    )
                if index + 1 < len(accounts):
                    time.sleep(max(0, int(os.getenv('ACCOUNT_HEALTH_ACCOUNT_DELAY_SECONDS', '5') or 5)))
            deleted = purge_expired_health_quarantine(db=db, now=now)
            summary = {
                'status': 'success',
                'processed': len(states),
                'healthy': states.count('healthy'),
                'suspect': states.count('suspect'),
                'transient': states.count('transient'),
                'quarantined': states.count('quarantined'),
                'deleted': deleted,
            }
            safe_console_print(f"[账号测活] {summary}")
            return summary
    finally:
        if token_lock_acquired:
            token_refresh_run_lock.release()
        account_health_run_lock.release()


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
