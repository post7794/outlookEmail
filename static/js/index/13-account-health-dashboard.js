/* global formatAbsoluteDateTime, hideModal, showModal, showToast */

let accountHealthRefreshTimer = null;
let accountHealthRequestSequence = 0;

function showAccountHealthDashboard() {
    const modal = document.getElementById('accountHealthModal');
    if (!modal) return;
    showModal('accountHealthModal');
    loadAccountHealthDashboard();
    clearInterval(accountHealthRefreshTimer);
    accountHealthRefreshTimer = setInterval(() => {
        if (modal.classList.contains('show')) loadAccountHealthDashboard({ silent: true });
    }, 30000);
}

function hideAccountHealthDashboard() {
    hideModal('accountHealthModal');
    clearInterval(accountHealthRefreshTimer);
    accountHealthRefreshTimer = null;
}

function healthNumber(value, fallback = 0) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : fallback;
}

function healthObject(value) {
    return value && typeof value === 'object' && !Array.isArray(value) ? value : {};
}

function normalizeHealthCounts(value) {
    if (Array.isArray(value)) {
        return Object.fromEntries(value.map(item => [
            String(item?.status || item?.error_code || item?.code || item?.name || 'unknown'),
            healthNumber(item?.count)
        ]));
    }
    return healthObject(value);
}

function healthStatusLabel(status) {
    const labels = {
        healthy: '健康', pending: '待检查', due: '待检查', unhealthy: '失败',
        suspect: '疑似失效', transient: '临时失败', quarantined: '已隔离',
        failed: '失败', error: '失败', success: '成功', partial_failed: '部分失败',
        disabled: '已停用', running: '检查中', unknown: '未知'
    };
    return labels[String(status || '').toLowerCase()] || String(status || '未知');
}

function formatHealthDate(value) {
    if (!value) return '--';
    if (typeof formatAbsoluteDateTime === 'function') {
        const formatted = formatAbsoluteDateTime(value);
        if (formatted && formatted !== '-') return formatted;
    }
    const normalized = typeof value === 'number' && value < 1e12 ? value * 1000 : value;
    const date = new Date(normalized);
    if (Number.isNaN(date.getTime())) return String(value);
    try {
        return new Intl.DateTimeFormat('zh-CN', {
            year: 'numeric', month: '2-digit', day: '2-digit',
            hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
        }).format(date);
    } catch (_) {
        return date.toLocaleString();
    }
}

function setAccountHealthAlert(message) {
    const alert = document.getElementById('accountHealthAlert');
    if (!alert) return;
    alert.textContent = message || '';
    alert.hidden = !message;
}

function setHealthText(id, value) {
    const node = document.getElementById(id);
    if (node) node.textContent = String(value);
}

function renderHealthBreakdown(containerId, counts, emptyLabel, clickable = false) {
    const container = document.getElementById(containerId);
    if (!container) return;
    container.replaceChildren();
    const entries = Object.entries(normalizeHealthCounts(counts)).filter(([, count]) => healthNumber(count) > 0);
    if (!entries.length) {
        const empty = document.createElement('span');
        empty.className = 'account-health-muted';
        empty.textContent = emptyLabel;
        container.appendChild(empty);
        return;
    }
    entries.sort((a, b) => healthNumber(b[1]) - healthNumber(a[1])).forEach(([name, count]) => {
        const chip = document.createElement(clickable ? 'button' : 'span');
        chip.className = 'account-health-chip';
        if (clickable) {
            chip.type = 'button';
            chip.title = '筛选此错误';
            chip.addEventListener('click', () => {
                const select = document.getElementById('accountHealthErrorFilter');
                if (select) select.value = name;
                loadAccountHealthDashboard();
            });
        }
        chip.append(document.createTextNode(`${clickable ? name : healthStatusLabel(name)} `));
        const strong = document.createElement('strong');
        strong.textContent = String(count);
        chip.appendChild(strong);
        container.appendChild(chip);
    });
}

function updateHealthErrorOptions(errorCounts) {
    const select = document.getElementById('accountHealthErrorFilter');
    if (!select) return;
    const selected = select.value;
    const entries = Object.entries(normalizeHealthCounts(errorCounts)).filter(([code, count]) => code && healthNumber(count) > 0);
    select.replaceChildren(new Option('全部错误', ''));
    entries.sort((a, b) => healthNumber(b[1]) - healthNumber(a[1])).forEach(([code, count]) => {
        select.add(new Option(`${code} (${count})`, code));
    });
    if (selected && entries.some(([code]) => code === selected)) select.value = selected;
}

function appendHealthCell(row, text, className = '') {
    const cell = document.createElement('td');
    if (className) cell.className = className;
    cell.textContent = text;
    row.appendChild(cell);
    return cell;
}

function renderHealthAccounts(accounts) {
    const tbody = document.getElementById('accountHealthTableBody');
    if (!tbody) return;
    tbody.replaceChildren();
    if (!Array.isArray(accounts) || !accounts.length) {
        const row = document.createElement('tr');
        const cell = appendHealthCell(row, '当前筛选条件下没有账号', 'account-health-empty');
        cell.colSpan = 7;
        tbody.appendChild(row);
        return;
    }

    accounts.forEach(account => {
        const row = document.createElement('tr');
        const emailCell = document.createElement('td');
        const email = document.createElement('span');
        email.className = 'account-health-email';
        email.textContent = String(account?.email || account?.account_email || account?.display_name || '未知账号');
        emailCell.appendChild(email);
        if (account?.id ?? account?.account_id) {
            const id = document.createElement('span');
            id.className = 'account-health-id';
            id.textContent = `#${account.id ?? account.account_id}`;
            emailCell.appendChild(id);
        }
        row.appendChild(emailCell);

        const status = String(account?.health_status || account?.status || 'unknown').toLowerCase();
        const statusCell = document.createElement('td');
        const badge = document.createElement('span');
        badge.className = `account-health-badge account-health-badge--${status.replace(/[^a-z0-9_-]/g, '')}`;
        badge.textContent = healthStatusLabel(status);
        statusCell.appendChild(badge);
        row.appendChild(statusCell);

        appendHealthCell(row, formatHealthDate(account?.last_health_check_at || account?.last_checked_at || account?.checked_at));
        appendHealthCell(row, formatHealthDate(account?.next_health_check_at || account?.next_check_at));
        appendHealthCell(row, String(healthNumber(account?.consecutive_auth_failures)));
        appendHealthCell(row, String(healthNumber(account?.transient_failure_count)));
        const errorCode = account?.last_health_error_code || account?.error_code || '';
        const errorMessage = account?.last_health_error || account?.last_check_error || account?.error_message || account?.message || '';
        appendHealthCell(row, [errorCode, errorMessage].filter(Boolean).join(' · ') || '--');
        tbody.appendChild(row);
    });
}

function renderHealthWorker(worker) {
    worker = healthObject(worker);
    const manualTask = healthObject(worker.manual_task);
    const lastRun = healthObject(worker.last_run);
    const lastStatus = String(worker.last_status || lastRun.status || '').toLowerCase();
    const inferredState = worker.busy || manualTask.running
        ? 'running'
        : (['failed', 'partial_failed'].includes(lastStatus) ? 'error' : (worker.enabled ? 'idle' : 'stopped'));
    const state = String(worker.state || worker.status || inferredState).toLowerCase();
    const dot = document.getElementById('healthWorkerDot');
    if (dot) dot.className = `account-health-worker-dot ${state === 'running' || state === 'idle' || state === 'healthy' ? 'is-running' : (state === 'error' || state === 'failed' ? 'is-error' : '')}`;
    const stateLabels = { running: '后台测活运行中', idle: '后台测活正常', healthy: '后台测活正常', stopped: '后台测活已停止', error: '后台测活异常', failed: '后台测活异常' };
    setHealthText('healthWorkerState', stateLabels[state] || '后台测活状态未知');
    const bits = [];
    if (manualTask.started_at || worker.last_run_at || worker.last_started_at) bits.push(`最近运行 ${formatHealthDate(manualTask.started_at || worker.last_run_at || worker.last_started_at)}`);
    const taskSummary = healthObject(manualTask.summary);
    if (taskSummary.processed != null || worker.last_processed != null || worker.processed != null) bits.push(`处理 ${taskSummary.processed ?? worker.last_processed ?? worker.processed} 个`);
    if (worker.poll_seconds != null) bits.push(`轮询 ${worker.poll_seconds} 秒`);
    if (worker.next_run_at) bits.push(`下次 ${formatHealthDate(worker.next_run_at)}`);
    if (manualTask.error || worker.error) bits.push(String(manualTask.error || worker.error));
    if (lastStatus === 'partial_failed') bits.push(`上轮有 ${healthNumber(lastRun.exception_count)} 个执行异常`);
    if (lastStatus === 'failed') bits.push(`上轮失败${lastRun.error_code ? `：${lastRun.error_code}` : ''}`);
    setHealthText('healthWorkerMeta', bits.join(' · ') || '服务端未提供任务详情');
    const runButton = document.getElementById('accountHealthRunBtn');
    if (runButton) {
        runButton.disabled = Boolean(!worker.enabled || worker.busy || manualTask.running);
        runButton.textContent = !worker.enabled ? '测活已停用' : (manualTask.running ? '检查中…' : '立即检查到期账号');
    }
}

function renderHealthRuns(runs) {
    const tbody = document.getElementById('accountHealthRunsBody');
    if (!tbody) return;
    tbody.replaceChildren();
    if (!Array.isArray(runs) || !runs.length) {
        const row = document.createElement('tr');
        const cell = appendHealthCell(row, '暂无运行记录', 'account-health-empty');
        cell.colSpan = 6;
        tbody.appendChild(row);
        return;
    }
    runs.forEach(run => {
        const row = document.createElement('tr');
        appendHealthCell(row, formatHealthDate(run?.started_at));
        appendHealthCell(row, run?.trigger_type === 'manual' ? '手动' : '定时');
        const statusCell = document.createElement('td');
        const status = String(run?.status || 'unknown').toLowerCase();
        const badge = document.createElement('span');
        badge.className = `account-health-badge account-health-badge--${status.replace(/[^a-z0-9_-]/g, '')}`;
        badge.textContent = healthStatusLabel(status);
        statusCell.appendChild(badge);
        row.appendChild(statusCell);
        appendHealthCell(row, `${healthNumber(run?.processed_count)}/${healthNumber(run?.selected_count)}`);
        appendHealthCell(row, String(healthNumber(run?.exception_count)));
        appendHealthCell(row, String(run?.error_code || '--'));
        tbody.appendChild(row);
    });
}

async function loadAccountHealthDashboard(options = {}) {
    const requestId = ++accountHealthRequestSequence;
    const refreshButton = document.getElementById('accountHealthRefreshBtn');
    const status = document.getElementById('accountHealthStatusFilter')?.value || '';
    const errorCode = document.getElementById('accountHealthErrorFilter')?.value || '';
    const query = document.getElementById('accountHealthQuery')?.value?.trim() || '';
    const params = new URLSearchParams({ limit: '100' });
    if (status) params.set('status', status);
    if (errorCode) params.set('error_code', errorCode);
    if (query) params.set('q', query);
    if (refreshButton && !options.silent) refreshButton.disabled = true;
    if (!options.silent) setAccountHealthAlert('');

    try {
        const response = await fetch(`/api/account-health/dashboard?${params}`, { cache: 'no-store' });
        const payload = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(payload?.error || payload?.message || `加载失败（HTTP ${response.status}）`);
        if (requestId !== accountHealthRequestSequence) return;
        const data = healthObject(payload.data || payload);
        const summary = healthObject(data.summary || data.overview);
        const statusCounts = normalizeHealthCounts(data.status_counts || summary.status_counts || data.by_status);
        const errorCounts = normalizeHealthCounts(data.error_counts || summary.error_counts || data.by_error);
        const total = healthNumber(summary.total ?? data.total ?? statusCounts.total, 0);
        const healthy = healthNumber(summary.healthy ?? statusCounts.healthy, 0);
        const failed = healthNumber(summary.quarantined ?? statusCounts.quarantined, 0);
        const pending = healthNumber(summary.pending ?? statusCounts.pending, 0)
            + healthNumber(summary.suspect ?? statusCounts.suspect, 0)
            + healthNumber(summary.transient ?? statusCounts.transient, 0);
        setHealthText('healthTotalCount', total);
        setHealthText('healthHealthyCount', healthy);
        setHealthText('healthPendingCount', pending);
        setHealthText('healthFailedCount', failed);
        renderHealthWorker(data.worker || data.worker_status);
        renderHealthBreakdown('accountHealthStatusBreakdown', statusCounts, '暂无数据');
        renderHealthBreakdown('accountHealthErrorBreakdown', errorCounts, '暂无错误', true);
        updateHealthErrorOptions(errorCounts);
        renderHealthAccounts(data.accounts || data.items || data.recent_accounts || []);
        renderHealthRuns(data.recent_runs || []);
        setHealthText('accountHealthUpdatedAt', `更新于 ${formatHealthDate(data.generated_at || data.updated_at || Date.now())}`);
    } catch (error) {
        if (requestId !== accountHealthRequestSequence) return;
        setAccountHealthAlert(error?.message || '测活数据加载失败');
        if (!options.silent) showToast(error?.message || '测活数据加载失败', 'error');
    } finally {
        if (refreshButton && requestId === accountHealthRequestSequence) refreshButton.disabled = false;
    }
}

async function runAccountHealthCheck(scope = 'due') {
    const button = document.getElementById('accountHealthRunBtn');
    let started = false;
    if (button) {
        button.disabled = true;
        button.textContent = '正在启动…';
    }
    setAccountHealthAlert('');
    try {
        const response = await fetch('/api/account-health/run', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ scope })
        });
        const payload = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(payload?.error || payload?.message || `启动失败（HTTP ${response.status}）`);
        started = true;
        showToast(payload?.message || '测活任务已启动', 'success');
        await loadAccountHealthDashboard({ silent: true });
    } catch (error) {
        setAccountHealthAlert(error?.message || '测活任务启动失败');
        showToast(error?.message || '测活任务启动失败', 'error');
    } finally {
        if (button && !started) {
            button.disabled = false;
            button.textContent = '立即检查到期账号';
        }
    }
}
