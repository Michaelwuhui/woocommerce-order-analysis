(function () {
    'use strict';

    const ACTIVE = new Set(['queued', 'running', 'recovering', 'cancelling']);
    const TERMINAL = new Set(['success', 'error', 'cancelled', 'interrupted']);
    const STORAGE_KEY = 'wooAnalysisActiveSyncRun';
    let runId = null;
    let pollTimer = null;
    let startedAt = null;

    function byId(id) {
        return document.getElementById(id);
    }

    function modal() {
        const element = byId('syncProgressModal');
        return element && window.bootstrap
            ? bootstrap.Modal.getOrCreateInstance(element)
            : null;
    }

    function resetUi() {
        const bar = byId('syncProgressBar');
        const status = byId('syncStatusText');
        const close = byId('closeSyncModalBtn');
        const cancel = byId('cancelSyncBtn');
        if (bar) {
            bar.style.width = '2%';
            bar.classList.add('progress-bar-animated', 'bg-primary');
            bar.classList.remove('bg-success', 'bg-danger', 'bg-warning');
        }
        if (status) {
            status.textContent = '正在创建同步批次...';
            status.classList.remove('text-danger', 'text-warning', 'text-success');
        }
        if (byId('syncLogConsole')) byId('syncLogConsole').textContent = '';
        if (close) {
            close.disabled = true;
            close.textContent = '完成';
            close.onclick = null;
        }
        if (cancel) cancel.disabled = true;
        startedAt = Date.now();
    }

    function setText(id, value) {
        const node = byId(id);
        if (node) node.textContent = value == null || value === '' ? '-' : String(value);
    }

    function updateTimer() {
        if (!startedAt) return;
        const elapsed = Math.max(0, Math.floor((Date.now() - startedAt) / 1000));
        const mins = String(Math.floor(elapsed / 60)).padStart(2, '0');
        const secs = String(elapsed % 60).padStart(2, '0');
        setText('syncTimeText', mins + ':' + secs);
    }

    function progressPercent(status) {
        const totalPages = Number(status.total_pages || 0);
        const completedPages = Number(status.completed_pages || 0);
        if (totalPages > 0) {
            return Math.min(99, Math.max(2, Math.round(completedPages * 100 / totalPages)));
        }
        const totalSites = Number(status.total_sites || 0);
        const completedSites = Number(status.completed_sites || 0);
        if (totalSites > 0) {
            return Math.min(99, Math.max(2, Math.round(completedSites * 100 / totalSites)));
        }
        return 2;
    }

    function renderLogs(logs) {
        const consoleNode = byId('syncLogConsole');
        if (!consoleNode) return;
        consoleNode.textContent = '';
        (logs || []).slice(-100).forEach(function (line) {
            const item = document.createElement('div');
            item.className = 'text-white-50 mb-1';
            item.textContent = String(line);
            consoleNode.appendChild(item);
        });
        consoleNode.scrollTop = consoleNode.scrollHeight;
    }

    function render(status) {
        const current = status.current_site || {};
        const labels = { quick: '快速同步', auto: '自动同步', deep: '深度同步' };
        setText('syncModeValue', labels[status.mode] || status.mode);
        setText(
            'syncSiteValue',
            current.url
                ? String(current.site_id) + ' · ' + current.url
                : (current.site_id || '-')
        );
        setText(
            'syncSitesValue',
            String(status.completed_sites || 0) + ' / ' + String(status.total_sites || 0)
        );
        setText('syncPageValue', status.current_page || current.current_page || 0);
        setText('syncFetchedValue', status.fetched_orders || 0);
        setText('syncWrittenValue', status.written_orders || 0);
        setText('syncRetryValue', status.retry_count || 0);
        setText('syncHeartbeatValue', status.heartbeat_at || '-');
        setText('syncStatusText', status.message || status.status);
        renderLogs(status.logs);

        const bar = byId('syncProgressBar');
        const statusNode = byId('syncStatusText');
        const close = byId('closeSyncModalBtn');
        const cancel = byId('cancelSyncBtn');
        if (bar) bar.style.width = progressPercent(status) + '%';

        if (status.interruption_state === 'recovering') {
            if (statusNode) {
                statusNode.textContent = '任务已中断/正在恢复';
                statusNode.classList.add('text-warning');
            }
            if (bar) {
                bar.classList.remove('bg-primary');
                bar.classList.add('bg-warning');
            }
        }

        if (ACTIVE.has(status.status)) {
            if (cancel) cancel.disabled = status.status === 'cancelling';
            return;
        }
        if (!TERMINAL.has(status.status)) return;

        stopPolling();
        localStorage.removeItem(STORAGE_KEY);
        if (bar) {
            bar.style.width = '100%';
            bar.classList.remove('progress-bar-animated', 'bg-primary', 'bg-warning');
            bar.classList.add(status.status === 'success' ? 'bg-success' : 'bg-danger');
        }
        if (statusNode) {
            statusNode.classList.add(
                status.status === 'success' ? 'text-success' : 'text-danger'
            );
        }
        if (cancel) cancel.disabled = true;
        if (close) {
            close.disabled = false;
            close.textContent = '完成并刷新';
            close.onclick = function () { window.location.reload(); };
        }
        localStorage.setItem('syncCompleted', String(Date.now()));
    }

    async function poll() {
        if (!runId) return;
        updateTimer();
        try {
            const response = await fetch(
                '/api/sync/status/' + encodeURIComponent(runId),
                { headers: { Accept: 'application/json' } }
            );
            const body = await response.json();
            if (!response.ok) throw new Error(body.error || '状态读取失败');
            render(body);
        } catch (error) {
            setText('syncStatusText', '状态暂时不可读，正在重试：' + error.message);
            const statusNode = byId('syncStatusText');
            if (statusNode) statusNode.classList.add('text-warning');
        }
    }

    function startPolling(id, createdAt) {
        runId = String(id);
        localStorage.setItem(STORAGE_KEY, runId);
        startedAt = createdAt ? Date.parse(createdAt) : Date.now();
        if (!Number.isFinite(startedAt)) startedAt = Date.now();
        stopPolling();
        pollTimer = window.setInterval(poll, 2000);
        poll();
    }

    function stopPolling() {
        if (pollTimer) window.clearInterval(pollTimer);
        pollTimer = null;
    }

    async function startGlobalSync() {
        const button = byId('syncAllBtn');
        if (button) button.disabled = true;
        resetUi();
        const instance = modal();
        if (instance) instance.show();
        try {
            const response = await fetch('/api/sync/all', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
                body: '{}'
            });
            const body = await response.json();
            if (!response.ok || !body.success) {
                throw new Error(body.error || body.message || '启动失败');
            }
            startPolling(body.run_id || body.sync_id, body.status && body.status.created_at);
            if (body.existing) {
                setText('syncStatusText', '已有任务运行，已恢复其真实进度');
            }
        } catch (error) {
            setText('syncStatusText', '启动失败：' + error.message);
            const statusNode = byId('syncStatusText');
            if (statusNode) statusNode.classList.add('text-danger');
            const close = byId('closeSyncModalBtn');
            if (close) close.disabled = false;
        } finally {
            if (button) button.disabled = false;
        }
    }

    async function cancelCurrent() {
        if (!runId) return;
        const cancel = byId('cancelSyncBtn');
        if (cancel) cancel.disabled = true;
        try {
            const response = await fetch(
                '/api/sync/' + encodeURIComponent(runId) + '/cancel',
                { method: 'POST', headers: { Accept: 'application/json' } }
            );
            const body = await response.json();
            if (!response.ok) throw new Error(body.error || '取消失败');
            render(body.status);
        } catch (error) {
            setText('syncStatusText', '取消请求失败：' + error.message);
            if (cancel) cancel.disabled = false;
        }
    }

    async function resumeActive() {
        try {
            const response = await fetch('/api/sync/active', {
                headers: { Accept: 'application/json' }
            });
            if (!response.ok) return;
            const body = await response.json();
            if (!body.active || !ACTIVE.has(body.active.status)) return;
            resetUi();
            const instance = modal();
            if (instance) instance.show();
            render(body.active);
            startPolling(body.active.run_id, body.active.created_at);
        } catch (_error) {
            const remembered = localStorage.getItem(STORAGE_KEY);
            if (remembered) startPolling(remembered);
        }
    }

    document.addEventListener('DOMContentLoaded', function () {
        const button = byId('syncAllBtn');
        if (button && !button.dataset.syncRunBound) {
            button.dataset.syncRunBound = '1';
            button.addEventListener('click', startGlobalSync);
        }
        const cancel = byId('cancelSyncBtn');
        if (cancel && !cancel.dataset.syncRunBound) {
            cancel.dataset.syncRunBound = '1';
            cancel.addEventListener('click', cancelCurrent);
        }
        resumeActive();
    });
})();
