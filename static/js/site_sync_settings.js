document.addEventListener('DOMContentLoaded', () => {
    const progressModalEl = document.getElementById('syncProgressModal');
    const progressModal = bootstrap.Modal.getOrCreateInstance(progressModalEl);
    const statusText = document.getElementById('syncStatusText');
    const timeText = document.getElementById('syncTimeText');
    const progressBar = document.getElementById('syncProgressBar');
    const logConsole = document.getElementById('syncLogConsole');
    const closeButton = document.getElementById('closeSyncModalBtn');
    let pollTimer = null;
    let clockTimer = null;
    let startedAt = 0;
    let activeButton = null;

    function sortSites() {
        const body = document.getElementById('sitesTableBody');
        if (!body) return;
        const rows = Array.from(body.querySelectorAll('tr[data-manager]'));
        rows.sort((a, b) => {
            const managerResult = (a.dataset.manager || '').localeCompare(
                b.dataset.manager || '', 'zh-Hans-CN', { sensitivity: 'base' }
            );
            return managerResult || (a.dataset.siteUrl || '').localeCompare(b.dataset.siteUrl || '');
        });
        rows.forEach(row => body.appendChild(row));
    }

    document.getElementById('sortSitesByManagerBtn')?.addEventListener('click', sortSites);

    async function parseJsonResponse(response, action) {
        const text = await response.text();
        let payload = {};
        if (text) {
            try {
                payload = JSON.parse(text);
            } catch (_error) {
                throw new Error(`${action}返回了非 JSON 响应（HTTP ${response.status}）`);
            }
        }
        if (!response.ok) {
            throw new Error(payload.error || `${action}失败（HTTP ${response.status}）`);
        }
        return payload;
    }

    function setLogs(logs) {
        logConsole.replaceChildren();
        (Array.isArray(logs) ? logs : []).forEach(line => {
            const item = document.createElement('div');
            item.textContent = String(line);
            logConsole.appendChild(item);
        });
        logConsole.scrollTop = logConsole.scrollHeight;
    }

    function stopTimers() {
        if (pollTimer) window.clearTimeout(pollTimer);
        if (clockTimer) window.clearInterval(clockTimer);
        pollTimer = null;
        clockTimer = null;
    }

    function finish(status, message, logs) {
        stopTimers();
        statusText.textContent = message || (status === 'success' ? '同步完成' : '同步失败');
        setLogs(logs);
        progressBar.style.width = '100%';
        progressBar.classList.remove('progress-bar-animated', 'bg-primary');
        progressBar.classList.add(status === 'success' ? 'bg-success' : 'bg-danger');
        closeButton.disabled = false;
        if (activeButton) activeButton.disabled = false;
        activeButton = null;
    }

    async function pollStatus(statusId) {
        try {
            const response = await fetch(`/api/sync/status/${encodeURIComponent(statusId)}`, {
                headers: { 'Accept': 'application/json' }
            });
            const data = await parseJsonResponse(response, '同步状态');
            statusText.textContent = data.message || '同步进行中...';
            setLogs(data.logs);
            if (['success', 'error', 'cancelled', 'interrupted'].includes(data.status)) {
                finish(data.status, data.message, data.logs);
                return;
            }
            if (data.status === 'unknown') {
                finish('error', '未找到同步任务状态，请重新发起', data.logs);
                return;
            }
            progressBar.style.width = '65%';
            pollTimer = window.setTimeout(() => pollStatus(statusId), 1200);
        } catch (error) {
            finish('error', error.message, []);
        }
    }

    function resetProgress(title, button) {
        stopTimers();
        activeButton = button;
        if (activeButton) activeButton.disabled = true;
        progressModalEl.querySelector('.modal-title').textContent = title;
        statusText.textContent = '正在提交同步任务...';
        timeText.textContent = '00:00';
        setLogs([]);
        closeButton.disabled = true;
        progressBar.style.width = '15%';
        progressBar.classList.remove('bg-success', 'bg-danger');
        progressBar.classList.add('progress-bar-animated', 'bg-primary');
        startedAt = Date.now();
        clockTimer = window.setInterval(() => {
            const seconds = Math.floor((Date.now() - startedAt) / 1000);
            timeText.textContent = `${String(Math.floor(seconds / 60)).padStart(2, '0')}:${String(seconds % 60).padStart(2, '0')}`;
        }, 1000);
        progressModal.show();
    }

    async function startSync(endpoint, siteId, title, button, includeBody) {
        resetProgress(title, button);
        try {
            const options = {
                method: 'POST',
                headers: { 'Accept': 'application/json' }
            };
            if (includeBody) {
                options.headers['Content-Type'] = 'application/json';
                options.body = JSON.stringify({ site_id: siteId });
            }
            const response = await fetch(endpoint, options);
            const data = await parseJsonResponse(response, title);
            // Durable PostgreSQL runs return run_id. Keep siteId only for the
            // explicitly supported SQLite rollback path.
            const statusId = data.run_id || data.sync_id || siteId;
            statusText.textContent = data.message || '同步任务已启动';
            progressBar.style.width = '35%';
            pollTimer = window.setTimeout(() => pollStatus(statusId), 500);
        } catch (error) {
            finish('error', error.message, []);
        }
    }

    document.querySelectorAll('.sync-btn').forEach(button => {
        button.addEventListener('click', () => {
            const siteId = Number(button.dataset.siteId);
            startSync('/api/sync', siteId, '快速同步', button, true);
        });
    });

    document.querySelectorAll('.deep-sync-btn').forEach(button => {
        button.addEventListener('click', () => {
            document.getElementById('deepSyncSiteId').value = button.dataset.siteId;
            bootstrap.Modal.getOrCreateInstance(document.getElementById('deepSyncConfirmModal')).show();
        });
    });

    document.querySelectorAll('.clean-sync-btn').forEach(button => {
        button.addEventListener('click', () => {
            document.getElementById('cleanSyncSiteId').value = button.dataset.siteId;
            bootstrap.Modal.getOrCreateInstance(document.getElementById('cleanSyncConfirmModal')).show();
        });
    });

    document.getElementById('confirmDeepSyncBtn')?.addEventListener('click', () => {
        const siteId = Number(document.getElementById('deepSyncSiteId').value);
        bootstrap.Modal.getInstance(document.getElementById('deepSyncConfirmModal'))?.hide();
        const button = document.querySelector(`.deep-sync-btn[data-site-id="${siteId}"]`);
        startSync(`/api/sync/deep/${siteId}`, siteId, '深度同步', button, false);
    });

    document.getElementById('confirmCleanSyncBtn')?.addEventListener('click', () => {
        const siteId = Number(document.getElementById('cleanSyncSiteId').value);
        bootstrap.Modal.getInstance(document.getElementById('cleanSyncConfirmModal'))?.hide();
        const button = document.querySelector(`.clean-sync-btn[data-site-id="${siteId}"]`);
        startSync(`/api/sync/clean/${siteId}`, siteId, '清理同步', button, false);
    });

    progressModalEl.addEventListener('hidden.bs.modal', () => {
        if (closeButton.disabled) return;
        stopTimers();
    });
});
