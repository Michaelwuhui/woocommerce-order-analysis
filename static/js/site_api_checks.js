(function (root) {
    'use strict';

    async function requestJson(fetchImpl, url) {
        const controller = new AbortController();
        const timer = setTimeout(() => controller.abort(), 45000);
        try {
            const response = await fetchImpl(url, {method: 'POST', signal: controller.signal});
            if (response.redirected || response.status === 401 || response.status === 403) {
                const error = new Error('登录已过期或无检测权限，请刷新页面重新登录');
                error.fatal = true;
                throw error;
            }
            const data = await response.json();
            if (!response.ok || !data || !data.success) {
                throw new Error((data && data.error) || `服务器错误 (HTTP ${response.status})`);
            }
            return data;
        } finally {
            clearTimeout(timer);
        }
    }

    async function runChecks(fetchImpl, onProgress = () => {}) {
        const manifest = await requestJson(fetchImpl, '/api/sites/check-all');
        if (manifest.mode !== 'sequential' || !Array.isArray(manifest.sites)) {
            throw new Error('检测接口已更新，请刷新系统设置页面后重试');
        }
        const results = [];
        for (const site of manifest.sites) {
            onProgress(results.length, manifest.sites.length, site);
            try {
                const data = await requestJson(fetchImpl, `/api/site/${site.id}/check`);
                results.push({site, ok: data.read === 'ok', message: data.message || '连接异常'});
            } catch (error) {
                if (error.fatal) {
                    error.message += `（已完成 ${results.length}/${manifest.sites.length}）`;
                    throw error;
                }
                results.push({site, ok: false, message: error.name === 'AbortError' ? '检测超时' : error.message});
            }
            onProgress(results.length, manifest.sites.length, site);
        }
        return results;
    }

    function bind(button) {
        if (!button || button.dataset.apiCheckBound) return;
        button.dataset.apiCheckBound = '1';
        button.addEventListener('click', async () => {
            if (button.disabled) return;
            const html = button.innerHTML;
            const title = button.title;
            const controls = Array.from(document.querySelectorAll('.check-api-btn'));
            const disabled = controls.map(control => control.disabled);
            controls.forEach(control => { control.disabled = true; });
            button.disabled = true;
            button.textContent = '正在准备只读检测…';
            try {
                const results = await runChecks(root.fetch.bind(root), (done, total, site) => {
                    button.textContent = `检测中 (${done}/${total})…`;
                    button.title = `正在检测 ${site.url}，请保持页面打开`;
                });
                const failed = results.filter(result => !result.ok);
                let message = `只读连接检测完成：正常 ${results.length - failed.length} 个，异常 ${failed.length} 个。`;
                if (!results.length) message = '暂无可检测的站点。';
                if (failed.length) message += '\n\n' + failed.map(result => `${result.site.url}: ${result.message}`).join('\n');
                message += '\n写权限未测试。';
                root.alert(message);
                if (results.length) root.location.reload();
            } catch (error) {
                root.alert('检测失败: ' + error.message);
            } finally {
                button.disabled = false;
                button.innerHTML = html;
                button.title = title;
                controls.forEach((control, index) => { control.disabled = disabled[index]; });
            }
        });
    }

    const api = {runChecks, bind};
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else root.SiteApiChecks = api;
})(typeof window !== 'undefined' ? window : globalThis);
