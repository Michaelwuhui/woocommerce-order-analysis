const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function element() {
    const classes = new Set();
    return {
        textContent: '', style: {}, children: [], disabled: false, hidden: false,
        dataset: {}, listeners: {},
        addEventListener(name, callback) { this.listeners[name] = callback; },
        classList: {add: (...names) => names.forEach(n => classes.add(n)),
                    remove: (...names) => names.forEach(n => classes.delete(n)),
                    contains: name => classes.has(name)},
        appendChild(node) { this.children.push(node); },
        replaceChildren() { this.children = []; },
        get childElementCount() { return this.children.length; },
    };
}

function runtime(finalStatus, buttonEndpoint) {
    const nodes = new Map(['syncProgressBar', 'syncStatusText', 'syncAvailabilityWarnings',
        'syncLogConsole', 'closeSyncModalBtn', 'cancelSyncBtn'].map(id => [id, element()]));
    const store = new Map();
    const requests = [];
    const ready = [];
    if (buttonEndpoint !== undefined) {
        const button = element();
        if (buttonEndpoint) button.dataset.syncEndpoint = buttonEndpoint;
        nodes.set('syncAllBtn', button);
    }
    const window = {setInterval: () => 1, clearInterval() {}, location: {reload() {}}};
    const context = {
        window, document: {getElementById: id => nodes.get(id), createElement: element,
            addEventListener(event, callback) { if (event === 'DOMContentLoaded') ready.push(callback); }},
        localStorage: {setItem: (k,v) => store.set(k,v), getItem: k => store.get(k), removeItem: k => store.delete(k)},
        fetch: async (url, options) => {
            requests.push({url, options});
            return {ok: true, json: async () => url.includes('/status/') ? finalStatus :
                {success: true, run_id: 'test-run', status: {created_at: '2026-10-02T05:00:00Z'}}};
        },
    };
    vm.runInNewContext(fs.readFileSync(require.resolve('../static/js/sync_runs.js'), 'utf8'), context);
    return {nodes, window, store, requests, ready};
}

for (const endpoint of ['/api/sync/own', '/api/sync/all', '']) {
    test(`quick sync button uses its authorized endpoint: ${endpoint || 'settings default'}`, async () => {
        const {nodes, requests, ready} = runtime(
            {status: 'success', sites: [], logs: []}, endpoint);
        ready.forEach(callback => callback());
        nodes.get('syncAllBtn').listeners.click();
        await new Promise(resolve => setImmediate(resolve));
        const posts = requests.filter(request => request.options?.method === 'POST');
        assert.deepEqual(posts.map(request => request.url), [endpoint || '/api/sync/all']);
        assert.equal(nodes.get('syncAllBtn').disabled, false);
        assert.equal(nodes.get('syncStatusText').classList.contains('text-danger'), false);
    });
}

test('partial sync finishes with an outage warning and does not keep polling', async () => {
    const status = {status: 'error', outcome: 'partial', total_sites: 2, completed_sites: 2,
        message: '1 个站点完成，1 个站点暂不可用，将自动复查',
        sites: [{url: 'https://bad.invalid', temporarily_unavailable: true,
            availability_message: 'DNS 解析失败；最早 14:00 自动复查 <img src=x>'}], logs: []};
    const {nodes, window, store} = runtime(status);
    await window.WooSyncRuns.start('/api/sync/all');
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(nodes.get('syncProgressBar').style.width, '100%');
    assert.equal(nodes.get('syncProgressBar').classList.contains('bg-warning'), true);
    assert.equal(nodes.get('syncProgressBar').classList.contains('bg-danger'), false);
    assert.equal(nodes.get('closeSyncModalBtn').disabled, false);
    assert.equal(store.has('wooAnalysisActiveSyncRun'), false);
    assert.equal(nodes.get('syncAvailabilityWarnings').hidden, false);
    assert.equal(nodes.get('syncAvailabilityWarnings').children[0].textContent,
        'https://bad.invalid：DNS 解析失败；最早 14:00 自动复查 <img src=x>');
});

test('healthy completion uses a success display without outage warnings', async () => {
    const {nodes, window} = runtime({status: 'success', outcome: 'success', sites: [], logs: []});
    await window.WooSyncRuns.start('/api/sync/all');
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(nodes.get('syncProgressBar').classList.contains('bg-success'), true);
    assert.equal(nodes.get('syncAvailabilityWarnings').hidden, true);
});

test('a recheck with every store still unavailable uses a warning and finishes', async () => {
    const {nodes, window} = runtime({status: 'error', outcome: 'unavailable', sites: [], logs: [],
        message: '站点自动复查已结束：0 个站点恢复，9 个站点暂不可用，将自动复查'});
    await window.WooSyncRuns.start('/api/sync/all');
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(nodes.get('syncProgressBar').classList.contains('bg-warning'), true);
    assert.equal(nodes.get('syncProgressBar').classList.contains('bg-danger'), false);
    assert.equal(nodes.get('closeSyncModalBtn').disabled, false);
    assert.match(nodes.get('syncStatusText').textContent, /站点自动复查已结束/);
});
