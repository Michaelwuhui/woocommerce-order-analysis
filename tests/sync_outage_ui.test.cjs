const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function element() {
    const classes = new Set();
    return {
        textContent: '', style: {}, children: [], disabled: false, hidden: false,
        classList: {add: (...names) => names.forEach(n => classes.add(n)),
                    remove: (...names) => names.forEach(n => classes.delete(n)),
                    contains: name => classes.has(name)},
        appendChild(node) { this.children.push(node); },
        replaceChildren() { this.children = []; },
        get childElementCount() { return this.children.length; },
    };
}

function runtime(finalStatus) {
    const nodes = new Map(['syncProgressBar', 'syncStatusText', 'syncAvailabilityWarnings',
        'syncLogConsole', 'closeSyncModalBtn', 'cancelSyncBtn'].map(id => [id, element()]));
    const store = new Map();
    const window = {setInterval: () => 1, clearInterval() {}, location: {reload() {}}};
    const context = {
        window, document: {getElementById: id => nodes.get(id), createElement: element, addEventListener() {}},
        localStorage: {setItem: (k,v) => store.set(k,v), getItem: k => store.get(k), removeItem: k => store.delete(k)},
        fetch: async url => ({ok: true, json: async () => url.includes('/status/') ? finalStatus :
            {success: true, run_id: 'test-run', status: {created_at: '2026-10-02T05:00:00Z'}}}),
    };
    vm.runInNewContext(fs.readFileSync(require.resolve('../static/js/sync_runs.js'), 'utf8'), context);
    return {nodes, window, store};
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
