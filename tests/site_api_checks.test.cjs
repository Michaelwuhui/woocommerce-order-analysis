const test = require('node:test');
const assert = require('node:assert/strict');
const {runChecks, bind} = require('../static/js/site_api_checks.js');

function response(data, status = 200) {
    return {ok: status >= 200 && status < 300, status, redirected: false, json: async () => data};
}
const sites = [{id: 11, url: 'https://one.example'}, {id: 22, url: 'https://two.example'}];
const manifest = () => response({success: true, mode: 'sequential', sites});

test('checks all sites sequentially and counts read success with unknown write permission', async () => {
    let active = 0;
    const calls = [], progress = [];
    const result = await runChecks(async (url, options) => {
        assert.equal(options.method, 'POST');
        calls.push(url);
        if (calls.length === 1) return manifest();
        assert.equal(active++, 0);
        await new Promise(resolve => setImmediate(resolve));
        active--;
        return response({success: true, read: 'ok', write: 'unknown'});
    }, (done, total) => progress.push([done, total]));
    assert.deepEqual(calls, ['/api/sites/check-all', '/api/site/11/check', '/api/site/22/check']);
    assert.equal(result.filter(r => r.ok).length, 2);
    assert.deepEqual(progress.at(-1), [2, 2]);
});

test('continues after a failing site and reports partial results accurately', async () => {
    let call = 0;
    const result = await runChecks(async () => ++call === 1 ? manifest() : response({
        success: true, read: call === 2 ? 'error' : 'ok', message: 'HTTP 429', write: 'unknown'
    }));
    assert.deepEqual(result.map(r => r.ok), [false, true]);
});

test('network failure does not prevent checking the next site', async () => {
    let call = 0;
    const result = await runChecks(async () => {
        if (++call === 1) return manifest();
        if (call === 2) throw new TypeError('network error');
        return response({success: true, read: 'ok'});
    });
    assert.deepEqual(result.map(r => r.ok), [false, true]);
});

test('session expiry stops the remaining requests and includes progress', async () => {
    let call = 0;
    await assert.rejects(runChecks(async () => ++call === 1 ? manifest() : response({}, 401)), /已完成 0\/2/);
    assert.equal(call, 2);
});

test('empty manifest finishes without invoking any site endpoint', async () => {
    let call = 0;
    assert.deepEqual(await runChecks(async () => {
        call++;
        return response({success: true, mode: 'sequential', sites: []});
    }), []);
    assert.equal(call, 1);
});

test('old backend response gives an actionable refresh message', async () => {
    await assert.rejects(runChecks(async () => response({success: true, check_id: 10})), /刷新/);
});

test('button failure restores controls and ignores duplicate clicks', async () => {
    const original = {fetch: global.fetch, document: global.document, alert: global.alert};
    let handler, calls = 0, release;
    const control = {disabled: false};
    const button = {dataset: {}, disabled: false, innerHTML: '检测全部', title: '检测',
        addEventListener: (_, listener) => { handler = listener; }};
    global.document = {querySelectorAll: () => [control]};
    global.alert = () => {};
    global.fetch = async () => {
        calls++;
        await new Promise(resolve => { release = resolve; });
        return response({error: 'unavailable'}, 503);
    };
    try {
        bind(button);
        const pending = handler();
        assert.equal(button.disabled, true);
        assert.equal(control.disabled, true);
        await handler();
        assert.equal(calls, 1);
        release();
        await pending;
        assert.equal(button.disabled, false);
        assert.equal(control.disabled, false);
        assert.equal(button.innerHTML, '检测全部');
    } finally {
        Object.assign(global, original);
    }
});
