// Independent scope/selection release contracts against the real JS modules.
const assert = require('node:assert/strict');
const {createController} = require('../static/js/product_manager_catalog.js');
const {createEditor} = require('../static/js/product_manager_catalog_editor.js');
const copy = value => JSON.parse(JSON.stringify(value));
const response = (data, status = 200) => ({ok: status < 400, status,
    headers: {get: () => 'application/json'}, json: async () => copy(data)});
const ids = sites => sites.map(site => String(site.id)).sort();
const loadedIds = controller => controller.getSnapshot().loadedScope.siteIds.slice().sort();

function fixture() {
    const sites = [
        {id: 1, manager: 'Alice', country: 'PL'}, {id: 3, manager: 'Alice', country: 'PL'},
        {id: 7, manager: 'Alice', country: 'CZ'}, {id: 9, manager: 'Bob', country: 'PL'},
    ].map(site => ({...site, url: `https://site-${site.id}.invalid`, configured: true}));
    const rows = new Map(sites.map(site => {
        const row = {site_id: String(site.id), product_id: 10, variation_id: 0,
            type: 'simple', name: 'Mint', product_name: 'Merrymi 12000', sku: `synthetic-${site.id}`,
            status: 'publish', flavors: ['Mint'], brands: ['Merrymi'], country: site.country,
            manage_stock: true, stock_quantity: 5, stock_status: 'instock', price: '10.00', regular_price: '10.00', sale_price: '',
            edit_identity: {sku: `synthetic-${site.id}`, type: 'simple', parent_id: 0, attributes: []},
            edit_before: {manage_stock: true, stock_quantity: 5, stock_status: 'instock', regular_price: '10.00', sale_price: ''}};
        return [String(site.id), row];
    }));
    const f = {sites, calls: [], writes: [], allowed: new Set(sites.map(site => site.id)), failed: new Set()};
    f.pageCalls = () => f.calls.filter(call => call.path.endsWith('/catalog-page'));
    f.fetch = async (url, options = {}) => {
        const target = new URL(url, 'https://synthetic-app.invalid'), params = target.searchParams;
        const call = {url, path: target.pathname, method: options.method || 'GET', params}; f.calls.push(call);
        assert.equal(target.origin, 'https://synthetic-app.invalid');
        if (call.path.endsWith('/catalog-sites')) {
            const wanted = params.has('site_ids') ? params.get('site_ids').split(',') : null;
            return response({sites: sites.filter(site => f.allowed.has(site.id) && (!wanted || wanted.includes(String(site.id))) &&
                (!params.get('manager') || params.get('manager') === site.manager) && (!params.get('country') || params.get('country') === site.country))});
        }
        if (call.path.endsWith('/catalog-page')) {
            const site = Number(params.get('site_id'));
            assert(!params.has('site_ids') || params.get('site_ids').split(',').includes(String(site)), 'each actual page is within the frozen checkbox selection');
            if (!f.allowed.has(site)) return response({error: 'Synthetic permission revoked'}, 403);
            if (f.failed.has(site)) return response({error: 'Synthetic unavailable'}, 504);
            assert(rows.has(String(site)), 'unknown sites cannot be read');
            return response({site_id: site, parent_id: 0, page: 1, per_page: 100,
                rows: [rows.get(String(site))], source_ids: [10], source_statuses: {'10': 'publish'}, variable_products: [],
                total: 1, total_pages: 1, scanned: 1, complete_page: true, has_more: false, next_page: null, warnings: []});
        }
        if (call.path.endsWith('/catalog-edit-config')) return response({csrf_token: 'synthetic-only',
            sites: sites.map(site => ({...site, can_edit: true, can_clone: true})), clone_targets: []});
        if (call.path.endsWith('/catalog-item')) {
            const site = params.get('site_id'), row = rows.get(site); assert(row);
            if (call.method === 'GET') return response({success: true, item: row, identity: row.edit_identity,
                before: row.edit_before, stock_bridge: {present: false, valid: true}, routing: {mode: 'direct'}});
            assert.equal(call.method, 'PUT'); assert.equal(options.headers['X-PM-CSRF'], 'synthetic-only');
            const body = JSON.parse(options.body);
            assert.equal(`${body.site_id}:${body.product_id}:${body.variation_id}`, `${site}:10:0`);
            assert.deepEqual(body.expected_identity, row.edit_identity);
            f.writes.push({site, body}); Object.assign(row, body.changes); Object.assign(row.edit_before, body.changes);
            return response({success: true, item: row, verification: {status: 'verified', direct_site: true}});
        }
        throw new Error('Unexpected synthetic URL ' + url);
    };
    f.controller = createController({fetch: f.fetch, emitInterval: 0});
    return f;
}

async function discontiguousSelectionIntersectionAndEmpty() {
    const f = fixture(), c = f.controller; await c.discoverSites();
    assert.equal(f.pageCalls().length, 0);
    c.setLoadScope({siteIds: ['1', '7'], manager: 'Alice'});
    assert.deepEqual(ids(c.getSnapshot().selectedSites), ['1', '7']);
    await c.load('', {queryMode: 'all'});
    assert.deepEqual(loadedIds(c), ['1', '7']);
    assert.deepEqual(f.pageCalls().map(call => call.params.get('site_id')).sort(), ['1', '7']);
    assert(f.pageCalls().every(call => call.params.get('site_ids') === '1,7'));
    const originalKeys = c.getSnapshot().rows.map(row => row.key), at = c.getSnapshot().loadedAt;
    const before = f.calls.length;
    c.setLoadScope({siteIds: [], country: 'PL'});
    assert.equal(c.getSnapshot().scopeCount, 0); await c.load('', {queryMode: 'all'});
    c.resetFilters();
    assert.equal(f.calls.length, before, 'clearing checkboxes never falls back to every site');
    assert.deepEqual(c.getSnapshot().rows.map(row => row.key), originalKeys); assert.equal(c.getSnapshot().loadedAt, at);
    c.setLoadScope({siteIds: ['1', '7'], country: 'PL'});
    assert.deepEqual(ids(c.getSnapshot().selectedSites), ['1']);
    await c.load('', {queryMode: 'all'}); assert.deepEqual(loadedIds(c), ['1']);
}

async function refreshRetryAndMetadataRevocationKeepBoundary() {
    const f = fixture(), c = f.controller; await c.discoverSites();
    c.setLoadScope({siteIds: ['1', '7'], manager: 'Alice'}); f.failed.add(7);
    await c.load('', {queryMode: 'all'});
    const at = c.getSnapshot().loadedAt, originalScope = copy(c.getSnapshot().loadedScope);
    c.setLoadScope({siteIds: ['9'], manager: 'Bob', country: 'PL'});
    assert.equal(c.getSnapshot().loadedAt, at); assert.deepEqual(c.getSnapshot().loadedScope, originalScope);
    const before = f.calls.length; f.failed.clear(); await c.retryIncomplete();
    assert(f.calls.slice(before).every(call => call.path.endsWith('/catalog-page') && call.params.get('site_id') === '7'));
    const refreshing = f.pageCalls().length; await c.refresh();
    assert.deepEqual(f.pageCalls().slice(refreshing).map(call => call.params.get('site_id')).sort(), ['1', '7']);

    c.setLoadScope({siteIds: ['1', '7'], manager: 'Alice', country: ''});
    f.allowed = new Set([3, 9]); await c.discoverSites();
    assert.deepEqual(c.getSnapshot().loadScope.siteIds, ['1', '7'], 'metadata refresh preserves explicit intent even when all checked IDs disappear');
    assert.equal(c.getSnapshot().scopeCount, 0);
    const noFallback = f.calls.length; await c.load('', {queryMode: 'all'}); await c.refresh();
    assert.equal(f.calls.length, noFallback, 'missing selected IDs cannot silently become newly visible metadata IDs');
    assert.deepEqual(loadedIds(c), ['1', '7']);
}

async function localFilteringPreservesDraftsAndActualWriteTargets() {
    const f = fixture(), c = f.controller; await c.discoverSites();
    c.setLoadScope({siteIds: ['1', '7']}); await c.load('', {queryMode: 'all'});
    const editor = createEditor({catalog: c, fetch: f.fetch, autoPoll: false, confirm: () => false}); await editor.ready();
    editor.select(['1:10:0']); editor.setDraft('1:10:0', 'regular_price', '12.00');
    const before = f.calls.length;
    for (const text of ['Mint', 'MERRYMI', 'no matching product', '']) {
        c.setFilters({text, country: 'value:cz'}); c.setPage(2); c.resetFilters();
    }
    c.setLoadScope({siteIds: ['3', '9']});
    assert.equal(f.calls.length, before); assert.deepEqual(editor.getSnapshot().selected, ['1:10:0']);
    assert.equal(editor.getDraft('1:10:0').regular_price, '12.00');
    await editor.saveDrafts(['1:10:0']);
    assert.deepEqual(f.writes.map(write => write.site), ['1']);
    editor.select(['7:10:0']); c.setFilters({country: 'value:pl'});
    await editor.bulk('price', {regular_price: '15.00'});
    assert.deepEqual(f.writes.map(write => write.site), ['1', '7']);
    assert(!f.calls.some(call => call.path.endsWith('/catalog-item') && ['3', '9'].includes(call.params.get('site_id'))));
}

(async () => {
    await discontiguousSelectionIntersectionAndEmpty(); await refreshRetryAndMetadataRevocationKeepBoundary();
    await localFilteringPreservesDraftsAndActualWriteTargets();
    process.stdout.write('Independent multiselect contracts passed: discontiguous/intersected/empty scope, original refresh/retry, vanished permissions, local filtering and exact selected writes.\n');
})().catch(error => { process.stderr.write(error.stack + '\n'); process.exitCode = 1; });
