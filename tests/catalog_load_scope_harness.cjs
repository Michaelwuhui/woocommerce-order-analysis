// Independent release contracts: real catalog controller/editor, synthetic I/O.
// No production requests, persistent cache or implementation replicas are used.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {createController, mount} = require('../static/js/product_manager_catalog.js');
const {createEditor} = require('../static/js/product_manager_catalog_editor.js');
const copy = value => JSON.parse(JSON.stringify(value));
const response = (data, status = 200) => ({ok: status < 400, status,
    headers: {get: () => 'application/json'}, json: async () => copy(data)});
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
const keys = snapshot => snapshot.rows.map(row => row.key).sort();
const selectedIds = snapshot => snapshot.selectedSites.map(site => String(site.id)).sort();
const siteIds = snapshot => snapshot.sites.map(site => String(site.id)).sort();
const loadedIds = snapshot => snapshot.loadedScope.siteIds.map(String).sort();

function product(site, id, index = 0) {
    const row = {site_id: String(site), product_id: id, variation_id: 0,
        product_name: 'Merrymi Test Device', name: index % 2 ? 'Berry' : 'Mint',
        sku: `synthetic-${site}-${id}`, brands: ['Merrymi'], flavors: [index % 2 ? 'Berry' : 'Mint'],
        type: 'simple', status: 'publish', manage_stock: true, stock_quantity: 5,
        stock_status: 'instock', regular_price: '10.00', sale_price: '', price: '10.00'};
    row.edit_identity = {sku: row.sku, type: 'simple', parent_id: 0, attributes: []};
    row.edit_before = {manage_stock: true, stock_quantity: 5, stock_status: 'instock', regular_price: '10.00', sale_price: ''};
    return row;
}

function fixture(settings = {}) {
    const sites = [
        {id: 1, url: 'https://one.invalid', manager: 'Alice', country: 'PL', configured: true},
        {id: 2, url: 'https://two.invalid', manager: 'Alice', country: 'CZ', configured: true},
        {id: 3, url: 'https://three.invalid', manager: 'Bob', country: 'PL', configured: true},
        {id: 4, url: 'https://four.invalid', manager: '', country: '', configured: true},
    ];
    const sourceRows = new Map(sites.map(site => [String(site.id), site.id === 1
        ? Array.from({length: 115}, (_, index) => product(1, index ? 1000 + index : 10, index))
        : [product(site.id, 10)]]));
    const actual = new Map([...sourceRows.values()].flat().map(row => [`${row.site_id}:${row.product_id}:0`, copy(row)]));
    const calls = [], writes = [];
    const f = {sites, calls, writes, allowed: new Set(sites.map(site => site.id)), fail: new Set(settings.fail || []), metadataError: false, pageGate: null, rowCountries: new Map()};
    f.productCalls = () => calls.filter(call => call.pathname.endsWith('/catalog-page'));
    f.metadataCalls = () => calls.filter(call => call.pathname.endsWith('/catalog-sites'));
    f.fetch = async (url, options = {}) => {
        const parsed = new URL(url, 'https://application.invalid'), p = parsed.searchParams;
        const call = {url, method: options.method || 'GET', pathname: parsed.pathname, params: p}; calls.push(call);
        assert(parsed.origin === 'https://application.invalid', 'test cannot request an external URL');
        if (parsed.pathname.endsWith('/catalog-sites')) {
            if (f.metadataError) return response({error: 'Synthetic metadata unavailable'}, 503);
            const ids = p.has('site_ids') ? new Set(p.get('site_ids').split(',').filter(Boolean)) : null;
            const visible = sites.filter(site => f.allowed.has(site.id) && (!ids || ids.has(String(site.id))) &&
                (!p.get('manager') || site.manager === p.get('manager')) &&
                (!p.get('country') || site.country === p.get('country')) &&
                (p.get('manager_mode') !== 'unassigned' || !site.manager) &&
                (p.get('country_mode') !== 'unassigned' || !site.country));
            return response({sites: visible});
        }
        if (parsed.pathname.endsWith('/catalog-page')) {
            assert.equal(call.method, 'GET');
            if (f.pageGate) await f.pageGate;
            const id = Number(p.get('site_id'));
            if (!f.allowed.has(id)) return response({error: 'Synthetic permission revoked'}, 403);
            if (f.fail.has(id)) return response({error: 'Synthetic upstream unavailable'}, 504);
            const rows = sourceRows.get(String(id)); assert(rows, 'unknown sites must never be scanned');
            const number = Number(p.get('page')), offset = (number - 1) * 100;
            const current = rows.slice(offset, offset + 100), more = offset + current.length < rows.length;
            return response({site_id: id, parent_id: 0, page: number, per_page: 100,
                rows: current.map(row => f.rowCountries.has(id) ? {...row, country: f.rowCountries.get(id)} : row), source_ids: current.map(row => row.product_id),
                source_statuses: Object.fromEntries(current.map(row => [String(row.product_id), 'publish'])),
                variable_products: [], total: rows.length, total_pages: Math.ceil(rows.length / 100),
                scanned: current.length, has_more: more, next_page: more ? number + 1 : null,
                complete_page: true, warnings: []});
        }
        if (parsed.pathname.endsWith('/catalog-edit-config')) return response({csrf_token: 'synthetic-token',
            sites: sites.map(site => ({...site, can_edit: true, can_clone: true})), clone_targets: []});
        if (parsed.pathname.endsWith('/catalog-item')) {
            const key = `${p.get('site_id')}:${p.get('product_id')}:${p.get('variation_id') || 0}`;
            const row = actual.get(key); assert(row, 'only existing exact rows can be edited');
            if (call.method === 'GET') return response({success: true, item: row,
                identity: row.edit_identity, before: row.edit_before, routing: {mode: 'direct'}, stock_bridge: {present: false, valid: true}});
            assert.equal(call.method, 'PUT');
            assert.equal(options.headers['X-PM-CSRF'], 'synthetic-token');
            const body = JSON.parse(options.body);
            assert.equal(`${body.site_id}:${body.product_id}:${body.variation_id}`, key);
            assert.deepEqual(body.expected_identity, row.edit_identity);
            writes.push({key, body}); Object.assign(row, body.changes); Object.assign(row.edit_before, body.changes);
            return response({success: true, item: row, verification: {status: 'verified', direct_site: true}});
        }
        throw new Error('Unrecognized synthetic endpoint: ' + url);
    };
    f.controller = createController({fetch: f.fetch, emitInterval: 0});
    return f;
}

async function initialMetadataAndIntersection() {
    const f = fixture(), c = f.controller;
    assert.equal(f.calls.length, 0);
    await c.discoverSites();
    assert.equal(f.metadataCalls().length, 1); assert.equal(f.productCalls().length, 0);
    assert.equal(c.getSnapshot().metadataPhase, 'ready');
    assert.equal(c.getSnapshot().availableSites.length, 4);
    c.setLoadScope({manager: 'Alice', country: 'PL', siteId: '1'});
    assert.deepEqual(selectedIds(c.getSnapshot()), ['1']);
    assert.equal(c.getSnapshot().scopeCount, 1);
    const before = f.calls.length;
    await c.load('', {queryMode: 'all'});
    assert.deepEqual(siteIds(c.getSnapshot()), ['1']);
    assert.deepEqual(loadedIds(c.getSnapshot()), ['1']);
    assert(f.productCalls().every(call => call.params.get('site_id') === '1'));
    const metadata = f.calls.slice(before).find(call => call.pathname.endsWith('/catalog-sites'));
    assert.equal(metadata.params.get('manager'), 'Alice'); assert.equal(metadata.params.get('country'), 'PL');
    assert.equal(metadata.params.get('site_ids'), '1');
    assert.match(c.getSnapshot().loadedAt, /^\d{4}-\d{2}-\d{2}T/);
}

async function explicitMultiselectAndMetadataChangesNeverExpand() {
    const f = fixture(), c = f.controller;
    await c.discoverSites();
    assert.equal(c.getSnapshot().loadScope.siteIds, undefined, 'legacy unspecified scope still means all');
    c.setLoadScope({siteIds: ['1', '3'], country: 'PL'});
    const copyScope = c.getSnapshot().loadScope;
    copyScope.siteIds.push('4');
    assert.deepEqual(c.getSnapshot().loadScope.siteIds, ['1', '3'], 'snapshot array cannot mutate pending selections');
    await c.load('', {queryMode: 'all'});
    assert.deepEqual(loadedIds(c.getSnapshot()), ['1', '3']);
    assert(f.productCalls().every(call => ['1', '3'].includes(call.params.get('site_id'))));
    const oldScope = copy(c.getSnapshot().loadedScope), at = c.getSnapshot().loadedAt, oldRows = keys(c.getSnapshot());
    let before = f.calls.length;
    c.setLoadScope({siteIds: []});
    assert.equal(c.getSnapshot().scopeCount, 0);
    await c.load('', {queryMode: 'all'});
    assert.equal(f.calls.length, before, 'clearing every checkbox must not fetch all sites');
    assert.deepEqual(c.getSnapshot().loadedScope, oldScope);
    assert.equal(c.getSnapshot().loadedAt, at);
    assert.deepEqual(keys(c.getSnapshot()), oldRows);
    c.setLoadScope({siteIds: ['1', '3']});
    f.allowed = new Set([2, 4]);
    await c.discoverSites();
    assert.deepEqual(c.getSnapshot().loadScope.siteIds, ['1', '3'], 'metadata refresh keeps the explicit selection boundary');
    assert.equal(c.getSnapshot().scopeCount, 0);
    before = f.calls.length;
    await c.load('', {queryMode: 'all'});
    assert.equal(f.calls.length, before, 'lost selected sites must not fall back to other visible sites');
    assert.deepEqual(c.getSnapshot().loadedScope, oldScope);
    assert.equal(c.getSnapshot().loadedAt, at);
    f.sites.push({id: 5, url: 'https://new.invalid', manager: 'Alice', country: 'PL', configured: true});
    f.allowed = new Set([3, 4, 5]);
    await c.discoverSites();
    assert.deepEqual(selectedIds(c.getSnapshot()), ['3'], 'new metadata sites are not implicitly checked');
    assert.deepEqual(c.getSnapshot().loadScope.siteIds, ['1', '3']);
}

async function emptyAndUnknownScopeNeverExpandOrRelabel() {
    const f = fixture(), c = f.controller;
    await c.discoverSites(); c.setLoadScope({manager: 'Alice', country: 'PL'});
    await c.load('', {queryMode: 'all'});
    const original = c.getSnapshot(), oldKeys = keys(original), loaded = copy(original.loadedScope), loadedAt = original.loadedAt;
    for (const scope of [{manager: 'Alice', country: 'PL', siteId: '3'}, {siteId: '9999'}]) {
        const before = f.calls.length;
        c.setLoadScope(scope); assert.equal(c.getSnapshot().scopeCount, 0);
        assert.deepEqual(keys(c.getSnapshot()), oldKeys);
        assert.deepEqual(c.getSnapshot().loadedScope, loaded); assert.equal(c.getSnapshot().loadedAt, loadedAt);
        await c.load('', {queryMode: 'all'});
        assert.equal(f.calls.length, before, 'known empty scope makes no request at all');
        assert.deepEqual(keys(c.getSnapshot()), oldKeys, 'empty load must preserve old proven rows');
        assert.deepEqual(c.getSnapshot().loadedScope, loaded);
        c.resetFilters(); assert.deepEqual(keys(c.getSnapshot()), oldKeys);
    }
    const broken = fixture(); broken.metadataError = true;
    await broken.controller.discoverSites();
    broken.controller.setLoadScope({manager: 'Alice', country: 'PL', siteId: '1'});
    await broken.controller.load('', {queryMode: 'all'});
    assert.equal(broken.productCalls().length, 0, 'unavailable metadata is never an all-sites grant');
}

async function localCountrySearchAndPaginationNeverFetch() {
    const f = fixture(), c = f.controller;
    await c.discoverSites(); await c.load('', {queryMode: 'all'});
    const before = f.calls.length, original = keys(c.getSnapshot());
    assert.equal(c.getSnapshot().pages, 2);
    assert(c.getSnapshot().options.country.some(option => option.value === 'value:pl'));
    for (let iteration = 0; iteration < 4; iteration++) {
        c.setFilters({country: 'value:pl', text: 'mint', brand: 'value:merrymi'});
        const filtered = c.getSnapshot().filtered;
        assert(filtered.length > 0 && filtered.every(row => row.site.country === 'PL' && row.flavors.includes('Mint')));
        c.setPage(2); c.resetFilters(); c.setFilters({country: 'value:cz', flavor: 'value:mint'});
        assert.deepEqual(c.getSnapshot().filtered.map(row => row.site_id), ['2']);
        c.resetFilters(); c.setFilters({country: 'missing:country'});
        assert.deepEqual(c.getSnapshot().filtered.map(row => row.site_id), ['4']);
        c.resetFilters(); c.setPage(2); assert.equal(c.getSnapshot().pageRows.length, 18);
    }
    assert.equal(f.calls.length, before); assert.deepEqual(keys(c.getSnapshot()), original);
    c.setLoadScope({manager: 'Alice', country: 'CZ', siteId: '2'});
    c.resetFilters(); assert.deepEqual(keys(c.getSnapshot()), original, 'clear local filters cannot load a new scope');
    assert.equal(f.calls.length, before);
}

async function freshCountryFacetAndLiteralScopeLabels() {
    const f = fixture(), c = f.controller;
    f.rowCountries.set(1, 'HU'); f.rowCountries.set(2, '');
    await c.discoverSites(); c.setLoadScope({manager: 'Alice'});
    await c.load('', {queryMode: 'all'});
    const before = f.calls.length;
    assert(c.getSnapshot().options.country.some(option => option.value === 'value:hu'));
    c.setFilters({country: 'value:hu'});
    assert.equal(c.getSnapshot().filtered.length, 115);
    assert(c.getSnapshot().filtered.every(row => row.site_id === '1' && row.country === 'HU'));
    c.resetFilters(); c.setFilters({country: 'missing:country'});
    assert.deepEqual(c.getSnapshot().filtered.map(row => row.site_id), ['2']);
    assert.equal(f.calls.length, before);

    const literal = fixture(); literal.sites[2].manager = 'all'; literal.sites[2].country = '__unassigned__';
    await literal.controller.discoverSites(); literal.controller.setLoadScope({manager: 'all', country: '__unassigned__'});
    assert.deepEqual(selectedIds(literal.controller.getSnapshot()), ['3']);
    await literal.controller.load('', {queryMode: 'all'});
    assert(literal.productCalls().every(call => call.params.get('site_id') === '3'));
    literal.controller.setLoadScope({managerMode: 'unassigned', countryMode: 'unassigned'});
    assert.deepEqual(selectedIds(literal.controller.getSnapshot()), ['4']);
}

async function retryRefreshAndPermissionChangeUseSubmittedScope() {
    const f = fixture({fail: [2]}), c = f.controller;
    await c.discoverSites(); c.setLoadScope({manager: 'Alice', siteIds: ['1', '2']});
    await c.load('', {queryMode: 'all'});
    assert.equal(c.getSnapshot().phase, 'incomplete');
    const loaded = copy(c.getSnapshot().loadedScope), at = c.getSnapshot().loadedAt;
    c.setLoadScope({manager: 'Bob', country: 'PL', siteIds: ['3', '4']});
    const before = f.calls.length; f.fail.clear();
    await c.retryIncomplete();
    assert(f.calls.slice(before).every(call => call.pathname.endsWith('/catalog-page') && call.params.get('site_id') === '2'));
    assert.deepEqual(c.getSnapshot().loadedScope, loaded); assert.equal(c.getSnapshot().loadedAt, at);
    const refreshStart = f.calls.length;
    await c.refresh();
    assert(f.calls.slice(refreshStart).filter(call => call.pathname.endsWith('/catalog-page')).every(call => ['1', '2'].includes(call.params.get('site_id'))));
    assert.deepEqual(loadedIds(c.getSnapshot()), ['1', '2']);
    f.allowed = new Set([3, 4]);
    const revokedStart = f.productCalls().length;
    const previousKeys = keys(c.getSnapshot());
    await c.refresh();
    assert.equal(f.productCalls().length, revokedStart, 'revoked original scope must not become newly visible sites');
    assert.deepEqual(keys(c.getSnapshot()), previousKeys, 'failed refresh retains explicitly labelled old data');
    assert.deepEqual(loadedIds(c.getSnapshot()), ['1', '2']);
    assert.match(c.getSnapshot().error, /权限|负责人|范围/);
}

async function preselectionDuringLoadDoesNotChangeInflightScan() {
    const f = fixture(), c = f.controller; await c.discoverSites();
    c.setLoadScope({manager: 'Alice'});
    let release; f.pageGate = new Promise(resolve => { release = resolve; });
    const pending = c.load('', {queryMode: 'all'});
    await until(() => f.productCalls().length > 0, 'product request started');
    c.setLoadScope({manager: 'Bob', country: 'PL', siteId: '3'});
    release(); await pending;
    assert.deepEqual(loadedIds(c.getSnapshot()), ['1', '2']);
    assert.deepEqual(selectedIds(c.getSnapshot()), ['3']);
    assert(f.productCalls().every(call => ['1', '2'].includes(call.params.get('site_id'))));
}

async function failedReplacementLeavesOriginalWorkRetryable() {
    const f = fixture(), c = f.controller; await c.discoverSites();
    c.setLoadScope({manager: 'Alice'});
    let release; f.pageGate = new Promise(resolve => { release = resolve; });
    const original = c.load('', {queryMode: 'all'});
    await until(() => f.productCalls().length > 0, 'original scan before replacement');
    const originalScope = copy(c.getSnapshot().loadedScope);
    c.setLoadScope({manager: 'Bob', siteId: '3'}); f.metadataError = true;
    await c.load('', {queryMode: 'all'});
    release(); await original;
    const state = c.getSnapshot();
    assert.deepEqual(state.loadedScope, originalScope);
    assert.equal(state.busy, false);
    assert.equal(state.counts.incompleteSites, 2, 'aborted original sites remain eligible for the retry button');
    assert(state.siteStates.every(site => !['pending', 'loading'].includes(site.status)), 'retained old progress cannot pretend requests are still running');
}

async function draftsAndBulkTargetsStayInLoadedRows() {
    const f = fixture(), c = f.controller; await c.discoverSites();
    c.setLoadScope({manager: 'Alice'}); await c.load('', {queryMode: 'all'});
    const editor = createEditor({catalog: c, fetch: f.fetch, autoPoll: false, confirm: () => false}); await editor.ready();
    const chosen = ['1:10:0', '2:10:0']; editor.select(chosen); editor.setDraft(chosen[0], 'regular_price', '14.50');
    const count = f.calls.length;
    c.setLoadScope({manager: 'Bob', country: 'PL', siteId: '3'});
    c.setFilters({country: 'value:cz'}); c.setPage(2); c.resetFilters();
    assert.deepEqual(editor.getSnapshot().selected.sort(), chosen);
    assert.equal(editor.getDraft(chosen[0]).regular_price, '14.50');
    assert.equal(f.calls.length, count); assert.equal(editor.canReload(), false, 'declined unsaved reload stays blocked');
    assert.equal(editor.getSnapshot().draftCount, 1);
    await editor.saveDrafts([chosen[0]]);
    assert.deepEqual(f.writes.map(write => write.key), [chosen[0]]);
    editor.select(chosen);
    await editor.bulk('price', {regular_price: '19.50'});
    assert.deepEqual(f.writes.slice(1).map(write => write.key).sort(), chosen);
    assert(f.writes.every(write => ['1', '2'].includes(String(write.body.site_id))));
    assert(!f.calls.some(call => call.pathname.endsWith('/catalog-item') && call.params.get('site_id') === '3'));
}

async function until(condition, label) {
    for (let index = 0; index < 300; index++) { if (condition()) return; await sleep(2); }
    throw new Error('Timed out waiting for ' + label);
}

// Minimal DOM surface. It stores attributes/options/events only; all scope,
// matching, loading, filtering and editing decisions remain production JS.
class Element {
    constructor(id) {
        this.id = id; this.value = ''; this.textContent = ''; this.disabled = false;
        this.dataset = {}; this.style = {}; this.options = []; this.listeners = new Map();
        const classes = new Set();
        this.classList = {toggle(name, enabled) { if (enabled) classes.add(name); else classes.delete(name); },
            add: name => classes.add(name), remove: name => classes.delete(name), contains: name => classes.has(name)};
    }
    set innerHTML(value) {
        this.html = value;
        this.options = [...value.matchAll(/<option\b[^>]*value="([^"]*)"[^>]*>([\s\S]*?)<\/option>/g)].map(match => ({value: match[1], textContent: match[2].replace(/<[^>]*>/g, '')}));
    }
    get innerHTML() { return this.html || ''; }
    get selectedOptions() { return this.options.filter(option => option.selected); }
    addEventListener(type, callback) { const listeners = this.listeners.get(type) || []; listeners.push(callback); this.listeners.set(type, listeners); }
    async fire(type) { for (const callback of this.listeners.get(type) || []) await callback({target: this, preventDefault() {}}); }
    querySelectorAll() { return []; }
    setAttribute(name, value) { this[name] = value; }
    removeAttribute(name) { delete this[name]; }
}

async function mountedPageLoadsOnlyMetadataAndExplainsReuse() {
    const f = fixture();
    const ids = ['pmCatalogPane', 'pmCatalogQueryMode', 'pmCatalogBrandInput', 'pmCatalogSearch',
        'pmCatalogFlavorInputWrap', 'pmCatalogBrandInputWrap', 'pmCatalogLoad', 'pmCatalogStop',
        'pmCatalogRetry', 'pmCatalogRefresh', 'pmCatalogStatus', 'pmCatalogCounts', 'pmCatalogProgress',
        'pmCatalogResults', 'pmCatalogResetFilters', 'pmCatalogLoadManager', 'pmCatalogLoadCountry',
        'pmCatalogLoadSite', 'pmCatalogScopeCount', 'pmCatalogScopeStatus', 'pmCatalogLoadedScope',
        ...['Manager', 'Country', 'Site', 'Brand', 'Product', 'Flavor', 'Stock', 'Publish', 'Text'].map(name => `pmCatalog${name}Filter`)];
    const nodes = new Map(ids.map(id => [id, new Element(id)]));
    const document = {getElementById: id => nodes.get(id) || null, activeElement: null};
    nodes.get('pmCatalogQueryMode').value = 'all';
    const synced = [];
    const previousSelects = global.ProductManagerSelects;
    global.ProductManagerSelects = {sync: element => synced.push(element.id)};
    const c = mount(document, {fetch: f.fetch});
    await until(() => c.getSnapshot().metadataPhase === 'ready', 'mount metadata'); await sleep(3);
    assert.equal(f.metadataCalls().length, 1); assert.equal(f.productCalls().length, 0);
    const siteElement = nodes.get('pmCatalogLoadSite');
    assert.equal(siteElement.multiple, true);
    assert.equal(siteElement.options.length, 4);
    assert(siteElement.options.every(option => option.value && option.selected), 'initial native options are actual sites and all checked');
    assert(synced.includes('pmCatalogLoadSite'), 'programmatic initial checkboxes synchronize the searchable widget');
    const selectSites = async ids => {
        siteElement.options.forEach(option => { option.selected = ids.includes(option.value); });
        await siteElement.fire('change');
    };
    const setOption = async (id, label) => {
        const element = nodes.get(id), option = element.options.find(option => option.textContent === label);
        assert(option, `${id} contains ${label}`); element.value = option.value; await element.fire('change');
    };
    await selectSites(['1', '2']);
    await setOption('pmCatalogLoadManager', 'Alice'); await setOption('pmCatalogLoadCountry', 'PL');
    assert.deepEqual(selectedIds(c.getSnapshot()), ['1']);
    assert.deepEqual(c.getSnapshot().loadScope.siteIds, ['1', '2']);
    assert.equal(c.getSnapshot().loadScope.siteId, '');
    assert.deepEqual(siteElement.selectedOptions.map(option => option.value), ['1', '2'], 'owner/country intersection does not uncheck other sites');
    await nodes.get('pmCatalogLoad').fire('click');
    await until(() => !c.getSnapshot().busy && c.getSnapshot().loadedAt, 'mount loaded'); await sleep(3);
    const range = nodes.get('pmCatalogLoadedScope').textContent;
    assert.match(range, /Alice/); assert.match(range, /PL/);
    assert.match(range, /本次勾选站点/); assert.doesNotMatch(range, /范围内全部站点/);
    const before = f.calls.length;
    const nativeOptions = siteElement.options;
    await setOption('pmCatalogLoadManager', 'Bob'); await sleep(3);
    assert.equal(nodes.get('pmCatalogLoadedScope').textContent, range, 'pending controls cannot relabel displayed results');
    nodes.get('pmCatalogCountryFilter').value = 'value:pl'; await nodes.get('pmCatalogCountryFilter').fire('change');
    nodes.get('pmCatalogTextFilter').value = 'mint'; await nodes.get('pmCatalogTextFilter').fire('input');
    await nodes.get('pmCatalogResetFilters').fire('click');
    await sleep(3);
    assert.equal(f.calls.length, before);
    assert.equal(siteElement.options, nativeOptions, 'local renders preserve native option identity and widget focus');
    assert.deepEqual(siteElement.selectedOptions.map(option => option.value), ['1', '2']);
    await selectSites([]); await sleep(3);
    assert.deepEqual(c.getSnapshot().loadScope.siteIds, []);
    assert.equal(c.getSnapshot().scopeCount, 0);
    assert.equal(nodes.get('pmCatalogLoad').disabled, true);
    assert.equal(nodes.get('pmCatalogLoadedScope').textContent, range);
    assert.equal(f.calls.length, before);
    await selectSites(['3']); await sleep(3);
    f.metadataError = true;
    await nodes.get('pmCatalogLoad').fire('click'); await sleep(3);
    assert.equal(nodes.get('pmCatalogLoadedScope').textContent, range);
    assert.match(nodes.get('pmCatalogStatus').textContent, /Synthetic metadata unavailable/);
    assert.equal(nodes.get('pmCatalogStatus').classList.contains('text-warning'), true);
    const template = fs.readFileSync(path.join(__dirname, '../templates/product_manager.html'), 'utf8');
    assert.match(template, /当前页|本页/); assert.match(template, /重新加载|手动刷新/);
    assert.match(template, /负责人/); assert.match(template, /国家/);
    assert.match(template, /id="pmCatalogLoadSite"[^>]*\bmultiple\b/);
    assert.match(template, /product_manager_selects\.css.*20261016/);
    assert.match(template, /product_manager_selects\.js.*20261016/);
    assert.equal((template.match(/<script\s*>/g) || []).length, 1, 'legacy inline script remains a single block');
    global.ProductManagerSelects = previousSelects;
}

(async () => {
    await initialMetadataAndIntersection(); await explicitMultiselectAndMetadataChangesNeverExpand(); await emptyAndUnknownScopeNeverExpandOrRelabel();
    await localCountrySearchAndPaginationNeverFetch(); await freshCountryFacetAndLiteralScopeLabels();
    await retryRefreshAndPermissionChangeUseSubmittedScope();
    await preselectionDuringLoadDoesNotChangeInflightScan(); await failedReplacementLeavesOriginalWorkRetryable();
    await draftsAndBulkTargetsStayInLoadedRows();
    await mountedPageLoadsOnlyMetadataAndExplainsReuse();
    process.stdout.write('Independent catalog scope contracts: metadata-only mount, intersection/empty boundaries, immutable loaded scope, local country/search/pages, original retry/refresh, revocation, drafts and exact bulk targets passed.\n');
})().catch(error => { process.stderr.write(error.stack + '\n'); process.exitCode = 1; });
