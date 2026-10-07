// Exercise the production cross-site controller with synthetic Woo pages only.
const assert = require('node:assert/strict');
const {createController, mount, resultHtml, safeUrl} = require('../static/js/product_manager_catalog.js');
const sleep = milliseconds => new Promise(resolve => setTimeout(resolve, milliseconds));
const sites = [1, 2, 3, 4].map(id => ({id, url: `https://site${id}.example`, manager: id < 3 ? 'Michael' : 'Anna'}));
const response = (data, status = 200, type = 'application/json') => ({status, ok: status < 400, headers: {get: () => type}, json: async () => data});
const row = (site, product, variation = 0, overrides = {}) => Object.assign({site_id: site, product_id: product, variation_id: variation,
    product_name: 'ELFBAR 600', name: variation ? 'Blueberry-Ice' : 'ELFBAR 600 Blueberry Ice', sku: `${site}-blue-${variation || product}`,
    flavors: ['Blueberry Ice'], brands: ['ELFBAR'], stock_status: 'instock', stock_quantity: 7, manage_stock: true,
    price: '12.00', status: 'publish', permalink: `https://site${site}.example/product/${product}`}, overrides);
const page = (site, parent, number, ids, rows, overrides = {}) => Object.assign({site_id: site, parent_id: parent, page: number, per_page: 50,
    source_ids: ids, source_statuses: Object.fromEntries(ids.map(id => [String(id), 'publish'])), scanned: ids.length, rows, variable_products: [], total: ids.length, total_pages: 1,
    has_more: false, next_page: null, complete_page: true, warnings: []}, overrides);

function abortableFetch(handler, delay = 2) {
    const calls = [], metrics = {active: 0, maxActive: 0, aborted: 0};
    const fetch = async (url, options) => {
        calls.push(url);
        metrics.active++;
        metrics.maxActive = Math.max(metrics.maxActive, metrics.active);
        try {
            await new Promise((resolve, reject) => {
                const timer = setTimeout(() => { options.signal.removeEventListener('abort', cancel); resolve(); }, delay);
                const cancel = () => { clearTimeout(timer); metrics.aborted++; reject(Object.assign(new Error('Aborted'), {name: 'AbortError'})); };
                if (options.signal.aborted) cancel();
                else options.signal.addEventListener('abort', cancel, {once: true});
            });
            return await handler(url, options);
        } finally { metrics.active--; }
    };
    return {fetch, calls, metrics};
}

async function crossSiteAndRetry() {
    let healed = false;
    const transitions = [];
    const fixture = abortableFetch(url => {
        if (url.endsWith('catalog-sites')) return response({sites});
        const params = new URL(url, 'https://app.example').searchParams;
        const site = Number(params.get('site_id')), parent = Number(params.get('parent_id')), number = Number(params.get('page'));
        assert.equal(params.get('search'), 'Blueberry Ice');
        if (site === 1 && parent === 0 && number === 1) return response(page(1, 0, 1, [10], [row(1, 10)], {total: 2, total_pages: 2, has_more: true, next_page: 2}));
        if (site === 1 && parent === 0 && number === 2) return response(page(1, 0, 2, [20], [], {total: 2, total_pages: 2, variable_products: [{id: 20, name: 'ELFBAR 600', variation_ids: [101, 102]}]}));
        if (site === 1 && parent === 20) return response(page(1, 20, number, [100 + number], [row(1, 20, 100 + number)], {total: 2, total_pages: 2, has_more: number === 1, next_page: number === 1 ? 2 : null}));
        if (site === 2) return response(page(2, 0, 1, [10], [row(2, 10, 0, {stock_status: 'outofstock', stock_quantity: 0})]));
        if (site === 3 && number === 1) return response(page(3, 0, 1, [10], [row(3, 10)], {total: 2, total_pages: 2, has_more: true, next_page: 2}));
        if (site === 3 && number === 2 && !healed) return response({error: '上游响应超时'}, 504);
        if (site === 3 && number === 2) return response(page(3, 0, 2, [11], [row(3, 11)], {total: 2, total_pages: 2}));
        if (site === 4) return response(page(4, 0, 1, [10], [row(4, 10)], healed ? {} : {total: null, total_pages: null, warnings: ['缺少分页头，按页长判断完整性。']}));
        throw new Error('Unexpected URL: ' + url);
    });
    const controller = createController({fetch: fixture.fetch, onChange: state => transitions.push({phase: state.phase, matched: state.counts.matched})});
    await controller.load('  Blueberry Ice  ');
    let state = controller.getSnapshot();
    assert.equal(state.phase, 'incomplete');
    assert.equal(state.counts.matched, 6);
    assert.equal(state.counts.completeSites, 2);
    assert.equal(state.counts.failedSites, 1);
    assert.equal(state.counts.incompleteSites, 2);
    assert.match(state.siteStates.find(site => site.site.id === '3').error, /HTTP 504.*超时/);
    assert.equal(new Set(state.rows.map(row => row.key)).size, 6, 'same product IDs on different sites must not collide');
    assert(transitions.some(update => update.phase === 'loading' && update.matched > 0), 'results stream before all sites finish');
    assert(fixture.metrics.maxActive <= 3);
    const requestsBeforeFilters = fixture.calls.length;
    controller.setFilters({manager: 'value:michael', brand: 'value:elfbar', product: 'value:elfbar 600', flavor: 'value:blueberry ice', publish: 'publish', stock: 'outofstock', text: 'Blueberry Ice'});
    state = controller.getSnapshot();
    assert.equal(state.filtered.length, 1);
    assert.equal(state.filtered[0].site_id, '2');
    controller.setFilters({site: '1'});
    assert.equal(controller.getSnapshot().filtered.length, 0, 'all selected filters are ANDed');
    controller.resetFilters();
    controller.setPage(2);
    assert.equal(fixture.calls.length, requestsBeforeFilters, 'local filters and pagination must not fetch');
    healed = true;
    const retryStart = fixture.calls.length;
    await controller.retryIncomplete();
    state = controller.getSnapshot();
    assert.equal(state.phase, 'complete');
    assert.equal(state.counts.matched, 7);
    assert.equal(state.counts.completeSites, 4);
    assert(fixture.calls.slice(retryStart).every(url => /site_id=[34]/.test(url)), 'retry only incomplete sites');
    assert(fixture.calls.slice(retryStart).some(url => /site_id=3.*page=1/.test(url)), 'retry starts incomplete site at its first page');
    assert.equal(fixture.metrics.active, 0);
}

async function completenessFailures() {
    const scenarios = {
        duplicate: number => page(1, 0, number, [10], [row(1, 10)], {total: 2, total_pages: 2, has_more: number === 1, next_page: number === 1 ? 2 : null}),
        changedTotal: number => page(1, 0, number, [9 + number], [row(1, 9 + number)], {total: number === 1 ? 2 : 3, total_pages: 2, has_more: number === 1, next_page: number === 1 ? 2 : null}),
        missingTotal: () => page(1, 0, 1, [10], [row(1, 10)], {total: 2}),
        skippedPage: () => page(1, 0, 1, [10], [row(1, 10)], {total: 3, total_pages: 3, has_more: true, next_page: 3}),
        missingCompleteness: () => page(1, 0, 1, [], [], {complete_page: false}),
        invalidIds: () => page(1, 0, 1, [0], []),
        wrongSite: () => page(2, 0, 1, [10], [row(2, 10)]),
    };
    for (const [scenario, makePage] of Object.entries(scenarios)) {
        const fixture = abortableFetch(url => url.endsWith('catalog-sites') ? response({sites: [sites[0]]}) : response(makePage(Number(new URL(url, 'https://app.example').searchParams.get('page')))));
        const controller = createController({fetch: fixture.fetch});
        await controller.load('blue');
        const state = controller.getSnapshot();
        assert.equal(state.phase, 'incomplete', scenario);
        assert.equal(state.counts.failedSites, 1, scenario);
        assert.match(state.siteStates[0].error, /不完整/, scenario);
    }
    const fixture = abortableFetch(url => {
        if (url.endsWith('catalog-sites')) return response({sites: [sites[0]]});
        const parent = Number(new URL(url, 'https://app.example').searchParams.get('parent_id'));
        return response(parent ? page(1, 10, 1, [101, 103], [row(1, 10, 101), row(1, 10, 103)]) : page(1, 0, 1, [10], [], {variable_products: [{id: 10, name: 'Parent', variation_ids: [101, 102]}]}));
    });
    const controller = createController({fetch: fixture.fetch});
    await controller.load('blue');
    assert.equal(controller.getSnapshot().phase, 'incomplete');
    assert.match(controller.getSnapshot().siteStates[0].error, /变体目录/);
}

async function childStatusAndWildcardSemantics() {
    const cases = [
        {name: 'draft-pending-future-extra', ids: [101, 102, 103, 104], statuses: {'101': 'publish', '102': 'draft', '103': 'pending', '104': 'future'}, expected: [101], complete: true},
        {name: 'private-listed', ids: [101, 102], statuses: {'101': 'publish', '102': 'private'}, expected: [101, 102], complete: true},
        {name: 'published-missing', ids: [101, 102], statuses: {'101': 'publish', '102': 'draft'}, expected: [101, 103], complete: false},
        {name: 'unexpected-published', ids: [101, 102], statuses: {'101': 'publish', '102': 'publish'}, expected: [101], complete: false},
        {name: 'missing-status-map', ids: [101], statuses: undefined, expected: [101], complete: false},
        {name: 'incomplete-status-map', ids: [101, 102], statuses: {'101': 'publish'}, expected: [101], complete: false},
        {name: 'unknown-status', ids: [101, 102], statuses: {'101': 'publish', '102': 'mystery'}, expected: [101], complete: false},
    ];
    for (const scenario of cases) {
        const fixture = abortableFetch(url => {
            if (url.endsWith('catalog-sites')) return response({sites: [sites[0]]});
            const parent = Number(new URL(url, 'https://app.example').searchParams.get('parent_id'));
            if (!parent) return response(page(1, 0, 1, [10], [], {variable_products: [{id: 10, name: 'Parent', variation_ids: scenario.expected}]}));
            return response(page(1, 10, 1, scenario.ids, scenario.ids.map(id => row(1, 10, id, {status: scenario.statuses && scenario.statuses[id] || 'publish'})), {source_statuses: scenario.statuses}));
        });
        const controller = createController({fetch: fixture.fetch});
        await controller.load('blue');
        const state = controller.getSnapshot();
        assert.equal(state.phase, scenario.complete ? 'complete' : 'incomplete', scenario.name);
        if (scenario.complete) assert.equal(state.counts.matched, scenario.ids.length, 'draft and scheduled rows remain visible');
    }
    const wildcardController = createController({fetch: async url => {
        if (url.endsWith('catalog-sites')) return response({sites: [sites[0]]});
        const parent = Number(new URL(url, 'https://app.example').searchParams.get('parent_id'));
        return response(parent ? page(1, 10, 1, [101], [row(1, 10, 101, {flavors: ['Blueberry Ice', 'Apple'], flavor_scope: 'any'})]) : page(1, 0, 1, [10], [], {variable_products: [{id: 10, name: 'Parent', variation_ids: [101]}]}));
    }});
    await wildcardController.load('blue');
    assert.equal(wildcardController.getSnapshot().phase, 'complete');
    assert.equal(wildcardController.getSnapshot().rows.length, 1, 'wildcard is one shared variation, not synthetic duplicates');
    wildcardController.setFilters({flavor: 'value:apple'});
    assert.equal(wildcardController.getSnapshot().filtered.length, 1);
    assert.match(resultHtml(wildcardController.getSnapshot()), /任意口味（共用此变体）/);
}

async function brandQueriesAndModes() {
    let healed = false;
    const fixture = abortableFetch(url => {
        if (url.endsWith('catalog-sites')) return response({sites});
        const params = new URL(url, 'https://app.example').searchParams;
        assert.equal(params.get('search'), '', 'brand queries do not require or reuse a flavor keyword');
        assert.equal(params.get('brand'), 'ELFBAR');
        assert.equal(params.get('query_mode'), 'brand');
        const site = Number(params.get('site_id')), parent = Number(params.get('parent_id'));
        if (site === 3 && !healed) return response({error: '站点暂不可用'}, 200);
        if (site === 1 && !parent) return response(page(1, 0, 1, [10], [], {variable_products: [{id: 10, name: 'ELFBAR 600', variation_ids: [101, 102]}]}));
        if (site === 1 && parent) return response(page(1, 10, 1, [101, 102], [row(1, 10, 101), row(1, 10, 102, {flavors: ['Apple'], name: 'Apple'})]));
        return response(page(site, 0, 1, [10], [row(site, 10, 0, {flavors: [site === 2 ? 'Grape' : 'Mango']})]));
    });
    const controller = createController({fetch: fixture.fetch});
    await controller.load('stale flavor text must be ignored', {queryMode: 'brand', brand: '  ELFBAR  '});
    let state = controller.getSnapshot();
    assert.equal(state.phase, 'incomplete');
    assert.equal(state.mode, 'brand');
    assert.equal(state.queryMode, 'brand');
    assert.equal(state.keyword, '');
    assert.equal(state.brand, 'ELFBAR');
    assert.deepEqual([...new Set(state.rows.flatMap(row => row.flavors))].sort(), ['Apple', 'Blueberry Ice', 'Grape', 'Mango']);
    const beforeFilter = fixture.calls.length;
    controller.setFilters({manager: 'value:michael'});
    assert.equal(controller.getSnapshot().filtered.length, 3);
    assert.equal(fixture.calls.length, beforeFilter, 'owner filtering after brand load is local');
    healed = true;
    const retryStart = fixture.calls.length;
    await controller.retryIncomplete();
    assert(fixture.calls.slice(retryStart).every(url => /site_id=3/.test(url)), 'brand retry only reads incomplete sites');
    assert.equal(controller.getSnapshot().phase, 'complete');
    await controller.refresh();
    state = controller.getSnapshot();
    assert.equal(state.mode, 'brand');
    assert.equal(state.brand, 'ELFBAR');
    assert.equal(state.keyword, '');
    assert.equal(state.filters.manager, 'value:michael', 'refresh preserves loaded result filters');
    assert.equal(state.filtered.length, 3);

    const modeCalls = [];
    const modes = createController({fetch: async url => {
        modeCalls.push(url);
        if (url.endsWith('catalog-sites')) return response({sites: [sites[0]]});
        return response(page(1, 0, 1, [10], [row(1, 10)]));
    }});
    await modes.load('unused', {queryMode: 'brand', brand: ''});
    assert.equal(modes.getSnapshot().phase, 'error');
    assert.match(modes.getSnapshot().error, /输入或选择一个品牌/);
    assert.equal(modeCalls.length, 0, 'empty brand never scans all sites');
    await modes.loadQuery({mode: 'brand', brand: 'R&M'});
    let params = new URL(modeCalls.at(-1), 'https://app.example').searchParams;
    assert.equal(params.get('brand'), 'R&M', 'free input and aliases are passed to backend without a forced dropdown selection');
    assert.equal(params.get('search'), '');
    assert.equal(params.get('query_mode'), 'brand');
    await modes.loadQuery({mode: 'all', keyword: 'ignored', brand: 'ignored'});
    params = new URL(modeCalls.at(-1), 'https://app.example').searchParams;
    assert.equal(params.get('brand'), '');
    assert.equal(params.get('search'), '');
    assert.equal(params.get('query_mode'), 'all');
    assert.equal(modes.getSnapshot().mode, 'all');
    await modes.load('Blueberry Ice'); // Existing public interface still means flavor search.
    params = new URL(modeCalls.at(-1), 'https://app.example').searchParams;
    assert.equal(params.get('brand'), '');
    assert.equal(params.get('search'), 'Blueberry Ice');
    assert.equal(params.get('query_mode'), 'flavor');
    assert.equal(modes.getSnapshot().mode, 'flavor');

    const stoppedFixture = abortableFetch(url => {
        if (url.endsWith('catalog-sites')) return response({sites});
        const params = new URL(url, 'https://app.example').searchParams;
        assert.equal(params.get('brand'), 'ELFBAR');
        assert.equal(params.get('search'), '');
        return response(page(Number(params.get('site_id')), 0, 1, [], []));
    }, 15);
    const stoppedBrand = createController({fetch: stoppedFixture.fetch});
    const running = stoppedBrand.load('', {queryMode: 'brand', brand: 'ELFBAR'});
    await sleep(20);
    stoppedBrand.stop();
    await running;
    assert.equal(stoppedBrand.getSnapshot().mode, 'brand');
    assert.equal(stoppedBrand.getSnapshot().brand, 'ELFBAR');
    await stoppedBrand.retryIncomplete();
    assert.equal(stoppedBrand.getSnapshot().phase, 'complete');

    for (const newMode of ['flavor', 'all']) {
        let releaseOld, oldStarted;
        const ready = new Promise(resolve => { oldStarted = resolve; });
        const gate = new Promise(resolve => { releaseOld = resolve; });
        const races = createController({fetch: async url => {
            if (url.endsWith('catalog-sites')) return response({sites: [sites[0]]});
            const params = new URL(url, 'https://app.example').searchParams;
            const old = params.get('brand') === 'ELFBAR';
            if (old) { oldStarted(); await gate; }
            return response(page(1, 0, 1, [10], [row(1, 10, 0, {product_name: old ? 'old brand' : 'new ' + newMode})]));
        }});
        const oldRun = races.load('', {queryMode: 'brand', brand: 'ELFBAR'});
        await ready;
        await races.load('Apple', {queryMode: newMode});
        releaseOld();
        await oldRun;
        assert.equal(races.getSnapshot().mode, newMode);
        assert.equal(races.getSnapshot().brand, '');
        assert.equal(races.getSnapshot().rows[0].product_name, 'new ' + newMode, 'late old brand pages cannot overwrite a switched mode');
    }
}

async function mountedQueryControls() {
    const ids = ['pmCatalogPane', 'pmCatalogSearch', 'pmCatalogLoad', 'pmCatalogStop', 'pmCatalogRetry', 'pmCatalogRefresh', 'pmCatalogStatus', 'pmCatalogCounts',
        'pmCatalogProgress', 'pmCatalogResults', 'pmCatalogManagerFilter', 'pmCatalogSiteFilter', 'pmCatalogBrandFilter', 'pmCatalogProductFilter',
        'pmCatalogFlavorFilter', 'pmCatalogStockFilter', 'pmCatalogPublishFilter', 'pmCatalogTextFilter', 'pmCatalogResetFilters',
        'pmCatalogQueryMode', 'pmCatalogBrandInput', 'pmCatalogFlavorInputWrap', 'pmCatalogBrandInputWrap'];
    const elements = Object.fromEntries(ids.map(id => {
        const classes = new Set(), handlers = new Map();
        return [id, {value: '', innerHTML: '', textContent: '', options: [], disabled: false,
            classList: {toggle(name, enabled) { if (enabled) classes.add(name); else classes.delete(name); }, contains: name => classes.has(name)},
            addEventListener(name, handler) { handlers.set(name, handler); },
            trigger(name, event = {}) { return handlers.get(name)(event); }}];
    }));
    elements.pmCatalogQueryMode.value = 'brand';
    elements.pmCatalogBrandInput.value = 'ELFBAR';
    elements.pmCatalogSearch.value = 'stale flavor';
    const calls = [];
    const controller = mount({getElementById: id => elements[id]}, {fetch: async url => {
        calls.push(url);
        if (url.endsWith('catalog-sites')) return response({sites: [sites[0]]});
        return response(page(1, 0, 1, [10], [row(1, 10)]));
    }});
    assert.equal(elements.pmCatalogQueryMode.value, 'brand', 'mount preserves brand default from template');
    assert(elements.pmCatalogFlavorInputWrap.classList.contains('d-none'));
    assert(!elements.pmCatalogBrandInputWrap.classList.contains('d-none'));
    let prevented = false;
    elements.pmCatalogBrandInput.trigger('keydown', {key: 'Enter', preventDefault() { prevented = true; }});
    await sleep(2);
    assert(prevented);
    assert.equal(controller.getSnapshot().mode, 'brand');
    assert.equal(controller.getSnapshot().brand, 'ELFBAR');
    assert.equal(new URL(calls.at(-1), 'https://app.example').searchParams.get('search'), '');
    elements.pmCatalogQueryMode.value = 'all';
    elements.pmCatalogQueryMode.trigger('change');
    assert(elements.pmCatalogFlavorInputWrap.classList.contains('d-none'));
    assert(elements.pmCatalogBrandInputWrap.classList.contains('d-none'));
    await elements.pmCatalogLoad.trigger('click');
    assert.equal(controller.getSnapshot().mode, 'all');
    elements.pmCatalogQueryMode.value = 'flavor';
    elements.pmCatalogSearch.value = 'Apple';
    elements.pmCatalogQueryMode.trigger('change');
    assert(!elements.pmCatalogFlavorInputWrap.classList.contains('d-none'));
    assert(elements.pmCatalogBrandInputWrap.classList.contains('d-none'));
    await elements.pmCatalogLoad.trigger('click');
    elements.pmCatalogQueryMode.value = 'brand';
    elements.pmCatalogBrandInput.value = 'unloaded brand';
    elements.pmCatalogRefresh.trigger('click');
    await sleep(2);
    assert.equal(elements.pmCatalogQueryMode.value, 'flavor', 'refresh restores the loaded mode instead of using unsent form edits');
    assert.equal(controller.getSnapshot().keyword, 'Apple');
    elements.pmCatalogQueryMode.value = 'brand';
    elements.pmCatalogBrandInput.value = '';
    const beforeEmptyBrand = calls.length;
    await elements.pmCatalogLoad.trigger('click');
    await sleep(2);
    assert.equal(calls.length, beforeEmptyBrand);
    assert.match(elements.pmCatalogStatus.textContent, /请输入或选择一个品牌/);
}

async function cancelAndRace() {
    const fixture = abortableFetch(url => url.endsWith('catalog-sites') ? response({sites}) : response(page(Number(new URL(url, 'https://app.example').searchParams.get('site_id')), 0, 1, [], [])), 15);
    const controller = createController({fetch: fixture.fetch});
    const pending = controller.load('stop');
    await sleep(20);
    controller.stop();
    await pending;
    const stopped = controller.getSnapshot();
    assert.equal(stopped.phase, 'stopped');
    assert.equal(stopped.counts.incompleteSites, 4);
    assert(fixture.metrics.aborted > 0);
    assert(fixture.metrics.maxActive <= 3);
    assert.equal(fixture.calls.length, 4, 'stop prevents queued fourth site from starting');
    await controller.retryIncomplete();
    assert.equal(controller.getSnapshot().phase, 'complete');

    let stopOnce = true, streamingController;
    const streamingFixture = abortableFetch(url => {
        if (url.endsWith('catalog-sites')) return response({sites});
        const id = Number(new URL(url, 'https://app.example').searchParams.get('site_id'));
        return response(page(id, 0, 1, [10], [row(id, 10)]));
    });
    streamingController = createController({fetch: streamingFixture.fetch, onChange: state => {
        if (stopOnce && state.phase === 'loading' && state.counts.matched) { stopOnce = false; streamingController.stop(); }
    }});
    await streamingController.load('blue');
    assert.equal(streamingController.getSnapshot().phase, 'stopped');
    assert.equal(streamingController.getSnapshot().rows.length, 1, 'stop retains pages already received');
    await streamingController.retryIncomplete();
    assert.equal(streamingController.getSnapshot().phase, 'complete');
    assert.equal(streamingController.getSnapshot().rows.length, 4);

    let releaseOld;
    const oldGate = new Promise(resolve => { releaseOld = resolve; });
    let discovery = 0;
    const staleController = createController({fetch: async url => {
        if (url.endsWith('catalog-sites')) {
            discovery++;
            if (discovery === 1) await oldGate; // Intentionally ignore AbortSignal to model a late response.
            return response({sites: [sites[0]]});
        }
        const query = new URL(url, 'https://app.example').searchParams.get('search');
        return response(page(1, 0, 1, [10], [row(1, 10, 0, {name: query, product_name: query})]));
    }});
    const oldRun = staleController.load('old');
    await staleController.load('new');
    releaseOld();
    await oldRun;
    assert.equal(staleController.getSnapshot().keyword, 'new');
    assert.equal(staleController.getSnapshot().rows[0].product_name, 'new', 'late old responses cannot replace a new query');

    let releasePage, oldPageStarted;
    const oldPageReady = new Promise(resolve => { oldPageStarted = resolve; });
    const oldPageGate = new Promise(resolve => { releasePage = resolve; });
    const latePageController = createController({fetch: async url => {
        if (url.endsWith('catalog-sites')) return response({sites: [sites[0]]});
        const query = new URL(url, 'https://app.example').searchParams.get('search');
        if (query === 'old') { oldPageStarted(); await oldPageGate; }
        return response(page(1, 0, 1, [10], [row(1, 10, 0, {name: query, product_name: query})]));
    }});
    const lateOldPage = latePageController.load('old');
    await oldPageReady;
    await latePageController.load('new');
    releasePage();
    await lateOldPage;
    assert.equal(latePageController.getSnapshot().rows[0].product_name, 'new', 'late page responses cannot overwrite new rows');
}

async function errorsAndRendering() {
    const slow = abortableFetch(() => response({sites: []}), 25);
    const slowController = createController({fetch: slow.fetch, requestTimeout: 5});
    await slowController.load('query');
    assert.equal(slowController.getSnapshot().phase, 'error');
    assert.match(slowController.getSnapshot().error, /超时/);
    for (const [type, data, pattern] of [['text/html', {}, /非 JSON.*HTTP 502/], ['application/json', {error: '读取被拒绝', complete_page: false}, /读取被拒绝/]]) {
        const controller = createController({fetch: async url => url.endsWith('catalog-sites') ? response({sites: [sites[0]]}) : response(data, type === 'text/html' ? 502 : 200, type)});
        await controller.load('no');
        const state = controller.getSnapshot();
        assert.equal(state.phase, 'incomplete');
        assert.match(state.siteStates[0].error, pattern);
        assert.match(resultHtml(state), /不能据此判断全部站点没有/);
    }
    const malicious = row(1, 10, 0, {product_name: '<img src=x onerror=alert(1)>', name: 'Jabłko Blueberry-Ice Mr. Blue', brands: ['<script>'], flavors: ['" onclick="alert(1)'], permalink: 'javascript:alert(1)'});
    const controller = createController({fetch: async url => url.endsWith('catalog-sites') ? response({sites: [Object.assign({}, sites[0], {manager: '<svg onload=x>'})]}) : response(page(1, 0, 1, [10], [malicious]))});
    await controller.load('query');
    const html = resultHtml(controller.getSnapshot());
    assert(!html.includes('<img'));
    assert(!html.includes('<script>'));
    assert(!html.includes('href="javascript:'));
    assert(html.includes('&lt;img'));
    assert(html.includes('&lt;svg'));
    assert(html.includes('价格（站点原币）'));
    assert.equal(safeUrl('javascript:alert(1)'), '');
    assert.equal(safeUrl('/relative/path'), '');
    assert.equal(safeUrl('https://example.com/path'), 'https://example.com/path');
    controller.setFilters({text: 'jablko blueberry ice'});
    assert.equal(controller.getSnapshot().filtered.length, 1, 'text filtering handles diacritics and separators');
    controller.setFilters({text: 'Mr Blue'});
    assert.equal(controller.getSnapshot().filtered.length, 1, 'text filtering handles punctuation like the main search');

    const largeController = createController({fetch: async url => {
        if (url.endsWith('catalog-sites')) return response({sites: [sites[0]]});
        return response(page(1, 0, 1, Array.from({length: 101}, (_, index) => index + 1), Array.from({length: 101}, (_, index) => row(1, index + 1))));
    }});
    await largeController.load('query');
    assert.equal(largeController.getSnapshot().pageRows.length, 100);
    largeController.setPage(2);
    assert.equal(largeController.getSnapshot().pageRows.length, 1);
}

async function boundedParentsTokensAndCaching() {
    const progress = [], activeParents = new Map(), peakParents = new Map();
    const fixture = abortableFetch(async url => {
        if (url.endsWith('catalog-sites')) return response({sites: sites.slice(0, 3)});
        const params = new URL(url, 'https://app.example').searchParams, site = Number(params.get('site_id')), parent = Number(params.get('parent_id'));
        if (!parent) {
            const descriptors = [10, 20, 30].map(id => ({id, name: 'Merrymi ' + id, variation_ids: [id * 10], parent_token: `synthetic-token-${site}-${id}`}));
            return response(page(site, 0, 1, [10, 20, 30], [], {variable_products: descriptors}));
        }
        assert.equal(params.get('parent_token'), `synthetic-token-${site}-${parent}`, 'variation scan must reuse the matching signed parent snapshot');
        const current = (activeParents.get(site) || 0) + 1; activeParents.set(site, current); peakParents.set(site, Math.max(peakParents.get(site) || 0, current));
        try { await sleep(8); return response(page(site, parent, 1, [parent * 10], [row(site, parent, parent * 10)])); }
        finally { activeParents.set(site, activeParents.get(site) - 1); }
    }, 2);
    const controller = createController({fetch: fixture.fetch, onChange: state => progress.push(state.phase)});
    await controller.load('Blue Ice');
    assert.equal(controller.getSnapshot().phase, 'complete'); assert.equal(progress.at(-1), 'complete', 'terminal callback is synchronous with completion');
    assert.equal(controller.getSnapshot().rows.length, 9); assert(fixture.metrics.maxActive <= 6); assert(fixture.metrics.maxActive >= 4, 'multiple parents really run concurrently');
    assert([...peakParents.values()].every(count => count === 2), 'each site caps variable parent workers at2');
    const first = controller.getSnapshot(), count = fixture.calls.length;
    controller.setPage(1); const pageSnapshot = controller.getSnapshot();
    assert.strictEqual(first.rows, pageSnapshot.rows); assert.strictEqual(first.options, pageSnapshot.options, 'page changes reuse facets');
    controller.setFilters({manager: 'value:michael'}); const filtered = controller.getSnapshot();
    assert.strictEqual(first.rows, filtered.rows); assert.strictEqual(first.options, filtered.options, 'local filters reuse rows and facets'); assert.equal(fixture.calls.length, count);
    controller.updateRows([Object.assign({}, first.rows[0], {regular_price: '31', stock_status: 'outofstock', stock_state: 'outofstock'})]);
    const edited = controller.getSnapshot(); assert.notStrictEqual(first.rows, edited.rows); assert.notStrictEqual(first.options, edited.options);
    assert.equal(edited.counts.matched, 9); assert.equal(edited.siteStates.find(state => state.site.id === first.rows[0].site_id).matched, 3, 'canonical updates do not double-count matches');
}
async function progressThrottling() {
    let notices = 0;
    const controller = createController({fetch: async url => {
        if (url.endsWith('catalog-sites')) return response({sites: [sites[0]]});
        const number = Number(new URL(url, 'https://app.example').searchParams.get('page'));
        return response(page(1, 0, number, [number], [row(1, number)], {total: 40, total_pages: 40, has_more: number < 40, next_page: number < 40 ? number + 1 : null}));
    }, onChange: () => { notices += 1; }});
    await controller.load('Blue Ice'); assert.equal(controller.getSnapshot().rows.length, 40); assert.equal(controller.getSnapshot().phase, 'complete');
    assert(notices <= 5, 'rapid pages coalesce progress instead of repeated full DOM renders');
    const terminalNotices = notices; await sleep(120); assert.equal(notices, terminalNotices, 'terminal emission clears the pending progress timer');
}

(async () => {
    await crossSiteAndRetry();
    await completenessFailures();
    await childStatusAndWildcardSemantics();
    await brandQueriesAndModes();
    await mountedQueryControls();
    await cancelAndRace();
    await errorsAndRendering();
    await boundedParentsTokensAndCaching();
    await progressThrottling();
    console.log('Cross-site catalog: streaming, pagination, variation integrity, local filters, retry, cancellation, race and rendering checks passed.');
})().catch(error => { console.error(error); process.exitCode = 1; });
