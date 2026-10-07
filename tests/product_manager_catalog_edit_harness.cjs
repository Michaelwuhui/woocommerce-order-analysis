// Exercise the real editor against isolated synthetic exact-site GET/PUT data.
const assert = require('node:assert/strict');
const {createEditor, validChanges, rowKey, sameField} = require('../static/js/product_manager_catalog_editor.js');
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
const copy = value => JSON.parse(JSON.stringify(value));
const response = (data, status = 200, type = 'application/json') => ({status, ok: status < 400, headers: {get: () => type}, json: async () => copy(data)});
function row(site = 1, product = 10, variation = 0, overrides = {}) {
    const value = Object.assign({site_id: String(site), product_id: product, variation_id: variation, key: `${site}:${product}:${variation}`,
        site: {id: String(site), url: `https://site${site}.example`, manager: site === 1 ? 'Michael' : 'Anna'}, sku: `sku-${product}-${variation}`,
        product_name: 'Merrymi 9000', name: 'Blue Ice', flavors: ['Blue Ice'], brands: ['Merrymi'], type: variation ? 'variation' : 'simple', status: 'publish',
        manage_stock: true, stock_quantity: 5, stock_status: 'instock', regular_price: '10.00', sale_price: '', price: '10.00'}, overrides);
    value.edit_identity = {sku: value.sku, type: value.type, parent_id: variation ? product : 0, attributes: variation ? [{id: 7, name: 'Flavor', options: ['Blue Ice']}] : []};
    value.edit_before = Object.fromEntries(['manage_stock', 'stock_quantity', 'stock_status', 'regular_price', 'sale_price'].map(field => [field, value[field]]));
    return value;
}
function fixture(inputRows, settings = {}) {
    let loaded = inputRows.map(copy), filteredKeys = null, page = 1, loading = false;
    const actual = new Map(inputRows.map(value => [rowKey(value), copy(value)])), calls = [], writes = [], jobs = [];
    const metrics = {active: 0, maxActive: 0, puts: 0, maxPuts: 0};
    const config = {csrf_token: 'synthetic-csrf', sites: [1, 2, 3, 4].map(id => ({id, url: `https://site${id}.example`, can_edit: true, can_clone: true})), clone_targets: [{id: 4, url: 'https://site4.example'}]};
    const catalog = {
        getSnapshot() {
            const filtered = filteredKeys ? loaded.filter(value => filteredKeys.has(rowKey(value))) : loaded;
            return {rows: loaded, filtered, pageRows: filtered.slice((page - 1) * 100, page * 100), page, pages: Math.max(1, Math.ceil(filtered.length / 100)), busy: loading};
        },
        updateRows(values) { const incoming = new Map(values.map(value => [rowKey(value), copy(value)])); loaded = loaded.map(value => incoming.get(rowKey(value)) || value); },
        setFiltered(keys) { filteredKeys = keys ? new Set(keys) : null; }, setPage(value) { page = value; }, setBusy(value) { loading = value; }
    };
    const fetch = async (url, options = {}) => {
        const method = options.method || 'GET'; calls.push({url, method});
        if (url.endsWith('catalog-edit-config')) return response(config);
        if (url.includes('clone-jobs/')) return response({job_id: url.split('/').at(-1), status: 'completed', terminal: true, completed_count: 1, total_count: 1, created_count: 1});
        if (url.endsWith('catalog-clone')) {
            assert.equal(options.headers['X-PM-CSRF'], 'synthetic-csrf');
            const payload = JSON.parse(options.body); jobs.push(payload);
            if (settings.cloneFailAt === jobs.length) return response({error: 'synthetic queued response lost', code: 'UNKNOWN_QUEUE'}, 503);
            return response({job_id: `synthetic-job-${jobs.length}`, status: 'queued', total_count: payload.product_ids.length}, 202);
        }
        assert(url.includes('/catalog-item?'), 'editor must never use legacy Master update endpoint');
        const params = new URL(url, 'https://app.example').searchParams, key = `${params.get('site_id')}:${params.get('product_id')}:${params.get('variation_id')}`;
        const actualRow = actual.get(key); assert(actualRow, 'exact child row must exist');
        metrics.active += 1; metrics.maxActive = Math.max(metrics.maxActive, metrics.active);
        if (method === 'PUT') { metrics.puts += 1; metrics.maxPuts = Math.max(metrics.maxPuts, metrics.puts); }
        try {
            if (settings.delay) await sleep(settings.delay);
            if (method === 'GET') {
                settings.onGet?.(key);
                if (settings.getFailure?.(key)) return response({error: 'synthetic upstream unavailable', code: 'READ_FAILED'}, 503);
                return response({success: true, identity: actualRow.edit_identity, before: actualRow.edit_before, item: actualRow, csrf_token: 'synthetic-csrf', routing: {mode: 'direct'}, stock_bridge: actualRow.stock_bridge || {present: false, valid: true, values: {}}});
            }
            assert.equal(method, 'PUT'); assert.equal(options.headers['X-PM-CSRF'], 'synthetic-csrf');
            const payload = JSON.parse(options.body); writes.push({key, payload});
            settings.onPut?.(key);
            assert.equal(`${payload.site_id}:${payload.product_id}:${payload.variation_id}`, key);
            assert.deepEqual(payload.expected_identity, actualRow.edit_identity);
            for (const field of Object.keys(payload.changes)) assert(sameField(field, payload.expected_before[field], actualRow.edit_before[field]), 'fresh changed-field proof must precede write');
            if (settings.reject?.(key)) return response({error: 'synthetic stale pre-write check', code: 'STALE_ITEM', write_started: false}, 409);
            Object.assign(actualRow, payload.changes); Object.assign(actualRow.edit_before, payload.changes);
            settings.afterApply?.(actualRow, payload);
            if (settings.uncertain?.(key)) return response({error: 'synthetic final read failed', code: 'WRITE_NOT_VERIFIED', write_started: true, verification_status: 'unconfirmed'}, 409);
            return response({success: true, item: actualRow, verification: {status: 'verified', direct_site: true}});
        } finally { metrics.active -= 1; if (method === 'PUT') metrics.puts -= 1; }
    };
    const editor = createEditor({catalog, fetch, uuid: () => 'synthetic-batch', autoPoll: false, confirm: () => true, storage: settings.storage});
    return {editor, catalog, actual, calls, writes, jobs, metrics};
}
async function selectionAndDirtyState() {
    const rows = Array.from({length: 105}, (_, index) => row(1, 100 + index)); rows.push(row(2, 100));
    const f = fixture(rows); await f.editor.ready();
    f.editor.select(f.catalog.getSnapshot().pageRows.map(rowKey)); f.catalog.setPage(2); f.editor.select([rowKey(rows[104]), rowKey(rows[105])]);
    f.editor.setDraft(rowKey(rows[0]), 'regular_price', '14.5');
    f.catalog.setFiltered([rowKey(rows[105])]);
    assert.equal(f.editor.getSnapshot().selectedCount, 102); assert.equal(f.editor.getSnapshot().selectedVisible, 1); assert.equal(f.editor.getSnapshot().selectedHidden, 101);
    assert.equal(f.editor.getDraft(rowKey(rows[0])).regular_price, '14.5', 'draft survives page and owner filter change');
    const count = f.calls.length; f.catalog.setFiltered(null); f.catalog.setPage(1); assert.equal(f.calls.length, count, 'filters do not request catalog or editable item data');
    await f.editor.saveDrafts([rowKey(rows[0])]);
    assert.equal(f.writes.length, 1); assert.equal(f.writes[0].key, rowKey(rows[0])); assert.equal(f.writes[0].payload.changes.regular_price, '14.5');
    assert.equal(f.actual.get(rowKey(rows[105])).regular_price, '10.00', 'same product ID on another site remains untouched');
    assert.equal(f.editor.getSnapshot().draftCount, 0); assert.equal(f.editor.getSnapshot().selectedCount, 101);
}
async function stalePreflightAndFailedRetry() {
    const rows = [row(1, 10, 101), row(2, 10, 101)], f = fixture(rows); await f.editor.ready();
    f.editor.setDraft(rowKey(rows[0]), 'regular_price', '20'); f.actual.get(rowKey(rows[0])).edit_before.regular_price = '12';
    const outcome = await f.editor.saveDrafts(); assert.equal(outcome[0].status, 'failed'); assert.equal(f.writes.length, 0);
    assert.equal(f.editor.getSnapshot().draftCount, 1, 'stale failure retains intended draft');
    f.actual.get(rowKey(rows[0])).edit_before.regular_price = '10.00';
    await f.editor.retryFailed(); assert.equal(f.writes.length, 1); assert.equal(f.editor.getSnapshot().draftCount, 0);
    const identity = fixture([row(1, 10, 101)]); await identity.editor.ready(); identity.editor.select(['1:10:101']);
    identity.actual.get('1:10:101').edit_identity.attributes[0].options = ['Different flavor'];
    const identityOutcome = await identity.editor.bulk('price', {sale_price: '3'}); assert.equal(identityOutcome[0].status, 'failed'); assert.equal(identity.writes.length, 0);
    let rejected = true;
    const failure = fixture(rows, {reject: key => rejected && key === '2:10:101'}); await failure.editor.ready(); failure.editor.select(rows.map(rowKey));
    await failure.editor.bulk('price', {regular_price: '15'});
    assert.equal(failure.editor.getSnapshot().selectedCount, 1); assert.deepEqual(failure.editor.getSnapshot().selected, ['2:10:101']);
    rejected = false; await failure.editor.retryFailed();
    assert.equal(failure.writes.filter(write => write.key === '1:10:101').length, 1, 'failed-only retry must never repeat a verified row');
    assert.equal(failure.writes.filter(write => write.key === '2:10:101').length, 2);
}
async function inheritedStockAndPayloads() {
    const value = row(1, 10, 101, {manage_stock: 'parent', stock_quantity: 50}), f = fixture([value]); await f.editor.ready();
    f.editor.setDraft(rowKey(value), 'manage_stock', true);
    assert.throws(() => f.editor.saveDrafts(), /填写/); assert.equal(f.writes.length, 0, 'parent total50 must never be copied to leaf from a toggle');
    f.editor.setDraft(rowKey(value), 'stock_quantity', '3'); await f.editor.saveDrafts();
    assert.deepEqual(f.writes[0].payload.changes, {manage_stock: true, stock_quantity: 3});
    const explicitParentNumber = fixture([value]); await explicitParentNumber.editor.ready();
    explicitParentNumber.editor.setDraft(rowKey(value), 'manage_stock', true); explicitParentNumber.editor.setDraft(rowKey(value), 'stock_quantity', '50');
    await explicitParentNumber.editor.saveDrafts(); assert.deepEqual(explicitParentNumber.writes[0].payload.changes, {manage_stock: true, stock_quantity: 50}, 'explicit leaf quantity equal to former parent total is still an explicit choice');
    const inheritedAgain = fixture([value]); await inheritedAgain.editor.ready(); inheritedAgain.editor.select([rowKey(value)]);
    assert.throws(() => inheritedAgain.editor.bulk('soldout'), /独立库存/); assert.equal(inheritedAgain.writes.length, 0);
    assert.throws(() => inheritedAgain.editor.bulk('stock', {stock_quantity: '4'}), /确认/);
    await inheritedAgain.editor.bulk('stock', {stock_quantity: '4', confirm_independent: true}); assert.equal(inheritedAgain.writes[0].payload.changes.stock_quantity, 4);
    const normal = fixture([row()]); await normal.editor.ready(); normal.editor.select(['1:10:0']); await normal.editor.bulk('soldout');
    assert.deepEqual(normal.writes[0].payload.changes, {manage_stock: false, stock_status: 'outofstock'});
    normal.editor.select(['1:10:0']); await normal.editor.bulk('restore', {stock_quantity: '8'}); assert.deepEqual(normal.writes[1].payload.changes, {manage_stock: true, stock_quantity: 8});
    normal.editor.select(['1:10:0']); await normal.editor.bulk('price', {regular_price: '', sale_price: '0'}); assert.deepEqual(normal.writes[2].payload.changes, {sale_price: ''});
    assert.throws(() => validChanges({stock_quantity: 2147483648}, value.edit_before), /2147483647/);
    assert.throws(() => validChanges({regular_price: '1e9'}, value.edit_before), /价格/);
    assert.throws(() => validChanges({sale_price: '1.123456789'}, value.edit_before), /价格/);
}
async function uncertainWriteCannotReplay() {
    const value = row(), f = fixture([value], {uncertain: () => true}); await f.editor.ready(); f.editor.select([rowKey(value)]);
    f.editor.setDraft(rowKey(value), 'regular_price', '18');
    const result = await f.editor.saveDrafts(); assert.equal(result[0].status, 'unconfirmed', '409 write_started:true is not a failed pre-write operation');
    assert.throws(() => f.editor.saveDrafts(), /核对/); assert.throws(() => f.editor.bulk('price', {regular_price: '18'}), /核对/);
    assert.throws(() => f.editor.setDraft(rowKey(value), 'regular_price', '19'), /核对/);
    await assert.rejects(() => f.editor.retryFailed(), /先选择/); assert.equal(f.writes.length, 1);
    await f.editor.verify(rowKey(value)); assert.equal(f.writes.length, 1, 'verification is GET-only'); assert.equal(f.editor.getSnapshot().draftCount, 0);
}
async function concurrencyAndPermissions() {
    const rows = Array.from({length: 6}, (_, index) => row(index % 2 + 1, 10 + index)), prices = fixture(rows, {delay: 4}); await prices.editor.ready(); prices.editor.select(rows.map(rowKey));
    await prices.editor.bulk('price', {regular_price: '22'}); assert.equal(prices.metrics.maxActive, 2); assert(prices.metrics.maxPuts <= 2);
    const stock = fixture(rows, {delay: 4}); await stock.editor.ready(); stock.editor.select(rows.map(rowKey)); await stock.editor.bulk('stock', {stock_quantity: '7'});
    assert.equal(stock.metrics.maxActive, 1, 'stock GET+PUT sequences remain serial');
    const forbidden = fixture([row(99)]); await forbidden.editor.ready(); forbidden.editor.select(['99:10:0']); forbidden.editor.setDraft('99:10:0', 'regular_price', '25');
    assert.equal(forbidden.editor.getSnapshot().selectedCount, 0); assert.equal(forbidden.editor.getSnapshot().draftCount, 0); assert.equal(forbidden.writes.length, 0);
    const external = fixture([row(1, 10, 0, {type: 'external'})]); await external.editor.ready(); external.editor.setDraft('1:10:0', 'regular_price', '12');
    assert.equal(external.editor.getSnapshot().draftCount, 0, 'external products remain clone-capable but cannot be edited'); assert.equal(external.writes.length, 0);
    const loading = fixture([row()]); await loading.editor.ready(); loading.editor.select(['1:10:0']); loading.catalog.setBusy(true);
    await assert.rejects(() => loading.editor.bulk('price', {regular_price: '2'}), /等待/); assert.equal(loading.writes.length, 0);
}
async function uncertainStockBridge() {
    const value = row(1, 10, 101, {stock_bridge: {present: true, valid: true, values: {wcms_stock_manage: 'yes', wcms_stock_qty: 5, wcms_stock_status: 'instock'}}});
    const f = fixture([value], {uncertain: () => true}); await f.editor.ready(); f.editor.select([rowKey(value)]);
    await f.editor.bulk('stock', {stock_quantity: '9'}); assert.equal(f.editor.getResult(rowKey(value)).status, 'unconfirmed');
    await assert.rejects(() => f.editor.verify(rowKey(value)), /WCMS/, 'five public fields alone cannot verify a stale bridge');
    assert.equal(f.writes.length, 1); assert.equal(f.editor.getSnapshot().selectedCount, 1);
    f.actual.get(rowKey(value)).stock_bridge.values.wcms_stock_qty = '9';
    f.actual.get(rowKey(value)).stock_bridge.valid = false;
    await assert.rejects(() => f.editor.verify(rowKey(value)), /WCMS/, 'invalid duplicate bridge metadata cannot be accepted');
    f.actual.get(rowKey(value)).stock_bridge.valid = true;
    await f.editor.verify(rowKey(value)); assert.equal(f.writes.length, 1); assert.equal(f.editor.getResult(rowKey(value)).status, 'verified');
}
async function stoppingFutureWrites() {
    const rows = Array.from({length: 5}, (_, index) => row(1, 10 + index));
    let f;
    f = fixture(rows, {delay: 2, onPut: () => f.editor.stopWrites()}); await f.editor.ready(); f.editor.select(rows.map(rowKey));
    const result = await f.editor.bulk('stock', {stock_quantity: '9'});
    assert.equal(f.writes.length, 1, 'stop prevents all later inventory PUTs'); assert.equal(result[0].status, 'verified', 'already sent PUT still reaches read-back completion');
    assert.equal(result.filter(item => item.status === 'stopped').length, 4); assert.equal(f.editor.getSnapshot().selectedCount, 4);
    assert.equal(f.calls.filter(call => call.url.includes('/catalog-item?') && call.method === 'GET').length, 1, 'no queued row starts a GET after stop');
    let duringPreview;
    duringPreview = fixture([row()], {onGet: () => duringPreview.editor.stopWrites()}); await duringPreview.editor.ready(); duringPreview.editor.setDraft('1:10:0', 'regular_price', '30');
    const stopped = await duringPreview.editor.saveDrafts(); assert.equal(stopped[0].status, 'stopped'); assert.equal(duringPreview.writes.length, 0, 'stop during fresh GET prevents its not-yet-sent PUT');
    assert.equal(duringPreview.editor.getSnapshot().draftCount, 1); assert.match(duringPreview.editor.getSnapshot().action, /已停止/);
}
async function finalWooStockThresholdBridge() {
    for (const finalStatus of ['instock', 'outofstock']) {
        const value = row(1, 10, 101, {manage_stock: false, stock_quantity: 0, stock_status: 'outofstock', stock_bridge: {present: true, valid: true, values: {wcms_stock_manage: 'no', wcms_stock_qty: 0, wcms_stock_status: 'outofstock'}}});
        const f = fixture([value], {uncertain: () => true, afterApply(item) {
            item.stock_status = item.edit_before.stock_status = finalStatus;
            item.stock_bridge.values = {wcms_stock_manage: 'yes', wcms_stock_qty: 3, wcms_stock_status: finalStatus};
        }});
        await f.editor.ready(); f.editor.select([rowKey(value)]); await f.editor.bulk('restore', {stock_quantity: '3'});
        await f.editor.verify(rowKey(value)); assert.equal(f.editor.getResult(rowKey(value)).status, 'verified'); assert.equal(f.writes.length, 1, 'final bridge status follows Woo custom quantity threshold, not old preview');
    }
}
async function cloneUnknownLockPersists() {
    const storageData = new Map(), storage = {getItem: key => storageData.get(key) || null, setItem: (key, value) => storageData.set(key, value), removeItem: key => storageData.delete(key)};
    const rows = Array.from({length: 51}, (_, index) => row(1, 100 + index));
    const f = fixture(rows, {cloneFailAt: 2, storage}); await f.editor.ready(); f.editor.select(rows.map(rowKey));
    const opts = {target_site_id: 4, collision_mode: 'clone_as_new'};
    await assert.rejects(() => f.editor.cloneSelected(opts), /lost/); assert.equal(f.jobs.length, 2); assert.equal(f.editor.getSnapshot().jobs.length, 1, 'known first durable job survives later lost response');
    assert(f.editor.getSnapshot().cloneSubmissionLock); assert(storageData.has('productManagerCatalogCloneSubmissionLock'));
    await assert.rejects(() => f.editor.cloneSelected(opts), /阻止重复/); assert.equal(f.jobs.length, 2, 'same confirmed dialog cannot repeat either known or unknown chunk');
    const refreshed = fixture(rows, {storage}); await refreshed.editor.ready(); refreshed.editor.select(rows.map(rowKey));
    await assert.rejects(() => refreshed.editor.cloneSelected(opts), /阻止重复/); assert.equal(refreshed.jobs.length, 0, 'refresh/sessionStorage does not automatically release clone lock');
    await f.editor.pollJobs(); assert.equal(f.jobs.length, 2, 'only known job polling continues after unknown submission');
    assert.equal(f.editor.acknowledgeCloneSubmission(), true); assert.equal(f.editor.getSnapshot().cloneSubmissionLock, null); assert(!storageData.has('productManagerCatalogCloneSubmissionLock'));
}
async function cloneGroupingAndDurableProgress() {
    const rows = [row(1, 10, 101), row(1, 10, 102), ...Array.from({length: 51}, (_, index) => row(1, 20 + index)), row(2, 10), row(4, 10)];
    const f = fixture(rows); await f.editor.ready(); f.editor.select(rows.map(rowKey));
    await f.editor.cloneSelected({target_site_id: 4, include_variations: true, include_images: false, status_on_target: 'publish', collision_mode: 'clone_as_new'});
    assert.equal(f.jobs.length, 3); assert.deepEqual(f.jobs.map(job => job.product_ids.length), [50, 2, 1]);
    assert(f.jobs.every(job => job.target_site_id === 4 && job.source_site_id !== 4 && job.status_on_target === 'draft' && job.include_images === false));
    assert.equal(f.jobs.flatMap(job => job.product_ids).filter(id => id === 10).length, 2, 'two flavors dedup within source; identical IDs on another source retain own job');
    assert.equal(f.editor.getSnapshot().cloneSkippedSelf, 1);
    const posted = f.jobs.length; await f.editor.pollJobs(); assert.equal(f.jobs.length, posted, 'polling durable jobs does not create duplicates'); assert(f.editor.getSnapshot().jobs.every(job => job.terminal));
}
(async () => {
    await selectionAndDirtyState(); await stalePreflightAndFailedRetry(); await inheritedStockAndPayloads(); await uncertainWriteCannotReplay();
    await concurrencyAndPermissions(); await uncertainStockBridge(); await stoppingFutureWrites(); await finalWooStockThresholdBridge();
    await cloneGroupingAndDurableProgress(); await cloneUnknownLockPersists();
    process.stdout.write('Cross-site editor: selection/drafts, exact-child preflight, inherited inventory, failure/retry, uncertain writes, concurrency and durable clone checks passed.\n');
})().catch(error => { process.stderr.write(error.stack + '\n'); process.exitCode = 1; });
