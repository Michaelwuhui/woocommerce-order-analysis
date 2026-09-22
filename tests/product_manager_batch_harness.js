// Execute the page's real batch functions against an endpoint-lease simulator.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const template = fs.readFileSync(process.argv[2], 'utf8');
const operation = process.argv[3];
const scenario = process.argv[4];
const functions = template.slice(template.indexOf('    async function parseProductManagerJson('),
    template.indexOf('    function showBatchResult('));
const rowKeyFunction = template.slice(template.indexOf('    function rowKeyToIds('),
    template.indexOf('    function findRowByKey('));
const calls = [];
let active = 0, maxActive = 0, summary, refreshes = 0;
const elements = {
    pmBatchApplyBtn: {dataset: {opType: operation}, disabled: false},
    pmBatchStockQty: {value: '7'}, pmBatchRestoreQty: {value: '9'},
    pmBatchRegularPrice: {value: '15'}, pmBatchSalePrice: {value: ''},
};
const context = {
    $: key => elements[key] || {},
    selectedRows: new Set(scenario === 'parent'
        ? ['prod-10', 'var-10-11', 'prod-20']
        : ['var-10-11', 'var-10-12', 'var-10-13']),
    products: [{id: 10, type: 'variable'}, {id: 20, type: 'simple'}],
    variationsCache: {}, expandedParents: new Set([10]),
    currentSiteId: 2, currentPage: 1,
    window: {crypto: {randomUUID: () => 'batch-test-id'}},
    bootstrap: {Modal: {getInstance: () => ({hide() {}})}},
    alert: message => {throw new Error(message);},
    showBatchResult: result => {summary = result;},
    loadProducts: () => {refreshes++;},
    fetch: async (url, options) => {
        if (!options) {
            assert.equal(url, '/api/product-manager/variations/2/10');
            return {ok: true, text: async () => JSON.stringify({variations: [{id: 11}, {id: 12}, {id: 13}]})};
        }
        assert.equal(options.method, 'PUT');
        const payload = JSON.parse(options.body);
        assert.equal(payload.batch_id, 'batch-test-id');
        assert(!('parent_id' in payload));
        assert(!('product_id' in payload));
        const isStock = 'manage_stock' in payload;
        const busy = isStock && active > 0;
        active++;
        maxActive = Math.max(maxActive, active);
        calls.push({url, payload});
        await new Promise(resolve => setTimeout(resolve, 5));
        active--;
        const rejected = scenario === 'failure' && url.endsWith('/12');
        const error = busy ? 'RESOURCE_BUSY' : rejected ? 'WC rejected stock update' : null;
        return {ok: !error, status: error ? 409 : 200, text: async () => JSON.stringify(
            error ? {success: false, error} : {success: true, variation: {id: Number(url.split('/').pop())}}
        )};
    },
};
vm.createContext(context);
vm.runInContext(rowKeyFunction + functions, context);
(async () => {
    await context.applyBatchOp();
    const expected = scenario === 'parent' ? 5 : 3;
    assert.equal(calls.length, expected);
    assert.equal(summary.success_count, expected - (scenario === 'failure' ? 1 : 0));
    assert.equal(summary.failed_count, scenario === 'failure' ? 1 : 0);
    assert.equal(maxActive, operation === 'price' ? 2 : 1);
    assert.equal(elements.pmBatchApplyBtn.disabled, false);
    assert.equal(refreshes, 1);
    assert.equal(Object.keys(context.variationsCache).length, 0);
    if (scenario === 'parent') {
        assert(calls.slice(0, 3).every(call => call.url.includes('/variations/')));
        assert.equal(new Set(calls.map(call => call.url)).size, expected);
        assert.equal(calls[3].url, '/api/product-manager/products/2/10');
        assert.equal(calls[4].url, '/api/product-manager/products/2/20');
    } else {
        assert(calls.every(call => call.url.includes('/10/variations/')));
    }
    if (scenario === 'failure') assert.match(summary.results.failed[0].error, /WC rejected/);
    console.log(JSON.stringify({operation, scenario, successes: summary.success_count, failures: summary.failed_count, maxActive}));
})().catch(error => {console.error(error); process.exitCode = 1;});
