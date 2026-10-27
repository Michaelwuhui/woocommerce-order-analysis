const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const view = require('../static/js/shipping_list_view.js');
const productFilter = require('../static/js/shipping_product_filter_compact1.js');
const template = fs.readFileSync(path.join(root, 'templates/shipping.html'), 'utf8');
const inlineScript = template.match(/<script>\s*([\s\S]*?)<\/script>/)[1];

class Element {
    constructor() {
        this.value = '';
        this.textContent = '';
        this.innerHTML = '';
        this.style = {};
        this.dataset = {};
        this.attributes = {};
        this.events = {};
        const classes = new Set();
        this.classList = {
            add: (...names) => names.forEach(name => classes.add(name)),
            remove: (...names) => names.forEach(name => classes.delete(name)),
            contains: name => classes.has(name),
            toggle: (name, enabled) => enabled ? classes.add(name) : classes.delete(name)
        };
    }
    addEventListener(name, handler) { this.events[name] = handler; }
    getAttribute(name) { return this.attributes[name] || ''; }
    setAttribute(name, value) { this.attributes[name] = value; }
    scrollIntoView() {}
    appendChild(child) { this.textContent += child.textContent; }
    append(...children) { children.forEach(child => this.appendChild(child)); }
    replaceChildren(...children) {
        this.innerHTML = '';
        this.textContent = children.map(child => child.textContent).join('');
    }
}

function makeOrder(index, productName = 'Common Flavor', quantity = 1) {
    return {
        id: `order-${index}`, number: `TEST${index}`, source: 'shop.example',
        source_url: 'https://shop.example', currency: 'PLN', total: 10,
        shipping_total: 0, customer_name: 'Synthetic customer',
        customer_email: 'buyer@example.invalid', customer_address: 'Test street',
        products: [{name: productName, quantity, total: 10}], product_count: quantity,
        parcels: [], fulfillment_warehouses: []
    };
}

function harness({orders = [], canShip = true, fetchImpl} = {}) {
    const elements = new Map();
    const get = id => {
        if (!elements.has(id)) elements.set(id, new Element());
        return elements.get(id);
    };
    const tabs = ['#pendingTab', '#shippedTab', '#outcomeTab'].map(href => {
        const element = new Element();
        element.setAttribute('href', href);
        return element;
    });
    let selectedTab = tabs[0];
    const document = {
        head: new Element(), documentElement: new Element(), events: {},
        getElementById: get,
        createElement: () => new Element(),
        addEventListener(name, handler) { this.events[name] = handler; },
        querySelector: selector => selector === '#shippingTabs .active' ? selectedTab : null,
        querySelectorAll: selector => selector === '#shippingTabs a' ? tabs : []
    };
    const calls = [];
    const context = vm.createContext({
        document, URLSearchParams, AbortController,
        ShippingListView: view, ShippingProductFilter: productFilter,
        flatpickr: () => ({}), console,
        setTimeout, clearTimeout,
        fetch: (url, options) => {
            calls.push({url, options});
            if (fetchImpl) return fetchImpl(url, options);
            const body = url.startsWith('/api/shipping/pending-outcome') ? []
                : url.startsWith('/api/shipping/shipped')
                    ? {orders: [], pagination: {page: 2, pages: 3, per_page: 50, total: 120}}
                    : orders;
            return Promise.resolve({ok: true, json: async () => body});
        }
    });
    context.window = context;
    vm.runInContext(inlineScript.replace(
        "{{ 'true' if current_user.can_ship() else 'false' }}", String(canShip)), context);
    return {
        context, document, get, calls,
        select: href => { selectedTab = tabs.find(tab => tab.getAttribute('href') === href); },
        renderedIds: () => Array.from(get('pendingOrdersBody').innerHTML.matchAll(/showOrderDetail\('([^']+)'\)/g), match => match[1])
    };
}

test('every pending order remains reachable exactly once through bounded pages', () => {
    const orders = Array.from({length: 137}, (_, index) => makeOrder(index));
    const seen = [];
    for (let page = 1; page <= 3; page += 1) {
        const result = view.paginate(orders, page);
        assert.equal(result.total, 137);
        assert.equal(result.pages, 3);
        assert.ok(result.orders.length <= 50);
        seen.push(...result.orders.map(order => order.id));
    }
    assert.deepEqual(seen, orders.map(order => order.id));
    assert.equal(new Set(seen).size, 137);
});

test('empty, invalid and out-of-range pages stay bounded without mutating input', () => {
    assert.deepEqual(view.paginate([], 9).orders, []);
    assert.equal(view.paginate(null, -1).page, 1);
    const orders = Array.from({length: 60}, (_, index) => makeOrder(index));
    assert.equal(view.paginate(orders, 99, 500).page, 2);
    assert.equal(view.paginate(orders, 99, 500).orders.length, 10);
    assert.equal(view.paginate(orders, -3).page, 1);
    assert.equal(view.paginate(orders, NaN).orders.length, 50);
    assert.equal(orders.length, 60);
});

test('default opening requests only pending and bounds row rendering, preserving global totals', async () => {
    const page = harness({orders: Array.from({length: 137}, (_, index) => makeOrder(index, 'Common Flavor', 2))});
    page.document.events.DOMContentLoaded();
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(page.calls.length, 1);
    assert.match(page.calls[0].url, /^\/api\/shipping\/pending\?/);
    assert.equal(page.renderedIds().length, 50);
    assert.equal(page.get('pendingCount').textContent, '待发货: 137 单 / 274 支');
    assert.match(page.get('pendingPageSummary').textContent, /第 1-50 单，共 137 单/);
    assert.equal(page.get('shippedOrdersBody').innerHTML, '');
    assert.equal(page.get('outcomeOrdersBody').innerHTML, '');
});

test('local pending pagination makes no additional API request and reaches the last order', async () => {
    const page = harness({orders: Array.from({length: 137}, (_, index) => makeOrder(index))});
    await page.context.loadOrders();
    page.context.changePendingPage(3);
    assert.equal(page.calls.length, 1);
    assert.equal(page.renderedIds().length, 37);
    assert.equal(page.renderedIds().at(-1), 'order-136');
    assert.equal(page.get('pendingCount').textContent, '待发货: 137 单 / 137 支');
});

test('product filtering searches beyond the first page and keeps full filtered quantities', async () => {
    const orders = Array.from({length: 137}, (_, index) => makeOrder(index));
    orders[100] = makeOrder(100, 'Off-page Flavor', 3);
    const page = harness({orders});
    await page.context.loadOrders();
    page.get('pendingProductFilter').value = 'Off-page Flavor';
    await page.context.loadOrders();
    assert.deepEqual(page.renderedIds(), ['order-100']);
    assert.equal(page.get('pendingCount').textContent, '待发货: 1 单 / 3 支');
    assert.match(page.get('pendingProductStats').textContent, /1 个订单 · 3 支待发货/);
    page.get('pendingProductFilter').value = '';
    await page.context.loadOrders();
    assert.equal(page.renderedIds().length, 50);
    assert.equal(page.get('pendingCount').textContent, '待发货: 137 单 / 139 支');
});

test('switching tabs and shipped page changes fetch only the selected list', async () => {
    const page = harness();
    page.select('#shippedTab');
    await page.context.loadOrders(2, true);
    assert.equal(page.calls.length, 1);
    assert.match(page.calls[0].url, /^\/api\/shipping\/shipped\?/);
    assert.match(page.calls[0].url, /page=2/);
    assert.equal(page.get('shippedCount').textContent, '已发货: 120');
    page.select('#outcomeTab');
    await page.context.loadOrders();
    assert.equal(page.calls.length, 2);
    assert.match(page.calls[1].url, /^\/api\/shipping\/pending-outcome\?/);
    assert.equal(page.get('outcomeCount').textContent, '待确认: 0');
});

test('changing scope sends filters to the API and clears unrelated stale counts', async () => {
    const page = harness({orders: [makeOrder(1)]});
    await page.context.loadOrders();
    page.get('shippedCount').textContent = '已发货: 120';
    page.get('outcomeCount').textContent = '待确认: 17';
    page.get('countryFilter').value = 'PL';
    page.get('siteFilter').value = 'https://scope.example';
    page.get('managerFilter').value = 'Synthetic manager';
    page.get('dateFrom').value = '2026-10-01';
    page.get('searchInput').value = 'TEST1';
    await page.context.loadOrders();
    const params = new URLSearchParams(page.calls.at(-1).url.split('?')[1]);
    assert.equal(params.get('source'), 'https://scope.example');
    assert.equal(params.get('country'), 'PL');
    assert.equal(params.get('manager'), 'Synthetic manager');
    assert.equal(params.get('start_date'), '2026-10-01');
    assert.equal(params.get('search'), 'TEST1');
    assert.equal(page.get('shippedCount').textContent, '已发货: —');
    assert.equal(page.get('outcomeCount').textContent, '待确认: —');
});

test('a slower obsolete response cannot overwrite a newer filter result', async () => {
    const pending = [];
    const page = harness({fetchImpl: () => new Promise(resolve => pending.push(resolve))});
    const first = page.context.loadOrders();
    page.get('siteFilter').value = 'https://new.example';
    const second = page.context.loadOrders();
    assert.equal(page.calls[0].options.signal.aborted, true);
    pending[1]({ok: true, json: async () => [makeOrder('new')]});
    await second;
    pending[0]({ok: true, json: async () => [makeOrder('old')]});
    await first;
    assert.deepEqual(page.renderedIds(), ['order-new']);
});

test('failed requests display a retry message without leaving a busy placeholder', async () => {
    const page = harness({fetchImpl: async () => ({ok: false, status: 503})});
    await page.context.loadOrders();
    assert.equal(page.get('pendingOrdersBody').textContent, '加载失败，请点击刷新重试');
    assert.equal(page.get('pendingPagination').classList.contains('d-none'), true);
});

test('the read-only operator guard remains applied', async () => {
    const page = harness({canShip: false, orders: [makeOrder(1)]});
    await page.context.loadOrders();
    assert.equal(page.document.documentElement.classList.contains('ship-readonly'), true);
    assert.match(page.document.head.textContent, /\.ship-readonly \.ship-action/);
});
