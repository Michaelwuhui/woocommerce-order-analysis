'use strict';

const assert = require('node:assert/strict');
const test = require('node:test');
const table = require('../static/js/customer_table.js');

function customer(overrides = {}) {
    return {
        name: 'Alice Customer', email: 'primary@example.test', phone: '+48 501 222 333',
        identity_emails: ['primary@example.test', 'secondary@example.test'],
        identity_email_count: 2, identity_matched_by: ['phone', 'address'],
        source: 'https://www.example.test', site_count: 2,
        site_list: [{short: 'example.test', manager: 'Manager One'}, {short: 'second.test', manager: 'Manager Two'}],
        tier: 'VIP', successful_orders: 12, total_spent_cny: 1250.75, total_loss: 20,
        undelivered_orders: 1, problem_return_orders: 1, total_orders: 15, refusal_rate: 6.7,
        shipping_loss_total: 5, product_loss_total: 15, currency: 'PLN',
        missing_exchange_rates: ['XYZ@2026-09'], last_order_date: '2026-10-06T09:00:00',
        actions: [{type: 'success', icon: 'gift', text: '专属礼遇'}],
        ...overrides,
    };
}

const noFilters = {query: '', lossOnly: false, multiSiteOnly: false, mergedOnly: false};
const config = {siteManagers: {'https://www.example.test': 'Manager One'}, sourceDisplayMode: 'all'};

test('search reaches all secondary emails and telephone in rows beyond page one', () => {
    const rows = Array.from({length: 80}, (_, i) => customer({
        name: 'Customer ' + i,
        email: `primary${i}@example.test`,
        identity_emails: [`primary${i}@example.test`, `SECONDARY${i}@example.test`],
        phone: '+48 500 200 ' + i,
    }));
    assert.equal(rows.filter(row => table.matchesCustomer(row, {...noFilters, query: 'secondary79@'})).length, 1);
    assert.equal(rows.filter(row => table.matchesCustomer(row, {...noFilters, query: 'secondary79@'}))[0], rows[79]);
    assert.equal(table.matchesCustomer(rows[79], {...noFilters, query: '+48 500 200 79'}), true);
    assert.equal(table.matchesCustomer(rows[79], {...noFilters, query: 'customer 79'}), true);
    assert.equal(table.matchesCustomer(rows[79], {...noFilters, query: 'missing@example.test'}), false);
    assert.equal(table.matchesCustomer(customer({identity_emails: []}), {...noFilters, query: 'primary@example.test'}), true);
    assert.equal(table.matchesCustomer(customer({identity_emails: null}), {...noFilters, query: 'primary@example.test'}), true);
});

test('loss, multisite, merged identity, and search conditions combine with AND', () => {
    const filters = {query: 'secondary', lossOnly: true, multiSiteOnly: true, mergedOnly: true};
    assert.equal(table.matchesCustomer(customer(), filters), true);
    for (const overrides of [
        {total_loss: 0}, {total_loss: -1}, {site_count: 1}, {identity_email_count: 1},
        {identity_emails: ['only@example.test']},
    ]) assert.equal(table.matchesCustomer(customer(overrides), filters), false);
    assert.equal(table.matchesCustomer(customer({total_loss: '20', site_count: '2', identity_email_count: '2'}), filters), true);
    assert.equal(table.matchesCustomer(customer({total_loss: 'NaN'}), {...noFilters, lossOnly: true}), false);
    assert.equal(table.matchesCustomer(customer({site_count: null}), {...noFilters, multiSiteOnly: true}), false);
    assert.equal(table.matchesCustomer(customer({identity_email_count: null}), {...noFilters, mergedOnly: true}), false);
    assert.equal(table.matchesCustomer(customer({total_loss: 0, site_count: 1, identity_email_count: 1}), noFilters), true);
});

test('spending, order, site, and loss sorting use native numbers rather than display strings', () => {
    const rows = [customer({name: 'low', total_spent_cny: 9, total_loss: 8}),
        customer({name: 'high', total_spent_cny: 1000, total_loss: 120}),
        customer({name: 'mid', total_spent_cny: 100, total_loss: 19})];
    for (const column of [2, 5, 6, 7]) {
        assert.equal(typeof table.sortValue(column, rows[0], config), 'number');
        assert.equal(typeof table.renderCell(column, 'sort', rows[0], config), 'number');
        assert.equal(typeof table.renderCell(column, 'type', rows[0], config), 'number');
    }
    assert.deepEqual(rows.slice().sort((a, b) => table.sortValue(6, b, config) - table.sortValue(6, a, config)).map(row => row.name),
        ['high', 'mid', 'low']);
    assert.deepEqual(rows.slice().sort((a, b) => table.sortValue(7, b, config) - table.sortValue(7, a, config)).map(row => row.name),
        ['high', 'mid', 'low']);
    assert.equal(table.sortValue(8, rows[0], config), '2026-10-06');
    assert.equal(table.renderCell(6, 'display', rows[1], config).startsWith('¥1,000.00'), true);
    assert.equal(table.renderCell(6, 'display', rows[0], config).includes('XYZ@2026-09'), true);
});

test('source display and sorting respect manager-only preference', () => {
    assert.equal(table.sortValue(1, customer(), config), 'Manager One example.test');
    const managerOnly = {...config, sourceDisplayMode: 'manager_only'};
    assert.equal(table.sortValue(1, customer(), managerOnly), 'Manager One');
    assert.equal(table.renderCell(1, 'display', customer(), managerOnly).includes('example.test'), false);
    assert.equal(table.renderCell(1, 'display', customer(), managerOnly).includes('Manager One'), true);
    assert.equal(table.renderCell(1, 'display', customer(), {siteManagers: {}, sourceDisplayMode: 'manager_only'}), '');
});

test('all external customer values are escaped in display HTML and details carry inert data', () => {
    const attack = '\"><img src=x onerror=globalThis.pwned=1><script>pwned()</script>\'&';
    const row = customer({name: attack, email: attack, phone: attack, source: attack, currency: attack,
        identity_emails: [attack, 'safe@example.test'], identity_matched_by: [attack],
        site_list: [{short: attack, manager: attack}], missing_exchange_rates: [attack],
        actions: [{type: attack, icon: attack, text: attack}], last_order_date: attack});
    const maliciousConfig = {siteManagers: {[attack]: attack}, sourceDisplayMode: 'all'};
    assert.equal(table.escapeHtml('<>&\"\''), '&lt;&gt;&amp;&quot;&#39;');
    for (let column = 0; column < 11; column++) {
        const html = table.renderCell(column, 'display', row, maliciousConfig);
        assert.equal(html.includes('<img'), false, 'raw image at column ' + column);
        assert.equal(html.includes('<script'), false, 'raw script at column ' + column);
        assert.equal(html.includes('onclick='), false, 'inline script at column ' + column);
        if (html.includes('onerror=')) assert.equal(html.includes('&lt;img'), true, 'attack may occur only as escaped text');
    }
    assert.match(table.renderCell(10, 'display', row, maliciousConfig), /data-customer-detail="&quot;&gt;&lt;img/);
    assert.match(table.renderCell(9, 'display', row, maliciousConfig), /bg-secondary/);
    assert.match(table.renderCell(9, 'display', row, maliciousConfig), /bi-info-circle/);
    assert.equal(table.renderCell(0, 'filter', row, maliciousConfig), attack);
});

test('column schema preserves numeric fields and non-orderable actions', () => {
    const columns = table.buildColumns(config);
    assert.equal(columns.length, 11);
    for (let i = 0; i < columns.length; i++) {
        assert.equal(columns[i].type, [2, 5, 6, 7].includes(i) ? 'num' : 'string');
        assert.equal(columns[i].orderable, ![3, 9, 10].includes(i));
        assert.equal(columns[i].render(null, 'sort', customer()), table.sortValue(i, customer(), config));
    }
});

function mockNode(id) {
    const handlers = {};
    const classes = new Set();
    return {
        id, handlers, value: '', innerHTML: '', style: {},
        addEventListener(type, fn) { handlers[type] = fn; },
        classList: {toggle(value, enabled) { if (enabled) classes.add(value); else classes.delete(value); }},
        contains(button) { return button.withinTable === true; },
        classes,
    };
}

function mountFixture(rows) {
    const nodes = Object.fromEntries(['customersTable', 'customerSearchBox', 'toggleLossOnlyBtn',
        'toggleMultiSiteBtn', 'toggleMergedBtn'].map(id => [id, mockNode(id)]));
    const filters = [];
    let options;
    const draws = [], orders = [];
    const dt = {draw() { draws.push(true); }, order(value) { orders.push(value); return dt; }};
    function jquery(element) {
        assert.equal(element, nodes.customersTable);
        return {DataTable(value) { options = value; return dt; }};
    }
    jquery.fn = {dataTable: {ext: {search: {push(fn) { filters.push(fn); }}}}};
    const document = {getElementById(id) { return nodes[id] || null; }};
    assert.equal(table.mount(document, jquery, config, rows), dt);
    return {nodes, filters, get options() { return options; }, draws, orders, dt};
}

test('mount retains all records while enabling deferred rows and skipping width scan', () => {
    const rows = Array.from({length: 150}, (_, i) => customer({name: 'Customer ' + i}));
    const fixture = mountFixture(rows);
    assert.equal(fixture.options.data, rows);
    assert.equal(fixture.options.deferRender, true);
    assert.equal(fixture.options.autoWidth, false);
    assert.equal(fixture.options.pageLength, 25);
    assert.deepEqual(fixture.options.order, [[6, 'desc']]);
    const predicate = fixture.filters[0];
    const settings = {nTable: fixture.nodes.customersTable, aoData: [{_aData: rows[149], anCells: null}]};
    assert.equal(predicate(settings, [], 0, rows[149]), true);
    assert.equal(predicate(settings, [], 0), true);
    assert.equal(predicate({nTable: {}}, [], 999), true, 'other tables remain unaffected');
});

test('toggle event wiring combines flags and flushes pending query before drawing', () => {
    const fixture = mountFixture([customer()]);
    fixture.nodes.customerSearchBox.value = '  SECONDARY@EXAMPLE.TEST  ';
    const settings = {nTable: fixture.nodes.customersTable};
    const matches = row => fixture.filters[0](settings, [], 0, row);
    fixture.nodes.toggleLossOnlyBtn.handlers.click();
    fixture.nodes.toggleMultiSiteBtn.handlers.click();
    fixture.nodes.toggleMergedBtn.handlers.click();
    assert.equal(matches(customer()), true);
    assert.equal(matches(customer({total_loss: 0})), false);
    assert.equal(matches(customer({site_count: 1})), false);
    assert.equal(matches(customer({identity_email_count: 1})), false);
    assert.equal(matches(customer({identity_emails: ['other@example.test']})), false);
    assert.equal(fixture.nodes.toggleMergedBtn.classes.has('active'), true);
    assert.deepEqual(fixture.orders, [[7, 'desc'], [2, 'desc']]);
    fixture.nodes.toggleLossOnlyBtn.handlers.click();
    assert.equal(matches(customer({total_loss: 0})), true);
    assert.deepEqual(fixture.orders.at(-1), [6, 'desc']);
    assert.equal(fixture.draws.length, 4);
});

test('search events debounce drawing and normalize query without inspecting DOM rows', () => {
    const saved = {setTimeout: globalThis.setTimeout, clearTimeout: globalThis.clearTimeout};
    const callbacks = new Map();
    let sequence = 0;
    globalThis.setTimeout = (fn, delay) => { assert.equal(delay, 300); callbacks.set(++sequence, fn); return sequence; };
    globalThis.clearTimeout = id => { callbacks.delete(id); };
    try {
        const fixture = mountFixture([customer()]);
        fixture.nodes.customerSearchBox.value = 'old';
        fixture.nodes.customerSearchBox.handlers.input();
        fixture.nodes.customerSearchBox.value = '  SECONDARY@EXAMPLE.TEST  ';
        fixture.nodes.customerSearchBox.handlers.input();
        assert.equal(callbacks.size, 1);
        assert.equal(fixture.draws.length, 0);
        callbacks.values().next().value();
        assert.equal(fixture.draws.length, 1);
        const predicate = fixture.filters[0];
        const settings = {nTable: fixture.nodes.customersTable};
        assert.equal(predicate(settings, [], 79, customer()), true);
        assert.equal(predicate(settings, [], 79, customer({identity_emails: ['nomatch@example.test']})), false);
    } finally {
        globalThis.setTimeout = saved.setTimeout;
        globalThis.clearTimeout = saved.clearTimeout;
    }
});

test('delegated detail events pass literal quoted email without eval and reject outside buttons', () => {
    const saved = globalThis.showCustomerDetail;
    const calls = [];
    globalThis.showCustomerDetail = email => calls.push(email);
    try {
        const fixture = mountFixture([customer()]);
        const email = "o'quote\"<script>@example.test";
        const button = {withinTable: true, getAttribute(name) { assert.equal(name, 'data-customer-detail'); return email; }};
        fixture.nodes.customersTable.handlers.click({target: {closest() { return button; }}});
        assert.deepEqual(calls, [email]);
        button.withinTable = false;
        fixture.nodes.customersTable.handlers.click({target: {closest() { return button; }}});
        fixture.nodes.customersTable.handlers.click({target: {closest() { return null; }}});
        assert.deepEqual(calls, [email]);
    } finally {
        if (saved === undefined) delete globalThis.showCustomerDetail;
        else globalThis.showCustomerDetail = saved;
    }
});
