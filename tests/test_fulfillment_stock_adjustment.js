const assert = require('node:assert/strict');
const {test} = require('node:test');
const {fulfillmentStocktakePayload: payload, fulfillmentStocktakeButtons: buttons, fulfillmentStocktakeReceipt: receipt} = require('../static/fulfillment_stock_adjustment.js');
const {fulfillmentShortageDetails: shortageDetails} = require('../static/fulfillment_shortages.js');
const context = {revision: 3, items: [
  {sku_id: 1, sku_code: 'ONE', on_hand: 10, reserved: 2, movement: 99},
  {sku_id: 2, sku_code: 'TWO', on_hand: 0, reserved: 0, movement: 0},
]};
test('only edited physical counts are submitted with their original snapshot', () => {
  const result = payload(context, {1: '10', 2: '4'}, ' physical count ', 'request-one');
  assert.deepEqual(result, {request_key: 'request-one', revision: 3, note: 'physical count', items: [
    {sku_id: 2, qty: 4, baseline: 0, baseline_reserved: 0, baseline_movement: 0},
  ]});
});
test('invalid counts, shortage of reserved stock, empty reason and no change are rejected', () => {
  for (const qty of ['', '-1', '1.5', '1e2', '100000001', '1']) {
    assert.throws(() => payload(context, {1: qty, 2: '0'}, 'counted', 'request-one'));
  }
  assert.throws(() => payload(context, {1: '12', 2: '0'}, ' ', 'request-one'));
  assert.throws(() => payload(context, {1: '10', 2: '0'}, 'counted', 'request-one'));
});
test('non-admin gets no quick-adjust buttons; joint dispatch preserves warehouse identities', () => {
  const targets = [{fulfillment_id: 'one', warehouse_name: 'North'}, {fulfillment_id: 'two', warehouse_name: 'South'}];
  assert.equal(buttons(targets, 'order', false, String), '');
  const html = buttons(targets, 'order', true, String);
  assert.match(html, /data-ff-stock-id="one"/);
  assert.match(html, /data-ff-stock-id="two"/);
  assert.match(html, /快速调整库存 · North/);
  assert.match(html, /快速调整库存 · South/);
});

test('saved stocktake distinguishes remaining shortages, missing status and resolved stock', () => {
  const remaining = receipt(7, {state: {has_shortage: true}, shortage_items: [{sku_code: 'MISSING', shortage_qty: 1}]});
  assert.equal(remaining.warning, true);
  assert.match(remaining.text, /MISSING 缺 1 件/);
  assert.equal(receipt(7, null).warning, true);
  assert.match(receipt(7, null).text, /库存已调整/);
  assert.equal(receipt(7, {state: {has_shortage: true}}).warning, true);
  assert.equal(receipt(7, {state: {has_shortage: false}, shortage_items: []}).warning, false);
});

test('unallocated shortage display escapes imported product names and distinguishes mapping', () => {
  const escape = value => String(value).replaceAll('<', '&lt;').replaceAll('>', '&gt;');
  const html = shortageDetails([{name: '<img onerror=alert(1)>', sku_code: '<SKU>', sku_id: null,
    ordered_qty: 1, allocated_qty: 0, shortage_qty: 1}], escape);
  assert.doesNotMatch(html, /<img/);
  assert.match(html, /&lt;SKU&gt;/);
  assert.match(html, /未建立 SKU 映射/);
  assert.equal(shortageDetails([], escape), '');
});
