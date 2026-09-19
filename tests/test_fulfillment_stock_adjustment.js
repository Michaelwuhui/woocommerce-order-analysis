const assert = require('node:assert/strict');
const {test} = require('node:test');
const {fulfillmentStocktakePayload: payload, fulfillmentStocktakeButtons: buttons} = require('../static/fulfillment_stock_adjustment.js');
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
