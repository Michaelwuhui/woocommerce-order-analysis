const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const html = fs.readFileSync('templates/shipping.html', 'utf8');
function extract(start, end) {
  return html.slice(html.indexOf(start), html.indexOf(end, html.indexOf(start)));
}
const buttons = Object.fromEntries(['confirmShipBtn', 'shipContinueBtn', 'shipFinalBtn'].map(id => [id, {}]));
const alerts = [], calls = [];
let refreshes = 0;
const context = vm.createContext({
  document: {getElementById: id => buttons[id] || ({bigOrderAck: {checked: true}, shipOrderId: {value: '20-123'}})[id]},
  alert: text => alerts.push(text),
  fetch: async (url, options) => {calls.push({url, options}); return {};},
  parseMutationJsonResponse: async () => ({success: true, message: '已安排后台核对'}),
  loadOrders: () => {refreshes++;},
  _shipReconciliationPending: true, _shipFulfillmentLoading: false,
  _shipUsesFulfillment: false, _shipFulfillmentId: null,
});
vm.runInContext(extract('async function recheckShipment(', 'function shipEscape('), context);
vm.runInContext(extract('function updateShipBtnState()', '// Populate the carrier dropdown'), context);
vm.runInContext(extract('function shipOrder(mode)', 'async function shipFulfillmentOrder('), context);
const base = fs.readFileSync('templates/base.html', 'utf8');
vm.runInContext(base.slice(base.indexOf('function renderShipmentReconciliation('), base.indexOf('function showOrderDetail(')), context);
(async () => {
  context.updateShipBtnState();
  assert.ok(Object.values(buttons).every(b => b.disabled));
  context.shipOrder('single');
  assert.equal(calls.length, 0);
  assert.match(alerts.pop(), /原运单正在核对/);
  const button = {};
  await context.recheckShipment(button);
  assert.equal(calls.length, 1);
  assert.equal(calls[0].url, '/api/shipping/reconcile');
  assert.deepEqual(JSON.parse(calls[0].options.body), {order_id: '20-123'});
  assert.equal(button.disabled, false);
  assert.equal(refreshes, 1);
  context._shipReconciliationPending = false;
  context.updateShipBtnState();
  assert.ok(Object.values(buttons).every(b => !b.disabled));
  const recovered = context.renderShipmentReconciliation({outcome: 'verified', label: '已补齐', reason: '<unsafe>'});
  assert.match(recovered, /alert-success/);
  assert.match(recovered, /已补齐/);
  assert.ok(!recovered.includes('<unsafe>'));
  assert.equal(context.renderShipmentReconciliation({outcome: 'completed'}), '');
  console.log('Shipment reconciliation UI: pending guard, read-back action, refresh and retry state passed');
})().catch(error => {console.error(error); process.exitCode = 1;});
