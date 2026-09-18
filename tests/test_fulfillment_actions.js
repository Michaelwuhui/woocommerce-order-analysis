const assert = require('node:assert/strict');
const {test} = require('node:test');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '..', 'templates', 'fulfillment.html'), 'utf8');
const fn = source.slice(source.indexOf('function fulfillmentWorkflowButtons('), source.indexOf('async function showFulfillment('));
function buttons(status, mode='internal', allowed=true) {
  const context = {ffCanManage:allowed, esc:s=>String(s)};
  vm.runInNewContext(fn, context);
  return context.fulfillmentWorkflowButtons({id:'test', status, mode}, 'test-order');
}
test('terminal or blocked fulfillments expose no pick/pack actions', () => {
  for (const state of ['cancelled','superseded','shipped','delivered','returned','stock_shortage','submission_unknown']) {
    assert.equal(buttons(state), '');
  }
});
test('picking and packing buttons match legal transitions', () => {
  assert.match(buttons('ready_to_pick'), /开始拣货/);
  assert.match(buttons('ready_to_pick'), /打包完成/);
  assert.doesNotMatch(buttons('picking'), /开始拣货/);
  assert.match(buttons('picking'), /打包完成/);
  assert.equal(buttons('packed'), '');
});
test('external warehouses and users without management permission get no workflow actions', () => {
  assert.equal(buttons('ready_to_pick','external_wms'), '');
  assert.equal(buttons('ready_to_pick','internal',false), '');
});
