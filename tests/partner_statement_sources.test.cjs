const test = require('node:test');
const assert = require('node:assert/strict');
const view = require('../static/js/partner_statement_sources.js');

test('transit goods clearly show zero capital and profit, with service currency separate', () => {
    const html = view.summary({revision:2,currency:'PLN',settlement:{ready:true,amounts_label:'完整核算',
        partner_profit:30,our_receivable:130,packing_fee_cny:20,packing_charges:1,transit_shipping_loss_separate:0},
        groups:[{mode:'transit_service',label:'我方备货 · 中转服务',orders:1,quantity:1,revenue:100,cost_recovery:0,partner_profit:0,our_receivable:100}]},false);
    assert.match(html,/不返还备货成本 · 不参与商品分成/);
    assert.match(html,/中转打包费 · 应付合伙人/);
    assert.match(html,/20\.00 <small class="fs-6">CNY/);
    assert.match(html,/第 2 版 · 冻结快照/);
    assert.match(html,/不自动抵扣/);
});

test('unverified amounts remain unknown and product references cannot inject markup', () => {
    const html=view.renderRows({rows:[{order_number:'<img onerror=x>',order_id:'id',product_name:'<script>x</script>',
        quantity:1,mode_label:'中转服务',cost_recovery:null,partner_profit:null,our_receivable:null,source_basis:'OMS',pending:['金额待核实']}]},'lines',false);
    assert.match(html,/待核实/);
    assert.ok(!html.includes('<script>') && !html.includes('<img'));
    assert.match(html,/&lt;script&gt;/);
    assert.equal(view.amount(null),'待核实');
    assert.equal(view.amount(0),'0.00');
});
