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

test('unassigned supply source displays unknown rights rather than zero', () => {
    const html=view.summary({currency:'PLN',revision:1,settlement:{ready:false,pending_orders:1,packing_pending:0},
        groups:[{mode:'pending',label:'来源待核实',orders:1,quantity:1,revenue:100,pending_count:1,cost_recovery:0,partner_profit:0,our_receivable:0}]},false);
    assert.equal((html.match(/<span class="text-warning">待核实<\/span>/g)||[]).length,3);
});

test('per-unit fee shows quantity times price, transit parcel contents and outcome policy', () => {
    const settlement={ready:true,packing_basis:'shipped_unit',packing_units:3,packing_charges:1,
        packing_unit_fee_cny:10,packing_fee_cny:30,packing_policy:'失败、取消、退款、退回及补发均收费'};
    const html=view.summary({currency:'PLN',revision:2,settlement,groups:[]},false);
    assert.match(html,/3 支 × 10 元\/支 · 1 个包裹/);
    assert.match(html,/退款、退回及补发均收费/);
    const rows=view.renderRows({rows:[{order_number:'123',order_status:'refunded',billing_basis:'shipped_unit',
        quantity:3,unit_fee_cny:10,amount_cny:30,shipment_id:'id',tracking_number:'TRACK',
        items:[{product_name:'<unsafe>',quantity:3,line_item_id:'11',fulfillment_item_id:7}],pending:[]}]},'fees',true);
    assert.match(rows,/3 支 × 10\.00 CNY\/支/);
    assert.match(rows,/refunded/);
    assert.match(rows,/&lt;unsafe&gt;/);
    assert.ok(!rows.includes('data-reship-decision'));
});
