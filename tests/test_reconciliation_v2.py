"""Acceptance tests exercise the real OMS -> owned stock -> ledger -> cash chain."""
import json
from decimal import Decimal
import pytest
import test_fulfillment as fulfillment_fixture
from reconciliation_schema import migrate
from reconciliation_core import *
from reconciliation_inventory import bind_batch, return_source
from reconciliation_ledger import *
from fulfillment_service import plan_order, create_shipment, transition_fulfillment


@pytest.fixture
def setup():
    helper = fulfillment_fixture.FulfillmentDomainTests()
    helper.setUp()
    db = helper.db
    migrate(db)
    ours = save_object(db,'party','本团队',{'internal':True,'currency':'PLN'},1)
    partner = save_object(db,'party','合伙人',{'currency':'CNY'},1)
    carrier = save_object(db,'party','物流商',{'currency':'PLN'},1)
    pools=[]; contracts={}
    for n, capital in enumerate(({ours:'1'},{ours:'.6',partner:'.4'})):
        pool=save_object(db,'pool','货权'+str(n),{'ownership':capital,'capital':capital,'ownership_method':'physical'},1)
        pools.append(pool)
        profit={ours:'1'} if n==0 else {ours:'.5',partner:'.5'}
        contract=save_object(db,'contract','合同'+str(n),{'pool_id':pool,'template':'joint','profit':profit,'loss':profit,
            'shipping':{ours:'1'},'currency':'PLN','capital_status':'unpaid','recognition':'shipped',
            'effective_from':'2020-01-01','effective_to':'2099-01-01'},1)
        contracts[pool]=contract
        db.execute('INSERT INTO inv_batches(id,warehouse_id,sku_id,batch_no,qty_received,qty_remaining,unit_cost) VALUES (?,1,1,?,10,10,10)',(n+1,'LOT'+str(n)))
    for n,pool in enumerate(pools):
        bind_batch(db,n+1,pool,'10','PLN',1)
    fee=save_object(db,'fee','每单20',{'provider_id':partner,'amount':'20','currency':'CNY','basis':'order',
        'event':'shipped','return_policy':'keep','reship_policy':'free','effective_from':'2020-01-01','effective_to':'2099-01-01'},1)
    cfg={'our_party_id':ours,'currency':'PLN','policy':'fefo','contracts':contracts,'fees':[fee],
         'provider_mode':'primary','primary_provider':partner,'fx':{'CNY/PLN':'.5'}}
    def order(oid='1-100',qty=12,shipping='25',total=None):
        helper.add_order(oid,'PL',qty=qty,shipping_total=shipping,currency='PLN',line_total=qty*30,order_total=total)
        configure_order(db,oid,dict(cfg),1)
        plan=plan_order(db,oid)
        fulfillment=db.execute("SELECT id FROM oms_fulfillments WHERE order_id=? AND status<>'superseded'",(oid,)).fetchone()['id']
        shipment=create_shipment(db,fulfillment,'TRACK-'+oid)
        return plan,shipment
    def bill(shipment,reference='first',**kw):
        data={'provider_id':carrier,'currency':'PLN','outbound':'8','reverse':'0','collection_fee':'2',
            'collected':'385','remitted':'0','collected_at':'2026-09-21','remitted_at':'','reference':reference,'proof':'测试凭证'}
        data.update(kw)
        record_bill(db,shipment['id'],data,1)
    yield {'db':db,'helper':helper,'ours':ours,'partner':partner,'carrier':carrier,'pools':pools,'contracts':contracts,'cfg':cfg,'order':order,'bill':bill}
    helper.tearDown()


def test_joint_stock_fee_and_cash_chain(setup):
    d=setup; db=d['db']; _,shipment=d['order'](); d['bill'](shipment)
    stock=db.execute('SELECT on_hand,reserved FROM inv_stock WHERE warehouse_id=1 AND sku_id=1').fetchone()
    assert tuple(stock)==(8,0)
    assert len(rows(db,'SELECT * FROM rec_sources'))==2
    result=recognize(db,'1-100',1)
    assert not result['pending']
    totals={(r['party_id'],r['category'],r['currency']):dec(r['amount']) for r in result['totals']}
    assert totals[d['partner'],'service','CNY']==20
    assert totals[d['partner'],'capital','PLN']==8
    assert totals[d['ours'],'shipping_profit','PLN']==15
    count=len(rows(db,'SELECT * FROM rec_entries'))
    recognize(db,'1-100',1)
    assert len(rows(db,'SELECT * FROM rec_entries'))==count
    sid=create_statement(db,d['partner'],'CNY','2020-01-01','2099-01-01','trade',1)
    for status in ('verified','confirmed','locked'): statement_transition(db,sid,status,1)
    cash=record_cash(db,d['partner'],'CNY','20','2026-09-21','BANK1','proof',1)
    allocate_cash(db,cash,sid,'20',1)
    assert db.execute('SELECT status FROM rec_statements WHERE id=?',(sid,)).fetchone()['status']=='paid'


def test_replan_and_cancel_release_once(setup):
    d=setup; db=d['db']; d['helper'].add_order('1-1','PL',qty=12,currency='PLN')
    configure_order(db,'1-1',d['cfg'],1)
    first=plan_order(db,'1-1'); second=plan_order(db,'1-1')
    assert second['action']=='noop'
    assert db.execute('SELECT reserved FROM inv_stock WHERE warehouse_id=1 AND sku_id=1').fetchone()['reserved']==12
    f=db.execute('SELECT id FROM oms_fulfillments WHERE order_id=?',('1-1',)).fetchone()['id']
    transition_fulfillment(db,f,'cancelled'); transition_fulfillment(db,f,'cancelled')
    assert db.execute('SELECT reserved FROM inv_stock WHERE warehouse_id=1 AND sku_id=1').fetchone()['reserved']==0
    assert sum(r['reserved'] for r in rows(db,'SELECT reserved FROM rec_batches'))==0


def test_missing_cost_fx_bill_block(setup):
    d=setup; db=d['db']; _,shipment=d['order']()
    assert economics(db,'1-100')['pending']
    with pytest.raises(ReconciliationError): recognize(db,'1-100',1)
    d['bill'](shipment)
    source=rows(db,'SELECT * FROM rec_sources')[0]; snap=json.loads(source['snapshot']);snap['unit_cost']=None
    db.execute('UPDATE rec_sources SET snapshot=? WHERE id=?',(dump(snap),source['id']))
    assert any('成本缺失' in p for p in economics(db,'1-100')['pending'])


def test_free_shipping_and_negative_profit(setup):
    d=setup; db=d['db']; _,shipment=d['order'](qty=1,shipping='0')
    d['bill'](shipment,outbound='50',collection_fee='0',collected='30')
    result=recognize(db,'1-100',1)
    assert next(r['amount'] for r in result['totals'] if r['category']=='shipping_profit')=='-50.00'
    assert result['customer_shipping']=='0.0'


def test_duplicate_tracking_no_second_stock_deduction(setup):
    d=setup; db=d['db']; _,shipment=d['order']()
    f=db.execute('SELECT id FROM oms_fulfillments WHERE order_id=?',('1-100',)).fetchone()['id']
    again=create_shipment(db,f,'TRACK-1-100')
    assert again['id']==shipment['id']
    assert db.execute('SELECT on_hand FROM inv_stock WHERE warehouse_id=1 AND sku_id=1').fetchone()['on_hand']==8


def test_return_original_pool_and_locked_difference(setup):
    d=setup; db=d['db']; _,shipment=d['order']();d['bill'](shipment);recognize(db,'1-100',1)
    sid=create_statement(db,d['partner'],'PLN','2020-01-01','2099-01-01','trade',1)
    for status in ('verified','confirmed','locked'): statement_transition(db,sid,status,1)
    before=db.execute('SELECT snapshot FROM rec_statements WHERE id=?',(sid,)).fetchone()['snapshot']
    source=rows(db,'SELECT * FROM rec_sources WHERE batch_id=2')[0]
    return_source(db,source['id'],1,True,'可销售退货','RET1',1)
    return_source(db,source['id'],1,True,'可销售退货','RET1',1)
    assert db.execute('SELECT qty_remaining FROM inv_batches WHERE id=2').fetchone()['qty_remaining']==9
    recognize(db,'1-100',1)
    assert db.execute('SELECT snapshot FROM rec_statements WHERE id=?',(sid,)).fetchone()['snapshot']==before
    assert rows(db,'SELECT * FROM rec_entries WHERE party_id=? AND currency=? AND statement_id IS NULL',(d['partner'],'PLN'))


def test_paid_capital_not_payable(setup):
    d=setup; db=d['db']; _,shipment=d['order']();d['bill'](shipment)
    for source in rows(db,'SELECT * FROM rec_sources'):
        snap=json.loads(source['snapshot']);snap['contract']['data']['capital_status']='paid'
        db.execute('UPDATE rec_sources SET snapshot=? WHERE id=?',(dump(snap),source['id']))
    assert not any(r['category']=='capital' for r in economics(db,'1-100')['totals'])


def test_currency_and_party_cannot_cross_clear(setup):
    d=setup;db=d['db'];_,shipment=d['order']();d['bill'](shipment);recognize(db,'1-100',1)
    sid=create_statement(db,d['partner'],'CNY','2020-01-01','2099-01-01','trade',1)
    for status in ('verified','confirmed','locked'):statement_transition(db,sid,status,1)
    cash=record_cash(db,d['ours'],'PLN','20','2026-09-21','BANK1','proof',1)
    with pytest.raises(ReconciliationError):allocate_cash(db,cash,sid,20,1)


def test_amount_conservation_and_unknown():
    for total in ('10.01','-10.01','0'):
        assert sum(allocate(total,{'a':1,'b':1,'c':1}).values())==dec(total)
    with pytest.raises(ReconciliationError):dec(None)
    with pytest.raises(ReconciliationError):dec('NaN')

def test_financial_correction_changes_back_to_previous_value(setup):
    d=setup;db=d['db'];_,shipment=d['order']();d['bill'](shipment);first=recognize(db,'1-100',1)
    financial_correction(db,'1-100',{'fx':{'CNY/PLN':'1'}},'FX1','核实汇率',1)
    second=recognize(db,'1-100',1)
    financial_correction(db,'1-100',{'fx':{'CNY/PLN':'.5'}},'FX2','撤回错误汇率',1)
    third=recognize(db,'1-100',1)
    assert first['totals']==third['totals'] and first['totals']!=second['totals']
    current=defaultdict(Decimal)
    for r in rows(db,'SELECT * FROM rec_entries'):
        current[r['party_id'],r['category'],r['currency']]+=dec(r['amount'])
    assert current=={(r['party_id'],r['category'],r['currency']):dec(r['amount']) for r in third['totals']}


def test_site_scoped_order_identity_and_fee_versions(setup):
    d=setup;db=d['db'];_,s1=d['order']('1-77',2);d['bill'](s1,collected='85');r1=recognize(db,'1-77',1)
    old_fee=object_(db,d['cfg']['fees'][0],'fee')['data'];newfee=save_object(db,'fee','新价格',{**old_fee,'amount':'30'},1)
    d['cfg']['fees']=[newfee]
    d['helper'].add_order('3-77','CZ',qty=2,shipping_total='200',currency='CZK')
    # Source-market mapping is still mandatory, even when an owner has stock.
    db.execute("INSERT INTO inv_market_warehouses(market_code,warehouse_id,priority,is_active) VALUES ('CZ',1,1,1) ON CONFLICT DO NOTHING")
    cfg={**d['cfg'],'currency':'CZK','fx':{'CNY/CZK':'3','PLN/CZK':'6'}}
    configure_order(db,'3-77',cfg,1);plan_order(db,'3-77')
    f=db.execute("SELECT id FROM oms_fulfillments WHERE order_id='3-77'").fetchone()['id'];s2=create_shipment(db,f,'CZ-77')
    d['bill'](s2,currency='CZK',collected='220')
    r2=recognize(db,'3-77',1)
    assert next(r['amount'] for r in r1['totals'] if r['category']=='service')=='20.00'
    assert next(r['amount'] for r in r2['totals'] if r['category']=='service')=='30.00'
    assert economics(db,'1-77')['fees'][0]['amount']=='20.00'
    assert r1['site']!=r2['site']


@pytest.mark.parametrize('shipping',['2800','3200','0'])
def test_actual_customer_shipping_never_fixed_country_override(setup,shipping):
    d=setup;_,parcel=d['order'](qty=1,shipping=shipping);d['bill'](parcel,collected=str(dec(shipping)+30))
    result=economics(d['db'],'1-100')
    assert dec(result['customer_shipping'])==dec(shipping)
    assert next(dec(r['amount']) for r in result['totals'] if r['category']=='shipping_profit')==dec(shipping)-10


def test_zero_unknown_and_decimal_cent_conservation(setup):
    d=setup;db=d['db'];d['helper'].add_order('1-cent','PL',qty=12,line_total='100.01',currency='PLN')
    configure_order(db,'1-cent',d['cfg'],1);plan_order(db,'1-cent')
    f=db.execute("SELECT id FROM oms_fulfillments WHERE order_id='1-cent'").fetchone()['id'];parcel=create_shipment(db,f,'CENT')
    d['bill'](parcel,collected='100.01',outbound='0',collection_fee='0')
    result=recognize(db,'1-cent',1)
    assert sum(dec(s['revenue']) for s in result['sources'])==Decimal('100.01')
    assert not result['pending']


def test_legacy_receipts_untouched(setup):
    d=setup;db=d['db'];db.execute('CREATE TABLE partner_receipts(id INTEGER PRIMARY KEY,partner_id INTEGER,amount TEXT)')
    db.execute("INSERT INTO partner_receipts VALUES (1,1,'999.99')")
    _,p=d['order']();d['bill'](p);recognize(db,'1-100',1)
    assert tuple(db.execute('SELECT * FROM partner_receipts').fetchone())==(1,1,'999.99')


def test_unknown_quote_never_wins_cost_route(setup):
    d=setup;db=d['db'];d['helper'].add_order('1-cost','PL',qty=1,currency='PLN')
    configure_order(db,'1-cost',{**d['cfg'],'policy':'cost','quotes':{'1':{'goods':'10','service':'0','outbound':'0'}}},1)
    result=plan_order(db,'1-cost')
    assert result['shortages'] and not result['assignments']
    assert db.execute('SELECT reserved FROM inv_stock WHERE warehouse_id=1 AND sku_id=1').fetchone()['reserved']==0


def test_oversell_and_guard_old_processor(setup):
    from inv_common import record_movement
    d=setup;db=d['db'];d['helper'].add_order('1-full','PL',qty=21,currency='PLN')
    configure_order(db,'1-full',d['cfg'],1);p=plan_order(db,'1-full')
    assert sum(x['qty'] for x in p['shortages'])==1
    assert db.execute('SELECT reserved FROM inv_stock WHERE warehouse_id=1 AND sku_id=1').fetchone()['reserved']==20
    with pytest.raises(ReconciliationError):record_movement(db,warehouse_id=1,sku_id=1,movement_type='sale_out',qty_delta=-1,ref_type='legacy')


def test_physical_transfer_preserves_owner_and_total(setup):
    from reconciliation_inventory import move_batch
    d=setup;db=d['db'];bid=move_batch(db,1,2,3,'TR1','目标仓已验收',1)
    assert db.execute('SELECT pool_id FROM rec_batches WHERE batch_id=?',(bid,)).fetchone()['pool_id']==d['pools'][0]
    assert sum(r['qty_remaining'] for r in rows(db,'SELECT qty_remaining FROM inv_batches'))==20
    assert db.execute('SELECT on_hand FROM inv_stock WHERE warehouse_id=1 AND sku_id=1').fetchone()['on_hand']==17
    assert db.execute('SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1').fetchone()['on_hand']==3
    move_batch(db,1,2,3,'TR1','目标仓已验收',1)
    assert sum(r['qty_remaining'] for r in rows(db,'SELECT qty_remaining FROM inv_batches'))==20


def test_partial_cash_and_zero_statement(setup):
    d=setup;db=d['db'];adjustment(db,d['partner'],20,'CNY','service','A1','费用',1)
    sid=create_statement(db,d['partner'],'CNY','2020-01-01','2099-01-01','trade',1)
    for status in ('verified','confirmed','locked'):statement_transition(db,sid,status,1)
    cash=record_cash(db,d['partner'],'CNY',20,'2026-09-21','BANK1','proof',1)
    allocate_cash(db,cash,sid,5,1);allocate_cash(db,cash,sid,15,1)
    assert db.execute('SELECT status FROM rec_statements WHERE id=?',(sid,)).fetchone()['status']=='paid'
    adjustment(db,d['partner'],0,'CNY','profit','ZERO','零活动',1)
    zero=create_statement(db,d['partner'],'CNY','2020-01-01','2099-01-01','trade',1)
    for status in ('verified','confirmed','locked'):statement_transition(db,zero,status,1)
    assert db.execute('SELECT status FROM rec_statements WHERE id=?',(zero,)).fetchone()['status']=='paid'


def test_partner_api_and_export_never_leak_other_pool_costs(setup,monkeypatch):
    from flask import Flask
    from flask_login import LoginManager,UserMixin
    import reconciliation_api
    d=setup;db=d['db'];_,parcel=d['order']();d['bill'](parcel);recognize(db,'1-100',1)
    sid=create_statement(db,d['partner'],'PLN','2020-01-01','2099-01-01','trade',1)
    ours=create_statement(db,d['ours'],'PLN','2020-01-01','2099-01-01','trade',1)
    db.execute('INSERT INTO rec_user_parties(user_id,party_id) VALUES (2,?)',(d['partner'],));db.commit()
    class NonClosing:
        def __getattr__(self,k):return getattr(db,k)
        def close(self):pass
    class Partner(UserMixin):
        id=2;username='partner'
        def can_view_reconciliation(self):return True
        def can_view_costs(self):return True # even a mistakenly enabled flag cannot bypass party scope
        def can_edit_reconciliation(self):return True
        def get_accessible_partner_ids(self):return None
    app=Flask(__name__);app.secret_key='test';app.register_blueprint(reconciliation_api.bp)
    manager=LoginManager(app);manager.user_loader(lambda _:Partner())
    monkeypatch.setattr(reconciliation_api,'get_conn',lambda:NonClosing())
    client=app.test_client()
    with client.session_transaction() as s:s['_user_id']='2';s['rec_csrf']='token'
    response=client.get('/api/reconciliation-v2/state');assert response.status_code==200
    payload=response.get_json();assert not payload['can_edit'] and 'pools' not in payload
    assert len(payload['parties'])==1
    text=response.get_data(as_text=True)
    assert 'source_totals' not in text and 'unit_cost' not in text
    assert client.get('/api/reconciliation-v2/export/'+ours).status_code==403
    assert client.get('/api/reconciliation-v2/export/'+sid).status_code==200
    assert client.get('/api/reconciliation-v2/order/1-100').status_code==403
    assert client.post('/api/reconciliation-v2/config',json={},headers={'X-Rec-CSRF':'token'}).status_code==403

@pytest.mark.parametrize('basis,expected',[('order','20.00'),('parcel','40.00')])
def test_two_parcels_mixed_owners_service_fee_basis(setup,basis,expected):
    d=setup;db=d['db'];db.execute('UPDATE inv_stock SET on_hand=1 WHERE warehouse_id=1 AND sku_id=1')
    db.execute('UPDATE inv_batches SET qty_remaining=CASE WHEN id=1 THEN 1 ELSE 0 END')
    db.execute('INSERT INTO inv_stock(warehouse_id,sku_id,on_hand,reserved) VALUES (2,1,1,0)')
    db.execute("UPDATE oms_warehouse_integrations SET provider='internal',inventory_authority='local' WHERE warehouse_id=2")
    db.execute("INSERT INTO inv_batches(id,warehouse_id,sku_id,batch_no,qty_received,qty_remaining) VALUES (3,2,1,'HU',1,1)")
    bind_batch(db,3,d['pools'][1],10,'PLN',1)
    fee=object_(db,d['cfg']['fees'][0],'fee')['data']
    new=save_object(db,'fee','计费模式',{**fee,'basis':basis},1)
    d['helper'].add_order('1-split','PL',qty=2,shipping_total=25,currency='PLN',line_total=60)
    configure_order(db,'1-split',{**d['cfg'],'fees':[new]},1);plan=plan_order(db,'1-split')
    assert len(plan['fulfillment_ids'])==2
    for n,fid in enumerate(plan['fulfillment_ids']):
        p=create_shipment(db,fid,'SPLIT'+str(n));d['bill'](p,collected='42.5',outbound='0',collection_fee='0')
    result=recognize(db,'1-split',1)
    assert next(r['amount'] for r in result['totals'] if r['category']=='service')==expected
    assert sum(dec(b['collected']) for b in result['bills'])==85


def test_delivery_fee_and_collection_are_separate_dates(setup):
    d=setup;db=d['db'];fee=object_(db,d['cfg']['fees'][0],'fee')['data']
    d['cfg']['fees']=[save_object(db,'fee','签收付费',{**fee,'event':'delivered'},1)]
    _,parcel=d['order'](qty=1);d['bill'](parcel,collected='55',remitted='0')
    before=recognize(db,'1-100',1)
    assert not any(r['category']=='service' for r in before['totals'])
    db.execute("UPDATE oms_shipments SET status='delivered',delivered_at='2026-09-22' WHERE id=?",(parcel['id'],))
    after=recognize(db,'1-100',1)
    assert next(r['amount'] for r in after['totals'] if r['category']=='service')=='20.00'
    assert not any(r['category']=='collection_remitted' for r in after['totals'])


def test_split_provider_contract_requires_consistent_total(setup):
    d=setup;db=d['db'];fee=object_(db,d['cfg']['fees'][0],'fee')['data']
    other=save_object(db,'fee','另一服务商',{**fee,'provider_id':d['carrier']},1)
    d['cfg'].update(provider_mode='split',provider_weights={d['partner']:'.5',d['carrier']:'.5'},fees=d['cfg']['fees']+[other])
    _,parcel=d['order'](qty=1);d['bill'](parcel,collected='55')
    result=recognize(db,'1-100',1)
    amounts=[dec(r['amount']) for r in result['totals'] if r['category']=='service']
    assert amounts==[Decimal('10'),Decimal('10')]


def test_shipping_tax_and_cod_fee_not_counted_twice(setup):
    d=setup;db=d['db'];db.execute('ALTER TABLE orders ADD COLUMN shipping_tax TEXT');db.execute('ALTER TABLE orders ADD COLUMN fee_lines TEXT')
    d['cfg']['customer_shipping_fee_ids']=['9']
    d['helper'].add_order('1-tax','PL',qty=1,line_total='30',line_tax='3',shipping_total='25',currency='PLN',order_total='65.5')
    db.execute("UPDATE orders SET shipping_tax='2.5',fee_lines=? WHERE id='1-tax'",(dump([{'id':9,'total':'5','total_tax':'0'}]),))
    configure_order(db,'1-tax',d['cfg'],1);plan=plan_order(db,'1-tax');p=create_shipment(db,plan['fulfillment_ids'][0],'TAX')
    d['bill'](p,collected='65.5')
    result=recognize(db,'1-tax',1)
    assert next(r['amount'] for r in result['totals'] if r['category']=='shipping_profit')=='20.00'
    assert result['expected_collection']=='65.50'
    assert result['collection_variance']=='0.00'

def test_return_after_ownership_transfer_restores_original_owner(setup):
    from reconciliation_inventory import transfer_ownership
    d=setup;db=d['db'];_,parcel=d['order'](qty=1)
    source=rows(db,'SELECT * FROM rec_sources')[0]
    transfer_ownership(db,1,d['pools'][1],'剩余存货转让',1)
    return_source(db,source['id'],1,True,'原订单退回','OWNER-RET',1)
    last=rows(db,'SELECT * FROM rec_batches ORDER BY batch_id')[-1]
    assert last['pool_id']==d['pools'][0] and last['batch_id']!=1
    assert sum(r['qty_remaining'] for r in rows(db,'SELECT qty_remaining FROM inv_batches'))==20


def test_authorized_offset_and_timezone_cutoff(setup):
    d=setup;db=d['db'];adjustment(db,d['carrier'],10,'PLN','carrier','EXP','运费',1)
    adjustment(db,d['carrier'],-10,'PLN','collection_due','COD','未回款',1)
    ids=[]
    for bucket in ('trade','collection'):
        sid=create_statement(db,d['carrier'],'PLN','2026-01-01','2027-01-01',bucket,1,'Europe/Warsaw')
        snap=json.loads(db.execute('SELECT snapshot FROM rec_statements WHERE id=?',(sid,)).fetchone()['snapshot'])
        assert snap['utc_start'].startswith('2025-12-31T23:00:00')
        for status in ('verified','confirmed','locked'):statement_transition(db,sid,status,1)
        ids.append(sid)
    authorized_offset(db,*ids,10,'OFFSET1','双方已确认的抵扣凭证',1)
    assert all(r['status']=='paid' for r in rows(db,'SELECT status FROM rec_statements'))
    assert all(dec(r['outstanding'])==0 for r in balances(db))


def test_pending_order_blocks_period_lock(setup):
    d=setup;db=d['db'];d['order'](qty=1)
    adjustment(db,d['partner'],20,'CNY','service','OLD','已知费用',1)
    sid=create_statement(db,d['partner'],'CNY','2020-01-01','2099-01-01','trade',1)
    with pytest.raises(ReconciliationError,match='待补全'):statement_transition(db,sid,'verified',1)

def test_extra_expenses_are_allocated_once_and_cash_retry_is_safe(setup):
    d=setup;db=d['db'];_,p=d['order'](qty=1);d['bill'](p,collected='55')
    financial_correction(db,'1-100',{'expenses':{'PACK':{'kind':'goods','party_id':d['carrier'],'amount':'3','currency':'PLN','proof':'包装费'}},'operating_costs_complete':True},'EX1','已核实',1)
    result=recognize(db,'1-100',1)
    assert result['profit_type']=='operating_confirmed'
    assert sum(dec(s['other_cost']) for s in result['sources'])==3
    assert next(r['amount'] for r in result['totals'] if r['category']=='other_expense')=='3.00'
    sid=create_statement(db,d['partner'],'CNY','2020-01-01','2099-01-01','trade',1)
    for status in ('verified','confirmed','locked'):statement_transition(db,sid,status,1)
    cash=record_cash(db,d['partner'],'CNY',20,'2026-09-21','IDEM-BANK','proof',1)
    assert record_cash(db,d['partner'],'CNY',20,'2026-09-21','IDEM-BANK','proof',1)==cash
    allocate_cash(db,cash,sid,20,1,'REQ1');allocate_cash(db,cash,sid,20,1,'REQ1')
    assert dec(rows(db,'SELECT amount FROM rec_cash_allocations')[0]['amount'])==20

def test_site_profile_automates_new_orders_without_rewriting_legacy(setup):
    d=setup;db=d['db'];d['helper'].add_order('1-auto','PL',qty=1,currency='PLN')
    save_object(db,'profile','试点站点',{'site':'https://pl.test','starts_on':'2020-01-01','enabled':True,'config':dict(d['cfg'])},1)
    result=plan_order(db,'1-auto')
    assert not result['shortages'] and order_config(db,'1-auto')['site_profile_id']
    saved=order_config(db,'1-auto')
    save_object(db,'profile','暂停新单',{'site':'https://pl.test','starts_on':'2020-01-01','enabled':False,'config':dict(d['cfg'])},1)
    assert order_config(db,'1-auto')==saved
    d['helper'].add_order('1-paused','PL',qty=1,currency='PLN')
    from reconciliation_core import apply_site_profile
    apply_site_profile(db,'1-paused')
    assert order_config(db,'1-paused') is None

def test_first_recognition_uses_business_dates_later_changes_are_adjustments(setup):
    d=setup;db=d['db'];_,p=d['order'](qty=1)
    db.execute("UPDATE oms_shipments SET shipped_at='2026-08-31T20:00:00+00:00' WHERE id=?",(p['id'],))
    d['bill'](p,collected='55',collected_at='2026-09-01',remitted='45',remitted_at='2026-09-03')
    recognize(db,'1-100',1)
    service=rows(db,"SELECT * FROM rec_entries WHERE category='service'")[0]
    assert service['occurred_at'].startswith('2026-08-31')
    due=rows(db,"SELECT * FROM rec_entries WHERE category='collection_due'")[0]
    remitted=rows(db,"SELECT * FROM rec_entries WHERE category='collection_remitted'")[0]
    assert due['occurred_at'].startswith('2026-09-01') and remitted['occurred_at'].startswith('2026-09-03')
    sid=create_statement(db,d['partner'],'CNY','2026-08-01','2026-09-01','trade',1)
    for status in ('verified','confirmed','locked'):statement_transition(db,sid,status,1)
    financial_correction(db,'1-100',{'expenses':{'X':{'kind':'goods','party_id':d['carrier'],'amount':'5','currency':'PLN','proof':'晚到成本'}}},'X1','补记成本',1)
    recognize(db,'1-100',1)
    assert db.execute('SELECT status FROM rec_statements WHERE id=?',(sid,)).fetchone()['status']=='locked'

def test_changing_configuration_requires_replan_before_shipping(setup):
    d=setup;db=d['db'];d['helper'].add_order('1-change','PL',qty=1,currency='PLN')
    configure_order(db,'1-change',d['cfg'],1);planned=plan_order(db,'1-change')
    changed={**d['cfg'],'policy':'own_first'}
    configure_order(db,'1-change',changed,1);db.commit()
    with pytest.raises(ReconciliationError,match='重新分仓'):
        create_shipment(db,planned['fulfillment_ids'][0],'STALE',commit=False)
    db.rollback()
    replacement=plan_order(db,'1-change')
    create_shipment(db,replacement['fulfillment_ids'][0],'FRESH')
    assert db.execute('SELECT on_hand FROM inv_stock WHERE warehouse_id=1 AND sku_id=1').fetchone()['on_hand']==19

def test_already_net_refund_is_not_deducted_again(setup):
    d=setup;db=d['db'];_,p=d['order'](qty=1);d['bill'](p,collected='55')
    financial_correction(db,'1-100',{'refunds':{'501':'5'},'refunds_in_order_totals':True},'NET1','订单金额已经是退款后净额',1)
    result=recognize(db,'1-100',1)
    assert sum(dec(s['revenue']) for s in result['sources'])==30
    financial_correction(db,'1-100',{'refunds_in_order_totals':False},'NET2','核实原金额尚未扣退款',1)
    result=recognize(db,'1-100',1)
    assert sum(dec(s['revenue']) for s in result['sources'])==25

def test_legacy_return_cannot_reassign_owned_stock(setup):
    from inv_batches import restock_batch
    from inv_common import record_movement
    d=setup;db=d['db']
    with pytest.raises(ReconciliationError):restock_batch(db,1,1)
    with pytest.raises(ReconciliationError):record_movement(db,warehouse_id=1,sku_id=1,movement_type='return_in',qty_delta=1,ref_type='legacy_return')
    assert db.execute('SELECT qty_remaining FROM inv_batches WHERE id=1').fetchone()['qty_remaining']==10
