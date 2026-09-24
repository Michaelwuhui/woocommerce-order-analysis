import copy
import json
import sqlite3
from decimal import Decimal

import pytest
from reconciliation_schema import migrate
from reconciliation_core import save_object, object_, ReconciliationError
from reconciliation_history import save_rule, preview, create_draft, rate_at, _calculate


@pytest.fixture
def db():
    c=sqlite3.connect(':memory:');c.row_factory=sqlite3.Row
    c.executescript('''
    CREATE TABLE warehouses(id INTEGER PRIMARY KEY,name TEXT);
    INSERT INTO warehouses VALUES(10,'中转仓');
    CREATE TABLE exchange_rates(year_month TEXT,currency TEXT,rate_to_cny TEXT);
    INSERT INTO exchange_rates VALUES('2026-08','PLN','1.81988');
    CREATE TABLE brands(id INTEGER PRIMARY KEY,name TEXT,aliases TEXT);
    INSERT INTO brands VALUES(9,'FUMOT','[]');
    CREATE TABLE series(id INTEGER,brand_id INTEGER,name TEXT);
    CREATE TABLE product_mappings(raw_name TEXT,source TEXT,brand_id INTEGER,series_id INTEGER,puff_count INTEGER,flavor TEXT);
    CREATE TABLE product_costs(id INTEGER PRIMARY KEY,warehouse_id INTEGER,brand_id INTEGER,series_id INTEGER,
      puff_count INTEGER,flavor TEXT,cost_price TEXT,cost_currency TEXT,effective_date TEXT);
    CREATE TABLE orders(id TEXT PRIMARY KEY,number TEXT,source TEXT,status TEXT,currency TEXT,
      is_undelivered INTEGER,is_problem_return INTEGER,line_items TEXT,fee_lines TEXT,refunds TEXT,
      total TEXT,shipping_total TEXT,total_tax TEXT);
    CREATE TABLE oms_fulfillments(id TEXT,order_id TEXT,warehouse_id INTEGER);
    CREATE TABLE oms_shipments(id TEXT,fulfillment_id TEXT,tracking_number TEXT,shipped_at TEXT,status TEXT);
    CREATE TABLE shipping_logs(order_id TEXT,tracking_number TEXT,shipped_at TEXT,status TEXT,is_reship INTEGER,reship_reason TEXT);
    CREATE TABLE oms_order_items(id INTEGER,woo_line_item_id TEXT);
    CREATE TABLE oms_fulfillment_items(id INTEGER,order_item_id INTEGER,fulfillment_id TEXT);
    CREATE TABLE oms_shipment_items(shipment_id TEXT,fulfillment_item_id INTEGER,quantity INTEGER);
    ''')
    migrate(c)
    team=save_object(c,'party','团队',{'internal':True,'currency':'PLN'},1)
    supplier=save_object(c,'party','供货方',{'currency':'AUD'},1)
    provider=save_object(c,'party','发货方',{'currency':'PLN'},1)
    pool=save_object(c,'pool','供货方独资',{'ownership':{supplier:'1'},'capital':{supplier:'1'},'ownership_method':'virtual'},1)
    contract=save_object(c,'contract','货值另账',{'pool_id':pool,'template':'margin','profit':{team:'1'},'loss':{team:'1'},'shipping':{team:'1'},'currency':'PLN','capital_status':'none','recognition':'shipped','effective_from':'2026-08-01','effective_to':'2099-01-01'},1)
    rule=save_rule(c,'核算规则',{'warehouse_id':10,'pool_id':pool,'contract_id':contract,'our_party_id':team,'supplier_id':supplier,'provider_id':provider,'effective_from':'2026-08-01','effective_to':'2099-01-01','management_cny':'20','freight':{'PLN':{'outbound':'25','reverse':'0'},'CZK':{'outbound':'200','reverse':'200'},'HUF':{'outbound':'2800','reverse':'2800'}}},1)
    def add(oid='1-1',returned=False):
        c.execute('INSERT INTO orders VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)',(oid,oid,'https://example.test','on-hold','PLN',int(returned),0,json.dumps([{'id':1,'quantity':2,'total':'100'}]),'[]','[]','125','25','0'))
        c.execute('INSERT INTO oms_fulfillments VALUES(?,?,10)',(oid,oid))
        c.execute('INSERT INTO oms_shipments VALUES(?,?,?,?,?)',(oid,oid,oid,'2026-08-10 01:00:00','returned' if returned else 'shipped'))
        c.execute('INSERT INTO shipping_logs VALUES(?,?,?,?,0,NULL)',(oid,oid,'2026-08-10','shipped'))
        n=c.execute('SELECT COUNT(*) FROM oms_order_items').fetchone()[0]+1
        c.execute('INSERT INTO oms_order_items VALUES(?,?)',(n,'1'))
        c.execute('INSERT INTO oms_fulfillment_items VALUES(?,?,?)',(n,n,oid))
        c.execute('INSERT INTO oms_shipment_items VALUES(?,?,2)',(oid,n))
    yield c,rule,add
    c.close()


def test_shipped_unpaid_revenue_return_zero_and_separate_cost(db):
    c,rule,add=db;add();add('1-2',True)
    before=c.total_changes
    p=preview(c,rule,'2026-08','PLN')
    assert c.total_changes==before
    assert p['totals']['revenue']=='125.00'
    assert p['totals']['revenue_cny']=='227.49'
    assert p['totals']['freight']=='50.00'
    assert p['totals']['management_cny']=='40.00'
    assert p['totals']['shipping_net']=='-25.00'
    assert p['supplier_statement']['goods_value'] is None
    assert p['rows'][0]['product_profit'] is None
    assert p['rows'][1]['product_profit']=='0.00'  # Returned goods are not billed to the team.
    assert p['can_lock'] is False and p['counts']['pending_orders']==0


def test_dated_warehouse_costs_show_real_partial_coverage_without_inventing_profit(db):
    c,rule,add=db;add()
    c.execute("UPDATE orders SET line_items=?", (json.dumps([
        {'id':1,'name':'Fumot Tornado 9000 Puffs - Grape','quantity':1,'total':'50'},
        {'id':2,'name':'Fumot Leopard 40000 Puffs - Mint','quantity':1,'total':'50'}]),))
    c.execute("INSERT INTO oms_order_items VALUES(2,'2')")
    c.execute("INSERT INTO oms_fulfillment_items VALUES(2,2,'1-1')")
    c.execute("UPDATE oms_shipment_items SET quantity=1")
    c.execute("INSERT INTO oms_shipment_items VALUES('1-1',2,1)")
    c.execute("INSERT INTO product_costs VALUES(1,10,9,NULL,9000,NULL,'39','CNY','2026-07-01')")
    c.execute("INSERT INTO product_costs VALUES(2,1,9,NULL,40000,NULL,'10','CNY','2026-07-01')")
    p=preview(c,rule,'2026-08','PLN');r=p['rows'][0]
    assert p['counts']['cost_matched_quantity']=='1' and p['counts']['cost_missing_quantity']=='1'
    assert p['counts']['cost_pending_orders']==1
    assert r['known_goods_value_cny']=='39.00' and r['supplier_goods_value'] is None
    assert r['product_profit'] is None and p['totals']['supplier_goods_value'] is None
    assert r['cost_lines'][0]['cost_id']==1 and r['cost_lines'][0]['cost_effective_date']=='2026-07-01'
    assert '40000' in r['cost_missing'][0]
    c.execute("INSERT INTO product_costs VALUES(3,10,9,NULL,40000,NULL,'50','CNY','2026-08-11')")
    assert preview(c,rule,'2026-08','PLN')['counts']['cost_missing_quantity']=='1'  # No future price.
    c.execute("UPDATE product_costs SET effective_date='2026-07-01' WHERE id=3")
    complete=preview(c,rule,'2026-08','PLN');r=complete['rows'][0]
    assert complete['counts']['cost_pending_orders']==0
    assert r['supplier_goods_value_cny']=='89.00'
    assert r['supplier_goods_value']=='48.90'
    assert r['product_profit']=='51.10'
    assert r['contribution_profit']=='40.11'


def test_return_does_not_bill_goods_and_saved_missing_cost_stays_immutable(db):
    c,rule,add=db;add();add('1-2',True)
    c.execute("UPDATE orders SET line_items=?",(json.dumps([{'id':1,'name':'Fumot 9000 Puffs','quantity':2,'total':'100'}]),))
    before=preview(c,rule,'2026-08','PLN')
    draft=create_draft(c,{'rule_id':rule,'month':'2026-08','currency':'PLN',
                          'request_key':'missing-cost-snapshot-001','expected_digest':before['digest']},1)
    c.execute("INSERT INTO product_costs VALUES(1,10,9,NULL,9000,NULL,'39','CNY','2026-07-01')")
    after=preview(c,rule,'2026-08','PLN')
    assert before['counts']['cost_pending_orders']==1 and after['counts']['cost_pending_orders']==0
    assert after['totals']['supplier_goods_value_cny']=='78.00'
    assert after['rows'][1]['supplier_goods_value_cny']=='0.00'
    assert object_(c,draft,'history_draft')['data']['snapshot']['totals']['supplier_goods_value'] is None
    with pytest.raises(ReconciliationError,match='变化'):
        create_draft(c,{'rule_id':rule,'month':'2026-08','currency':'PLN',
                        'request_key':'fresh-after-cost-0001','expected_digest':before['digest']},1)


def test_market_currency_cost_keeps_pln_and_cny_values_separate(db):
    c,rule,add=db;add()
    c.execute("UPDATE orders SET line_items=?",(json.dumps([
        {'id':1,'name':'Fumot 15000 Puffs','quantity':2,'total':'100'}]),))
    c.execute("INSERT INTO product_costs VALUES(1,10,9,NULL,15000,NULL,'19','PLN','2026-07-01')")
    snap=preview(c,rule,'2026-08','PLN')
    row=snap['rows'][0]
    assert row['supplier_goods_value']=='38.00'
    assert row['supplier_goods_value_cny']=='69.16'
    assert row['cost_lines'][0]['cost_currency']=='PLN'
    assert row['cost_lines'][0]['cost_fx_month']=='2026-08'
    assert row['product_profit']=='62.00'


def test_multiple_pln_cost_lines_do_not_gain_a_cent_from_round_trip_fx(db):
    c,rule,add=db;add()
    c.execute("UPDATE orders SET line_items=?",(json.dumps([
        {'id':1,'name':'Fumot 15000 Puffs','quantity':2,'total':'30'},
        {'id':2,'name':'Fumot 40000 Puffs','quantity':2,'total':'30'},
        {'id':3,'name':'Fumot 80000 Puffs','quantity':2,'total':'40'}]),))
    for ident in (2,3):
        c.execute('INSERT INTO oms_order_items VALUES(?,?)',(ident,str(ident)))
        c.execute("INSERT INTO oms_fulfillment_items VALUES(?,?,'1-1')",(ident,ident))
        c.execute("INSERT INTO oms_shipment_items VALUES('1-1',?,2)",(ident,))
    for ident,puffs,price in ((1,15000,'19'),(2,40000,'46'),(3,80000,'57')):
        c.execute('INSERT INTO product_costs VALUES(?,10,9,NULL,?,NULL,?,\'PLN\',\'2026-08-01\')',
                  (ident,puffs,price))
    row=preview(c,rule,'2026-08','PLN')['rows'][0]
    assert row['supplier_goods_value']=='244.00'
    assert row['supplier_goods_value_cny']=='444.05'
    assert row['product_profit']=='-144.00'


def test_fx_backward_only_and_exact_decimal(db):
    c,_,_=db
    assert rate_at(c,'PLN','2026-09')['rate']=='1.81988'
    assert rate_at(c,'PLN','2026-09')['fallback'] is True
    assert rate_at(c,'PLN','2026-07') is None
    c.execute("INSERT INTO exchange_rates VALUES('2026-09','PLN','1.7')")
    assert rate_at(c,'PLN','2026-09')['rate']=='1.7'


def test_draft_idempotent_snapshot_and_no_financial_postings(db):
    c,rule,add=db;add()
    p=preview(c,rule,'2026-08','PLN')
    request={'rule_id':rule,'month':'2026-08','currency':'PLN','request_key':'test-request-key-0001','expected_digest':p['digest']}
    ident=create_draft(c,request,1)
    c.execute("UPDATE orders SET is_undelivered=1")
    assert create_draft(c,request,1)==ident
    assert object_(c,ident,'history_draft')['data']['snapshot']['totals']['revenue']=='125.00'
    for t in ('rec_entries','rec_statements','rec_cash','rec_batches','rec_sources'):
        assert c.execute('SELECT COUNT(*) FROM '+t).fetchone()[0]==0
    with pytest.raises(ReconciliationError,match='变化'):
        create_draft(c,{**request,'request_key':'test-request-key-0002'},1)


def test_partial_shipments_never_claim_whole_order_revenue(db):
    c,rule,add=db;add()
    c.execute('UPDATE oms_shipment_items SET quantity=1')
    p=preview(c,rule,'2026-08','PLN')
    assert p['rows'][0]['revenue'] is None
    assert p['counts']['pending_orders']==1


def test_split_parcels_charge_once(db):
    c,rule,add=db;add()
    c.execute('UPDATE oms_shipment_items SET quantity=1')
    c.execute("INSERT INTO oms_shipments VALUES('split','1-1','split','2026-08-11','shipped')")
    c.execute("INSERT INTO shipping_logs VALUES('1-1','split','2026-08-11','shipped',0,NULL)")
    c.execute("INSERT INTO oms_shipment_items VALUES('split',1,1)")
    p=preview(c,rule,'2026-08','PLN')
    assert p['counts']['pending_orders']==0
    assert p['totals']['management_cny']=='20.00' and p['totals']['freight']=='25.00'


@pytest.mark.parametrize('reason,fee,pending',[('ordinary','40.00',0),('warehouse_omission','20.00',0),('unknown','0.00',1)])
def test_reship_reason_and_original_order_fee(db,reason,fee,pending):
    c,rule,add=db;add()
    c.execute("INSERT INTO oms_shipments VALUES('reship','1-1','reship','2026-08-11','shipped')")
    c.execute("INSERT INTO shipping_logs VALUES('1-1','reship','2026-08-11','shipped',1,'free text is not authorization')")
    c.execute("INSERT INTO oms_shipment_items VALUES('reship',1,2)")
    p=preview(c,rule,'2026-08','PLN',{'reship':{'reason':reason,'evidence':'工单记录'}})
    assert p['counts']['pending_orders']==pending
    assert p['totals']['management_cny']==fee
    if not pending: assert p['totals']['freight']=='25.00'


def test_cross_month_reship_no_repeat_revenue_or_freight(db):
    c,rule,add=db;add()
    c.execute("INSERT INTO oms_shipments VALUES('reship','1-1','reship','2026-09-11','shipped')")
    c.execute("INSERT INTO shipping_logs VALUES('1-1','reship','2026-09-11','shipped',1,'ordinary')")
    c.execute("INSERT INTO oms_shipment_items VALUES('reship',1,2)")
    p=preview(c,rule,'2026-09','PLN',{'reship':{'reason':'ordinary','evidence':'工单'}})
    assert p['totals']['revenue']=='0.00' and p['totals']['freight']=='0.00'
    assert p['totals']['management_cny']=='20.00'


@pytest.mark.parametrize('unit,total',[('PLN','25.00'),('CZK','400.00'),('HUF','5600.00')])
def test_return_freight_by_market(db,unit,total):
    c,rule,add=db;add(returned=True)
    c.execute('UPDATE orders SET currency=?',(unit,))
    if unit!='PLN': c.execute('INSERT INTO exchange_rates VALUES(?,?,?)',('2026-08',unit,'.02'))
    p=preview(c,rule,'2026-08',unit)
    assert p['totals']['freight']==total and p['totals']['revenue']=='0.00'


def test_unmatched_scope_and_missing_rate(db):
    c,rule,add=db;add()
    with pytest.raises(ReconciliationError): preview(c,rule,'2026-08','PLN',{'foreign':{'reason':'ordinary','evidence':'x'}})
    with pytest.raises(ReconciliationError): preview(c,rule,'2026-07','PLN')
    c.execute('DELETE FROM exchange_rates')
    p=preview(c,rule,'2026-08','PLN')
    assert p['rows'][0]['revenue_cny'] is None and p['counts']['pending_orders']==1


def test_permission_and_csrf(db):
    import reconciliation_api as api_module
    from flask import Flask
    from flask_login import LoginManager, UserMixin
    c,rule,add=db;add();c.commit()
    app=Flask(__name__);app.secret_key='isolated-test';app.register_blueprint(api_module.bp)
    lm=LoginManager(app)
    class User(UserMixin):
        id='1';username='outsider'
        def can_view_reconciliation(self):return True
        def can_view_costs(self):return True
        def can_edit_reconciliation(self):return True
        def get_accessible_partner_ids(self):return []
    @lm.user_loader
    def load(_):return User()
    class Connection:
        def __getattr__(self,key):return getattr(c,key)
        def close(self):pass
    old=api_module.get_conn;api_module.get_conn=lambda:Connection()
    try:
        client=app.test_client()
        with client.session_transaction() as s:s['_user_id']='1';s['rec_csrf']='test'
        for route in ('state','draft/nope','export/nope'):
            assert client.get('/api/reconciliation-v2/history/'+route).status_code==403
        assert client.post('/api/reconciliation-v2/history/preview',json={}).status_code==403
        User.username='admin'
        assert client.post('/api/reconciliation-v2/history/preview',json={}).status_code==403
        p=preview(c,rule,'2026-08','PLN')
        ident=create_draft(c,{'rule_id':rule,'month':'2026-08','currency':'PLN','request_key':'admin-export-0000001','expected_digest':p['digest']},1)
        response=client.get('/api/reconciliation-v2/history/export/'+ident)
        assert response.status_code==200
        assert 'attachment;' in response.headers['Content-Disposition']
        assert '227.49' in response.data.decode('utf-8-sig')
        assert '待成本' in response.data.decode('utf-8-sig')
        old=object_(c,ident,'history_draft')['data']
        old['snapshot']['version']=1
        for row in old['snapshot']['rows']:
            for field in ('supplier_goods_value_cny','known_goods_value_cny','contribution_profit',
                          'cost_status','cost_missing','cost_lines'):
                row.pop(field,None)
        c.execute('UPDATE rec_objects SET data=? WHERE id=?',(json.dumps(old,ensure_ascii=False),ident))
        legacy=client.get('/api/reconciliation-v2/history/export/'+ident)
        assert legacy.status_code==200
        assert '旧快照未计算成本' in legacy.data.decode('utf-8-sig')
        c.execute("UPDATE orders SET line_items=?",(json.dumps([
            {'id':1,'name':'Fumot 9000 Puffs','quantity':2,'total':'100'}]),))
        c.execute("INSERT INTO product_costs VALUES(1,10,9,NULL,9000,NULL,'39','CNY','2026-07-01')")
        fresh=preview(c,rule,'2026-08','PLN')
        fresh_id=create_draft(c,{'rule_id':rule,'month':'2026-08','currency':'PLN',
                                 'request_key':'admin-export-cost-001','expected_digest':fresh['digest']},1)
        updated=client.get('/api/reconciliation-v2/history/export/'+fresh_id)
        assert updated.status_code==200
        assert '78.00' in updated.data.decode('utf-8-sig')
        assert '本仓成本已匹配' in updated.data.decode('utf-8-sig')
    finally:api_module.get_conn=old


def test_history_template_parses():
    from jinja2 import Environment
    from pathlib import Path
    Environment().parse((Path(__file__).parents[1]/'templates/reconciliation_history.html').read_text(encoding='utf-8'))
