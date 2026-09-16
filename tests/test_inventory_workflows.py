import concurrent.futures
import os
import sqlite3
import uuid
from pathlib import Path

import pytest
from flask import Flask
from flask_login import LoginManager, UserMixin
from jinja2 import ChoiceLoader, DictLoader, FileSystemLoader

import inv_common
import inv_inventory
import inv_workflows as workflows
from inv_workflows_schema import migrate
from inv_temporary_access_schema import migrate as migrate_temporary


@pytest.fixture
def system(tmp_path, monkeypatch):
    path = tmp_path / 'inventory.db'
    test_database = os.environ.get('INVENTORY_WORKFLOW_TEST_PG_DATABASE')
    schema = 'inventory_test_' + uuid.uuid4().hex

    if test_database:
        import db_backend
        assert test_database.startswith('woo_inventory_test_')
        assert os.environ.get('WOO_DB_NAME_OVERRIDE') == test_database
        assert db_backend.is_postgres_backend()
        root = db_backend.connect()
        assert root.execute('SELECT current_database()').fetchone()[0] == test_database
        root.execute('CREATE SCHEMA ' + schema)
        root.commit()
        root.close()

    def connect():
        if test_database:
            c = db_backend.connect()
            c.execute('SET search_path TO ' + schema)
            return c
        c = sqlite3.connect(path, timeout=10)
        c.row_factory = sqlite3.Row
        c.execute('PRAGMA foreign_keys=ON')
        return c

    c = connect()
    fixture_sql = '''
        CREATE TABLE users(id INTEGER PRIMARY KEY, username TEXT, name TEXT,
            can_view_inventory INTEGER DEFAULT 0, can_manage_inventory INTEGER DEFAULT 0);
        INSERT INTO users VALUES(1,'admin','Admin',0,0),(2,'shipper','Shipper',0,0),
            (3,'reviewer','Reviewer',1,1),(4,'other','Other',0,0),(5,'scoped','Scoped',0,0);
        CREATE TABLE warehouses(id INTEGER PRIMARY KEY,name TEXT,code TEXT,country TEXT,is_active INTEGER);
        INSERT INTO warehouses VALUES(1,'Source','SOURCE','PL',1),(2,'Transit','TRANSIT','PL',1),
            (3,'Other','OTHER','PL',1),(4,'WMS','WMS','PL',1),(5,'Manual','MANUAL','PL',1);
        CREATE TABLE oms_warehouse_integrations(warehouse_id INTEGER PRIMARY KEY,inventory_authority TEXT);
        INSERT INTO oms_warehouse_integrations VALUES(1,'local'),(2,'local'),(3,'local'),(4,'external_wms'),(5,'manual_partner');
        CREATE TABLE oms_warehouse_user_permissions(user_id INTEGER,warehouse_id INTEGER,can_view INTEGER,can_ship INTEGER);
        INSERT INTO oms_warehouse_user_permissions VALUES(2,2,1,1),(4,3,1,1);
        CREATE TABLE partner_users(user_id INTEGER,partner_id INTEGER);
        CREATE TABLE inv_skus(id INTEGER PRIMARY KEY,sku_code TEXT,name TEXT,is_active INTEGER);
        INSERT INTO inv_skus VALUES(1,'ONE','One',1),(2,'TWO','Two',1),(3,'OFF','Inactive',0),(4,'UNMAPPED','Unmapped',1);
        CREATE TABLE oms_sku_warehouses(warehouse_id INTEGER,sku_id INTEGER,is_enabled INTEGER);
        INSERT INTO oms_sku_warehouses VALUES(1,1,1),(1,2,1),(2,1,1),(2,2,1),(3,1,1),(4,1,1);
        CREATE TABLE inv_stock(warehouse_id INTEGER,sku_id INTEGER,on_hand INTEGER,reserved INTEGER,
            updated_at TEXT,UNIQUE(warehouse_id,sku_id));
        INSERT INTO inv_stock VALUES(1,1,20,3,NULL),(1,2,1,0,NULL),(2,1,10,2,NULL),(3,1,10,0,NULL);
        CREATE TABLE inv_movements(id INTEGER PRIMARY KEY AUTOINCREMENT,ts TEXT DEFAULT CURRENT_TIMESTAMP,
            warehouse_id INTEGER,sku_id INTEGER,batch_id INTEGER,movement_type TEXT,qty_delta INTEGER,
            reserved_delta INTEGER,qty_before INTEGER,qty_after INTEGER,reserved_before INTEGER,reserved_after INTEGER,
            ref_type TEXT,ref_id TEXT,order_id TEXT,operator_id INTEGER,operator_name TEXT,note TEXT);
    '''
    if test_database:
        fixture_sql = fixture_sql.replace('INTEGER PRIMARY KEY AUTOINCREMENT', 'BIGSERIAL PRIMARY KEY')
        for statement in fixture_sql.split(';'):
            if statement.strip():
                c.execute(statement)
    else:
        c.executescript(fixture_sql)
    migrate(c)
    migrate_temporary(c)
    c.close()
    for module in (inv_common, inv_inventory, workflows):
        monkeypatch.setattr(module, 'get_conn', connect)
    app = Flask(__name__)
    app.config.update(SECRET_KEY='test', TESTING=True)
    app.jinja_loader = ChoiceLoader([
        DictLoader({'base.html':'{% block content %}{% endblock %}'}),
        FileSystemLoader(str(Path(__file__).resolve().parents[1] / 'templates')),
    ])
    login = LoginManager(app)

    class User(UserMixin):
        def __init__(self, row):
            self.id, self.username, self.name = row['id'], row['username'], row['name']

    @login.user_loader
    def load_user(uid):
        with connect() as db:
            row = db.execute('SELECT * FROM users WHERE id=?', (uid,)).fetchone()
        return User(row) if row else None

    app.register_blueprint(inv_inventory.inv_inv_bp)
    return app, connect


def client(system, uid=2):
    c = system[0].test_client()
    with c.session_transaction() as session:
        session['_user_id'] = str(uid)
        session['_fresh'] = True
    return c


def post(system, payload, uid=2, expected=200):
    response = client(system, uid).post('/api/inv/operations', json=payload)
    assert response.status_code == expected, response.get_json()
    return response.get_json()


def payload(kind='replenishment', **overrides):
    result = dict(kind=kind, warehouse_id=2, reference=str(uuid.uuid4()),
                  request_key=str(uuid.uuid4()), note='counted physical goods', items=[dict(sku_id=1,qty=5)])
    result.update(overrides)
    return result


def review(system, did, uid=3, decision='approve', expected=200):
    r = client(system, uid).post(f'/api/inv/operations/{did}/review', json={'decision':decision,'note':'checked documents'})
    assert r.status_code == expected, r.get_json()
    return r.get_json()


def value(system, sql, args=()):
    c = system[1]()
    try:
        return c.execute(sql,args).fetchone()[0]
    finally:
        c.close()


def detail(system, did, uid=2):
    r=client(system,uid).get(f'/api/inv/operations/{did}')
    assert r.status_code == 200, r.get_json()
    return r.get_json()


def receipt(system, parent, qty, damaged=0, short=0, **kw):
    return post(system,payload('receipt',parent_id=parent,items=[dict(sku_id=1,qty=qty,damaged_qty=damaged,short_qty=short)],**kw))['id']


def test_receipt_partial_damage_shortage_and_audit(system):
    parent=post(system,payload())['id']
    first=receipt(system,parent,2)
    assert value(system,'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1') == 10
    review(system,first)
    assert detail(system,parent)['status']=='partial'
    second=receipt(system,parent,1,damaged=1,short=1)
    review(system,second)
    d=detail(system,parent)
    assert d['status']=='completed_with_difference'
    assert d['remaining']['1']==0
    assert value(system,'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1') == 13
    assert value(system,'SELECT reserved FROM inv_stock WHERE warehouse_id=2 AND sku_id=1') == 2
    assert value(system,'SELECT COUNT(*) FROM inv_movements') == 2
    assert value(system,'SELECT operator_id FROM inv_movements LIMIT 1') == 3
    assert len(detail(system,second)['events'])==2


def test_request_replay_mismatch_and_duplicate_reference(system):
    p=payload()
    first=post(system,p)
    assert post(system,p)['id']==first['id']
    changed=dict(p,note='changed')
    post(system,changed,expected=409)
    post(system,dict(p,request_key=str(uuid.uuid4())),expected=409)
    assert value(system,'SELECT COUNT(*) FROM inv_documents')==1


def test_review_replay_is_noop(system):
    parent=post(system,payload())['id']
    did=receipt(system,parent,5)
    review(system,did)
    assert review(system,did)['replayed']
    review(system,did,decision='reject',expected=409)
    assert value(system,'SELECT COUNT(*) FROM inv_movements')==1


def test_self_review_and_unprivileged_approval_denied(system):
    parent=post(system,payload())['id']
    did=receipt(system,parent,5)
    review(system,did,uid=2,expected=403)
    did2=post(system,payload('stocktake',items=[dict(sku_id=1,qty=9,baseline=10,baseline_reserved=2,baseline_movement=0)]),uid=3)['id']
    review(system,did2,uid=3,expected=403)
    assert value(system,'SELECT COUNT(*) FROM inv_movements')==0


def test_scope_and_explicit_non_partner_inventory_entry(system):
    shipper=client(system)
    assert shipper.get('/inventory/operations').status_code==200
    assert [w['id'] for w in shipper.get('/api/inv/operations/context').get_json()['warehouses']]==[2]
    assert shipper.get('/api/inv/operations/catalog/3').status_code==403
    post(system,payload(warehouse_id=3),expected=403)
    did=post(system,payload())['id']
    assert client(system,4).get(f'/api/inv/operations/{did}').status_code==403
    assert client(system,4).get('/api/inv/operations').get_json()==[]
    assert client(system,4).post(f'/api/inv/operations/{did}/cancel',json={'note':'bad'}).status_code==403


@pytest.mark.parametrize('items', [[],[{'sku_id':1,'qty':-1}],[{'sku_id':1,'qty':1.5}],
    [{'sku_id':1,'qty':True}],[{'sku_id':1,'qty':0}],[{'sku_id':4,'qty':1}],
    [{'sku_id':3,'qty':1}],[{'sku_id':1,'qty':1},{'sku_id':1,'qty':2}]])
def test_invalid_items(system,items):
    post(system,payload(items=items),expected=400)
    assert value(system,'SELECT COUNT(*) FROM inv_documents')==0


def test_manual_partner_cannot_be_quantity_managed(system):
    post(system,payload(warehouse_id=5),uid=1,expected=400)


def test_overreceipt_pending_duplicate_and_reject_recovery(system):
    parent=post(system,payload())['id']
    post(system,payload('receipt',parent_id=parent,items=[dict(sku_id=1,qty=6)]),expected=400)
    did=receipt(system,parent,3,reference='delivery-1')
    post(system,payload('receipt',parent_id=parent),expected=409)
    review(system,did,decision='reject')
    fixed=receipt(system,parent,5,reference='delivery-1')
    review(system,fixed)
    assert detail(system,parent)['status']=='completed'
    assert value(system,'SELECT COUNT(*) FROM inv_movements')==1


def test_stocktake_uses_on_hand_not_available(system):
    did=post(system,payload('stocktake',items=[dict(sku_id=1,qty=9,baseline=10,baseline_reserved=2,baseline_movement=0)]))['id']
    review(system,did)
    assert value(system,'SELECT qty_delta FROM inv_movements')==-1
    assert value(system,'SELECT reserved FROM inv_stock WHERE warehouse_id=2 AND sku_id=1')==2


@pytest.mark.parametrize('stage',['before_submit','before_review'])
def test_stocktake_concurrent_movement_requires_recount(system,stage):
    p=payload('stocktake',items=[dict(sku_id=1,qty=9,baseline=10,baseline_reserved=2,baseline_movement=0)])
    did=post(system,p)['id'] if stage=='before_review' else None
    c=system[1]()
    inv_common.record_movement(c,warehouse_id=2,sku_id=1,movement_type='reserve',reserved_delta=1)
    c.commit();c.close()
    if did:
        review(system,did,expected=409)
        assert detail(system,did)['status']=='submitted'
    else:
        post(system,p,expected=409)
    assert value(system,'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1')==10


def test_stocktake_below_reserved_and_zero_difference(system):
    p=payload('stocktake',items=[dict(sku_id=1,qty=1,baseline=10,baseline_reserved=2,baseline_movement=0)])
    post(system,p,expected=409)
    p=payload('stocktake',items=[dict(sku_id=1,qty=10,baseline=10,baseline_reserved=2,baseline_movement=0)])
    did=post(system,p)['id'];review(system,did)
    assert detail(system,did)['status']=='approved'
    assert value(system,'SELECT COUNT(*) FROM inv_movements')==0


def test_transfer_out_transit_then_receipt(system):
    did=post(system,payload('transfer',source_warehouse_id=1),uid=3)['id']
    post(system,payload('receipt',parent_id=did),expected=409)
    review(system,did,uid=1)
    assert value(system,'SELECT on_hand FROM inv_stock WHERE warehouse_id=1 AND sku_id=1')==15
    assert value(system,'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1')==10
    assert detail(system,did)['status']=='in_transit'
    incoming=receipt(system,did,5)
    review(system,incoming)
    assert value(system,'SELECT SUM(on_hand) FROM inv_stock WHERE sku_id=1 AND warehouse_id IN (1,2)')==30
    assert detail(system,did)['status']=='completed'
    assert client(system,3).post(f'/api/inv/operations/{did}/cancel',json={'note':'bad'}).status_code==409


def test_transfer_all_or_nothing_and_reserved_protection(system):
    did=post(system,payload('transfer',source_warehouse_id=1,items=[dict(sku_id=1,qty=5),dict(sku_id=2,qty=2)]),uid=3)['id']
    review(system,did,uid=1,expected=409)
    assert value(system,'SELECT on_hand FROM inv_stock WHERE warehouse_id=1 AND sku_id=1')==20
    assert value(system,'SELECT COUNT(*) FROM inv_movements')==0
    did=post(system,payload('transfer',source_warehouse_id=1,items=[dict(sku_id=1,qty=18)]),uid=3)['id']
    review(system,did,uid=1,expected=409)


def test_external_transfer_never_changes_external_stock(system):
    did=post(system,payload('transfer',source_warehouse_id=4),uid=3)['id']
    review(system,did,uid=1)
    assert value(system,'SELECT COUNT(*) FROM inv_stock WHERE warehouse_id=4')==0
    assert value(system,"SELECT COUNT(*) FROM inv_movements WHERE warehouse_id=4")==0
    assert value(system,"SELECT COUNT(*) FROM inv_document_events WHERE action='external_dispatch_confirmed'")==1
    incoming=receipt(system,did,5);review(system,incoming)
    assert value(system,'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1')==15


def test_permissions_scoped_and_audited(system):
    path='/api/inv/operations/permissions'
    assert client(system,3).put(path,json={}).status_code==403
    grant={'user_id':5,'warehouse_id':2,'can_approve':1}
    assert client(system,1).put(path,json=grant).status_code==200
    scoped=client(system,5)
    assert scoped.get('/inventory/operations').status_code==200
    assert [w['id'] for w in scoped.get('/api/inv/operations/context').get_json()['warehouses']]==[2]
    post(system,payload(),uid=5,expected=403)
    parent=post(system,payload())['id'];did=receipt(system,parent,5)
    review(system,did,uid=5)
    assert value(system,"SELECT COUNT(*) FROM inv_document_events WHERE action='permission_changed'")==1
    c=system[1]();migrate(c);c.close()
    assert value(system,'SELECT can_approve FROM inv_operation_permissions WHERE user_id=5 AND warehouse_id=2')==1


def test_legacy_direct_adjust_and_receive_cannot_bypass_review(system):
    c=client(system,1)
    assert c.post('/api/inv/adjust',json={'warehouse_id':2,'sku_id':1,'qty_delta':5,'reason':'bypass'}).status_code==409
    assert c.post('/api/inv/purchase-orders/1/receive',json={}).status_code==409


def test_concurrent_duplicate_create_and_review(system):
    p=payload()
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
        results=list(pool.map(lambda _:post(system,p),range(2)))
    assert results[0]['id']==results[1]['id']
    did=receipt(system,results[0]['id'],5)
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
        list(pool.map(lambda _:review(system,did),range(2)))
    assert value(system,'SELECT COUNT(*) FROM inv_movements')==1


def test_cancel_and_migration_preserve_stock(system):
    did=post(system,payload())['id']
    r=client(system).post(f'/api/inv/operations/{did}/cancel',json={'note':'wrong reference'})
    assert r.status_code==200
    post(system,payload('receipt',parent_id=did),expected=409)
    c=system[1]();migrate(c);c.close()
    assert value(system,'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1')==10


def test_concurrent_transfer_and_order_reservation(system):
    from fulfillment_service import _apply_managed_transfer_stock_once, DomainError
    did=post(system,payload('transfer',source_warehouse_id=1,items=[dict(sku_id=1,qty=17)]),uid=3)['id']

    def reserve():
        c=system[1]()
        try:
            c.execute('BEGIN IMMEDIATE')
            _apply_managed_transfer_stock_once(c,warehouse_id=1,sku_id=1,quantity=1,
                movement_type='reserve',ref_id='test-race',order_id='test-order',actor=None,note='test reservation')
            c.commit()
            return 'reserved'
        except DomainError:
            c.rollback()
            return 'insufficient'
        finally:
            c.close()

    def approve():
        r=client(system,1).post(f'/api/inv/operations/{did}/review',json={'decision':'approve','note':'physical transfer checked'})
        assert r.status_code in (200,409),r.get_json()
        return r.status_code

    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
        a=pool.submit(reserve); b=pool.submit(approve)
        result=a.result(),b.result()
    assert result in (('reserved',409),('insufficient',200))
    assert value(system,'SELECT on_hand-reserved FROM inv_stock WHERE warehouse_id=1 AND sku_id=1')>=0


def test_transfer_authority_change_blocks_approval(system):
    did=post(system,payload('transfer',source_warehouse_id=1),uid=3)['id']
    c=system[1]();c.execute("UPDATE oms_warehouse_integrations SET inventory_authority='external_wms' WHERE warehouse_id=1");c.commit();c.close()
    review(system,did,uid=1,expected=409)
    assert value(system,'SELECT COUNT(*) FROM inv_movements')==0


def test_label_only_partner_transfer_uses_destination_skus_and_evidence(system):
    did=post(system,payload('transfer',source_warehouse_id=5),uid=3)['id']
    review(system,did,uid=1)
    assert value(system,'SELECT COUNT(*) FROM inv_stock WHERE warehouse_id=5')==0
    incoming=receipt(system,did,5);review(system,incoming)
    assert value(system,'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1')==15


@pytest.mark.skipif(not os.environ.get('INVENTORY_WORKFLOW_BROWSER'), reason='opt-in isolated browser acceptance')
def test_browser_receiving_review_and_mobile_layout(system, tmp_path):
    import threading
    import requests
    from flask import Response
    from playwright.sync_api import sync_playwright, expect
    from werkzeug.serving import make_server

    app = system[0]
    # Same Bootstrap version and script ordering as the real base template.
    assets = {}
    for key, suffix in [('css','css/bootstrap.min.css'),('js','js/bootstrap.bundle.min.js')]:
        response = requests.get('https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/' + suffix, timeout=30)
        response.raise_for_status()
        assets[key] = response.content
    app.jinja_loader = ChoiceLoader([
        DictLoader({'base.html':'''<!doctype html><html><head><meta charset="UTF-8">
            <meta name="viewport" content="width=device-width,initial-scale=1">
            <link rel="stylesheet" href="/test-assets/css"></head><body class="bg-dark text-white">
            {% block content %}{% endblock %}<script src="/test-assets/js"></script></body></html>'''}),
        FileSystemLoader(str(Path(__file__).resolve().parents[1] / 'templates')),
    ])

    @app.get('/test-assets/<kind>')
    def test_assets(kind):
        return Response(assets[kind], mimetype='text/css' if kind=='css' else 'application/javascript')

    server = make_server('127.0.0.1', 0, app, threaded=True)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    base = f'http://127.0.0.1:{server.server_port}'
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(executable_path=os.environ['INVENTORY_WORKFLOW_BROWSER'], headless=True)
            def browser_user(uid):
                context = browser.new_context(viewport={'width':1440,'height':1000})
                cookie = client(system,uid).get_cookie('session')
                context.add_cookies([{'name':'session','value':cookie.value,'url':base}])
                return context.new_page()
            page = browser_user(2)
            errors=[]
            page.on('pageerror',lambda error:errors.append(str(error)))
            page.goto(base+'/inventory/operations')
            page.get_by_role('button',name='新建补货单',exact=True).click()
            page.locator('#formReference').fill('browser-replenishment')
            page.locator('#formNote').fill('physical arrival list')
            page.locator('tr[data-sku="1"] .choose').check()
            page.locator('tr[data-sku="1"] .qty').fill('5')
            page.get_by_role('button',name='保存补货单（不增库存）',exact=True).click()
            expect(page.locator('#opTitle')).to_contain_text('待到货')
            page.get_by_role('button',name='登记本批实收',exact=True).click()
            page.locator('#formReference').fill('browser-delivery-1')
            page.locator('#formNote').fill('received three, two pending')
            page.locator('tr[data-sku="1"] .choose').check()
            page.locator('tr[data-sku="1"] .qty').fill('3')
            page.get_by_role('button',name='提交审核（暂不入账）',exact=True).click()
            expect(page.locator('#opTitle')).to_contain_text('待审核')
            expect(page.get_by_role('button',name='审核通过并入账',exact=True)).to_have_count(0)
            assert value(system,'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1')==10
            reviewer=browser_user(3)
            reviewer.goto(base+'/inventory/operations')
            reviewer.locator('#opRows [data-detail="2"]').click()
            reviewer.locator('#reviewNote').fill('count and delivery checked')
            reviewer.on('dialog',lambda dialog:dialog.accept())
            reviewer.get_by_role('button',name='审核通过并入账',exact=True).click()
            expect(reviewer.locator('#opTitle')).to_contain_text('已入账')
            assert value(system,'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1')==13
            reviewer.screenshot(path=str(tmp_path/'receipt-reviewed.png'),full_page=True)
            page.set_viewport_size({'width':390,'height':844})
            page.screenshot(path=str(tmp_path/'receipt-mobile.png'),full_page=True)
            assert page.evaluate('document.documentElement.scrollWidth <= window.innerWidth')
            assert errors == []
            browser.close()
    finally:
        server.shutdown()
        thread.join(timeout=5)
