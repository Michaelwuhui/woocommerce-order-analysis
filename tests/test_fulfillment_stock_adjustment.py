import concurrent.futures
import json
import os
import uuid

import pytest
from jinja2 import ChoiceLoader, DictLoader

import fulfillment_api
import inv_workflows
from test_inventory_workflows import system, client, value, detail, payload, post, review
from test_temporary_inventory_access import grant


URL = '/api/inv/operations/fulfillment/fulfill-one/stocktake'


@pytest.fixture
def fulfillment_system(system, monkeypatch):
    conn = system[1]()
    statements = '''
        CREATE TABLE orders(id TEXT PRIMARY KEY,number TEXT,status TEXT);
        INSERT INTO orders VALUES ('order-one','TEST-1','processing'),('order-two','TEST-2','processing');
        CREATE TABLE oms_order_fulfillment_state(order_id TEXT PRIMARY KEY,revision INTEGER);
        INSERT INTO oms_order_fulfillment_state VALUES ('order-one',1),('order-two',1);
        CREATE TABLE oms_fulfillments(id TEXT PRIMARY KEY,order_id TEXT,warehouse_id INTEGER,
            mode TEXT,status TEXT,revision INTEGER);
        INSERT INTO oms_fulfillments VALUES
            ('fulfill-one','order-one',2,'internal','stock_shortage',1),
            ('fulfill-two','order-two',2,'internal','ready_to_pick',1);
        CREATE TABLE oms_order_items(id INTEGER PRIMARY KEY,order_id TEXT,sku_id INTEGER,name TEXT,
            shortage_qty INTEGER,ordered_qty INTEGER);
        INSERT INTO oms_order_items VALUES (1,'order-one',1,'Allocated item',0,2),
            (2,'order-one',2,'Unallocated item',1,1),
            (3,'order-one',NULL,'Missing mapping',1,1),(4,'order-two',1,'Other order',0,1);
        CREATE TABLE oms_fulfillment_items(id INTEGER PRIMARY KEY,fulfillment_id TEXT,order_item_id INTEGER);
        INSERT INTO oms_fulfillment_items VALUES (1,'fulfill-one',1),(2,'fulfill-two',4);
        INSERT INTO inv_skus VALUES (5,'UNRELATED','Unrelated managed product',1);
        INSERT INTO oms_sku_warehouses VALUES (2,5,1);
    '''
    for statement in statements.split(';'):
        if statement.strip():
            conn.execute(statement)
    conn.commit()
    conn.close()
    monkeypatch.setattr(fulfillment_api, 'get_conn', system[1])
    app = system[0]
    app.jinja_loader = ChoiceLoader([DictLoader({'base.html':
        '{% block content %}{% endblock %}{% block scripts %}{% endblock %}'}), app.jinja_loader])
    app.register_blueprint(fulfillment_api.fulfillment_bp)
    return system


def stock_payload(system, quantities=None):
    quantities = {1: 12} if quantities is None else quantities
    context = client(system, 1).get(URL).get_json()
    return {'request_key': str(uuid.uuid4()), 'note': 'Verified physical count', 'revision': context['revision'],
            'items': [{'sku_id': item['sku_id'], 'qty': quantities[item['sku_id']],
                       'baseline': item['on_hand'], 'baseline_reserved': item['reserved'],
                       'baseline_movement': item['movement']}
                      for item in context['items'] if item['sku_id'] in quantities]}


def adjust(system, data=None, uid=1, expected=200, url=URL):
    response = client(system, uid).post(url, json=data if data is not None else stock_payload(system))
    assert response.status_code == expected, response.get_json()
    return response.get_json()


def test_only_builtin_admin_sees_entry_and_can_read_or_write(fulfillment_system):
    system = fulfillment_system
    grant(system, uid=2)
    data = stock_payload(system)
    for uid in (2, 3, 4, 5):
        assert client(system, uid).get(URL).status_code == 403
        adjust(system, data, uid=uid, expected=403)
    assert system[0].test_client().get(URL).status_code == 401
    admin = client(system, 1).get('/fulfillment').get_data(as_text=True)
    manager = client(system, 3).get('/fulfillment').get_data(as_text=True)
    assert 'const ffCanAdjustStock=true' in admin and 'fulfillment_stock_adjustment.js' in admin
    assert 'const ffCanAdjustStock=false' in manager and 'fulfillment_stock_adjustment.js' not in manager
    assert value(system, 'SELECT COUNT(*) FROM inv_documents') == 0


def test_preview_includes_unallocated_shortages_without_creating_stock(fulfillment_system):
    system = fulfillment_system
    response = client(system, 1).get(URL)
    assert response.status_code == 200
    context = response.get_json()
    assert [item['sku_id'] for item in context['items']] == [1, 2]
    assert context['items'][1]['shortage_qty'] == 1
    assert context['items'][0]['on_hand'] == 10 and context['items'][0]['reserved'] == 2
    assert context['unavailable'] == [{'name': 'Missing mapping', 'shortage_qty': 1}]
    assert value(system, 'SELECT COUNT(*) FROM inv_stock WHERE warehouse_id=2 AND sku_id=2') == 0


def test_stocktake_is_atomic_audited_and_preserves_reserved_and_order(fulfillment_system):
    system = fulfillment_system
    result = adjust(system, stock_payload(system, {1: 7, 2: 3}))
    assert result['status'] == 'approved'
    doc = detail(system, result['id'], uid=1)
    assert doc['created_by'] == doc['reviewed_by'] == 1
    assert 'admin' in doc['review_note']
    assert doc['events'][-1]['action'] == 'admin_stocktake_approved'
    audit = json.loads(doc['events'][-1]['detail'])
    assert audit['fulfillment_id'] == 'fulfill-one' and audit['authorization'] == 'builtin_admin'
    assert value(system, 'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1') == 7
    assert value(system, 'SELECT reserved FROM inv_stock WHERE warehouse_id=2 AND sku_id=1') == 2
    assert value(system, 'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=2') == 3
    assert value(system, "SELECT COUNT(*) FROM inv_movements WHERE operator_id=1 AND order_id='order-one'") == 2
    assert value(system, "SELECT status FROM orders WHERE id='order-one'") == 'processing'
    assert value(system, "SELECT status FROM oms_fulfillments WHERE id='fulfill-one'") == 'stock_shortage'
    assert value(system, 'SELECT COUNT(*) FROM inv_temporary_stock_grants') == 0


def test_replay_survives_replan_and_changed_payload_is_rejected(fulfillment_system):
    system = fulfillment_system
    data = stock_payload(system)
    first = adjust(system, data)
    conn = system[1]()
    conn.execute("UPDATE oms_order_fulfillment_state SET revision=2 WHERE order_id='order-one'")
    conn.execute("UPDATE oms_fulfillments SET status='superseded' WHERE id='fulfill-one'")
    conn.commit()
    conn.close()
    assert adjust(system, data)['id'] == first['id']
    assert adjust(system, data)['replayed']
    adjust(system, dict(data, note='changed'), expected=409)
    adjust(system, dict(data, request_key=str(uuid.uuid4())), expected=409)
    assert value(system, 'SELECT COUNT(*) FROM inv_movements') == 1


@pytest.mark.parametrize('change', [
    {'qty': 1}, {'baseline': 9}, {'baseline_reserved': 1}, {'baseline_movement': 99},
])
def test_stale_or_below_reserved_is_rejected(fulfillment_system, change):
    system = fulfillment_system
    data = stock_payload(system)
    data['items'][0].update(change)
    adjust(system, data, expected=409)
    assert value(system, 'SELECT COUNT(*) FROM inv_documents') == 0
    assert value(system, 'SELECT COUNT(*) FROM inv_movements') == 0


@pytest.mark.parametrize('qty', [-1, 1.5, True, None, '', 100000001])
def test_invalid_or_omitted_count_never_resets_stock(fulfillment_system, qty):
    system = fulfillment_system
    data = stock_payload(system)
    data['items'][0]['qty'] = qty
    adjust(system, data, expected=400)
    assert value(system, 'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1') == 10


def test_warehouse_and_sku_cannot_be_substituted(fulfillment_system):
    system = fulfillment_system
    data = stock_payload(system)
    adjust(system, dict(data, warehouse_id=1), expected=400)
    data['items'][0]['sku_id'] = 5
    adjust(system, data, expected=403)
    assert value(system, 'SELECT COUNT(*) FROM inv_documents') == 0


@pytest.mark.parametrize('sql,expected', [
    ("UPDATE oms_fulfillments SET mode='external_wms' WHERE id='fulfill-one'", 409),
    ("UPDATE oms_warehouse_integrations SET inventory_authority='manual_partner' WHERE warehouse_id=2", 400),
    ("UPDATE warehouses SET is_active=0 WHERE id=2", 400),
    ("UPDATE oms_fulfillments SET status='cancelled' WHERE id='fulfill-one'", 409),
])
def test_external_inactive_or_cancelled_scope_cannot_write(fulfillment_system, sql, expected):
    system = fulfillment_system
    data = stock_payload(system)
    conn = system[1](); conn.execute(sql); conn.commit(); conn.close()
    adjust(system, data, expected=expected)
    assert client(system, 1).get(URL).status_code == expected
    assert value(system, 'SELECT COUNT(*) FROM inv_movements') == 0


def test_audit_failure_rolls_back_every_stock_row(fulfillment_system, monkeypatch):
    system = fulfillment_system
    original = inv_workflows.event
    def fail(conn, did, action, detail, **kwargs):
        if action == 'admin_stocktake_approved':
            raise RuntimeError('audit unavailable')
        return original(conn, did, action, detail, **kwargs)
    monkeypatch.setattr(inv_workflows, 'event', fail)
    with pytest.raises(RuntimeError):
        adjust(system, stock_payload(system, {1: 12, 2: 3}))
    assert value(system, 'SELECT COUNT(*) FROM inv_documents') == 0
    assert value(system, 'SELECT COUNT(*) FROM inv_movements') == 0
    assert value(system, 'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1') == 10
    assert value(system, 'SELECT COUNT(*) FROM inv_stock WHERE warehouse_id=2 AND sku_id=2') == 0


def test_simultaneous_same_request_applies_once(fulfillment_system):
    system = fulfillment_system
    data = stock_payload(system)
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
        results = list(pool.map(lambda _: adjust(system, data), range(2)))
    assert results[0]['id'] == results[1]['id']
    assert value(system, 'SELECT COUNT(*) FROM inv_documents') == 1
    assert value(system, 'SELECT COUNT(*) FROM inv_movements') == 1


def test_competing_counts_do_not_overwrite_each_other(fulfillment_system):
    system = fulfillment_system
    first = stock_payload(system)
    second = stock_payload(system, {1: 14})
    def submit(data):
        return client(system, 1).post(URL, json=data).status_code
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
        statuses = list(pool.map(submit, [first, second]))
    assert sorted(statuses) == [200, 409]
    assert value(system, 'SELECT COUNT(*) FROM inv_movements') == 1


def test_ordinary_admin_stocktake_still_requires_separate_review(fulfillment_system):
    system = fulfillment_system
    data = payload('stocktake', admin_approve=True,
                   items=stock_payload(system)['items'])
    result = post(system, data, uid=1)
    assert result['status'] == 'submitted'
    review(system, result['id'], uid=1, expected=403)
    assert value(system, 'SELECT COUNT(*) FROM inv_movements') == 0


def test_real_fulfillment_detail_limits_stock_targets_to_admin_and_local_warehouse(monkeypatch):
    from flask import Flask
    from flask_login import LoginManager, UserMixin
    from test_fulfillment import FulfillmentDomainTests
    from fulfillment_service import plan_order

    domain = FulfillmentDomainTests()
    domain.setUp()
    class Connection:
        def execute(self, *args, **kwargs):
            return domain.db.execute(*args, **kwargs)
        def close(self):
            pass
    class User(UserMixin):
        def __init__(self, uid):
            self.id = uid
            self.username = 'admin' if uid == '1' else 'manager'
    try:
        domain.add_order('quick-stock-order', 'PL', qty=2)
        plan_order(domain.db, 'quick-stock-order')
        domain.db.execute("UPDATE oms_fulfillments SET mode='internal'")
        domain.db.execute("UPDATE oms_warehouse_integrations SET inventory_authority='local' WHERE warehouse_id=1")
        domain.db.commit()
        app = Flask(__name__)
        app.config.update(SECRET_KEY='synthetic-detail-only', TESTING=True)
        login = LoginManager(app)
        login.user_loader(lambda uid: User(uid))
        app.register_blueprint(fulfillment_api.fulfillment_bp)
        monkeypatch.setattr(fulfillment_api, 'get_conn', Connection)
        monkeypatch.setattr(fulfillment_api, '_allowed_warehouse_ids', lambda capability='can_view': None)
        def read(uid):
            browser = app.test_client()
            with browser.session_transaction() as session:
                session['_user_id'] = uid
                session['_fresh'] = True
            response = browser.get('/api/fulfillment/order/quick-stock-order')
            assert response.status_code == 200, response.get_json()
            return response.get_json()['fulfillments'][0]['stock_adjustment_targets']
        assert read('1')[0]['warehouse_id'] == 1
        assert read('2') == []
        domain.db.execute("UPDATE oms_warehouse_integrations SET inventory_authority='manual_partner' WHERE warehouse_id=1")
        assert read('1') == []
    finally:
        domain.tearDown()


@pytest.mark.skipif(not os.environ.get('FULFILLMENT_STOCK_BROWSER'), reason='opt-in isolated browser acceptance')
def test_browser_admin_count_and_lost_response_replay(fulfillment_system, tmp_path):
    import threading
    from pathlib import Path
    from flask import jsonify
    from playwright.sync_api import sync_playwright, expect
    from werkzeug.serving import make_server

    system = fulfillment_system
    app = system[0]
    app.static_folder = str(Path(__file__).resolve().parents[1] / 'static')
    app.jinja_loader = ChoiceLoader([DictLoader({'base.html':'''<!doctype html><html><head><meta charset="UTF-8">
        <meta name="viewport" content="width=device-width,initial-scale=1"></head><body>
        {% block content %}{% endblock %}<script>window.bootstrap={Modal:{getOrCreateInstance:()=>({show(){}})}};</script>
        {% block scripts %}{% endblock %}</body></html>'''}), app.jinja_loader])
    # Synthetic fulfillment display; stock preview and POST use the real API and test database.
    row = {'id': 'fulfill-one', 'order_id': 'order-one', 'order_number': 'TEST-1',
           'warehouse_name': 'Transit', 'warehouse_id': 2, 'mode': 'internal',
           'status': 'stock_shortage', 'has_shortage': 1, 'manual_review': 0,
           'items': [], 'shipments': [], 'stock_adjustment_targets': [
               {'fulfillment_id': 'fulfill-one', 'warehouse_id': 2, 'warehouse_name': 'Transit'}]}
    app.view_functions['fulfillment.list_fulfillment_orders'] = lambda: jsonify(
        items=[row], summary={'stock_shortage': 1}, warehouses=[{'id': 2, 'name': 'Transit'}])
    app.view_functions['fulfillment.fulfillment_order_detail'] = lambda order_id: jsonify(
        state={'has_shortage': 1}, fulfillments=[row])
    server = make_server('127.0.0.1', 0, app, threaded=True)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    base = f'http://127.0.0.1:{server.server_port}'
    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=True)
            errors = []
            def page_for(uid):
                context = browser.new_context()
                context.add_cookies([{'name': 'session', 'value': client(system, uid).get_cookie('session').value,
                                     'url': base}])
                page = context.new_page()
                page.on('pageerror', lambda error: errors.append(str(error)))
                page.on('dialog', lambda dialog: dialog.accept())
                page.goto(base + '/fulfillment')
                page.get_by_role('button', name='处理', exact=True).click()
                return page
            page = page_for(1)
            page.get_by_role('button', name='快速调整库存', exact=True).click()
            expect(page.locator('#ffStockQty2')).to_have_value('0')
            page.locator('#ffStockQty1').fill('12')
            page.locator('#ffStockNote').fill('Isolated browser physical count')
            intercepted = []
            def lose_first_response(route):
                if route.request.method != 'POST':
                    route.continue_()
                    return
                intercepted.append(route.request.post_data_json)
                if len(intercepted) == 1:
                    response = route.fetch()
                    assert response.status == 200
                    route.abort()
                else:
                    route.continue_()
            page.route('**/stocktake', lose_first_response)
            page.get_by_role('button', name='确认调整库存', exact=True).click()
            expect(page.get_by_role('button', name='重试核对本次调整', exact=True)).to_be_visible()
            expect(page.locator('#ffStockQty1')).to_be_disabled()
            page.get_by_role('button', name='重试核对本次调整', exact=True).click()
            expect(page.locator('#ffStockAdjustmentPanel')).to_contain_text('库存已调整')
            assert intercepted[0] == intercepted[1]
            assert value(system, 'SELECT COUNT(*) FROM inv_documents') == 1
            assert value(system, 'SELECT COUNT(*) FROM inv_movements') == 1
            assert value(system, 'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1') == 12
            manager = page_for(3)
            expect(manager.get_by_role('button', name='快速调整库存', exact=True)).to_have_count(0)
            assert errors == []
            browser.close()
    finally:
        server.shutdown()
        thread.join(timeout=5)
