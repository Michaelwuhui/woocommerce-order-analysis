import pytest
from flask import Flask
from flask_login import LoginManager, UserMixin

import fulfillment_api
from fulfillment_service import plan_order
import test_fulfillment as domain_tests


@pytest.fixture
def shortage_api(monkeypatch):
    domain = domain_tests.FulfillmentDomainTests()
    domain.setUp()
    domain.add_order('shortage-display', 'PL')
    plan_order(domain.db, 'shortage-display')
    domain.db.execute("UPDATE oms_order_items SET ordered_qty=2,shortage_qty=1")
    domain.db.execute("UPDATE oms_order_fulfillment_state SET has_shortage=1,aggregate_status='stock_shortage'")
    domain.db.execute("UPDATE oms_fulfillments SET status='stock_shortage'")
    for sid in (2, 3):
        domain.db.execute('INSERT INTO inv_skus(id,sku_code,name,is_active) VALUES (?,?,?,1)',
                          (sid, f'SKU{sid}', f'Product {sid}'))
        domain.db.execute('INSERT INTO oms_sku_warehouses(sku_id,warehouse_id,is_enabled) VALUES (?,?,1)',
                          (sid, sid-1))
    for sid in (2, 3, None):
        domain.db.execute('''INSERT INTO oms_order_items
            (order_id,woo_line_item_id,sku_id,name,ordered_qty,shortage_qty)
            VALUES ('shortage-display',?,?,?,1,1)''',
            (str(sid), sid, f'Unallocated {sid}'))
    domain.db.commit()

    class Connection:
        def execute(self, *args, **kwargs):
            return domain.db.execute(*args, **kwargs)
        def close(self):
            pass
    class User(UserMixin):
        id = '1'
        username = 'admin'
    app = Flask(__name__)
    app.config.update(SECRET_KEY='isolated-shortage-display', TESTING=True)
    LoginManager(app).user_loader(lambda uid: User())
    app.register_blueprint(fulfillment_api.fulfillment_bp)
    scope = {'allowed': None}
    monkeypatch.setattr(fulfillment_api, 'get_conn', Connection)
    monkeypatch.setattr(fulfillment_api, '_allowed_warehouse_ids', lambda *args: scope['allowed'])
    client = app.test_client()
    with client.session_transaction() as session:
        session['_user_id'] = '1'
        session['_fresh'] = True
    yield domain.db, client, scope
    domain.tearDown()


def test_list_and_detail_include_fully_unallocated_and_unmapped_lines(shortage_api):
    conn, client, scope = shortage_api
    before = conn.total_changes
    detail = client.get('/api/fulfillment/order/shortage-display').get_json()
    listing = client.get('/api/fulfillment/orders?search=shortage-display').get_json()
    assert {i['sku_id'] for i in detail['shortage_items']} == {1, 2, 3, None}
    assert listing['items'][0]['shortage_items'] == detail['shortage_items']
    assert len(detail['fulfillments'][0]['items']) == 1
    assert detail['shortage_items'][0]['allocated_qty'] == 1
    assert detail['shortage_items'][1]['allocated_qty'] == 0
    assert conn.total_changes == before  # Reading never replans or changes stock.


def test_warehouse_scope_does_not_reveal_other_warehouse_shortages(shortage_api):
    conn, client, scope = shortage_api
    scope['allowed'] = [1]
    detail = client.get('/api/fulfillment/order/shortage-display').get_json()
    listing = client.get('/api/fulfillment/orders?search=shortage-display').get_json()
    assert {i['sku_id'] for i in detail['shortage_items']} == {1, 2}
    assert listing['items'][0]['shortage_items'] == detail['shortage_items']
    scope['allowed'] = [2]
    assert client.get('/api/fulfillment/order/shortage-display').status_code == 403
    assert client.get('/api/fulfillment/orders?search=shortage-display').get_json()['items'] == []


def test_resolved_shortages_disappear_and_empty_scope_is_rejected(shortage_api):
    conn, client, scope = shortage_api
    conn.execute('UPDATE oms_order_items SET shortage_qty=0')
    assert client.get('/api/fulfillment/order/shortage-display').get_json()['shortage_items'] == []
    scope['allowed'] = []
    assert client.get('/api/fulfillment/order/shortage-display').status_code == 403
