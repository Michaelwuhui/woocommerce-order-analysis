"""Exercise the real pending endpoint SQL without importing production startup."""
import ast
import html
import json
import sqlite3
from pathlib import Path
from types import SimpleNamespace

import pytest
from flask import Flask, jsonify, request


ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture
def pending(tmp_path):
    path = tmp_path / 'pending.db'

    def connect():
        conn = sqlite3.connect(path)
        conn.row_factory = sqlite3.Row
        return conn

    conn = connect()
    conn.executescript('''
        CREATE TABLE sites(url TEXT, manager TEXT, country TEXT);
        CREATE TABLE warehouses(id INTEGER, name TEXT);
        CREATE TABLE orders(
            id TEXT, number TEXT, status TEXT, total REAL, currency TEXT,
            date_created TEXT, payment_method TEXT, source TEXT, billing TEXT,
            shipping TEXT, line_items TEXT, meta_data TEXT, shipping_total REAL,
            shipping_lines TEXT, customer_note TEXT, warehouse_id INTEGER,
            is_undelivered INTEGER, is_problem_return INTEGER, delivery_confirmed INTEGER
        );
        CREATE TABLE order_notes(id INTEGER, order_id TEXT, customer_note INTEGER,
            note TEXT, date_created TEXT, author TEXT);
        CREATE TABLE oms_order_fulfillment_state(order_id TEXT, revision INTEGER,
            aggregate_status TEXT, has_shortage INTEGER, manual_review INTEGER, manual_reason TEXT);
        CREATE TABLE oms_warehouse_user_permissions(user_id INTEGER, warehouse_id INTEGER, can_view INTEGER);
        CREATE TABLE oms_fulfillments(id TEXT, order_id TEXT, warehouse_id INTEGER, revision INTEGER, status TEXT);
        CREATE TABLE oms_order_items(id INTEGER PRIMARY KEY, order_id TEXT,
            woo_line_item_id TEXT, name TEXT, raw_json TEXT, shortage_qty INTEGER);
        CREATE TABLE oms_fulfillment_items(id INTEGER PRIMARY KEY, fulfillment_id TEXT,
            order_item_id INTEGER, allocated_qty INTEGER);
        CREATE TABLE shipping_logs(id INTEGER, order_id TEXT, tracking_number TEXT,
            carrier_slug TEXT, shipped_at TEXT, items_json TEXT, is_partial INTEGER, status TEXT);
        CREATE TABLE shipping_carriers(slug TEXT, name TEXT);
        INSERT INTO sites VALUES('https://vapeprofi.cz', 'manager', 'CZ');
        INSERT INTO warehouses VALUES(10, 'Transit'), (20, 'Other warehouse');
    ''')
    conn.commit()
    conn.close()
    user = SimpleNamespace(id=1, username='admin', is_admin=lambda: True, is_viewer=lambda: False)
    tree = ast.parse((ROOT / 'app.py').read_text(encoding='utf-8'))
    node = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == 'get_pending_orders')
    node.decorator_list = []
    namespace = {
        'html': html,
        'request': request, 'jsonify': jsonify, 'current_user': user,
        'get_db_connection': connect,
        'get_big_order_thresholds': lambda conn: (0, 0),
        'get_big_order_native_amount_threshold': lambda *args: 0,
        'evaluate_big_order': lambda *args, **kwargs: (False, []),
        'get_user_allowed_sources': lambda *args: None,
        '_manual_partner_only_warehouse_scope': lambda *args: False,
        'parse_json_field': lambda value: json.loads(value) if value else {},
        'partition_shipping_logs': lambda rows: (rows, []),
        'is_pending_shipping_candidate': lambda *args: True,
        '_build_risk_index': lambda conn: {},
        '_assess_customer_risk': lambda *args, **kwargs: {},
        'extract_custom_billing_fields': lambda value: {'customer_inpost_id': '', 'customer_social': ''},
        '_is_high_risk_shipping_postcode': lambda value: False,
        '_compose_address': lambda addr: '', '_au_state_mismatch': lambda addr: None,
    }
    helpers = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in {
        'extract_flavor_from_meta', 'extract_puffs_from_meta',
        'get_full_product_name', 'get_display_product_name',
    }]
    exec(compile(ast.Module(body=helpers + [node], type_ignores=[]), 'app.py', 'exec'), namespace)
    app = Flask(__name__)

    def call():
        with app.test_request_context('/api/shipping/pending?source=https://vapeprofi.cz'):
            return namespace['get_pending_orders']().get_json()

    def add_order(number, quantities, allocated, unit_price=378, planned=True):
        conn = connect()
        order_id = '30-' + str(number)
        names = ['Strawberry Ice', 'BLUEBERRY RASPBERRY', 'Kiwi Passion Fruit Guava',
                 'Blue Razz Gummy Bear', 'Mixed Berries', 'STRAWBERRY WATERMELON', 'Watermelon Bubblegum']
        products = [dict(id=28+i, product_id=29668, variation_id=29670+i,
                         name=names[i], quantity=qty, total=f'{qty*unit_price:.2f}')
                    for i, qty in enumerate(quantities)]
        conn.execute('''INSERT INTO orders VALUES(?, ?, 'processing', ?, 'CZK',
            '2026-09-08T12:34:37', 'cod', 'https://vapeprofi.cz', '{}', '{}', ?, '[]',
            200, '[]', '', 10, 0, 0, 0)''',
            (order_id, str(number), round(sum(quantities)*unit_price+200, 2), json.dumps(products)))
        if planned:
            conn.execute("INSERT INTO oms_order_fulfillment_state VALUES(?, 1, 'stock_shortage', 1, 1, 'Shortage')", (order_id,))
            conn.execute("INSERT INTO oms_fulfillments VALUES(?, ?, 10, 1, 'stock_shortage')", (order_id, order_id))
            for item, qty in zip(products, allocated):
                cursor = conn.execute('''INSERT INTO oms_order_items(order_id, woo_line_item_id,
                    name, raw_json, shortage_qty) VALUES(?, ?, ?, ?, ?)''',
                    (order_id, str(item['id']), item['name'], json.dumps(item), item['quantity']-qty))
                if qty:
                    conn.execute('''INSERT INTO oms_fulfillment_items(fulfillment_id,
                        order_item_id, allocated_qty) VALUES(?, ?, ?)''', (order_id, cursor.lastrowid, qty))
        conn.commit()
        conn.close()
        return products

    return SimpleNamespace(call=call, add=add_order, connect=connect, user=user, namespace=namespace)


@pytest.mark.parametrize('number,allocated,price,shortages', [
    (32842, [1, 0, 1, 2, 1, 1, 1], 378, [0, 2, 0, 1, 0, 0, 0]),
    (32843, [1, 0, 1, 0, 1, 1, 1], 341.10, [0, 2, 0, 3, 0, 0, 0]),
])
def test_shortage_orders_show_all_ordered_lines_and_amounts(pending, number, allocated, price, shortages):
    source = pending.add(number, [1, 2, 1, 3, 1, 1, 1], allocated, price)
    order, = pending.call()
    assert order['product_count'] == 10
    assert len(order['products']) == 7
    for shown, original, shortage in zip(order['products'], source, shortages):
        assert shown['item_id'] == original['id']
        assert shown['quantity'] == original['quantity']
        assert shown['total'] == float(original['total'])
        assert shown['shortage_quantity'] == shortage
    assert sum(p['total'] for p in order['products']) + order['shipping_total'] == pytest.approx(order['total'])
    assert order['has_shortage'] and order['manual_review']


def test_site_manager_without_warehouse_restriction_sees_full_order(pending):
    pending.user.username = 'site_manager'
    pending.user.is_admin = lambda: False
    pending.add(32842, [1, 2], [1, 0])
    assert pending.call()[0]['product_count'] == 3


@pytest.mark.parametrize('warehouse_scope', [False, True])
def test_pending_products_include_flavor_for_each_variation(pending, warehouse_scope):
    products = pending.add(1878, [1, 2], [1, 2])
    conn = pending.connect()
    for product, flavor in zip(products, ['Green Apple', 'Dragon Fruit Ice']):
        product['name'] = 'Merry Mi Blade 30000 — 21 smaków'
        product['meta_data'] = [{
            'key': 'pa_smak', 'value': flavor.lower().replace(' ', '-'),
            'display_key': 'Smak', 'display_value': flavor,
        }]
        conn.execute('UPDATE oms_order_items SET name=?, raw_json=? WHERE woo_line_item_id=?',
                     (product['name'], json.dumps(product), str(product['id'])))
    conn.execute('UPDATE orders SET line_items=? WHERE id=?', (json.dumps(products), '30-1878'))
    if warehouse_scope:
        conn.execute('INSERT INTO oms_warehouse_user_permissions VALUES(1, 10, 1)')
        pending.user.username = 'warehouse_operator'
        pending.user.is_admin = lambda: False
    conn.commit()
    conn.close()

    order, = pending.call()
    suffix = '（Transit）' if warehouse_scope else ''
    assert [p['name'] for p in order['products']] == [
        f'Merry Mi Blade 30000 — 21 smaków - {flavor}{suffix}'
        for flavor in ['Green Apple', 'Dragon Fruit Ice']
    ]
    assert [p['quantity'] for p in order['products']] == [1, 2]


def test_warehouse_operator_sees_market_shortages_not_other_allocations(pending):
    pending.add(32842, [1, 2, 1], [1, 0, 0])
    conn = pending.connect()
    conn.execute('INSERT INTO oms_warehouse_user_permissions VALUES(1, 10, 1)')
    conn.execute("INSERT INTO oms_fulfillments VALUES('hidden', '30-32842', 20, 1, 'ready')")
    item = conn.execute("SELECT id FROM oms_order_items WHERE woo_line_item_id='30'").fetchone()[0]
    conn.execute("INSERT INTO oms_fulfillment_items(fulfillment_id, order_item_id, allocated_qty) VALUES('hidden', ?, 1)", (item,))
    conn.execute('UPDATE oms_order_items SET shortage_qty=0 WHERE id=?', (item,))
    conn.commit()
    conn.close()
    pending.user.username = 'warehouse_operator'
    pending.user.is_admin = lambda: False
    order, = pending.call()
    assert order['product_count'] == 3
    assert len(order['products']) == 2
    assert order['products'][0]['name'] == 'Strawberry Ice（Transit）'
    assert order['products'][1]['shortage_quantity'] == 2
    assert order['fulfillment_warehouses'] == ['Transit']


def test_wholly_unallocated_order_and_unplanned_order_remain_complete(pending):
    pending.add(32842, [1, 2], [0, 0])
    pending.add(32843, [1, 2], [0, 0], planned=False)
    orders = {o['number']: o for o in pending.call()}
    assert orders['32842']['product_count'] == orders['32843']['product_count'] == 3
    assert [p['shortage_quantity'] for p in orders['32842']['products']] == [1, 2]
    assert [p['shortage_quantity'] for p in orders['32843']['products']] == [0, 0]


def test_source_permissions_still_hide_unauthorized_orders(pending):
    pending.add(32842, [1, 2], [1, 0])
    pending.namespace['get_user_allowed_sources'] = lambda *args: ['https://other.example']
    assert pending.call() == []


@pytest.mark.parametrize('username', ['jinyi', 'jingu', 'qianpinhong', 'new_market_shipper'])
def test_wholly_shortage_order_visible_without_any_allocated_fulfillment(pending, username):
    pending.add(14568, [1, 2], [0, 0])
    conn = pending.connect()
    conn.execute('DELETE FROM oms_fulfillments')
    conn.execute('INSERT INTO oms_warehouse_user_permissions VALUES(1, 10, 1)')
    conn.commit(); conn.close()
    pending.user.username = username
    pending.user.is_admin = lambda: False
    pending.namespace['get_user_allowed_sources'] = lambda *args: ['https://vapeprofi.cz']
    order, = pending.call()
    assert order['product_count'] == 3
    assert [p['shortage_quantity'] for p in order['products']] == [1, 2]
    assert order['fulfillment_warehouses'] == []
    pending.namespace['get_user_allowed_sources'] = lambda *args: []
    assert pending.call() == []


def test_partial_allocation_and_shortage_are_one_line_not_extra_product(pending):
    pending.add(14568, [3, 1, 1], [1, 1, 1])
    conn=pending.connect()
    conn.execute('INSERT INTO oms_warehouse_user_permissions VALUES(1, 10, 1)')
    conn.commit(); conn.close()
    pending.user.username='shipper'
    pending.user.is_admin=lambda:False
    order, = pending.call()
    assert len(order['products']) == 3
    assert order['product_count'] == 5
    assert order['products'][0]['quantity'] == 3
    assert order['products'][0]['shortage_quantity'] == 2
    assert order['products'][0]['total'] == 3*378


def test_live_country_site_exclusion_resolver(tmp_path):
    path=tmp_path/'scopes.db'
    def connect():
        conn=sqlite3.connect(path);conn.row_factory=sqlite3.Row;return conn
    c=connect()
    c.executescript('''CREATE TABLE user_country_permissions(user_id INTEGER,country TEXT);
        CREATE TABLE user_site_permissions(user_id INTEGER,site_id INTEGER);
        CREATE TABLE user_site_exclusions(user_id INTEGER,site_id INTEGER);
        CREATE TABLE sites(id INTEGER,url TEXT,country TEXT);
        INSERT INTO sites VALUES(1,'cz-one','CZ'),(2,'cz-excluded','CZ'),(3,'pl-one','PL'),(4,'hu-one','HU');
        INSERT INTO user_country_permissions VALUES(18,'CZ'),(18,'PL');
        INSERT INTO user_site_exclusions VALUES(18,2);''')
    c.commit();c.close()
    tree=ast.parse((ROOT/'app.py').read_text(encoding='utf-8'))
    node=next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='_resolve_user_sites')
    ns={'get_db_connection':connect}
    exec(compile(ast.Module(body=[node],type_ignores=[]),'app.py','exec'),ns)
    assert set(ns['_resolve_user_sites'](18)) == {'cz-one','pl-one'}
    c=connect();c.execute("INSERT INTO sites VALUES(5,'cz-new','CZ')");c.commit();c.close()
    assert set(ns['_resolve_user_sites'](18)) == {'cz-one','pl-one','cz-new'}
