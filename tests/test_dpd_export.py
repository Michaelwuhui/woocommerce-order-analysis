import ast
import io
import json
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from openpyxl import load_workbook

from dpd_export import find_recent_duplicate_order_ids, site_order_reference
from shipping_export import build_dpd_shipping_workbook


ROOT = Path(__file__).resolve().parents[1]


def _record(order_id, created, email='', phone='', address='Street 1', postcode='00-001', city='Warszawa'):
    return {
        'id': order_id,
        'date_created': created.isoformat(),
        'billing': {
            'email': email,
            'phone': phone,
        },
        'shipping': {
            'address_1': address,
            'postcode': postcode,
            'city': city,
            'phone': phone,
        },
        'meta_data': [],
    }


def test_duplicate_window_flags_both_pending_orders_at_72_hours():
    start = datetime(2026, 8, 20, 10, 0)
    rows = [
        _record('first', start, email='buyer@example.com', phone='+48 501 002 003'),
        _record('second', start + timedelta(hours=72), email='buyer@example.com', phone='501002003'),
    ]

    assert find_recent_duplicate_order_ids(rows, {'first', 'second'}) == {'first', 'second'}


def test_duplicate_window_ignores_orders_after_72_hours_and_placeholder_phones():
    start = datetime(2026, 8, 20, 10, 0)
    rows = [
        _record('first', start, phone='123456789', address='Street 1'),
        _record('later', start + timedelta(hours=72, seconds=1), phone='123456789', address='Street 2'),
    ]

    assert find_recent_duplicate_order_ids(rows, {'first', 'later'}) == set()


def test_duplicate_window_matches_dpd_fallback_address():
    start = datetime(2026, 8, 20, 10, 0)
    rows = [
        _record('first', start, address='', postcode='', city=''),
        _record('second', start + timedelta(hours=1), address='', postcode='', city=''),
    ]
    for row in rows:
        row['shipping'] = {}
        row['meta_data'] = [
            {'key': '_billing_adres_dpd', 'value': 'Dluga'},
            {'key': '_billing_numer_domu', 'value': '10'},
            {'key': '_billing_kod_pocztowy', 'value': '00-001'},
            {'key': '_billing_miejscowosc', 'value': 'Warszawa'},
        ]

    assert find_recent_duplicate_order_ids(rows, {'first', 'second'}) == {'first', 'second'}


def test_customer_number_uses_requested_vapesklep_abbreviation():
    assert site_order_reference('https://www.vapesklep.pl', '#12345') == 'vsklep12345'
    assert site_order_reference('https://vapeklub.pl', 'PL-8') == 'vapeklubPL-8'


def test_shipping_template_exposes_filtered_poland_dpd_export():
    template = (ROOT / 'templates' / 'shipping.html').read_text(encoding='utf-8')

    assert 'id="exportDpdBtn"' in template
    assert "'/api/shipping/export/dpd'" in template
    assert 'country && country !== \'PL\'' in template


def test_dpd_endpoint_excludes_big_orders_and_reports_counts(tmp_path):
    db_path = tmp_path / 'orders.db'
    conn = sqlite3.connect(db_path)
    conn.executescript(
        """
        CREATE TABLE sites (url TEXT PRIMARY KEY, country TEXT);
        CREATE TABLE orders (
            id TEXT PRIMARY KEY, date_created TEXT, billing TEXT, shipping TEXT,
            meta_data TEXT, status TEXT, source TEXT
        );
        INSERT INTO sites(url, country) VALUES ('https://vapesklep.pl', 'PL');
        """
    )
    for order_id, hour in [('eligible', 0), ('big', 100)]:
        conn.execute(
            """INSERT INTO orders(id, date_created, billing, shipping, meta_data, status, source)
               VALUES (?, ?, ?, ?, '[]', 'processing', 'https://vapesklep.pl')""",
            (
                order_id,
                (datetime(2026, 8, 20) + timedelta(hours=hour)).isoformat(),
                json.dumps({'email': f'{order_id}@example.com', 'phone': '501002003'}),
                json.dumps({
                    'first_name': 'Jan', 'last_name': order_id,
                    'address_1': f'Street {order_id}', 'postcode': '00-001',
                    'city': 'Warszawa', 'phone': f'5010020{hour:02d}',
                }),
            ),
        )
    conn.commit()
    conn.close()

    pending_orders = [{
        'id': order_id,
        'number': '101' if order_id == 'eligible' else '102',
        'source_url': 'https://vapesklep.pl',
        'shipping_method': 'DPD',
        'currency': 'PLN',
        'payment_method': 'cod',
        'is_big_order': order_id == 'big',
        'manual_review': False,
        'has_shortage': False,
        'shipment_sync_pending': False,
        'parcels_shipped': 0,
        'customer_name': f'Jan {order_id}',
        'customer_city': 'Warszawa',
        'customer_street_address': f'Street {order_id}',
        'customer_phone': f'5010020{hour:02d}',
        'customer_email': f'{order_id}@example.com',
        'customer_postcode': '00-001',
        'total': 249.9,
    } for order_id, hour in [('eligible', 0), ('big', 100)]]

    tree = ast.parse((ROOT / 'app.py').read_text(encoding='utf-8'))
    node = next(
        item for item in tree.body
        if isinstance(item, ast.FunctionDef) and item.name == 'export_dpd_pending_list'
    )
    node.decorator_list = []

    def connect():
        db = sqlite3.connect(db_path)
        db.row_factory = sqlite3.Row
        return db

    namespace = {
        'request': request,
        'jsonify': jsonify,
        'send_file': send_file,
        'datetime': datetime,
        'get_pending_orders': lambda: jsonify(pending_orders),
        'get_db_connection': connect,
        'find_recent_duplicate_order_ids': find_recent_duplicate_order_ids,
        'site_order_reference': site_order_reference,
        'build_dpd_shipping_workbook': build_dpd_shipping_workbook,
    }
    exec(compile(ast.Module(body=[node], type_ignores=[]), 'app.py', 'exec'), namespace)

    flask_app = Flask('dpd-export-test')
    with flask_app.test_request_context('/api/shipping/export/dpd?country=PL'):
        response = namespace['export_dpd_pending_list']()

    assert response.status_code == 200
    assert response.headers['X-Export-Row-Count'] == '1'
    assert response.headers['X-DPD-Excluded-Count'] == '1'
    assert response.headers['X-DPD-Excluded-Big-Order-Count'] == '1'
    response.direct_passthrough = False
    workbook = load_workbook(io.BytesIO(response.get_data()), data_only=False)
    worksheet = workbook['下单模板']
    assert worksheet['A2'].value == 'vsklep101'
    assert worksheet['T2'].value == 249.9
