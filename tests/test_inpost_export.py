import ast
import json
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path

from flask import Flask, jsonify, request, send_file

from dpd_export import find_recent_duplicate_order_ids, site_order_reference
from inpost_export import (
    build_inpost_txt,
    extract_inpost_locker_code,
    format_collection_amount,
    normalize_polish_phone,
)


ROOT = Path(__file__).resolve().parents[1]


def test_inpost_txt_matches_supplied_sample_exactly():
    output = build_inpost_txt([
        {
            'email': 'bskacper555@gmail.com',
            'phone': '+48510661436',
            'customer_number': 'sk-13094',
            'cod_amount': '327.40',
            'locker_code': 'SOS01N',
        },
        {
            'email': 'martynakoselskaaa@gmail.comm',
            'phone': '+48666227344',
            'customer_number': 'sk-13093',
            'cod_amount': '100.60',
            'locker_code': 'ZUVA01M',
        },
    ])

    assert output.getvalue() == (
        b'bskacper555@gmail.com;+48510661436;A;paczkomat;sk-13094;327.4;SOS01N\r\n'
        b'martynakoselskaaa@gmail.comm;+48666227344;A;paczkomat;sk-13093;100.6;ZUVA01M'
    )


def test_inpost_fields_are_normalized_for_partner_import():
    assert normalize_polish_phone('510 661 436') == '+48510661436'
    assert normalize_polish_phone('48 510-661-436') == '+48510661436'
    assert normalize_polish_phone('0048 510 661 436') == '+48510661436'
    assert normalize_polish_phone('123') == ''
    assert extract_inpost_locker_code('SOS01N, Locker Street 1') == 'SOS01N'
    assert extract_inpost_locker_code('Paczkomat ZUVA01M, Address') == 'ZUVA01M'
    assert format_collection_amount('327.40') == '327.4'
    assert format_collection_amount('100.00') == '100'
    assert format_collection_amount(100) == '100'
    assert format_collection_amount('invalid') == ''


def test_shipping_template_exposes_filtered_poland_inpost_export():
    template = (ROOT / 'templates' / 'shipping.html').read_text(encoding='utf-8')

    assert 'id="exportInpostBtn"' in template
    assert "'/api/shipping/export/inpost'" in template
    assert "country && country !== 'PL'" in template


def test_inpost_endpoint_excludes_review_and_incomplete_orders(tmp_path):
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
    order_ids = ('eligible', 'review', 'missing-locker')
    for index, order_id in enumerate(order_ids):
        conn.execute(
            """INSERT INTO orders(id, date_created, billing, shipping, meta_data, status, source)
               VALUES (?, ?, ?, ?, '[]', 'processing', 'https://vapesklep.pl')""",
            (
                order_id,
                (datetime(2026, 8, 20) + timedelta(days=index * 5)).isoformat(),
                json.dumps({
                    'email': f'{order_id}@example.com',
                    'phone': f'51066143{index}',
                }),
                json.dumps({
                    'address_1': f'Street {index}',
                    'postcode': f'00-00{index}',
                    'city': f'City {index}',
                    'phone': f'51066143{index}',
                }),
            ),
        )
    conn.commit()
    conn.close()

    pending_orders = []
    for index, order_id in enumerate(order_ids):
        pending_orders.append({
            'id': order_id,
            'number': str(101 + index),
            'source_url': 'https://vapesklep.pl',
            'shipping_method': 'InPost Paczkomat',
            'currency': 'PLN',
            'payment_method': 'cod',
            'is_big_order': False,
            'manual_review': order_id == 'review',
            'has_shortage': False,
            'shipment_sync_pending': False,
            'high_risk_postcode': '',
            'parcels_shipped': 0,
            'customer_phone': f'51066143{index}',
            'customer_email': f'{order_id}@example.com',
            'customer_inpost_id': '' if order_id == 'missing-locker' else 'SOS01N, Address',
            'total': 249.90,
        })

    tree = ast.parse((ROOT / 'app.py').read_text(encoding='utf-8'))
    node = next(
        item for item in tree.body
        if isinstance(item, ast.FunctionDef) and item.name == 'export_inpost_pending_list'
    )
    node.decorator_list = []

    def connect():
        database = sqlite3.connect(db_path)
        database.row_factory = sqlite3.Row
        return database

    namespace = {
        'request': request,
        'jsonify': jsonify,
        'send_file': send_file,
        'datetime': datetime,
        'get_pending_orders': lambda: jsonify(pending_orders),
        'get_db_connection': connect,
        'find_recent_duplicate_order_ids': find_recent_duplicate_order_ids,
        'site_order_reference': site_order_reference,
        'build_inpost_txt': build_inpost_txt,
        'extract_inpost_locker_code': extract_inpost_locker_code,
        'format_collection_amount': format_collection_amount,
        'normalize_polish_phone': normalize_polish_phone,
    }
    exec(compile(ast.Module(body=[node], type_ignores=[]), 'app.py', 'exec'), namespace)

    flask_app = Flask('inpost-export-test')
    with flask_app.test_request_context('/api/shipping/export/inpost?country=PL'):
        response = namespace['export_inpost_pending_list']()

    assert response.status_code == 200
    assert response.headers['X-Export-Row-Count'] == '1'
    assert response.headers['X-InPost-Candidate-Count'] == '3'
    assert response.headers['X-InPost-Excluded-Count'] == '2'
    assert response.headers['X-InPost-Excluded-Manual-Review-Count'] == '1'
    assert response.headers['X-InPost-Excluded-Incomplete-Count'] == '1'
    response.direct_passthrough = False
    assert response.get_data() == (
        b'eligible@example.com;+48510661430;A;paczkomat;vsklep101;249.9;SOS01N'
    )
