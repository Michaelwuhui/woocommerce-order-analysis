"""Real Flask route and PostgreSQL ledger; all external requests are replaced."""
import copy
import itertools
import json
import os
import threading
import time
from concurrent.futures import ThreadPoolExecutor

import pytest
import requests
import db_backend as db

pytestmark = pytest.mark.skipif(not db.is_postgres_backend(), reason='isolated PostgreSQL required')
_ids = itertools.count(900000)


class Response:
    def __init__(self, data, status=200):
        self.data = data
        self.status_code = status
        self.text = json.dumps(data)

    def json(self):
        return copy.deepcopy(self.data)


@pytest.fixture()
def case(monkeypatch):
    from sync_service import get_connection
    test_database = os.environ.get('WOO_DB_NAME_OVERRIDE', '')
    assert test_database.startswith('woo_return_loss_test_'), 'Use an empty isolated test database'
    monkeypatch.setattr(requests.sessions.Session, 'request', lambda *a, **k: pytest.fail('real network forbidden'))
    import app as m
    m.app.config.update(TESTING=True)
    woo_id = next(_ids)
    oid = f'20-{woo_id}'
    items = [{'id': 1179, 'product_id': 31263, 'variation_id': 31268, 'quantity': 1, 'sku': 'RETRY-TEST', 'name': 'Test item'}]
    c = get_connection()
    assert c.execute('SELECT current_database()').fetchone()[0] == test_database
    c.execute("INSERT INTO sites(id,url,consumer_key,consumer_secret,country) "
              "VALUES (20,'https://shipping-test.invalid','synthetic-key','synthetic-secret','PL') "
              "ON CONFLICT(id) DO NOTHING")
    for uid, role, ship in [(10, 'admin', True), (11, 'user', True), (12, 'user', False)]:
        c.execute("INSERT INTO users(id,username,password_hash,name,role,can_ship) "
                  "VALUES (?,?,'unusable','Shipping retry test',?,?) ON CONFLICT(id) DO NOTHING",
                  (uid, f'ship-retry-test-{uid}', role, ship))
    c.execute("INSERT INTO user_country_permissions(user_id,country) VALUES (12,'PL') "
              "ON CONFLICT DO NOTHING")
    for table in ('order_notes', 'shipping_logs', 'external_operations', 'orders'):
        c.execute(f'DELETE FROM {table} WHERE {"id" if table == "orders" else "order_id"}=?', (oid,))
    c.execute('''INSERT INTO orders(id,number,source,status,line_items,meta_data,shipping_lines,billing,shipping)
                 VALUES (?,?,?,'processing',?,'[]','[]','{}','{}')''', (oid, str(woo_id), 'https://shipping-test.invalid', json.dumps(items)))
    c.commit(); c.close()
    monkeypatch.setattr(m, 'detect_site_tracking_format', lambda *a: 'ast')
    data = {'id': woo_id, 'status': 'processing', 'line_items': items, 'meta_data': []}
    state = {'remote': data, 'reject': True, 'get_status': 200, 'writes': 0, 'reads': 0, 'payloads': [], 'delay': 0}
    lock = threading.Lock()

    def get(url, **kwargs):
        assert url == f'https://shipping-test.invalid/wp-json/wc/v3/orders/{woo_id}'
        with lock:
            state['reads'] += 1
            return Response(state['remote'], state['get_status'])

    def post(url, **kwargs):
        assert url == f'https://shipping-test.invalid/wp-json/wc-ast-pro/v3/orders/{woo_id}/shipment-trackings', 'unexpected email or external write'
        with lock:
            state['writes'] += 1
            state['payloads'].append(kwargs['json'])
            reject = state['reject']
        if state['delay']: time.sleep(state['delay'])
        if reject: return Response({'code': 'rest_forbidden', 'message': 'rejected'}, 401)
        with lock:
            state['remote']['status'] = 'shipped'
            state['remote']['meta_data'] = [{
                'key': '_wc_shipment_tracking_items', 'value': [{'tracking_number': f'TRACK-{woo_id}'}],
            }]
        return Response({'tracking_id': 'synthetic-tracking'})

    monkeypatch.setattr(requests, 'get', get)
    monkeypatch.setattr(requests, 'post', post)
    monkeypatch.setattr(requests, 'put', lambda *a, **k: pytest.fail('unexpected PUT'))
    payload = {'order_id': oid, 'tracking_number': f'TRACK-{woo_id}', 'carrier_slug': 'dpd', 'send_email': False}

    def submit(changes=None, uid='10'):
        client = m.app.test_client()
        with client.session_transaction() as session:
            session['_user_id'] = uid; session['_fresh'] = True
        return client.post('/api/shipping/ship', json={**payload, **(changes or {})})

    def snapshot():
        c = get_connection()
        try:
            return {
                'logs': [dict(r) for r in c.execute('SELECT id,status,tracking_number,shipped_at FROM shipping_logs WHERE order_id=?', (oid,))],
                'ops': [dict(r) for r in c.execute('SELECT status,attempts FROM external_operations WHERE order_id=?', (oid,))],
                'order_status': c.execute('SELECT status FROM orders WHERE id=?', (oid,)).fetchone()[0],
            }
        finally: c.close()

    def failed():
        response = submit()
        assert response.status_code == 422, response.get_json()
        assert snapshot()['ops'][0]['status'] == 'failed'
        c = get_connection()
        c.execute("UPDATE shipping_logs SET shipped_at='2020-01-02 10:00:00' WHERE order_id=?", (oid,))
        c.commit(); c.close()
        state['reject'] = False

    return {'submit': submit, 'failed': failed, 'state': state, 'snapshot': snapshot, 'oid': oid, 'connection': get_connection}


def test_rejected_submission_can_retry_original_parcel_once(case):
    case['failed']()
    response = case['submit']()
    assert response.status_code == 200, response.get_json()
    after = case['snapshot']()
    assert after['ops'] == [{'status': 'local_committed', 'attempts': 2}]
    assert len(after['logs']) == 1 and after['logs'][0]['status'] == 'shipped'
    assert str(after['logs'][0]['shipped_at']) == '2020-01-02 10:00:00'
    assert case['state']['payloads'][-1]['date_shipped'] == '2020-01-02'
    assert after['order_status'] == 'shipped'
    assert case['submit']().get_json()['idempotent'] is True
    assert case['state']['writes'] == 2


@pytest.mark.parametrize('status', ['pending', 'reconciliation_required', 'external_success', 'local_committed', 'notified'])
def test_non_failed_ledger_cannot_replay_pending_parcel(case, status):
    case['failed']()
    c = case['connection'](); c.execute('UPDATE external_operations SET status=? WHERE order_id=?', (status, case['oid'])); c.commit(); c.close()
    assert case['submit']().status_code == 409
    assert case['state']['writes'] == 1


def test_pending_record_without_ledger_cannot_be_retried(case):
    case['failed']()
    c = case['connection'](); c.execute('DELETE FROM external_operations WHERE order_id=?', (case['oid'],)); c.commit(); c.close()
    assert case['submit']().status_code == 409
    assert case['state']['writes'] == 1


@pytest.mark.parametrize('change', [{'tracking_number': 'DIFFERENT'}, {'carrier_slug': 'inpost'}, {'new_parcel': True}, {'more_batches': True}, {'reship_reason': 'new parcel'}, {'shipped_items': [{'item_id': '1179', 'qty': 2}]}])
def test_retry_cannot_be_changed_into_another_shipment(case, change):
    case['failed']()
    assert case['submit'](change).status_code in {400, 409}
    assert case['state']['writes'] == 1


def test_preflight_reconciles_saved_parcel_without_write_or_email(case):
    case['failed']()
    case['state']['remote']['status'] = 'shipped'
    case['state']['remote']['meta_data'] = [{'key': '_tracking_number', 'value': 'TRACK-'+case['oid'].split('-')[1]}]
    response = case['submit']({'send_email': True})
    assert response.status_code == 200, response.get_json()
    assert case['state']['writes'] == 1
    assert case['snapshot']()['logs'][0]['status'] == 'shipped'


@pytest.mark.parametrize('change', ['quantity', 'tracking', 'http_error', 'cancelled'])
def test_preflight_failure_does_not_send_a_new_shipment(case, change):
    case['failed']()
    if change == 'quantity': case['state']['remote']['line_items'][0]['quantity'] = 2
    elif change == 'tracking': case['state']['remote']['meta_data'] = [{'key': '_tracking_number', 'value': 'OTHER'}]
    elif change == 'http_error': case['state']['get_status'] = 503
    else: case['state']['remote']['status'] = 'cancelled'
    assert case['submit']().status_code == 409
    assert case['state']['writes'] == 1
    assert case['snapshot']()['logs'][0]['status'] == 'pending_sync'


def test_country_and_shipping_permissions_still_apply(case):
    assert case['submit'](uid='11').status_code == 403
    assert case['submit'](uid='12').status_code == 403
    assert case['state']['writes'] == 0


def test_two_concurrent_retries_create_only_one_parcel(case):
    case['failed'](); case['state']['delay'] = 0.15
    with ThreadPoolExecutor(max_workers=2) as pool:
        responses = list(pool.map(lambda _: case['submit'](), range(2)))
    assert any(r.status_code == 200 for r in responses)
    assert all(r.status_code in {200, 409} for r in responses)
    assert case['state']['writes'] == 2
    assert len(case['snapshot']()['logs']) == 1
