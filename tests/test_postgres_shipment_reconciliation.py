"""Run only against an empty, isolated PostgreSQL schema copy."""
import concurrent.futures
from copy import deepcopy
import json
import os
import threading

import pytest
import requests

import db_backend as db
import shipment_reconciliation as rec
from external_operations import begin_operation
from sync_service import get_connection
from test_shipment_reconciliation import parcel, source

pytestmark = pytest.mark.skipif(not db.is_postgres_backend(), reason='isolated PostgreSQL required')
OID = '99001-123'


@pytest.fixture
def operation(monkeypatch):
    assert os.environ.get('WOO_DB_NAME_OVERRIDE', '').startswith('woo_reconcile_test_')
    monkeypatch.setattr(requests.sessions.Session, 'request',
                        lambda *a, **k: pytest.fail('Real network calls are forbidden in tests'))
    c = get_connection()
    for table in ('order_notes', 'shipping_logs', 'external_operations'):
        c.execute(f'DELETE FROM {table} WHERE order_id=?', (OID,))
    c.execute('DELETE FROM orders WHERE id=?', (OID,))
    c.execute("INSERT INTO sites(id,url,consumer_key,consumer_secret,country) VALUES"
              "(99001,'https://reconciliation-test.invalid','test','test','PL') ON CONFLICT(id) DO NOTHING")
    c.execute("INSERT INTO users(id,username,password_hash,name,role,can_ship,can_view_shipping) VALUES"
              "(99001,'reconcile-test','unusable','Test','user',true,true) ON CONFLICT(id) DO NOTHING")
    c.execute("INSERT INTO orders(id,number,source,status,line_items,meta_data,shipping_lines,date_modified) "
              "VALUES (?,'123','https://reconciliation-test.invalid','processing',?,'[]','[]','2026-09-16T10:00:00')",
              (OID, json.dumps(source()['line_items'])))
    p = parcel(); p.update(order_id=OID, site_id=99001)
    op = begin_operation(c, operation_type='ship_order', order_id=OID, site_id=99001,
                         request_payload=p, created_by='99001')
    c.execute("UPDATE external_operations SET status='reconciliation_required',created_at=CURRENT_TIMESTAMP-interval '1 hour',"
              "updated_at=CURRENT_TIMESTAMP-interval '10 minutes' WHERE operation_id=?", (op['operation_id'],))
    c.execute("INSERT INTO shipping_logs(order_id,woo_order_id,source,tracking_number,carrier_slug,shipped_by,"
              "shipped_at,items_json,is_partial,status) VALUES (?,'123','https://reconciliation-test.invalid',"
              "'TRACK-123','inpost',99001,'2026-09-17 03:00:00',?,0,'pending_sync')", (OID, json.dumps(p['items'])))
    c.commit(); c.close()
    yield op['operation_id']


class Response:
    status_code = 200
    def __init__(self, body=None): self.body = body if body is not None else source()
    def json(self): return deepcopy(self.body)


def rows():
    c = get_connection()
    try:
        return {'order': dict(c.execute('SELECT * FROM orders WHERE id=?', (OID,)).fetchone()),
                'logs': [dict(r) for r in c.execute('SELECT * FROM shipping_logs WHERE order_id=? ORDER BY id', (OID,))],
                'op': dict(c.execute('SELECT * FROM external_operations WHERE order_id=?', (OID,)).fetchone()),
                'notes': [dict(r) for r in c.execute('SELECT * FROM order_notes WHERE order_id=?', (OID,))],
                'movements': c.execute('SELECT count(*) FROM inv_movements WHERE order_id=?', (OID,)).fetchone()[0]}
    finally: c.close()


def execute(sql, args=()):
    c = get_connection()
    try: c.execute(sql, args); c.commit()
    finally: c.close()


def test_verified_remote_is_atomic_and_repeatable_without_side_effects(operation):
    before = rows(); calls = []
    def fetch(url, **kwargs):
        calls.append(url)
        assert url.endswith('/orders/123') and kwargs['allow_redirects'] is False
        return Response()
    assert rec.reconcile_operation(operation, fetch=fetch)['outcome'] == 'verified'
    after = rows()
    assert after['order']['status'] == 'on-hold'
    assert after['op']['status'] == 'local_committed'
    assert len(after['logs']) == 1 and after['logs'][0]['status'] == 'shipped'
    assert after['logs'][0]['shipped_at'] == before['logs'][0]['shipped_at']
    assert after['logs'][0]['shipped_by'] == before['logs'][0]['shipped_by']
    assert after['movements'] == before['movements'] == 0
    assert len(after['notes']) == 1 and not after['notes'][0]['customer_note']
    assert rec.reconcile_operation(operation, fetch=fetch)['outcome'] == 'already_closed'
    assert len(calls) == 1 and rows() == after


@pytest.mark.parametrize('field', ["status='completed'", 'delivery_confirmed=true'])
def test_does_not_downgrade_delivered_or_completed_order(operation, field):
    execute('UPDATE orders SET ' + field + ' WHERE id=?', (OID,))
    assert rec.reconcile_operation(operation, fetch=lambda *a, **k: Response())['outcome'] == 'verified'
    assert rows()['order']['status'] == 'completed'


def test_recovers_missing_local_log_after_process_interruption(operation):
    execute('DELETE FROM shipping_logs WHERE order_id=?', (OID,))
    assert rec.reconcile_operation(operation, fetch=lambda *a, **k: Response())['outcome'] == 'verified'
    state = rows()
    assert len(state['logs']) == 1 and state['logs'][0]['shipped_by'] == 99001
    assert state['logs'][0]['items_json']
    assert state['logs'][0]['shipped_at'].replace(' ', 'T') == state['op']['created_at'][:19]


def test_absent_remote_keeps_uncertainty_and_backs_off(operation):
    remote = source('none'); remote['status'] = 'processing'
    before = rows()
    assert operation in rec.due_operations([OID])
    assert rec.reconcile_operation(operation, fetch=lambda *a, **k: Response(remote))['outcome'] == 'remote_absent'
    after = rows()
    assert after['order'] == before['order'] and after['logs'] == before['logs'] and not after['notes']
    assert after['op']['status'] == 'reconciliation_required'
    c = get_connection()
    try:
        visible = rec.public_statuses(c, [OID])[OID]
        assert visible['label'] == '需处理' and not visible['retry_allowed']
    finally: c.close()
    assert operation not in rec.due_operations([OID])
    assert rec.reconcile_operation(operation, fetch=lambda *a, **k: pytest.fail('backoff ignored'))['outcome'] == 'not_due'


@pytest.mark.parametrize('failure', ['timeout', 'http429', 'http401', 'bad_json'])
def test_transient_reads_never_change_business_records(operation, failure):
    before = rows()
    def fetch(*a, **k):
        if failure == 'timeout': raise requests.Timeout('secret URL must not be logged')
        r = Response()
        if failure.startswith('http'): r.status_code = int(failure[4:])
        if failure == 'bad_json': r.json = lambda: json.loads('bad secret payload')
        return r
    result = rec.reconcile_operation(operation, fetch=fetch)
    assert result['outcome'] in {'read_failed', 'needs_review'}
    after = rows()
    assert after['order'] == before['order'] and after['logs'] == before['logs'] and not after['notes']
    assert 'secret' not in after['op']['last_error']


def test_fresh_request_and_definitive_failure_are_not_auto_replayed(operation):
    execute('UPDATE external_operations SET updated_at=CURRENT_TIMESTAMP WHERE operation_id=?', (operation,))
    assert rec.due_operations([OID]) == []
    assert rec.reconcile_operation(operation, fetch=lambda *a, **k: pytest.fail('too early'))['outcome'] == 'not_due'
    execute("UPDATE external_operations SET status='failed' WHERE operation_id=?", (operation,))
    assert rec.reconcile_operation(operation, fetch=lambda *a, **k: pytest.fail('failed replay'))['outcome'] == 'already_closed'


def test_concurrent_order_edit_during_network_read_defers_recovery(operation):
    def fetch(*a, **k):
        execute("UPDATE orders SET status='cancelled' WHERE id=?", (OID,))
        return Response()
    assert rec.reconcile_operation(operation, fetch=fetch)['outcome'] == 'changed_during_check'
    assert rows()['logs'][0]['status'] == 'pending_sync'


def test_duplicate_workers_only_complete_original_once(operation):
    entered, release = threading.Event(), threading.Event()
    def fetch(*a, **k): entered.set(); assert release.wait(8); return Response()
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
        first = pool.submit(rec.reconcile_operation, operation, fetch=fetch)
        assert entered.wait(8)
        try: assert rec.reconcile_operation(operation, fetch=lambda *a, **k: pytest.fail('duplicate fetch'))['outcome'] == 'busy'
        finally: release.set()
        assert first.result(8)['outcome'] == 'verified'
    assert len(rows()['notes']) == 1


def test_failure_after_local_updates_rolls_everything_back(operation, monkeypatch):
    original = rec.get_connection
    class Faulty:
        def __init__(self): self.c = original()
        def __getattr__(self, key): return getattr(self.c, key)
        def execute(self, sql, args=()):
            if 'INSERT INTO order_notes' in sql: raise RuntimeError('simulated commit path failure')
            return self.c.execute(sql, args)
    before = rows()
    monkeypatch.setattr(rec, 'get_connection', Faulty)
    with pytest.raises(RuntimeError): rec.reconcile_operation(operation, fetch=lambda *a, **k: Response())
    assert rows() == before


def test_older_source_version_is_visible_and_does_not_overwrite_mirror(operation):
    r = source(); r['date_modified'] = '2026-09-15T00:00:00'
    assert rec.reconcile_operation(operation, fetch=lambda *a, **k: Response(r))['outcome'] == 'needs_review'
    assert rows()['order']['status'] == 'processing'


def test_celery_sweep_publishes_due_operations_to_fetch_queue(operation, monkeypatch):
    import shipment_reconciliation_tasks as tasks
    from celery_app import celery_app
    sent = []
    monkeypatch.setattr(tasks.reconcile_shipment, 'apply_async', lambda **kwargs: sent.append(kwargs))
    assert tasks.scan_shipments.run()['queued'] == 1
    assert sent == [{'args': [operation], 'expires': 300}]
    assert celery_app.conf.beat_schedule['shipment-result-reconciliation']['schedule'] == 60
    assert celery_app.conf.task_routes[tasks.reconcile_shipment.name]['queue'] == 'sync_fetch'


def test_broker_outage_keeps_persisted_operation_discoverable(operation, monkeypatch):
    import shipment_reconciliation_tasks as tasks
    def unavailable(**kwargs): raise ConnectionError('test broker offline')
    monkeypatch.setattr(tasks.reconcile_shipment, 'apply_async', unavailable)
    tasks.enqueue_after_uncertain(operation)
    assert operation in rec.due_operations()


def test_regular_sync_queues_reconciliation_after_commit(operation, monkeypatch):
    import sync_utils
    import shipment_reconciliation_tasks as tasks
    monkeypatch.setattr(sync_utils, '_enqueue_fulfillment_plans', lambda *a, **k: None)
    monkeypatch.setattr(sync_utils, '_enqueue_order_notifications', lambda *a, **k: None)
    seen = []
    monkeypatch.setattr(tasks, 'enqueue_orders', lambda ids: seen.extend(ids))
    sync_utils.run_post_commit_sync_actions([{'order_id': OID}], strict=True)
    assert seen == [OID]
