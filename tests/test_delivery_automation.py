from unittest.mock import Mock

import pytest
import requests

import auto_confirm
import delivery_automation as da
from external_operations import ALLOWED_TRANSITIONS
from test_auto_confirm_returned import _database, _insert_order


class Connection:
    def __init__(self, db):
        self.db = db

    def execute(self, sql, params=()):
        if "pg_try_advisory_lock" in sql or "pg_advisory_unlock" in sql:
            return self.db.execute("SELECT 1")
        if "EXTRACT(EPOCH" in sql:
            return self.db.execute("SELECT 600")
        return self.db.execute(sql, params)

    def __getattr__(self, name):
        return getattr(self.db, name)


@pytest.fixture
def setup(monkeypatch):
    db = _database()
    for sql in [
        "ALTER TABLE sites ADD COLUMN id INTEGER DEFAULT 1",
        "ALTER TABLE sites ADD COLUMN consumer_key TEXT DEFAULT 'test-key'",
        "ALTER TABLE sites ADD COLUMN consumer_secret TEXT DEFAULT 'test-secret'",
        "ALTER TABLE orders ADD COLUMN delivery_confirmed_at TEXT",
        "ALTER TABLE orders ADD COLUMN delivery_confirmed_by INTEGER",
        "ALTER TABLE oms_order_fulfillment_state ADD COLUMN completion_sync_status TEXT",
        "ALTER TABLE oms_order_fulfillment_state ADD COLUMN updated_at TEXT",
    ]:
        db.execute(sql)
    db.executemany("INSERT INTO settings VALUES (?,?)", [
        (auto_confirm.ENABLE_KEY, "1"), (auto_confirm.SINCE_KEY, "2026-01-01")])
    _insert_order(db, "1-123", carrier_status="delivered")
    db.commit()
    ledger = {}

    def begin(conn, **kwargs):
        created = not ledger
        resumed = ledger.get("status") == "failed"
        if created or resumed:
            ledger.update(operation_id="operation", status="pending")
        return {**ledger, "should_execute": created or resumed}

    def transition(conn, op_id, target, **kwargs):
        current = ledger["status"]
        assert target == current or target in ALLOWED_TRANSITIONS[current]
        ledger.update(status=target, **kwargs)

    monkeypatch.setattr(da, "begin_operation", begin)
    monkeypatch.setattr(da, "transition_operation", transition)
    monkeypatch.setattr(da, "_record_phase", lambda conn, op, phase: ledger.update(phase=phase))
    monkeypatch.setattr(da, "_interrupted_before_write", lambda conn, op: ledger.get('phase') == 'preflight')
    yield Connection(db), ledger
    db.close()


def response(status="completed", http=200, order_id=123):
    return Mock(status_code=http, json=Mock(return_value={"id": order_id, "status": status}))


def test_completed_store_needs_no_write_and_repeated_task_adds_one_note(setup):
    conn, ledger = setup
    session = Mock(get=Mock(return_value=response()))
    assert da.process_delivered_order(conn, "1-123", session)["result"] == "confirmed"
    assert da.process_delivered_order(conn, "1-123", session)["result"] == "ineligible"
    assert session.put.call_count == 0
    assert ledger["status"] == "local_committed"
    assert conn.execute("SELECT COUNT(*) FROM order_notes").fetchone()[0] == 1
    assert conn.execute("SELECT status,delivery_confirmed FROM orders").fetchone()[:] == ("completed", 1)


@pytest.mark.parametrize("timeout", [False, True])
def test_single_write_is_confirmed_only_after_get_readback(setup, timeout):
    conn, ledger = setup
    session = Mock(get=Mock(side_effect=[response("on-hold"), response()]))
    session.put.side_effect = requests.Timeout() if timeout else None
    session.put.return_value = response()
    assert da.process_delivered_order(conn, "1-123", session)["result"] == "confirmed"
    assert session.put.call_count == 1
    assert session.get.call_count == 2
    assert ledger["status"] == "local_committed"


def test_lost_response_requires_read_reconciliation_never_second_put(setup):
    conn, ledger = setup
    session = Mock(get=Mock(side_effect=[response("on-hold"), requests.Timeout()]))
    session.put.side_effect = requests.Timeout()
    assert da.process_delivered_order(conn, "1-123", session)["result"] == "reconciliation_required"
    assert conn.execute("SELECT delivery_confirmed FROM orders").fetchone()[0] == 0
    retry = Mock(get=Mock(return_value=response("on-hold")))
    assert da.process_delivered_order(conn, "1-123", retry)["result"] == "reconciliation_required"
    retry.put.assert_not_called()
    retry.get.return_value = response()
    assert da.process_delivered_order(conn, "1-123", retry)["result"] == "confirmed"
    retry.put.assert_not_called()


@pytest.mark.parametrize("status", ["cancelled", "refunded", "failed", "pending"])
def test_remote_status_conflict_keeps_queue_untouched(setup, status):
    conn, ledger = setup
    session = Mock(get=Mock(return_value=response(status)))
    assert da.process_delivered_order(conn, "1-123", session)["result"] == "status_conflict"
    session.put.assert_not_called()
    assert conn.execute("SELECT delivery_confirmed FROM orders").fetchone()[0] == 0


def test_wrong_order_response_cannot_confirm(setup):
    conn, ledger = setup
    session = Mock(get=Mock(return_value=response(order_id=999)))
    assert da.process_delivered_order(conn, "1-123", session)["result"] == "read_failed"
    session.put.assert_not_called()


def test_definitive_429_preserves_local_state(setup):
    conn, ledger = setup
    session = Mock(get=Mock(return_value=response("shipped")), put=Mock(return_value=response(http=429)))
    assert da.process_delivered_order(conn, "1-123", session)["result"] == "write_rejected"
    assert ledger["status"] == "failed"
    assert conn.execute("SELECT delivery_confirmed FROM orders").fetchone()[0] == 0


@pytest.mark.parametrize("guard", ["switch", "manual_review", "return_flag", "in_transit"])
def test_noneligible_orders_never_contact_store(setup, guard):
    conn, ledger = setup
    if guard == "switch":
        conn.execute("UPDATE settings SET value='0' WHERE key=?", (auto_confirm.ENABLE_KEY,))
    elif guard == "manual_review":
        conn.execute("INSERT INTO oms_order_fulfillment_state(order_id,revision,aggregate_status) VALUES ('1-123',1,'manual_review')")
    elif guard == "return_flag":
        conn.execute("UPDATE orders SET is_problem_return=1")
    else:
        conn.execute("UPDATE orders SET carrier_status='in_transit'")
    conn.commit()
    session = Mock()
    assert da.process_delivered_order(conn, "1-123", session)["result"] == "ineligible"
    session.get.assert_not_called()
    session.put.assert_not_called()


def test_manual_exception_during_write_is_preserved(setup):
    conn, ledger = setup
    session = Mock(get=Mock(side_effect=[response("on-hold"), response()]))
    def write(*args, **kwargs):
        conn.execute("UPDATE orders SET is_problem_return=1")
        conn.commit()
        return response()
    session.put.side_effect = write
    assert da.process_delivered_order(conn, "1-123", session)["result"] == "eligibility_changed"
    assert conn.execute("SELECT delivery_confirmed,is_problem_return FROM orders").fetchone()[:] == (0, 1)
    assert ledger["status"] == "reconciliation_required"


def test_canonical_www_redirect_preserves_auth_and_write_uses_verified_url(setup):
    conn, ledger = setup
    redirect = Mock(status_code=301, headers={"Location": "https://www.pl.example/wp-json/wc/v3/orders/123"})
    session = Mock(get=Mock(side_effect=[redirect, response("on-hold"), response()]), put=Mock(return_value=response()))
    assert da.process_delivered_order(conn, "1-123", session)["result"] == "confirmed"
    assert session.get.call_args_list[1].kwargs["auth"] == ("test-key", "test-secret")
    assert session.put.call_args.args[0] == "https://www.pl.example/wp-json/wc/v3/orders/123"
    assert session.put.call_args.kwargs["allow_redirects"] is False


@pytest.mark.parametrize("target", [
    "https://attacker.example/wp-json/wc/v3/orders/123",
    "http://www.pl.example/wp-json/wc/v3/orders/123",
    "https://www.pl.example/wp-json/wc/v3/orders/999",
    "https://www.pl.example:444/wp-json/wc/v3/orders/123",
])
def test_untrusted_redirect_never_receives_credentials(setup, target):
    conn, ledger = setup
    session = Mock(get=Mock(return_value=Mock(status_code=301, headers={"Location": target})))
    assert da.process_delivered_order(conn, "1-123", session)["result"] == "read_failed"
    assert session.get.call_count == 1
    session.put.assert_not_called()
    assert ledger["error"] == "preflight:read_unsafe_redirect"


def test_sync_observing_completed_during_write_still_confirms_once(setup):
    conn, ledger = setup
    session = Mock(get=Mock(side_effect=[response("on-hold"), response()]))
    def write(*args, **kwargs):
        conn.execute("UPDATE orders SET status='completed'")
        conn.commit()
        return response()
    session.put.side_effect = write
    assert da.process_delivered_order(conn, "1-123", session)["result"] == "confirmed"
    assert conn.execute("SELECT delivery_confirmed FROM orders").fetchone()[0] == 1
    assert ledger["status"] == "local_committed"


@pytest.mark.parametrize("already_queued,expected", [(80, 0), (45, 3), (0, 48)])
def test_dispatcher_keeps_queue_bounded_for_regular_sync(monkeypatch, already_queued, expected):
    import redis
    import sync_tasks
    import carrier_outcome_bridge
    import fulfillment_outcome_recovery
    connection = Mock()
    broker = Mock(llen=Mock(return_value=already_queued), set=Mock(return_value=True))
    publish = Mock()
    monkeypatch.setattr(sync_tasks, "get_connection", lambda: connection)
    monkeypatch.setattr(auto_confirm, "is_enabled", lambda conn: True)
    monkeypatch.setattr(redis.Redis, "from_url", lambda url: broker)
    monkeypatch.setattr(carrier_outcome_bridge, "reconcile_cached_outcomes", lambda conn: 0)
    monkeypatch.setattr(fulfillment_outcome_recovery, "recover_terminal_states", lambda conn, **kw: [])
    monkeypatch.setattr(da, "dispatch_candidates", lambda conn, limit: [f"1-{i}" for i in range(100)])
    monkeypatch.setattr(sync_tasks.confirm_delivered_order, "apply_async", publish)
    result = sync_tasks.auto_confirm_delivered.run()
    assert result["dispatched"] == expected
    assert publish.call_count == expected
    connection.close.assert_called_once()


def test_worker_interrupted_in_preflight_is_safe_to_retry(setup):
    conn, ledger = setup
    ledger.update(operation_id='operation', status='pending', phase='preflight')
    session = Mock(get=Mock(side_effect=[response('on-hold'), response()]), put=Mock(return_value=response()))
    assert da.process_delivered_order(conn, '1-123', session)['result'] == 'retry_ready'
    session.get.assert_not_called()
    session.put.assert_not_called()
    assert ledger['status'] == 'failed'
    assert da.process_delivered_order(conn, '1-123', session)['result'] == 'confirmed'
    assert session.put.call_count == 1


def test_worker_interrupted_after_write_boundary_only_reconciles(setup):
    conn, ledger = setup
    ledger.update(operation_id='operation', status='pending', phase='write_started')
    session = Mock(get=Mock(return_value=response('on-hold')))
    assert da.process_delivered_order(conn, '1-123', session)['result'] == 'reconciliation_required'
    session.put.assert_not_called()
    assert ledger['status'] == 'reconciliation_required'
