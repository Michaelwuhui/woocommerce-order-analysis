import sqlite3
from datetime import datetime, timedelta
from pathlib import Path

import auto_confirm
from celery_app import celery_app


ROOT = Path(__file__).resolve().parents[1]


def _database():
    conn = sqlite3.connect(":memory:")
    conn.row_factory = sqlite3.Row
    conn.executescript(
        """
        CREATE TABLE settings (key TEXT PRIMARY KEY, value TEXT);
        CREATE TABLE sites (url TEXT PRIMARY KEY, country TEXT);
        CREATE TABLE orders (
            id TEXT PRIMARY KEY,
            number TEXT,
            source TEXT,
            status TEXT,
            payment_method TEXT,
            date_created TEXT,
            date_modified TEXT,
            shipping_total REAL,
            currency TEXT,
            is_undelivered INTEGER DEFAULT 0,
            shipping_loss_amount REAL DEFAULT 0,
            undelivered_at TEXT,
            undelivered_by INTEGER,
            undelivered_note TEXT,
            is_problem_return INTEGER DEFAULT 0,
            delivery_confirmed INTEGER DEFAULT 0,
            carrier_status TEXT,
            carrier_status_at TEXT
        );
        CREATE TABLE shipping_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id TEXT,
            tracking_number TEXT,
            shipped_at TEXT
        );
        CREATE TABLE order_notes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id TEXT,
            note TEXT,
            date_created TEXT,
            customer_note INTEGER,
            author TEXT,
            added_by_user INTEGER
        );
        CREATE TABLE oms_order_fulfillment_state (
            order_id TEXT PRIMARY KEY,
            revision INTEGER,
            aggregate_status TEXT,
            manual_review INTEGER DEFAULT 0,
            has_shortage INTEGER DEFAULT 0
        );
        CREATE TABLE oms_domain_events (
            aggregate_type TEXT, aggregate_id TEXT, event_type TEXT,
            to_status TEXT, actor_type TEXT
        );
        CREATE TABLE oms_fulfillments (
            id TEXT PRIMARY KEY,
            order_id TEXT,
            revision INTEGER,
            status TEXT
        );
        CREATE TABLE oms_shipments (
            id TEXT PRIMARY KEY,
            fulfillment_id TEXT,
            status TEXT
        );
        """
    )
    conn.execute("INSERT INTO sites(url,country) VALUES ('https://pl.example','PL')")
    conn.execute(
        "INSERT INTO settings(key,value) VALUES (?,?)",
        (auto_confirm.RETURNED_ENABLE_KEY, '1'),
    )
    conn.execute(
        "INSERT INTO settings(key,value) VALUES (?,?)",
        (auto_confirm.RETURNED_SINCE_KEY, '2026-09-01 00:00:00'),
    )
    return conn


def _insert_order(
    conn,
    order_id,
    *,
    carrier_status='returned',
    shipping=25,
    currency='PLN',
    source='https://pl.example',
    age_days=10,
    is_problem_return=0,
    delivery_confirmed=0,
):
    stamp = (datetime.now() - timedelta(days=age_days)).strftime('%Y-%m-%d %H:%M:%S')
    conn.execute(
        """
        INSERT INTO orders (
            id,number,source,status,payment_method,date_created,date_modified,
            shipping_total,currency,is_problem_return,delivery_confirmed,
            carrier_status,carrier_status_at
        ) VALUES (?,?,?,'shipped','cod',?,?,?,?,?,?,?,?)
        """,
        (
            order_id,
            order_id,
            source,
            stamp,
            stamp,
            shipping,
            currency,
            is_problem_return,
            delivery_confirmed,
            carrier_status,
            stamp,
        ),
    )
    conn.execute(
        "INSERT INTO shipping_logs(order_id,tracking_number,shipped_at) VALUES (?,?,?)",
        (order_id, f'TRACK-{order_id}', stamp),
    )


def _add_oms(conn, order_id, shipment_statuses):
    conn.execute(
        "INSERT INTO oms_order_fulfillment_state(order_id,revision,aggregate_status) VALUES (?,1,'shipped')",
        (order_id,),
    )
    fulfillment_id = f'ful-{order_id}'
    conn.execute(
        "INSERT INTO oms_fulfillments(id,order_id,revision,status) VALUES (?,?,1,'shipped')",
        (fulfillment_id, order_id),
    )
    for index, status in enumerate(shipment_statuses):
        conn.execute(
            "INSERT INTO oms_shipments(id,fulfillment_id,status) VALUES (?,?,?)",
            (f'ship-{order_id}-{index}', fulfillment_id, status),
        )


def test_only_unambiguous_carrier_returns_are_candidates():
    conn = _database()
    _insert_order(conn, 'legacy-return', shipping=25)
    _insert_order(conn, 'attention', carrier_status='attention')
    _insert_order(conn, 'problem-return', is_problem_return=1)
    _insert_order(conn, 'confirmed-return', delivery_confirmed=1)
    _insert_order(conn, 'fresh-return', age_days=0)

    _insert_order(conn, 'legacy-split')
    conn.execute(
        "INSERT INTO shipping_logs(order_id,tracking_number,shipped_at) "
        "SELECT order_id,'SECOND-TRACK',shipped_at FROM shipping_logs WHERE order_id='legacy-split' LIMIT 1"
    )

    _insert_order(conn, 'oms-return', shipping=40)
    _add_oms(conn, 'oms-return', ['returned'])
    _insert_order(conn, 'oms-mixed')
    _add_oms(conn, 'oms-mixed', ['returned', 'in_transit'])
    conn.commit()

    candidates = auto_confirm.find_returnable_orders(
        conn, auto_confirm.get_returned_since(conn)
    )
    assert [row['id'] for row in candidates] == ['legacy-return', 'oms-return']
    assert auto_confirm.returnable_stats(conn) == {
        'candidates': 2,
        'shipping_loss': 65.0,
    }
    assert auto_confirm.returnable_stats(conn, [])['candidates'] == 0


def test_enforce_returned_matches_manual_loss_default_and_is_idempotent():
    conn = _database()
    _insert_order(conn, 'return-1', shipping=25)
    _insert_order(conn, 'return-2', shipping=30)
    conn.commit()

    dry = auto_confirm.enforce_returned(conn, dry_run=True)
    assert dry['marked'] == 2
    assert dry['shipping_loss'] == 55.0
    assert conn.execute(
        "SELECT COUNT(*) FROM orders WHERE is_undelivered=1"
    ).fetchone()[0] == 0

    result = auto_confirm.enforce_returned(conn)
    assert result['checked'] == 2
    assert result['marked'] == 2
    assert result['shipping_loss'] == 55.0

    rows = conn.execute(
        """SELECT id,status,delivery_confirmed,is_undelivered,
                  shipping_total,shipping_loss_amount,undelivered_by,undelivered_note
           FROM orders ORDER BY id"""
    ).fetchall()
    assert all(row['status'] == 'shipped' for row in rows)
    assert all(row['delivery_confirmed'] == 0 for row in rows)
    assert all(row['is_undelivered'] == 1 for row in rows)
    assert [row['shipping_loss_amount'] for row in rows] == [25.0, 30.0]
    assert all(row['undelivered_by'] is None for row in rows)
    assert all('系统自动确认' in row['undelivered_note'] for row in rows)
    assert conn.execute("SELECT COUNT(*) FROM order_notes").fetchone()[0] == 2

    repeated = auto_confirm.enforce_returned(conn)
    assert repeated['checked'] == 0
    assert repeated['marked'] == 0
    assert conn.execute("SELECT COUNT(*) FROM order_notes").fetchone()[0] == 2


def test_operator_batch_is_scoped_attributed_and_independent_of_auto_switch():
    conn = _database()
    conn.execute("UPDATE settings SET value='0' WHERE key=?", (auto_confirm.RETURNED_ENABLE_KEY,))
    conn.execute("INSERT INTO sites(url,country) VALUES ('https://au.example','AU')")
    _insert_order(conn, 'pl-return-1', shipping=25)
    _insert_order(conn, 'pl-return-2', shipping=9, currency='EUR')
    _insert_order(
        conn,
        'au-return',
        shipping=40,
        currency='AUD',
        source='https://au.example',
        age_days=20,
    )
    conn.commit()

    preview = auto_confirm.confirm_returned_batch(
        conn,
        actor_name='测试操作员',
        actor_user_id=7,
        allowed_sources=['https://pl.example'],
        dry_run=True,
    )
    assert preview['checked'] == 2
    assert preview['marked'] == 2
    assert preview['shipping_loss_by_currency'] == {'EUR': 9.0, 'PLN': 25.0}
    assert conn.execute(
        "SELECT COUNT(*) FROM orders WHERE is_undelivered=1"
    ).fetchone()[0] == 0

    result = auto_confirm.confirm_returned_batch(
        conn,
        actor_name='测试操作员',
        actor_user_id=7,
        allowed_sources=['https://pl.example'],
    )
    assert result['marked'] == 2
    assert result['skipped'] == 0
    rows = conn.execute(
        "SELECT id,is_undelivered,undelivered_by,undelivered_note FROM orders ORDER BY id"
    ).fetchall()
    by_id = {row['id']: row for row in rows}
    assert by_id['pl-return-1']['is_undelivered'] == 1
    assert by_id['pl-return-1']['undelivered_by'] == 7
    assert '测试操作员' in by_id['pl-return-1']['undelivered_note']
    assert '批量确认' in by_id['pl-return-1']['undelivered_note']
    assert by_id['au-return']['is_undelivered'] == 0
    notes = conn.execute("SELECT author,note FROM order_notes ORDER BY id").fetchall()
    assert all(note['author'] == '测试操作员' for note in notes)
    assert len(notes) == 2

    repeated = auto_confirm.confirm_returned_batch(
        conn,
        actor_name='测试操作员',
        actor_user_id=7,
        allowed_sources=['https://pl.example'],
    )
    assert repeated['checked'] == 0
    assert repeated['marked'] == 0


def test_return_switch_and_periodic_task_are_independent_from_delivered_switch():
    conn = _database()
    conn.execute(
        "INSERT INTO settings(key,value) VALUES (?,?)",
        (auto_confirm.ENABLE_KEY, '0'),
    )
    assert auto_confirm.is_returned_enabled(conn) is True
    assert auto_confirm.is_enabled(conn) is False

    assert (
        celery_app.conf.task_routes['woo_sync.auto_confirm_returned']['queue']
        == 'sync_write'
    )
    schedule = celery_app.conf.beat_schedule['auto-confirm-carrier-returns']
    assert schedule['task'] == 'woo_sync.auto_confirm_returned'
    assert schedule['schedule'] == 60.0

    template = (ROOT / 'templates' / 'shipping.html').read_text(encoding='utf-8')
    app_source = (ROOT / 'app.py').read_text(encoding='utf-8')
    assert 'id="autoConfirmReturnedToggle"' in template
    assert 'toggleAutoConfirmReturned' in template
    assert 'id="batchConfirmReturnedBtn"' in template
    assert 'batchConfirmReturned' in template
    assert '/api/shipping/pending-outcome/confirm-returned-batch' in template
    assert 'def batch_confirm_returned_orders' in app_source
    assert 'auto_confirm_returned_enabled' in app_source
    assert 'auto_confirm_returned_since' in app_source
