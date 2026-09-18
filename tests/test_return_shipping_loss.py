import json

import pytest

import auto_confirm
from return_shipping_loss import (
    PolicyConflict, load_policy, money, quote_loss, save_policy,
)
from test_auto_confirm_returned import _database, _insert_order


def rule(warehouse_id=None, country='CZ', currency='CZK', outbound=200, inbound=200):
    return dict(warehouse_id=warehouse_id, destination_country=country,
                currency=currency, outbound_amount=outbound, return_amount=inbound)


@pytest.fixture
def conn():
    db = _database()
    db.executescript('''
        CREATE TABLE warehouses(id INTEGER PRIMARY KEY, name TEXT);
        INSERT INTO warehouses VALUES (1,'Poland'),(10,'Transit');
        ALTER TABLE orders ADD COLUMN warehouse_id INTEGER;
        ALTER TABLE orders ADD COLUMN shipping TEXT;
        ALTER TABLE orders ADD COLUMN billing TEXT;
        ALTER TABLE oms_fulfillments ADD COLUMN warehouse_id INTEGER;
        ALTER TABLE oms_fulfillments ADD COLUMN shipped_at TEXT;
        INSERT INTO sites VALUES ('https://cz.example','CZ');
    ''')
    _insert_order(db, 'cz-order', currency='CZK', shipping=0, source='https://cz.example')
    db.execute('UPDATE orders SET shipping=?,billing=? WHERE id=?',
               (json.dumps({'country': 'CZ'}), json.dumps({'country': 'PL'}), 'cz-order'))
    yield db
    db.close()


def configure(db, rules):
    result = save_policy(db, rules, load_policy(db)['version'])
    db.commit()
    return result


def quote(db):
    return quote_loss(db, db.execute("SELECT * FROM orders WHERE id='cz-order'").fetchone())


def dispatched(db, warehouse=10, revision=1, status='shipped'):
    db.execute("INSERT OR IGNORE INTO oms_order_fulfillment_state(order_id,revision,aggregate_status) VALUES ('cz-order',1,'shipped')")
    fid = f'f-{warehouse}-{revision}'
    db.execute('''INSERT INTO oms_fulfillments(id,order_id,revision,status,warehouse_id)
                  VALUES (?,'cz-order',?,?,?)''', (fid, revision, status, warehouse))
    if status != 'planned':
        db.execute("INSERT INTO oms_shipments VALUES (?,?, 'returned')", (f's-{fid}', fid))


def test_czech_zero_shipping_uses_round_trip_policy(conn):
    configure(conn, [rule()])
    result = quote(conn)
    assert result['amount'] == 400
    assert result['currency'] == 'CZK'
    assert result['destination_country'] == 'CZ'  # Shipping beats billing.
    assert result['matched'] and result['warehouse_id'] is None


def test_actual_dispatch_beats_order_warehouse_and_common_rule(conn):
    conn.execute("UPDATE orders SET warehouse_id=1")
    configure(conn, [rule(), rule(1, outbound=100, inbound=100), rule(10, outbound=150, inbound=150)])
    dispatched(conn)
    assert quote(conn)['amount'] == 300
    assert quote(conn)['warehouse_id'] == 10


def test_planned_and_obsolete_fulfillments_do_not_override_order(conn):
    conn.execute("UPDATE orders SET warehouse_id=1")
    configure(conn, [rule(), rule(1, outbound=75, inbound=25)])
    dispatched(conn, status='planned')
    dispatched(conn, revision=2)
    assert quote(conn)['amount'] == 100


def test_multiwarehouse_requires_explicit_manual_amount(conn):
    configure(conn, [rule()])
    dispatched(conn, 1)
    dispatched(conn, 10)
    result = quote(conn)
    assert result['amount'] is None
    assert '多个仓库' in result['message']


def test_currency_mismatch_is_not_written_as_order_currency(conn):
    configure(conn, [rule(currency='PLN')])
    result = quote(conn)
    assert result['amount'] is None
    assert '币种' in result['message']
    result = auto_confirm.enforce_returned(conn)
    assert result['marked'] == 0 and result['skipped'] == 1
    assert conn.execute("SELECT is_undelivered FROM orders").fetchone()[0] == 0


def test_unknown_warehouse_with_only_specific_rule_requires_review(conn):
    configure(conn, [rule(10)])
    assert quote(conn)['amount'] is None
    conn.execute('UPDATE orders SET warehouse_id=10')
    assert quote(conn)['amount'] == 400


def test_no_match_keeps_existing_amount_and_zero_is_valid(conn):
    configure(conn, [rule(country='PL', currency='PLN', outbound=25, inbound=0)])
    conn.execute('UPDATE orders SET shipping_total=37.50')
    assert quote(conn)['amount'] == 37.5
    configure(conn, [rule(outbound=0, inbound=0)])
    assert quote(conn)['amount'] == 0
    configure(conn, [])
    assert quote(conn)['amount'] == 37.5


def test_address_fallback_and_region_are_independent_of_warehouse(conn):
    configure(conn, [rule(), rule(1, country='PL', currency='PLN', outbound=25, inbound=0)])
    conn.execute("UPDATE orders SET warehouse_id=1,shipping='{}',billing='{}'")
    assert quote(conn)['amount'] == 400  # Site country only when address absent.
    conn.execute("UPDATE orders SET shipping=?,currency='PLN'", (json.dumps({'country': 'PL'}),))
    assert quote(conn)['amount'] == 25


@pytest.mark.parametrize('value', ['NaN', 'Infinity', '-Infinity', -1, 1000001, '', None, True, '1.001'])
def test_rejects_invalid_money(value):
    with pytest.raises(ValueError):
        money(value)


@pytest.mark.parametrize('rules', [
    [rule(), rule()], [rule(999)], [rule(True)], [rule(country='CZE')],
    [rule(currency='CZ')], [rule(outbound=None)], [rule(outbound=999999, inbound=2)], {},
])
def test_invalid_rules_do_not_replace_saved_rules(conn, rules):
    before = configure(conn, [rule()])
    with pytest.raises(ValueError):
        save_policy(conn, rules, before['version'])
    assert load_policy(conn) == before


def test_stale_save_does_not_overwrite_newer_config(conn):
    before = load_policy(conn)
    latest = configure(conn, [rule()])
    with pytest.raises(PolicyConflict):
        save_policy(conn, [], before['version'])
    assert load_policy(conn) == latest


def test_policy_change_does_not_rewrite_historical_losses(conn):
    conn.execute('UPDATE orders SET is_undelivered=1,shipping_loss_amount=80')
    configure(conn, [rule()])
    assert conn.execute('SELECT shipping_loss_amount FROM orders').fetchone()[0] == 80


@pytest.mark.parametrize('batch', [False, True])
def test_preview_and_confirmation_use_same_loss_and_are_idempotent(conn, batch):
    configure(conn, [rule(), rule(country='PL', currency='PLN', outbound=25, inbound=0)])
    _insert_order(conn, 'pl-order', shipping=25)
    def run(dry):
        if batch:
            return auto_confirm.confirm_returned_batch(conn, actor_name='Tester', actor_user_id=7, dry_run=dry)
        return auto_confirm.enforce_returned(conn, dry_run=dry)
    assert auto_confirm.returnable_stats(conn) == {'candidates': 2, 'shipping_loss': 425.0}
    preview = run(True)
    assert preview['marked'] == 2 and preview['shipping_loss'] == 425
    assert conn.execute('SELECT SUM(is_undelivered) FROM orders').fetchone()[0] == 0
    result = run(False)
    assert result['marked'] == 2 and result['shipping_loss'] == 425
    if batch:
        assert result['shipping_loss_by_currency'] == {'CZK': 400, 'PLN': 25}
    assert conn.execute("SELECT shipping_loss_amount FROM orders WHERE id='cz-order'").fetchone()[0] == 400
    assert conn.execute('SELECT COUNT(*) FROM order_notes').fetchone()[0] == 2
    assert run(False)['marked'] == 0
