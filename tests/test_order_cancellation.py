import pytest
import test_fulfillment as fixtures

from fulfillment_service import DomainError, plan_order, recompute_order_status, transition_fulfillment
from inv_migrations import up_017
from order_cancellation import cancellation_guard, stop_local_fulfillments


@pytest.fixture
def case():
    fixture = fixtures.FulfillmentDomainTests()
    fixture.setUp()
    c = fixture.db
    c.execute("INSERT INTO settings(key,value) VALUES ('oms_managed_product_isolation_enabled','1')")
    up_017(c)
    fixture.add_order('cancel-test', 'PL', 2, product_id=903, sku='', name='Fumot RandM Tornado 9000 Puffs - Grape')
    fid = plan_order(c, 'cancel-test')['fulfillment_ids'][0]
    row = c.execute('SELECT * FROM oms_fulfillments WHERE id=?', (fid,)).fetchone()
    yield c, fid, row['warehouse_id']
    fixture.tearDown()


def stop(c, **kwargs):
    return stop_local_fulfillments(c, 'cancel-test', actor={'id': None, 'name': 'Test'}, **kwargs)


def test_whole_cancellation_releases_stock_once_and_prevents_replanning(case):
    c, fid, warehouse = case
    before = tuple(c.execute('SELECT SUM(on_hand),SUM(reserved) FROM inv_stock WHERE warehouse_id=?', (warehouse,)).fetchone())
    assert before[1] == 2
    assert stop(c) == [fid]
    assert stop(c) == []
    after = tuple(c.execute('SELECT SUM(on_hand),SUM(reserved) FROM inv_stock WHERE warehouse_id=?', (warehouse,)).fetchone())
    assert after == (before[0], 0)
    assert c.execute("SELECT count(*) FROM inv_movements WHERE order_id='cancel-test' AND movement_type='release'").fetchone()[0] == 1
    assert c.execute('SELECT cancelled_qty FROM oms_fulfillment_items WHERE fulfillment_id=?', (fid,)).fetchone()[0] == 2
    assert recompute_order_status(c, 'cancel-test')['aggregate_status'] == 'cancelled'
    assert plan_order(c, 'cancel-test')['action'] == 'locked_noop'
    assert c.execute("SELECT status FROM orders WHERE id='cancel-test'").fetchone()[0] == 'processing'


def test_already_cancelled_warehouse_does_not_block_order_or_release_twice(case):
    c, fid, _ = case
    transition_fulfillment(c, fid, 'cancelled')
    # Reproduce the historic cancellation's missing item count.
    c.execute('UPDATE oms_fulfillment_items SET cancelled_qty=0 WHERE fulfillment_id=?', (fid,))
    assert stop(c, allowed_warehouse_ids=[]) == []
    assert c.execute('SELECT cancelled_qty FROM oms_fulfillment_items WHERE fulfillment_id=?', (fid,)).fetchone()[0] == 2
    assert c.execute("SELECT count(*) FROM inv_movements WHERE order_id='cancel-test' AND movement_type='release'").fetchone()[0] == 1


@pytest.mark.parametrize('status', ['picking', 'packed'])
def test_internal_picking_or_packing_can_stop_before_dispatch(case, status):
    c, fid, _ = case
    transition_fulfillment(c, fid, status)
    stop(c)
    assert c.execute('SELECT status FROM oms_fulfillments WHERE id=?', (fid,)).fetchone()[0] == 'cancelled'


@pytest.mark.parametrize('status', ['submitting', 'submission_unknown', 'accepted', 'cancel_pending', 'cancel_rejected', 'shipped', 'delivered', 'returned', 'exception'])
def test_one_blocked_warehouse_prevents_partial_cancellation(case, status):
    c, fid, _ = case
    c.execute("""INSERT INTO oms_fulfillments(id,order_id,warehouse_id,revision,status,mode,provider,idempotency_key,submitted_at)
                 VALUES ('other','cancel-test',2,1,?,'external_wms','hungary_wms','other','2026-09-18')""", (status,))
    c.commit()
    with pytest.raises(DomainError):
        stop(c)
    assert c.execute('SELECT status FROM oms_fulfillments WHERE id=?', (fid,)).fetchone()[0] == 'ready_to_pick'
    assert c.execute("SELECT count(*) FROM inv_movements WHERE order_id='cancel-test' AND movement_type='release'").fetchone()[0] == 0


def test_missing_warehouse_permission_keeps_reservation(case):
    c, fid, warehouse = case
    with pytest.raises(DomainError, match='取消权限'):
        stop(c, allowed_warehouse_ids=[])
    assert c.execute('SELECT SUM(reserved) FROM inv_stock WHERE warehouse_id=?', (warehouse,)).fetchone()[0] == 2


def test_every_safe_warehouse_is_cancelled_together(case):
    c, fid, _ = case
    c.execute("""INSERT INTO oms_fulfillments(id,order_id,warehouse_id,revision,status,mode,provider,idempotency_key)
                 VALUES ('other','cancel-test',2,1,'ready_to_submit','external_wms','hungary_wms','other')""")
    assert set(stop(c)) == {fid, 'other'}
    assert set(r[0] for r in c.execute("SELECT status FROM oms_fulfillments WHERE order_id='cancel-test'")) == {'cancelled'}


def test_fulfilled_quantity_blocks_even_when_fulfillment_status_is_stale(case):
    c, fid, _ = case
    c.execute('UPDATE oms_fulfillment_items SET fulfilled_qty=1 WHERE fulfillment_id=?', (fid,))
    with pytest.raises(DomainError, match='出库记录'):
        cancellation_guard(c, 'cancel-test')


def test_fully_cancelled_recompute_overrides_old_shortage_flags(case):
    c, fid, _ = case
    transition_fulfillment(c, fid, 'cancelled')
    c.execute("UPDATE oms_order_fulfillment_state SET has_shortage=1,manual_review=1 WHERE order_id='cancel-test'")
    assert recompute_order_status(c, 'cancel-test')['aggregate_status'] == 'cancelled'


def test_cancelled_unallocated_order_stays_stopped(case):
    c, fid, _ = case
    stop(c)
    c.execute('DELETE FROM oms_fulfillment_financials WHERE fulfillment_id=?', (fid,))
    c.execute('DELETE FROM oms_fulfillment_items WHERE fulfillment_id=?', (fid,))
    c.execute('DELETE FROM oms_fulfillments WHERE id=?', (fid,))
    assert recompute_order_status(c, 'cancel-test')['aggregate_status'] == 'cancelled'
    assert plan_order(c, 'cancel-test')['action'] == 'locked_noop'
