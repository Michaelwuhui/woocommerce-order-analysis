import json

import pytest

import auto_confirm
from fulfillment_service import plan_order, mark_manual_review, record_event, enqueue_job
from fulfillment_outcome_recovery import (
    STALE_PLAN_REASON, recover_terminal_state, recover_terminal_states, pending_outcome_blockers,
)
import test_fulfillment as fixtures


@pytest.fixture
def case():
    case = fixtures.FulfillmentDomainTests(); case.setUp()
    c = case.db
    for column in ["meta_data TEXT DEFAULT '[]'", "shipping_lines TEXT DEFAULT '[]'",
                   "date_modified TEXT DEFAULT '2020-01-01'", "carrier_status TEXT",
                   "carrier_status_at TEXT", "delivery_confirmed INTEGER DEFAULT 0",
                   "is_undelivered INTEGER DEFAULT 0", "is_problem_return INTEGER DEFAULT 0"]:
        c.execute("ALTER TABLE orders ADD COLUMN " + column)
    c.execute("CREATE TABLE shipping_logs(id INTEGER PRIMARY KEY,order_id TEXT,tracking_number TEXT,carrier_slug TEXT,shipped_at TEXT)")
    c.executemany("INSERT INTO settings(key,value) VALUES (?,?)", [
        (auto_confirm.ENABLE_KEY,"1"), (auto_confirm.SINCE_KEY,"2020-01-01"),
        (auto_confirm.RETURNED_ENABLE_KEY,"1"), (auto_confirm.RETURNED_SINCE_KEY,"2020-01-01")])
    yield case
    case.tearDown()


def add_managed(case, oid="1-123", *, partial=False):
    c=case.db
    case.add_order(oid,"PL",qty=2)
    plan_order(c,oid)
    f=c.execute("SELECT * FROM oms_fulfillments WHERE order_id=?",(oid,)).fetchone()
    fi=c.execute("SELECT * FROM oms_fulfillment_items WHERE fulfillment_id=?",(f['id'],)).fetchone()
    qty=1 if partial else 2
    c.execute("UPDATE oms_fulfillments SET status='delivered' WHERE id=?",(f['id'],))
    c.execute("UPDATE oms_fulfillment_items SET allocated_qty=?,fulfilled_qty=? WHERE id=?",(qty,qty,fi['id']))
    c.execute("""INSERT INTO oms_shipments(id,fulfillment_id,status,tracking_number,shipped_at)
        VALUES (?,?,'delivered',?,'2020-01-01')""",('ship-'+oid,f['id'],'TRACK-'+oid))
    c.execute("INSERT INTO oms_shipment_items VALUES (?,?,?)",('ship-'+oid,fi['id'],qty))
    if partial:
        c.execute("""INSERT INTO oms_fulfillments(id,order_id,warehouse_id,revision,status,idempotency_key)
            VALUES (?,?,2,?,'ready_to_pick',?)""",('unshipped-'+oid,oid,f['revision'],'unshipped-'+oid))
        c.execute("""INSERT INTO oms_fulfillment_items(fulfillment_id,order_item_id,sku_id,allocated_qty)
            VALUES (?,?,1,1)""",('unshipped-'+oid,fi['order_item_id']))
    c.execute("""UPDATE orders SET status='on-hold',carrier_status='delivered',carrier_status_at='2020-02-01',
        date_created='2020-01-01',meta_data=? WHERE id=?""",
        (json.dumps([{'key':'_tracking_number','value':'TRACK-'+oid}]),oid))
    c.execute("INSERT INTO shipping_logs(order_id,tracking_number,carrier_slug,shipped_at) VALUES (?,?,'dpd','2020-01-01')",(oid,'TRACK-'+oid))
    mark_manual_review(c,oid,STALE_PLAN_REASON)
    return f['id'],fi['id']


def add_legacy(case, oid="1-124", outcome="returned"):
    c=case.db
    case.add_order(oid,'CZ',qty=1,product_id=999,sku='UNMAPPED')
    plan_order(c,oid)
    assert c.execute("SELECT COUNT(*) FROM oms_fulfillments WHERE order_id=?",(oid,)).fetchone()[0]==0
    c.execute("UPDATE oms_order_items SET shortage_qty=0 WHERE order_id=?",(oid,))
    c.execute("""UPDATE oms_order_fulfillment_state SET aggregate_status='shipped',has_shortage=0,
        manual_review=0,manual_reason=NULL WHERE order_id=?""",(oid,))
    c.execute("""UPDATE orders SET status='shipped',carrier_status=?,carrier_status_at='2020-02-01',
        date_created='2020-01-01',meta_data=? WHERE id=?""",
        (outcome,json.dumps([{'key':'_tracking_number','value':'TRACK-'+oid}]),oid))
    record_event(c,'order',oid,'legacy_terminal_shortage_reconciled',from_status='stock_shortage',to_status='shipped')
    c.commit()


def test_recovers_old_system_review_without_inventory_or_duplicate_completion_job(case):
    c=case.db; add_managed(case)
    stock=[tuple(r) for r in c.execute('SELECT * FROM inv_stock')]
    movements=c.execute('SELECT COUNT(*) FROM inv_movements').fetchone()[0]
    assert recover_terminal_states(c,outcome='delivered')==[
        {'order_id':'1-123','kind':'stale_plan_review','status':'delivered'}]
    assert [r['id'] for r in auto_confirm.find_confirmable_orders(c,'2020-01-01')]==['1-123']
    assert c.execute("SELECT COUNT(*) FROM oms_integration_jobs WHERE job_type='COMPLETE_WOOCOMMERCE_ORDER'").fetchone()[0]==0
    assert [tuple(r) for r in c.execute('SELECT * FROM inv_stock')]==stock
    assert c.execute('SELECT COUNT(*) FROM inv_movements').fetchone()[0]==movements
    assert recover_terminal_states(c,outcome='delivered')==[]
    assert c.execute("SELECT COUNT(*) FROM oms_domain_events WHERE event_type='stale_plan_review_recovered'").fetchone()[0]==1


def test_partial_delivery_keeps_unshipped_items_and_cannot_confirm_whole_order(case):
    c=case.db; add_managed(case,partial=True)
    before=[tuple(r) for r in c.execute('SELECT * FROM oms_fulfillment_items ORDER BY id')]
    assert recover_terminal_states(c,outcome='delivered')[0]['status']=='partially_delivered'
    assert auto_confirm.find_confirmable_orders(c,'2020-01-01')==[]
    assert [tuple(r) for r in c.execute('SELECT * FROM oms_fulfillment_items ORDER BY id')]==before
    assert '部分包裹已签收' in pending_outcome_blockers(c,['1-123'])['1-123']['reason']
    assert c.execute("SELECT status FROM oms_fulfillments WHERE id='unshipped-1-123'").fetchone()[0]=='ready_to_pick'


@pytest.mark.parametrize('guard',[
    'human_hold','human_same_reason','changed_items','shortage','missing_quantity','missing_shipment',
    'open_plan_job','return_flag','second_tracking','changed_tracking_date','failed_fulfillment',
])
def test_does_not_release_other_holds_or_inconsistent_evidence(case,guard):
    c=case.db; fid,fi=add_managed(case)
    if guard=='human_hold':mark_manual_review(c,'1-123','客户要求暂停',actor={'id':1,'name':'operator'})
    elif guard=='human_same_reason':mark_manual_review(c,'1-123',STALE_PLAN_REASON,actor={'id':1,'name':'operator'})
    elif guard=='changed_items':
        item=json.loads(c.execute('SELECT line_items FROM orders').fetchone()[0]);item[0]['quantity']=3
        c.execute('UPDATE orders SET line_items=?',(json.dumps(item),))
    elif guard=='shortage':c.execute('UPDATE oms_order_fulfillment_state SET has_shortage=1')
    elif guard=='missing_quantity':c.execute('UPDATE oms_shipment_items SET quantity=1')
    elif guard=='missing_shipment':c.execute("UPDATE oms_shipments SET tracking_number='DIFFERENT'")
    elif guard=='open_plan_job':enqueue_job(c,'PLAN_ORDER','order','1-123','new-plan',{'order_id':'1-123'})
    elif guard=='return_flag':c.execute('UPDATE orders SET is_problem_return=1')
    elif guard=='second_tracking':c.execute("INSERT INTO shipping_logs(order_id,tracking_number) VALUES ('1-123','SECOND')")
    elif guard=='changed_tracking_date':c.execute("UPDATE orders SET date_modified='2020-03-01'")
    elif guard=='failed_fulfillment':c.execute("UPDATE oms_fulfillments SET last_error_code='manual_exception'")
    c.commit()
    assert recover_terminal_state(c,'1-123') is None
    assert c.execute('SELECT manual_review FROM oms_order_fulfillment_state').fetchone()[0]==1


@pytest.mark.parametrize('outcome',['delivered','returned'])
def test_legacy_manual_shipment_uses_audited_single_tracking_without_inventing_parcel(case,outcome):
    c=case.db;add_legacy(case,outcome=outcome)
    assert recover_terminal_states(c,outcome=outcome)[0]['status']==outcome
    rows=(auto_confirm.find_confirmable_orders(c,'2020-01-01') if outcome=='delivered'
          else auto_confirm.find_returnable_orders(c,'2020-01-01'))
    assert [r['id'] for r in rows]==['1-124']
    assert c.execute('SELECT COUNT(*) FROM oms_shipments').fetchone()[0]==0
    assert c.execute('SELECT COUNT(*) FROM oms_fulfillments').fetchone()[0]==0
    assert recover_terminal_states(c,outcome=outcome)==[]


@pytest.mark.parametrize('guard',['no_migration_event','missing_tracking','multiple_tracking','manual_hold',
                                  'active_plan','allocated_item','cancelled_order'])
def test_legacy_compatibility_is_narrow(case,guard):
    c=case.db;add_legacy(case)
    if guard=='no_migration_event':c.execute("DELETE FROM oms_domain_events WHERE event_type='legacy_terminal_shortage_reconciled'")
    elif guard=='missing_tracking':c.execute("UPDATE orders SET meta_data='[]'")
    elif guard=='multiple_tracking':c.execute("INSERT INTO shipping_logs(order_id,tracking_number) VALUES ('1-124','SECOND')")
    elif guard=='manual_hold':mark_manual_review(c,'1-124','需要核对')
    elif guard=='active_plan':enqueue_job(c,'PLAN_ORDER','order','1-124','new-plan',{'order_id':'1-124'})
    elif guard=='allocated_item':c.execute('UPDATE oms_order_items SET allocated_qty=1')
    elif guard=='cancelled_order':c.execute("UPDATE orders SET status='cancelled'")
    c.commit()
    assert recover_terminal_state(c,'1-124') is None
    assert auto_confirm.find_returnable_orders(c,'2020-01-01')==[]


def test_recovery_and_audit_rollback_together(case):
    c=case.db;add_managed(case)
    assert recover_terminal_state(c,'1-123')['status']=='delivered'
    c.rollback()
    assert c.execute('SELECT aggregate_status FROM oms_order_fulfillment_state').fetchone()[0]=='manual_review'
    assert c.execute("SELECT COUNT(*) FROM oms_domain_events WHERE event_type='stale_plan_review_recovered'").fetchone()[0]==0
