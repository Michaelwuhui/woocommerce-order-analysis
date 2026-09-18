import json

import pytest

import test_fulfillment as domain_tests
import inv_workflows
import inventory_replan as replan
from fulfillment_service import plan_order
from inv_common import record_movement
from inv_migrations import up_017


@pytest.fixture
def inventory():
    system = domain_tests.FulfillmentDomainTests()
    system.setUp()
    c = system.db
    c.execute("INSERT OR REPLACE INTO settings VALUES ('oms_managed_product_isolation_enabled','1')")
    up_017(c)
    c.execute("INSERT OR REPLACE INTO settings VALUES ('oms_fulfillment_enabled','1')")
    wid = c.execute("SELECT id FROM warehouses WHERE code='PL-JYJG-TRANSIT'").fetchone()[0]
    sid = c.execute("SELECT id FROM inv_skus WHERE sku_code='40K-CI'").fetchone()[0]
    c.execute('UPDATE inv_stock SET on_hand=0,reserved=0 WHERE warehouse_id=? AND sku_id=?', (wid, sid))
    c.commit()
    system.warehouse_id, system.sku_id = wid, sid
    yield system
    system.tearDown()


def add_order(system, order_id, qty=1, date='2026-09-01', **kwargs):
    system.add_order(order_id, 'PL', qty, product_id=901, sku='',
                     name='Fumot Leopard 40000 Puffs - Cola Ice', **kwargs)
    system.db.execute('UPDATE orders SET date_created=? WHERE id=?', (date, order_id))
    plan_order(system.db, order_id)


def receive(system, qty, **kwargs):
    return record_movement(system.db, warehouse_id=system.warehouse_id,
                           sku_id=system.sku_id, movement_type=kwargs.pop('movement_type', 'adjust'),
                           qty_delta=qty, **kwargs)


def jobs(c, kind):
    return [dict(row) for row in c.execute('SELECT * FROM oms_integration_jobs WHERE job_type=? ORDER BY id', (kind,))]


def dispatch(c, job):
    handler = replan.handle_restocked_warehouse if job['job_type']==replan.WAREHOUSE_JOB else replan.handle_restocked_order
    result = handler(c, job, json.loads(job['payload_json']))
    c.execute("UPDATE oms_integration_jobs SET status='succeeded' WHERE id=?", (job['id'],))
    c.commit()
    return result


def balance(system):
    return tuple(system.db.execute('SELECT on_hand,reserved FROM inv_stock WHERE warehouse_id=? AND sku_id=?',
                                    (system.warehouse_id,system.sku_id)).fetchone())


def test_stocktake_preserves_existing_reservation_and_recovers_shortage(inventory, monkeypatch):
    s, c = inventory, inventory.db
    receive(s, 1)
    add_order(s, 'partly-reserved', 3)
    assert balance(s)==(1,1)
    baseline_movement = c.execute('SELECT MAX(id) FROM inv_movements WHERE warehouse_id=? AND sku_id=?',
                                  (s.warehouse_id,s.sku_id)).fetchone()[0]
    monkeypatch.setattr(inv_workflows,'actor',lambda:(None,'test stocktake operator'))
    inv_workflows.apply_stocktake(c, {'id':123,'warehouse_id':s.warehouse_id,'reference':'count','created_name':'test'},
        [{'sku_id':s.sku_id,'qty':3,'baseline':1,'baseline_reserved':1,'baseline_movement':baseline_movement}], reviewer='系统自动审批')
    assert balance(s)==(3,1)
    assert len(jobs(c,replan.WAREHOUSE_JOB))==1
    dispatch(c,jobs(c,replan.WAREHOUSE_JOB)[0])
    child=jobs(c,replan.ORDER_JOB)[0]
    assert dispatch(c,child)['aggregate_status']=='allocated'
    assert balance(s)==(3,3)
    item=c.execute("SELECT allocated_qty,shortage_qty FROM oms_order_items WHERE order_id='partly-reserved'").fetchone()
    assert tuple(item)==(3,0)
    revision=c.execute("SELECT revision FROM oms_order_fulfillment_state WHERE order_id='partly-reserved'").fetchone()[0]
    movements=c.execute('SELECT COUNT(*) FROM inv_movements').fetchone()[0]
    assert dispatch(c,child)['action']=='skipped'
    assert balance(s)==(3,3)
    assert c.execute('SELECT COUNT(*) FROM inv_movements').fetchone()[0]==movements
    assert c.execute("SELECT revision FROM oms_order_fulfillment_state WHERE order_id='partly-reserved'").fetchone()[0]==revision
    assert c.execute('SELECT COUNT(*) FROM oms_shipments').fetchone()[0]==0
    assert c.execute("SELECT COUNT(*) FROM oms_integration_jobs WHERE job_type LIKE 'SUBMIT_%'").fetchone()[0]==0


def test_restock_queues_oldest_first_and_keeps_real_shortage(inventory):
    s,c=inventory,inventory.db
    add_order(s,'newer',2,date='2026-09-02')
    add_order(s,'older',3,date='2026-09-01')
    receive(s,4)
    parent=jobs(c,replan.WAREHOUSE_JOB)[0]
    assert dispatch(c,parent)['order_ids']==['older','newer']
    assert dispatch(c,parent)['order_ids']==['older','newer']
    assert len(jobs(c,replan.ORDER_JOB))==2
    for job in jobs(c,replan.ORDER_JOB):dispatch(c,job)
    assert balance(s)==(4,4)
    assert [tuple(row) for row in c.execute('SELECT order_id,allocated_qty,shortage_qty FROM oms_order_items ORDER BY order_id')]==[
        ('newer',1,1),('older',3,0)]


@pytest.mark.parametrize('kind',['purchase_in','transfer_in','adjust','return_in'])
def test_positive_stock_movements_queue_replan(inventory,kind):
    add_order(inventory,'waiting')
    receive(inventory,1,movement_type=kind)
    assert len(jobs(inventory.db,replan.WAREHOUSE_JOB))==1


def test_document_dedup_and_rollback_are_atomic(inventory):
    s,c=inventory,inventory.db
    add_order(s,'waiting',3)
    receive(s,1,ref_type='inventory_document',ref_id='9')
    receive(s,1,ref_type='inventory_document',ref_id='9')
    assert len(jobs(c,replan.WAREHOUSE_JOB))==1
    c.rollback()
    assert balance(s)==(0,0)
    assert jobs(c,replan.WAREHOUSE_JOB)==[]
    receive(s,1,ref_type='inventory_document',ref_id='9')
    c.commit()
    assert len(jobs(c,replan.WAREHOUSE_JOB))==1


@pytest.mark.parametrize('kwargs',[
    {'qty_delta':0,'movement_type':'reserve','reserved_delta':1},
    {'qty_delta':0,'movement_type':'release','reserved_delta':-1},
    {'qty_delta':-1,'movement_type':'adjust'},
    {'qty_delta':1,'movement_type':'adjust','apply_stock':False},
])
def test_non_restock_changes_do_not_create_jobs(inventory,kwargs):
    add_order(inventory,'waiting')
    record_movement(inventory.db,warehouse_id=inventory.warehouse_id,sku_id=inventory.sku_id,**kwargs)
    assert jobs(inventory.db,replan.WAREHOUSE_JOB)==[]


@pytest.mark.parametrize('change', ['manual_reason','manual_hold','shipped','routing_disabled','feature_disabled'])
def test_worker_rechecks_manual_holds_shipments_and_routing(inventory,change):
    s,c=inventory,inventory.db
    receive(s,1)
    add_order(s,'waiting',2)
    receive(s,1)
    dispatch(c,jobs(c,replan.WAREHOUSE_JOB)[0])
    child=jobs(c,replan.ORDER_JOB)[0]
    if change=='manual_reason':
        c.execute("UPDATE oms_order_fulfillment_state SET manual_reason='客户要求暂停' WHERE order_id='waiting'")
    elif change=='manual_hold':
        c.execute("UPDATE oms_fulfillments SET status='manual_hold' WHERE order_id='waiting'")
    elif change=='shipped':
        c.execute("UPDATE orders SET status='shipped' WHERE id='waiting'")
    elif change=='routing_disabled':
        c.execute("UPDATE settings SET value='0' WHERE key='oms_jyjg_transit_routing_enabled'")
    else:
        c.execute("UPDATE settings SET value='0' WHERE key='oms_fulfillment_enabled'")
    c.commit()
    before=balance(s)
    assert dispatch(c,child)['action']=='skipped'
    assert balance(s)==before


def test_enqueue_failure_rolls_back_stock(inventory,monkeypatch):
    s,c=inventory,inventory.db
    add_order(s,'waiting')
    def fail(*args,**kwargs):raise RuntimeError('queue unavailable')
    monkeypatch.setattr(replan,'enqueue_job',fail)
    with pytest.raises(RuntimeError):receive(s,1)
    c.rollback()
    assert balance(s)==(0,0)
    assert jobs(c,replan.WAREHOUSE_JOB)==[]


def test_different_warehouse_stock_does_not_replan(inventory):
    s,c=inventory,inventory.db
    add_order(s,'waiting')
    record_movement(c,warehouse_id=1,sku_id=s.sku_id,movement_type='adjust',qty_delta=10)
    parents=jobs(c,replan.WAREHOUSE_JOB)
    assert len(parents)==1
    assert dispatch(c,parents[0])['order_ids']==[]
    assert jobs(c,replan.ORDER_JOB)==[]
