"""Opt-in, loopback-only PostgreSQL tests with the real empty OMS schema."""
import json
import os
import uuid

import pytest

from test_reconciliation_postgres import pg as original_pg, add_order
from test_reconciliation_v2 import setup
from fulfillment_service import plan_order, create_shipment
from partner_statement_sources import build_snapshot, init_sqlite_schema, read_version, save_version


@pytest.fixture
def pg(setup,monkeypatch):
    if os.getenv('REC_TEST_POSTGRES')!='1':
        pytest.skip('Opt in with REC_TEST_POSTGRES=1 and an isolated local test cluster')
    from scripts import reconciliation_dev
    original_environment=reconciliation_dev.environment
    port=int(os.getenv('REC_SOURCE_TEST_PORT','55439'))
    assert 1024<=port<=65535
    def environment():
        original_environment()
        os.environ['WOO_DB_PORT']=str(port)
    monkeypatch.setattr(reconciliation_dev,'environment',environment)
    try:
        yield from original_pg.__wrapped__(setup)
    finally:
        import db_backend
        db_backend.close_pools()


def insert_copy(conn,table,row,**changes):
    data=dict(row);data.update(changes)
    if data.get('id') is None:data.pop('id',None)
    cols=','.join(data);marks=','.join('?' for _ in data)
    return conn.execute(f'INSERT INTO {table}({cols}) VALUES({marks})',list(data.values())).lastrowid


def test_real_postgres_joint_parcel_source_fee_and_immutable_revision(pg):
    data,connect=pg;c=connect()
    assert c.execute('SELECT current_database()').fetchone()[0]=='woo_reconciliation_test_dev'
    schema=c.execute('SELECT current_schema()').fetchone()[0]
    assert schema.startswith('rec_test_')
    c.execute('CREATE TABLE effective_partner_sites(partner_id bigint,site_id bigint)')
    c.execute("INSERT INTO partners(id,name,code,currency,cost_ratio,partner_profit_ratio,our_profit_ratio) VALUES(999,'Synthetic','poland_jin','PLN',.5,.25,.25)")
    c.execute('INSERT INTO effective_partner_sites VALUES(999,1)')
    c.execute('UPDATE sites SET cod_on_hold_is_shipped=? WHERE id=1',(1,))
    add_order(data,c,'1-source-pg',2)
    plan=plan_order(c,'1-source-pg');fid=plan['fulfillment_ids'][0]
    shipment=create_shipment(c,fid,'SOURCE-PG-TRACK')
    c.execute("UPDATE orders SET date_created='2026-09-10',status='completed',total=225,shipping_total=25 WHERE id='1-source-pg'")
    items=json.loads(c.execute("SELECT line_items FROM orders WHERE id='1-source-pg'").fetchone()[0])
    items[0]['total']='200';items[0]['quantity']=2
    c.execute("UPDATE orders SET line_items=? WHERE id='1-source-pg'",(json.dumps(items),))
    c.execute('UPDATE oms_shipments SET shipped_at=? WHERE id=?',('2026-09-11',shipment['id']))
    regular=dict(c.execute('SELECT * FROM oms_fulfillments WHERE id=?',(fid,)).fetchone())
    transit=c.execute("SELECT id FROM warehouses WHERE code='PL-JYJG-TRANSIT'").fetchone()
    if not transit:
        c.execute("INSERT INTO warehouses(id,code,name,country) VALUES(999,'PL-JYJG-TRANSIT','Synthetic transit','PL')")
        warehouse_id=999
    else:warehouse_id=transit[0]
    transit_fid=str(uuid.uuid4())
    insert_copy(c,'oms_fulfillments',regular,id=transit_fid,warehouse_id=warehouse_id,idempotency_key=regular['idempotency_key']+'-transit')
    item=dict(c.execute('SELECT * FROM oms_fulfillment_items WHERE fulfillment_id=?',(fid,)).fetchone())
    c.execute('UPDATE oms_fulfillment_items SET allocated_qty=1,fulfilled_qty=1 WHERE id=?',(item['id'],))
    second=insert_copy(c,'oms_fulfillment_items',item,id=None,fulfillment_id=transit_fid,allocated_qty=1,fulfilled_qty=1)
    c.execute('UPDATE oms_shipment_items SET quantity=1 WHERE shipment_id=?',(shipment['id'],))
    c.execute('INSERT INTO oms_shipment_items(shipment_id,fulfillment_item_id,quantity) VALUES(?,?,1)',(shipment['id'],second))
    c.execute('INSERT INTO shipping_logs(order_id,woo_order_id,source,tracking_number,shipped_at,status,is_reship) VALUES(?,?,?,?,?,?,?)',
              ('1-source-pg',123,'https://example.test','SOURCE-PG-TRACK','2026-09-11','shipped',False))
    rule={'warehouse_id':warehouse_id,'management_cny':'20','provider_id':data['partner'],'effective_from':'2020-01-01','effective_to':'2099-01-01'}
    c.execute('INSERT INTO rec_objects(id,kind,name,data,created_at,actor) VALUES(?,?,?,?,?,?)',
              (str(uuid.uuid4()),'history_rule','Synthetic packing',json.dumps(rule),'2026-01-01','1'))
    init_sqlite_schema(c)
    partner=dict(c.execute('SELECT * FROM partners WHERE id=999').fetchone())
    value=build_snapshot(c,partner,2026,9,'actual',{},lambda *args:{'unit_cost':40})
    assert value['settlement']['our_receivable']==130,value
    assert value['settlement']['partner_profit']==30
    assert value['settlement']['packing_fee_cny']==20
    assert value['settlement']['ready'] is True
    save_version(c,123,value,1);c.commit()
    c.close();c=connect()
    c.execute("UPDATE orders SET total=900 WHERE id='1-source-pg'")
    assert read_version(c,123,details=True)['settlement']==value['settlement']
    assert read_version(c,123,details=True)['service_lines'][0]['shipment_id']==shipment['id']
    c.rollback();c.close()
