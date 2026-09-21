"""Real PostgreSQL acceptance. Opt in to the loopback-only disposable dev cluster."""
import concurrent.futures
import json
import os
import uuid
import pytest
from test_reconciliation_v2 import setup
from reconciliation_core import rows, configure_order, ReconciliationError
from reconciliation_schema import migrate
from reconciliation_ledger import recognize, record_bill
from fulfillment_service import plan_order, create_shipment


@pytest.fixture
def pg(setup):
    if os.getenv('REC_TEST_POSTGRES')!='1':
        pytest.skip('Set REC_TEST_POSTGRES=1 after initializing scripts/reconciliation_dev.py')
    from scripts.reconciliation_dev import environment
    environment()
    import db_backend
    schema='rec_test_'+uuid.uuid4().hex
    root=db_backend.connect()
    assert root.execute('SELECT current_database()').fetchone()[0]=='woo_reconciliation_test_dev'
    root.execute('CREATE SCHEMA '+schema)
    tables=[r[0] for r in root.execute("SELECT tablename FROM pg_tables WHERE schemaname='public'").fetchall()]
    for table in tables:
        assert table.replace('_','').isalnum()
        root.execute(f'CREATE TABLE {schema}.{table} (LIKE public.{table} INCLUDING ALL)')
    root.commit();root.close()
    def connect():
        c=db_backend.connect();c.execute('SET search_path TO '+schema+',public');return c
    c=connect()
    source=setup['db']
    for row in source.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'").fetchall():
        table=row[0]
        if table not in tables:continue
        columns={r['column_name']:r['data_type'] for r in rows(c,'SELECT column_name,data_type FROM information_schema.columns WHERE table_schema=? AND table_name=?',(schema,table))}
        for record in source.execute('SELECT * FROM '+table).fetchall():
            names=[k for k in record.keys() if k in columns]
            values=[bool(record[k]) if columns[k]=='boolean' and record[k] is not None else record[k] for k in names]
            if table=='sites':
                names.extend(['consumer_key','consumer_secret']);values.extend(['disabled','disabled'])
            placeholders=','.join(['?']*len(names))
            c.execute(f"INSERT INTO {table}({','.join(names)}) VALUES ({placeholders})",values)
        if 'id' in columns:
            seq=c.execute('SELECT pg_get_serial_sequence(?,?)',(schema+'.'+table,'id')).fetchone()[0]
            if seq:
                c.execute(f"SELECT setval(?,GREATEST(COALESCE((SELECT MAX(id) FROM {table}),0),1),EXISTS(SELECT 1 FROM {table}))",(seq,))
    c.commit();c.close()
    yield setup,connect
    root=db_backend.connect()
    assert schema.startswith('rec_test_') and len(schema)==41
    root.execute('DROP SCHEMA '+schema+' CASCADE');root.commit();root.close()


def add_order(d,conn,oid,qty):
    d['helper'].add_order(oid,'PL',qty=qty,shipping_total=25,currency='PLN',line_total=qty*30)
    row=d['db'].execute('SELECT * FROM orders WHERE id=?',(oid,)).fetchone()
    conn.execute('INSERT INTO orders('+','.join(row.keys())+') VALUES ('+','.join('?' for _ in row)+')',list(row))
    configure_order(conn,oid,d['cfg'],1)


def test_postgres_real_chain(pg):
    d,connect=pg;c=connect();add_order(d,c,'1-pg',12);c.commit()
    result=plan_order(c,'1-pg');assert not result['shortages']
    f=result['fulfillment_ids'][0];parcel=create_shipment(c,f,'PG-CHAIN')
    record_bill(c,parcel['id'],{'provider_id':d['carrier'],'currency':'PLN','outbound':'8','reverse':'0',
        'collection_fee':'2','collected':'385','remitted':'375','collected_at':'2026-09-21','remitted_at':'2026-09-21','reference':'PG1','proof':'synthetic'},1)
    result=recognize(c,'1-pg',1);c.commit()
    assert not result['pending']
    assert c.execute('SELECT on_hand FROM inv_stock WHERE warehouse_id=1 AND sku_id=1').fetchone()['on_hand']==8
    assert len(rows(c,'SELECT * FROM rec_sources'))==2
    assert next(r['amount'] for r in result['totals'] if r['category']=='service')=='20.00'
    c.close()


def test_postgres_concurrent_reservations_never_oversell(pg):
    d,connect=pg;c=connect()
    for oid in ('1-a','1-b'):add_order(d,c,oid,12)
    c.commit();c.close()
    def reserve(oid):
        c=connect()
        try:
            result=plan_order(c,oid);return result['aggregate_status']
        except ReconciliationError:
            c.rollback();return 'stock_changed'
        finally:c.close()
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
        results=list(pool.map(reserve,['1-a','1-b']))
    c=connect();stock=c.execute('SELECT on_hand,reserved FROM inv_stock WHERE warehouse_id=1 AND sku_id=1').fetchone()
    assert 0<=stock['reserved']<=stock['on_hand']==20
    assert sum(r['reserved'] for r in rows(c,'SELECT reserved FROM rec_batches'))==stock['reserved']
    assert 'allocated' in results
    assert c.execute('SELECT COUNT(*) FROM rec_sources WHERE reserved<0').fetchone()[0]==0
    c.close()
