"""Exercise real authentication, transactions and retries using synthetic data."""
import copy
import itertools
import json
import os
import threading
from concurrent.futures import ThreadPoolExecutor

import pytest
import requests
import db_backend as db
from fulfillment_service import MANAGED_TRANSFER_ROUTING, transition_fulfillment

pytestmark = pytest.mark.skipif(not db.is_postgres_backend(), reason='isolated PostgreSQL required')
ids = itertools.count(980001)


class Response:
    def __init__(self, payload, status=200):
        self.payload, self.status_code = payload, status
    def json(self):
        return copy.deepcopy(self.payload)


@pytest.fixture
def case(monkeypatch):
    database = os.environ.get('WOO_DB_NAME_OVERRIDE', '')
    assert database.startswith('woo_return_loss_test_cancel_')
    monkeypatch.setattr(requests.sessions.Session, 'request', lambda *a, **kw: pytest.fail('real network forbidden'))
    import app as module
    module.app.config.update(TESTING=True)
    wc_id = next(ids)
    oid, fid = f'99188-{wc_id}', f'cancel-test-{wc_id}'
    c = db.connect()
    assert c.execute('SELECT current_database()').fetchone()[0] == database
    c.execute("INSERT INTO sites(id,url,consumer_key,consumer_secret,country) VALUES"
              "(99188,'https://cancel-test.invalid','synthetic','synthetic','PL') ON CONFLICT(id) DO NOTHING")
    for uid, role, manage in [(99188,'admin',True),(99189,'admin',False),(99190,'viewer',False),(99191,'user',True)]:
        c.execute("""INSERT INTO users(id,username,password_hash,name,role,can_manage_inventory)
                     VALUES (?,?,'unusable','Cancellation test',?,?) ON CONFLICT(id) DO NOTHING""",
                  (uid, f'cancel-user-{uid}', role, manage))
    c.execute("INSERT INTO user_country_permissions(user_id,country) VALUES (99191,'CZ') ON CONFLICT DO NOTHING")
    c.execute("INSERT INTO warehouses(id,name,code,country) VALUES (99188,'Cancellation test','CANCEL-TEST','PL') ON CONFLICT(id) DO NOTHING")
    c.execute("""INSERT INTO oms_warehouse_integrations(warehouse_id,provider,config_json,is_enabled)
                 VALUES (99188,'internal',?,0) ON CONFLICT(warehouse_id) DO NOTHING""",
              (json.dumps({'routing_policy':MANAGED_TRANSFER_ROUTING}),))
    c.execute("INSERT INTO inv_skus(id,sku_code,name,is_active) VALUES (99188,'CANCEL-TEST','Test item',1) ON CONFLICT(id) DO NOTHING")
    c.execute("""INSERT INTO inv_stock(warehouse_id,sku_id,on_hand,reserved) VALUES (99188,99188,20,2)
                 ON CONFLICT(warehouse_id,sku_id) DO UPDATE SET on_hand=20,reserved=2""")
    c.execute("""INSERT INTO orders(id,number,source,status,line_items,billing,shipping,meta_data)
                 VALUES (?,?,'https://cancel-test.invalid','processing','[]','{}','{}','[]')""",(oid,str(wc_id)))
    item = c.execute("""INSERT INTO oms_order_items(order_id,woo_line_item_id,sku_id,name,ordered_qty,allocated_qty)
                        VALUES (?,'1',99188,'Test item',2,2) RETURNING id""", (oid,)).fetchone()[0]
    c.execute("INSERT INTO oms_order_fulfillment_state(order_id,revision,aggregate_status) VALUES (?,1,'allocated')",(oid,))
    c.execute("""INSERT INTO oms_fulfillments(id,order_id,warehouse_id,status,mode,provider,idempotency_key)
                 VALUES (?,?,99188,'ready_to_pick','internal','internal',?)""",(fid,oid,fid))
    fi = c.execute("""INSERT INTO oms_fulfillment_items(fulfillment_id,order_item_id,sku_id,allocated_qty)
                      VALUES (?,?,99188,2) RETURNING id""",(fid,item)).fetchone()[0]
    c.execute("""INSERT INTO inv_movements(warehouse_id,sku_id,movement_type,qty_delta,reserved_delta,ref_type,ref_id,order_id)
                 VALUES (99188,99188,'reserve',0,2,'oms_fulfillment_item',?,?)""",(f'{fi}:reserve',oid))
    c.commit(); c.close()
    state = {'remote':{'id':wc_id,'status':'processing','date_modified':'2026-09-18T12:00:00','meta_data':[]},
             'writes':0,'reads':0,'outcome':'success','read_status':200,'entered':threading.Event(),'release':threading.Event()}

    def get(url, **kwargs):
        assert url == f'https://cancel-test.invalid/wp-json/wc/v3/orders/{wc_id}'
        state['reads'] += 1
        if state.get('read_hook'):
            hook = state.pop('read_hook')
            hook()
        return Response(state['remote'], state['read_status'])

    def put(url, **kwargs):
        assert url == f'https://cancel-test.invalid/wp-json/wc/v3/orders/{wc_id}'
        assert kwargs['json'] == {'status':'cancelled'}
        state['writes'] += 1
        if state['outcome'] == 'reject':
            return Response({'code':'rest_forbidden'},403)
        if state['outcome'] == 'unknown':
            raise requests.ReadTimeout('synthetic unknown response')
        if state['outcome'] == 'wrong_id':
            return Response({'id':wc_id+1,'status':'cancelled'})
        if state['outcome'] == 'wait':
            state['entered'].set()
            assert state['release'].wait(10)
        state['remote']['status'] = 'cancelled'
        if state['outcome'] == 'timeout_applied':
            raise requests.ReadTimeout('synthetic committed response')
        return Response(state['remote'])

    monkeypatch.setattr(requests,'get',get)
    monkeypatch.setattr(requests,'put',put)
    monkeypatch.setattr(requests,'post',lambda *a,**kw:pytest.fail('unexpected external note/email/write'))

    def submit(uid=99188):
        client=module.app.test_client()
        if uid:
            with client.session_transaction() as session:
                session['_user_id']=str(uid); session['_fresh']=True
        return client.post(f'/api/order/{oid}/status',json={'status':'cancelled'})

    def snapshot():
        c=db.connect()
        result={
            'order':dict(c.execute('SELECT status FROM orders WHERE id=?',(oid,)).fetchone()),
            'fulfillment':dict(c.execute('SELECT status FROM oms_fulfillments WHERE id=?',(fid,)).fetchone()),
            'state':dict(c.execute('SELECT * FROM oms_order_fulfillment_state WHERE order_id=?',(oid,)).fetchone()),
            'stock':dict(c.execute('SELECT on_hand,reserved FROM inv_stock WHERE warehouse_id=99188 AND sku_id=99188').fetchone()),
            'releases':c.execute("SELECT count(*) FROM inv_movements WHERE order_id=? AND movement_type='release'",(oid,)).fetchone()[0],
            'operations':[dict(r) for r in c.execute('SELECT status,attempts FROM external_operations WHERE order_id=?',(oid,))],
        }
        c.close()
        return result
    return {'id':oid,'fid':fid,'state':state,'submit':submit,'snapshot':snapshot}


def assert_cancelled(case):
    snap=case['snapshot']()
    assert snap['order']['status']=='cancelled'
    assert snap['fulfillment']['status']=='cancelled'
    assert snap['state']['aggregate_status']=='cancelled' and not snap['state']['manual_review']
    assert snap['stock']=={'on_hand':20,'reserved':0}
    assert snap['releases']==1
    assert len(snap['operations'])==1 and snap['operations'][0]['status']=='local_committed'


def test_cancel_from_order_details_is_idempotent(case):
    assert case['submit']().status_code==200
    assert case['submit']().status_code==200
    assert case['state']['writes']==1
    assert_cancelled(case)


def test_cancel_previously_cancelled_fulfillment_matches_reported_bug(case):
    c=db.connect(); transition_fulfillment(c,case['fid'],'cancelled'); c.commit(); c.close()
    assert case['submit'](99189).status_code==200
    assert_cancelled(case)


def test_timed_out_mutation_is_read_back_without_second_write(case):
    case['state']['outcome']='timeout_applied'
    assert case['submit']().status_code==200
    assert case['state']['writes']==1 and case['state']['reads']==2
    assert_cancelled(case)


def test_unknown_result_is_retained_and_retry_only_reads(case):
    case['state']['outcome']='unknown'
    assert case['submit']().status_code==409
    assert case['snapshot']()['order']['status']=='processing'
    assert case['snapshot']()['state']['manual_review']
    assert case['submit']().status_code==409
    assert case['state']['writes']==1
    case['state']['remote']['status']='cancelled'
    assert case['submit']().status_code==200
    assert case['state']['writes']==1
    assert_cancelled(case)


def test_definite_rejection_allows_explicit_retry_without_releasing_twice(case):
    case['state']['outcome']='reject'
    assert case['submit']().status_code==422
    assert case['snapshot']()['order']['status']=='processing'
    case['state']['outcome']='success'
    assert case['submit']().status_code==200
    assert case['state']['writes']==2
    assert_cancelled(case)


def test_concurrent_cancel_does_not_duplicate_remote_write_or_stock_release(case):
    case['state']['outcome']='wait'
    with ThreadPoolExecutor(max_workers=2) as pool:
        first=pool.submit(case['submit'])
        try:
            assert case['state']['entered'].wait(10)
            second=pool.submit(case['submit']).result(timeout=10)
            assert second.status_code==409
        finally:
            case['state']['release'].set()
        assert first.result(timeout=10).status_code==200
    assert case['state']['writes']==1
    assert_cancelled(case)


@pytest.mark.parametrize('remote', [
    {'status':'completed'}, {'status':'shipped'}, {'status':'on-hold'},
    {'meta_data':[{'key':'_wc_shipment_tracking_items','value':[{'tracking_number':'TEST'}]}]},
])
def test_remote_shipping_evidence_blocks_before_any_change(case,remote):
    case['state']['remote'].update(remote)
    assert case['submit']().status_code==409
    assert case['state']['writes']==0
    assert case['snapshot']()['stock']['reserved']==2
    assert case['snapshot']()['fulfillment']['status']=='ready_to_pick'


@pytest.mark.parametrize('uid', [None,99189,99190,99191])
def test_login_role_site_and_warehouse_permissions(case,uid):
    result=case['submit'](uid)
    assert result.status_code in (302,403)
    assert case['state']['writes']==0
    assert case['snapshot']()['stock']['reserved']==2


def test_preflight_failure_does_not_stop_warehouses(case):
    case['state']['read_status']=503
    assert case['submit']().status_code==422
    assert case['state']['writes']==0
    assert case['snapshot']()['stock']['reserved']==2


def test_wrong_order_identity_never_reports_success(case):
    case['state']['outcome']='wrong_id'
    assert case['submit']().status_code==409
    assert case['snapshot']()['order']['status']=='processing'


def test_warehouse_started_during_preflight_is_rechecked(case):
    def started():
        c=db.connect()
        c.execute("UPDATE oms_fulfillments SET mode='external_wms',status='submitting',submitted_at=CURRENT_TIMESTAMP WHERE id=?",(case['fid'],))
        c.commit();c.close()
    case['state']['read_hook']=started
    assert case['submit']().status_code==409
    assert case['state']['writes']==0
    assert case['snapshot']()['stock']['reserved']==2


@pytest.mark.parametrize('status',['picking','packed'])
def test_cancel_one_warehouse_stops_work_but_does_not_cancel_commercial_order(case,status):
    c=db.connect()
    c.execute('UPDATE oms_fulfillments SET status=? WHERE id=?',(status,case['fid']))
    c.commit();c.close()
    import app as module
    client=module.app.test_client()
    with client.session_transaction() as session:
        session['_user_id']='99188';session['_fresh']=True
    response=client.post(f"/api/fulfillment/{case['fid']}/cancel",json={'reason':'Synthetic cancellation'})
    assert response.status_code==200
    snap=case['snapshot']()
    assert snap['order']['status']=='processing'
    assert snap['fulfillment']['status']=='cancelled'
    assert snap['state']['aggregate_status']=='cancelled'
    assert snap['releases']==1 and snap['stock']['reserved']==0
    assert case['state']['writes']==0


def test_local_finalize_failure_recovers_without_repeating_remote_mutation(case,monkeypatch):
    import order_cancellation as cancellation
    original=cancellation._finish
    monkeypatch.setattr(cancellation,'_finish',lambda *a,**kw:(_ for _ in ()).throw(RuntimeError('synthetic local failure')))
    assert case['submit']().status_code==409
    assert case['state']['remote']['status']=='cancelled'
    assert case['snapshot']()['order']['status']=='processing'
    monkeypatch.setattr(cancellation,'_finish',original)
    assert case['submit']().status_code==200
    assert case['state']['writes']==1
    assert_cancelled(case)
