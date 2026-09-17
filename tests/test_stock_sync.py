import concurrent.futures
import copy
from datetime import timedelta
import json

import pytest

from stock_sync_fixtures import system
from stock_sync_common import SyncError, now, dumps, stamp
from stock_sync_policy import supply_result, resolve_target_stock, matches
from stock_sync_worker import recover
from stock_sync_jobs import claim, acquire
from stock_sync_guard import legacy_write
from stock_sync_schema import migrate


@pytest.mark.parametrize('available,safety,unit,expected',[(13,3,1,10),(25,0,10,2),(0,0,1,0),(8,3,10,0)])
def test_quantity_rounds_down_and_safety_once(available,safety,unit,expected):
    s={'pool_id':'one','authority':'local','available':available}
    assert supply_result([s,s],safety,unit)['quantity']==expected


def test_multisource_deduplicates_and_handles_unknown():
    sources=[{'pool_id':'a','authority':'local','available':8},{'pool_id':'b','authority':'local','available':12}]
    assert supply_result(sources,3,1)['quantity']==17
    sources[0]['available']=0
    assert supply_result(sources,0,1)['quantity']==12
    sources[0]['available']=None
    assert supply_result(sources,0,1)['quantity'] is None


@pytest.mark.parametrize('quantity,manual,status',[(8,'unknown','instock'),(0,'unknown','unknown'),(0,'outofstock','outofstock'),(0,'instock','instock')])
def test_mixed_source_uses_status(quantity,manual,status):
    r=supply_result([{'pool_id':'a','authority':'local','available':quantity},{'pool_id':'b','authority':'manual_partner','status':manual}],0,1)
    assert r['mode']=='status' and r['quantity'] is None and r['status']==status


def test_controls_precedence_and_no_fake_restore():
    s={'mode':'quantity','quantity':0}
    hold=[{'kind':'manual_hold','stock_status':'outofstock'},{'kind':'reference_snapshot','stock_status':'instock'}]
    p,r=resolve_target_stock({'mode':'quantity'},s,hold)
    assert p['stock_status']=='outofstock' and r=='MANUAL_HOLD_ACTIVE'
    p,r=resolve_target_stock({'mode':'quantity'},s,hold[1:])
    assert p['manage_stock'] and p['stock_quantity']==0
    with pytest.raises(SyncError):resolve_target_stock({'mode':'quantity'},{'mode':'status','status':'instock'},hold[1:])
    assert not matches({'manage_stock':'parent'},{'manage_stock':True})
    assert not matches({'stock_quantity':None},{'stock_quantity':0})


@pytest.mark.parametrize('source,target,count',[(20,3,3),(3,20,3),(5,2,2)])
def test_reference_intersection_preserves_unselected(system,source,target,count):
    system.seed(source,target)
    outside=copy.deepcopy({k:v for k,v in system.http.products.items() if 'outside' in k or ('target' in k and v['id']>2000+count)})
    p=system.plan('reference_status')
    assert p['status']=='ready' and p['summary']['change']==count
    j=system.execute(p)
    assert j['status']=='succeeded'
    assert len(system.http.puts)==count
    assert all(system.http.products[k]==v for k,v in outside.items())
    assert system.sql('SELECT * FROM inv_movements')==[]


def test_cross_page_all_exclusions_filter_and_new_items(system):
    system.seed(105,105)
    p=system.plan('reference_status',selection={'mode':'all','excluded_catalog_item_ids':[105]})
    assert p['summary']['target_resources']==104
    # Source catalog additions cannot expand the confirmed target set.
    system.http.products['https://reference.test/wp-json/wc/v3/products/9999']={**system.http.products['https://reference.test/wp-json/wc/v3/products/1001'],'id':9999}
    j=system.execute(p);assert len(j['items'])==104
    assert not any('/2105' in url for url,_ in system.http.puts)
    p=system.plan('reference_status',selection={'mode':'filtered_all','filter':{'search':'Style 10'}})
    assert p['summary']['selected_leaves']==7


def test_partial_selection_and_no_match(system):
    system.seed(5,2)
    p=system.plan('reference_status',selection={'mode':'explicit','catalog_item_ids':[1,3,4]})
    assert p['summary']['target_resources']==1 and p['summary']['unmatched_pairs']==2


def test_incomplete_catalog_blocks_all_not_explicit(system):
    system.seed(101,101)
    system.http.fail_get=lambda url,kw:url.endswith('/products') and kw['params']['page']==2
    r=system.call('/catalog-scans','POST',{'source_site_id':1});id_=r.get_json()['id'];system.work()
    snap=system.call('/catalog-scans/'+id_).get_json()
    assert not snap['complete'] and snap['progress']==100
    base={'operation':'reference_status','source_site_id':1,'catalog_snapshot_id':id_,'target_scope':{'mode':'explicit_sites','site_ids':[2]},'reason':'test'}
    r=system.call('/plans','POST',{**base,'selection':{'mode':'all'}})
    assert r.status_code==409 and r.get_json()['code']=='SOURCE_INCOMPLETE'
    r=system.call('/plans','POST',{**base,'selection':{'mode':'explicit','catalog_item_ids':[1]}})
    assert r.status_code==202


def test_owner_isolation_shared_reference_and_csrf(system):
    system.seed();system.login(2)
    assert {s['id'] for s in system.call('/options').get_json()['target_sites']}=={2}
    r=system.call('/catalog-scans','POST',{'target_scope':{'mode':'explicit_sites','site_ids':[3]}})
    assert r.status_code==403
    assert system.call('/catalog-scans','POST',{'source_site_id':1}).status_code==403
    system.login(1);assert system.call('/reference-sites/1','PUT',{'enabled':True}).status_code==200
    system.login(2);p=system.plan('reference_status');assert p['status']=='ready'
    body=json.dumps(p)
    assert 'fixture-cs' not in body and 'on_hand' not in body and 'consumer_key' not in body
    r=system.client.post('/api/product-manager/stock-sync/jobs',json={})
    assert r.status_code==403
    assert system.call('/reference-sites/1','PUT',{'enabled':False}).status_code==403


@pytest.mark.parametrize('change', ['owner','account','product_permission','shared_reference'])
def test_current_permissions_rechecked_at_execution(system,change):
    system.seed();system.call('/reference-sites/1','PUT',{'enabled':True});system.login(2)
    p=system.plan('reference_status')
    r=system.call('/jobs','POST',{'plan_id':p['id'],'plan_version':1,'idempotency_key':'key','accepted_item_ids':[i['id'] for i in p['items']]});j=r.get_json()['id']
    if change=='owner':system.sql("UPDATE sites SET manager='Changed' WHERE id=2")
    elif change=='account':system.sql('UPDATE users SET is_active=0 WHERE id=2')
    elif change=='product_permission':system.sql('UPDATE users SET can_manage_products=0 WHERE id=2')
    else:system.sql('UPDATE stock_sync_reference_sites SET enabled=0 WHERE site_id=1')
    system.work()
    assert not system.http.puts
    assert all(i['status']=='conflict' for i in system.sql('SELECT status FROM stock_sync_job_items WHERE job_id=?',(j,)))


def test_manual_hold_reference_recovery_and_layered_release(system):
    system.seed()
    j=system.execute(system.plan('manual_hold'))
    assert j['status']=='succeeded'
    system.execute(system.plan('reference_status'))
    assert all(p['stock_status']=='outofstock' for url,p in system.http.products.items() if 'target' in url)
    controls=system.call('/controls').get_json()['items'];super_ids=[c['id'] for c in controls if c['kind']=='manual_hold']
    system.login(2);system.execute(system.plan('manual_hold'))
    # A regular admin role still cannot clear a superadmin hold.
    scan=system.call('/catalog-scans','POST',{}).get_json()['id'];system.work()
    r=system.call('/plans','POST',{'operation':'release_hold','reason':'test','catalog_snapshot_id':scan,'selection':{'mode':'all'},'target_scope':{'mode':'explicit_sites','site_ids':[2]},'control_ids':super_ids})
    assert r.status_code==403
    own=[c['id'] for c in system.call('/controls').get_json()['items'] if c['kind']=='manual_hold' and c['protection']=='owner']
    j=system.execute(system.plan('release_hold',extra={'control_ids':own,'confirm_available':True}))
    assert j['status']=='succeeded'
    assert all(i['after']['stock_status']=='outofstock' for i in j['items'])
    system.login(1)
    j=system.execute(system.plan('release_hold',extra={'control_ids':super_ids,'confirm_available':True}))
    assert j['status']=='succeeded' and all(i['after']['stock_status']=='instock' for i in j['items'])


def test_reference_hold_replaced_without_clearing_manual(system):
    system.seed(1,1)
    source=system.http.products['https://reference.test/wp-json/wc/v3/products/1001'];source['stock_status']='outofstock'
    system.execute(system.plan('reference_status'))
    source['stock_status']='instock'
    system.execute(system.plan('reference_status'))
    active=system.sql("SELECT * FROM stock_sync_controls WHERE active=1 AND kind='reference_snapshot'")
    assert len(active)==1 and active[0]['stock_status']=='instock'


def quantity_setup(s,count=1):
    s.seed(count,count)
    s.sql('UPDATE oms_sku_warehouses SET warehouse_id=2')
    for k in range(1,count+1):s.sql('INSERT INTO inv_stock VALUES(2,?,20,7)',(k,))
    s.sql("INSERT INTO inv_site_sync_config VALUES(2,'off','quota',3)")


def test_quantity_reserved_pack_safety_and_zero_restore(system):
    quantity_setup(system)
    p=system.plan('quantity');assert p['items'][0]['intended']['stock_quantity']==10
    system.execute(p)
    system.execute(system.plan('manual_hold'))
    system.sql('UPDATE inv_stock SET on_hand=7')
    controls=[x['id'] for x in system.call('/controls').get_json()['items'] if x['kind']=='manual_hold']
    j=system.execute(system.plan('release_hold',extra={'control_ids':controls}))
    assert j['items'][0]['after']['stock_quantity']==0 and j['items'][0]['after']['manage_stock'] is True
    assert system.sql('SELECT on_hand,reserved FROM inv_stock')==[{'on_hand':7,'reserved':7}]


@pytest.mark.parametrize('kind,code',[('missing','INVENTORY_UNKNOWN'),('backlog','ORDER_SYNC_BACKLOG'),('quota','PUBLISH_STRATEGY_CONFLICT'),('unit','INVALID_UNIT')])
def test_quantity_conflicts(system,kind,code):
    quantity_setup(system)
    if kind=='missing':system.sql('DELETE FROM inv_stock')
    elif kind=='backlog':
        system.sql("INSERT INTO sync_runs VALUES('backlog','running','2026-09-16')")
        system.sql("INSERT INTO sync_site_progress VALUES('backlog',1,'fetching')")
    elif kind=='quota':system.sql("UPDATE inv_site_sync_config SET mode='live'")
    else:system.sql('UPDATE inv_site_sku_map SET qty_per_item=0 WHERE site_id=2')
    p=system.plan('quantity');assert p['items'][0]['reason']==code
    assert not system.http.puts


@pytest.mark.parametrize('kind,code',[('unit','UNIT_NOT_COMPARABLE'),('backorder','BACKORDER_POLICY_CONFLICT'),('draft','UNSUPPORTED_PRODUCT'),('master','PHYSICAL_SCOPE_CONFLICT'),('remap','MAPPING_CHANGED'),('duplicate','AMBIGUOUS_MAPPING')])
def test_reference_conflicts(system,kind,code):
    system.seed(1,1)
    if kind=='unit':system.sql('UPDATE inv_site_sku_map SET qty_per_item=10 WHERE site_id=2')
    elif kind=='backorder':system.http.products['https://reference.test/wp-json/wc/v3/products/1001']['stock_status']='onbackorder'
    elif kind=='draft':system.http.products['https://target.test/wp-json/wc/v3/products/2001']['status']='draft'
    elif kind=='master':system.sql('UPDATE sites SET product_master_id=9 WHERE id=2')
    elif kind=='duplicate':system.sql('INSERT INTO inv_site_sku_map SELECT 9999,site_id,wc_product_id,wc_variation_id,sku_id,qty_per_item,is_active,updated_at FROM inv_site_sku_map WHERE id=2001')
    else:
        system.execute(system.plan('manual_hold'))
        system.sql("UPDATE inv_site_sku_map SET updated_at='changed' WHERE id=2001")
    p=system.plan('reference_status');assert p['items'][0]['reason']==code


def test_source_alias_disagreement(system):
    system.seed(1,1)
    system.sql('INSERT INTO inv_site_sku_map VALUES(1999,1,1999,0,1,1,TRUE,NULL)')
    system.http.products['https://reference.test/wp-json/wc/v3/products/1999']={**system.http.products['https://reference.test/wp-json/wc/v3/products/1001'],'id':1999,'stock_status':'outofstock'}
    p=system.plan('reference_status');assert p['items'][0]['reason']=='SOURCE_STATUS_CONFLICT'


def test_variation_parent_and_wcms_bridge(system):
    system.seed(1,1)
    url='https://target.test/wp-json/wc/v3/products/2001'
    system.http.products[url].update(type='variable',manage_stock=True,variations=[4001])
    system.http.products[url+'/variations/4001']={'id':4001,'parent_id':2001,'type':'variation','status':'publish','manage_stock':'parent','stock_quantity':None,'stock_status':'instock','backorders':'no'}
    system.sql('UPDATE inv_site_sku_map SET wc_variation_id=4001 WHERE id=2001')
    assert system.plan()['items'][0]['reason']=='PARENT_STOCK_SHARED'
    system.http.products[url]['manage_stock']=False
    system.http.products[url+'/variations/4001'].update(manage_stock=False,meta_data=[{'key':'wcms_stock_manage','value':'no'}])
    system.http.bridge=True
    j=system.execute(system.plan());assert j['status']=='succeeded'
    for _,body in system.http.puts:
        assert all(m['key']!='wcms_stock_qty' for m in body.get('meta_data',[]))


def test_idempotent_confirmation_concurrent_claim_and_lease(system):
    system.seed(1,1);p=system.plan()
    body={'plan_id':p['id'],'plan_version':1,'idempotency_key':'same','accepted_item_ids':[i['id'] for i in p['items']]}
    first=system.call('/jobs','POST',body);second=system.call('/jobs','POST',body)
    assert first.get_json()==second.get_json()
    assert system.call('/jobs','POST',{**body,'accepted_item_ids':['changed']}).status_code==409
    def claim_one(n):
        c=system.connect()
        try:return claim(c,'worker'+str(n))
        finally:c.close()
    with concurrent.futures.ThreadPoolExecutor(2) as ex:
        results=list(ex.map(claim_one,range(2)))
    assert sum(x is not None for x in results)==1
    c=system.connect()
    try:
        acquire(c,['physical'],'one')
        with pytest.raises(SyncError):acquire(c,['physical'],'two')
    finally:c.close()


@pytest.mark.parametrize('kind,expected',[('timeout','verified_success'),('mismatch','failed'),('429','failed'),('401','failed')])
def test_http_results_require_authoritative_readback(system,kind,expected):
    system.seed(1,1)
    system.http.products['https://target.test/wp-json/wc/v3/products/2001']['stock_status']='instock'
    p=system.plan()
    if kind=='timeout':system.http.timeout_after_put=True
    elif kind=='mismatch':system.http.ignore_put=True
    else:system.http.fail_put=int(kind)
    j=system.execute(p)
    assert j['items'][0]['status']==expected
    assert len(system.http.puts)==1


@pytest.mark.parametrize('change',['target','source','inventory','mapping'])
def test_stale_plan_cannot_overwrite_new_inputs(system,change):
    quantity_setup(system)
    p=system.plan('quantity' if change=='inventory' else 'reference_status')
    if change=='target':system.http.products['https://target.test/wp-json/wc/v3/products/2001']['stock_quantity']=5
    elif change=='source':system.http.products['https://reference.test/wp-json/wc/v3/products/1001']['stock_status']='outofstock'
    elif change=='inventory':system.sql('UPDATE inv_stock SET reserved=8')
    else:system.sql('UPDATE inv_site_sku_map SET qty_per_item=2 WHERE id=2001')
    j=system.execute(p)
    assert j['items'][0]['status']=='conflict' and not system.http.puts


def test_cancel_after_inflight_finishes_and_retry_excludes_success(system):
    system.seed(3,3)
    p=system.plan('reference_status')
    def cancel_after_one(url):
        system.sql('UPDATE stock_sync_jobs SET cancel_requested=1')
    system.http.on_put=cancel_after_one
    j=system.execute(p)
    assert j['counts']=={'verified_success':1,'cancelled':2}
    failed=[i['id'] for i in j['items'] if i['status']=='cancelled']
    r=system.call('/jobs/'+j['id']+'/retry-plan','POST',{'item_ids':failed});assert r.status_code==202
    system.work();p2=system.call('/plans/'+r.get_json()['id']).get_json()
    assert p2['summary']['target_resources']==2
    system.http.on_put=None
    assert system.execute(p2)['status']=='succeeded'
    assert len(system.http.puts)==3


def test_unknown_write_holds_lease_then_readonly_recovery(system):
    system.seed(1,1)
    p=system.plan('reference_status')
    system.http.on_put=lambda url:setattr(system.http,'fail_get',lambda u,kw:'target' in u)
    j=system.execute(p)
    assert j['status']=='requires_review' and j['items'][0]['controls_saved']
    assert system.sql('SELECT * FROM stock_sync_resource_leases')
    assert system.call('/jobs/'+j['id']+'/retry-plan','POST',{'item_ids':[j['items'][0]['id']]}).status_code==409
    system.http.fail_get=None
    c=system.connect()
    try:recover(c,system.worker_id,system.woo)
    finally:c.close()
    assert len(system.http.puts)==1
    assert system.call('/jobs/'+j['id']).get_json()['status']=='succeeded'
    assert system.sql('SELECT * FROM stock_sync_resource_leases')==[]


def test_legacy_writes_cannot_bypass_and_price_is_allowed(system):
    system.seed(1,1);system.execute(system.plan())
    url='https://target.test/wp-json/wc/v3/products/2001'
    with pytest.raises(SyncError):
        with legacy_write(url,{'stock_status':'instock'}):pytest.fail('must block')
    with legacy_write(url,{'regular_price':'12'}):pass
    assert system.sql('SELECT * FROM stock_sync_resource_leases')==[]


def test_feature_off_and_migration_idempotency(system,monkeypatch):
    c=system.connect();migrate(c);migrate(c);c.close()
    assert len(system.sql('SELECT * FROM stock_sync_schema_migrations'))==1
    monkeypatch.setenv('STOCK_SYNC_ENABLED','0')
    assert system.call('/options').status_code==503
