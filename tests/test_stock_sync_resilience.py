import concurrent.futures
import copy
import json
from datetime import timedelta

import pytest

from stock_sync_fixtures import system
from stock_sync_common import now,stamp,SyncError
from stock_sync_jobs import confirm
from stock_sync_permissions import actor
from stock_sync_worker import recover
from stock_sync_admin import capability,binding_preview,rebind
from stock_sync_guard import legacy_write
from stock_sync_policy import supply_result
import stock_sync_worker as worker


@pytest.mark.parametrize('contract,age,expected',[(False,0,'EXTERNAL_RESERVATION_UNCONFIRMED'),(True,20000,'EXTERNAL_STOCK_STALE'),(True,0,None)])
def test_external_reservations_freshness_and_no_double_deduction(system,contract,age,expected):
    system.seed(1,1)
    system.sql('UPDATE oms_sku_warehouses SET warehouse_id=3')
    system.sql('UPDATE oms_warehouse_integrations SET config_json=? WHERE warehouse_id=3',(json.dumps({'stock_sync_available_includes_oms_reservations':contract}),))
    system.sql('INSERT INTO oms_external_stock VALUES(3,1,17,?,?)',(stamp(-age),stamp(-age)))
    p=system.plan('quantity')
    if expected:assert p['items'][0]['reason']==expected
    else:
        assert p['items'][0]['intended']['stock_quantity']==17
        assert system.execute(p)['status']=='succeeded'


def test_same_key_concurrent_confirmation_returns_one_job(system):
    system.seed(1,1);p=system.plan()
    request={'plan_id':p['id'],'plan_version':1,'idempotency_key':'simultaneous','accepted_item_ids':[i['id'] for i in p['items']]}
    def submit(n):
        c=system.connect()
        try:return confirm(c,actor(c,1),request)
        finally:c.close()
    with concurrent.futures.ThreadPoolExecutor(2) as ex:
        results=list(ex.map(submit,[1,2]))
    assert len(set(results))==1
    assert len(system.sql('SELECT * FROM stock_sync_jobs'))==1
    system.work();assert len(system.sql('SELECT * FROM stock_sync_controls'))==1


def test_saved_intent_survives_result_storage_failure(system,monkeypatch):
    system.seed(1,1);p=system.plan('reference_status')
    request={'plan_id':p['id'],'plan_version':1,'idempotency_key':'storage-failure','accepted_item_ids':[i['id'] for i in p['items']]}
    job=system.call('/jobs','POST',request).get_json()['id']
    original=worker.save_result
    def fail(c,j,i,status,error,after=None):
        if status=='verified_success':raise RuntimeError('database temporarily unavailable')
        return original(c,j,i,status,error,after)
    monkeypatch.setattr(worker,'save_result',fail)
    with pytest.raises(RuntimeError):system.work()
    assert system.sql('SELECT intent_json FROM stock_sync_job_items')[0]['intent_json']
    assert system.sql('SELECT quarantined FROM stock_sync_resource_leases')[0]['quarantined']
    monkeypatch.setattr(worker,'save_result',original)
    c=system.connect()
    try:recover(c,system.worker_id,system.woo)
    finally:c.close()
    assert system.call('/jobs/'+job).get_json()['status']=='succeeded'
    assert len(system.http.puts)==1


def test_newer_hold_supersedes_old_reference_restore(system):
    system.seed(1,1)
    old=system.plan('reference_status')
    system.execute(system.plan('manual_hold'))
    job=system.execute(old)
    assert job['items'][0]['status']=='conflict'
    assert system.http.products['https://target.test/wp-json/wc/v3/products/2001']['stock_status']=='outofstock'


def test_source_changes_mid_job_stop_remaining_same_sku(system):
    system.seed(1,1)
    p=system.plan('reference_status',targets=[2,3])
    system.http.on_put=lambda url:system.http.products['https://reference.test/wp-json/wc/v3/products/1001'].update(stock_status='outofstock')
    j=system.execute(p)
    assert j['counts']=={'verified_success':1,'conflict':1}
    assert len(system.http.puts)==1


def test_quantity_multisource_and_no_ledger_side_effect(system):
    system.seed(1,1)
    system.sql("UPDATE oms_warehouse_integrations SET inventory_authority='local' WHERE warehouse_id=1")
    system.sql('INSERT INTO oms_sku_warehouses VALUES(1,2,TRUE)')
    system.sql('INSERT INTO inv_stock VALUES(1,1,8,0)')
    system.sql('INSERT INTO inv_stock VALUES(2,1,17,5)')
    system.sql("INSERT INTO inv_site_sync_config VALUES(2,'off','mirror',3)")
    system.sql('UPDATE inv_site_sku_map SET qty_per_item=10 WHERE site_id=2')
    before=system.sql('SELECT * FROM inv_stock')
    p=system.plan('quantity');assert p['items'][0]['intended']['stock_quantity']==1
    system.execute(p)
    assert system.sql('SELECT * FROM inv_stock')==before


def test_shared_physical_endpoint_conflict_and_verified_child_capability(system):
    system.seed(1,1)
    system.sql("UPDATE sites SET url='https://target.test' WHERE id=3")
    p=system.plan();assert p['items'][0]['reason']=='PHYSICAL_SCOPE_CONFLICT'
    system.sql("UPDATE sites SET url='https://outside.test' WHERE id=3")
    system.sql('UPDATE sites SET product_master_id=9 WHERE id=2')
    c=system.connect()
    review=capability(c,1,2,'isolated fixture proof')
    capability(c,1,2,'isolated fixture proof',confirmation=review['confirm_hash']);c.close()
    p=system.plan();assert p['items'][0]['decision']=='unchanged'
    assert system.execute(p)['status']=='succeeded'


def test_mapped_stock_metadata_never_changes_price_or_content(system):
    system.seed(1,1)
    before=copy.deepcopy(system.http.products['https://target.test/wp-json/wc/v3/products/2001'])
    system.execute(system.plan('reference_status'))
    after=system.http.products['https://target.test/wp-json/wc/v3/products/2001']
    for key in ('regular_price','description','images','status','sku'):assert after[key]==before[key]
    for _,payload in system.http.puts:assert not set(payload).intersection({'price','images','description','status','sku'})


def test_retry_after_is_durable_and_blocks_new_preview_http(system):
    system.seed(1,1);p=system.plan('reference_status')
    system.http.fail_put=429
    job=system.execute(p)
    assert job['items'][0]['status']=='failed'
    cooldown=system.sql('SELECT * FROM stock_sync_rate_limits')[0]
    from stock_sync_common import parse_time
    assert (parse_time(cooldown['retry_at'])-now()).total_seconds()>110
    calls=len(system.http.calls)
    retry=system.call('/jobs/'+job['id']+'/retry-plan','POST',{'item_ids':[job['items'][0]['id']]}).get_json()
    system.work()
    assert system.call('/plans/'+retry['id']).get_json()['items'][0]['reason']=='REMOTE_RATE_LIMITED'
    assert len(system.http.calls)==calls
    system.sql('UPDATE stock_sync_rate_limits SET retry_at=?',(stamp(-1),))
    system.http.fail_put=None
    fresh=system.call('/jobs/'+job['id']+'/retry-plan','POST',{'item_ids':[job['items'][0]['id']]}).get_json()
    system.work();preview=system.call('/plans/'+fresh['id']).get_json()
    assert system.execute(preview)['status']=='succeeded'


def test_legacy_guard_retains_deleted_resource(system):
    system.seed(1,1);system.execute(system.plan())
    system.sql('DELETE FROM inv_site_sku_map WHERE id=2001')
    with pytest.raises(SyncError,match='已接入'):
        with legacy_write('https://target.test/wp-json/wc/v3/products/2001',{'stock_status':'instock'}):
            pytest.fail('historical resource must remain protected')


def test_rebind_requires_review_archives_controls_and_never_publishes(system):
    system.seed(2,2);system.execute(system.plan())
    system.sql('UPDATE inv_site_sku_map SET sku_id=2 WHERE id=2001')
    c=system.connect();calls=len(system.http.puts)
    try:
        with pytest.raises(SyncError):binding_preview(c,2,2001,'mapping correction',system.woo)
        review=binding_preview(c,1,2001,'mapping correction',system.woo)
        assert len(review['review']['archive_controls'])==1
        with pytest.raises(SyncError):rebind(c,1,2001,'mapping correction','invalid',system.woo)
        result=rebind(c,1,2001,'mapping correction',review['confirm_hash'],system.woo)
        assert result['saved'] and not result['website_changed']
    finally:c.close()
    assert len(system.http.puts)==calls
    assert system.sql('SELECT * FROM stock_sync_controls WHERE map_id=2001 AND active=1')==[]
    assert system.sql("SELECT * FROM stock_sync_events WHERE kind='mapping_rebound'")
    assert system.plan()['summary']['conflict']==0


def test_capability_invalidated_when_site_identity_changes(system):
    system.seed(1,1);system.sql('UPDATE sites SET product_master_id=9 WHERE id=2')
    c=system.connect()
    review=capability(c,1,2,'fixture scope and readback checked')
    capability(c,1,2,'fixture scope and readback checked',confirmation=review['confirm_hash']);c.close()
    assert system.plan()['summary']['conflict']==0
    system.sql("UPDATE sites SET consumer_key='rotated' WHERE id=2")
    assert system.plan()['items'][0]['reason']=='PHYSICAL_SCOPE_CONFLICT'


def test_same_physical_pool_alias_is_counted_once():
    pools=[{'pool_id':'shared','warehouse_id':i,'authority':'local','available':20,'on_hand':25,'reserved':5} for i in (1,2)]
    assert supply_result(pools,3,1)['quantity']==17
    pools[1]['available']=15
    with pytest.raises(SyncError):supply_result(pools,0,1)


def test_post_commit_reservation_backlog_blocks_quantity(system):
    system.seed(1,1)
    system.sql('UPDATE oms_sku_warehouses SET warehouse_id=2')
    system.sql('INSERT INTO inv_stock VALUES(2,1,20,7)')
    system.sql('CREATE TABLE sync_page_receipts(site_id INTEGER,post_commit_status TEXT)')
    system.sql("INSERT INTO sync_page_receipts VALUES(1,'pending')")
    assert system.plan('quantity')['items'][0]['reason']=='ORDER_SYNC_BACKLOG'


@pytest.mark.parametrize('payload',[None,[],True])
def test_api_rejects_non_object_input_without_internal_error(system,payload):
    r=system.client.post('/api/product-manager/stock-sync/plans',data=json.dumps(payload),content_type='application/json',headers={'X-CSRF-Token':system.token})
    assert r.status_code==400


def test_partial_failure_retry_only_failed_resource(system):
    system.seed(3,3);p=system.plan('reference_status')
    system.http.fail_put=lambda url:401 if url.endswith('/2002') else None
    job=system.execute(p)
    assert job['status']=='partial_failed' and job['counts']=={'failed':1,'verified_success':2}
    failed=[i['id'] for i in job['items'] if i['status']=='failed']
    request=system.call('/jobs/'+job['id']+'/retry-plan','POST',{'item_ids':failed}).get_json()
    system.work();preview=system.call('/plans/'+request['id']).get_json()
    assert preview['summary']['target_resources']==1
    system.http.fail_put=None
    assert system.execute(preview)['status']=='succeeded'
    assert [u for u,p in system.http.puts].count('https://target.test/wp-json/wc/v3/products/2001')==1
    assert [u for u,p in system.http.puts].count('https://target.test/wp-json/wc/v3/products/2002')==2


def test_partial_flavors_do_not_touch_parent_or_siblings(system):
    system.seed(8,8)
    for sid,host in ((1,'reference'),(2,'target')):
        root=f'https://{host}.test/wp-json/wc/v3/products/'
        parent={**system.http.products[root+str(sid*1000+1)],'id':sid*1000,'type':'variable','manage_stock':False,
                'variations':[sid*1000+k for k in range(1,9)]}
        for k in range(1,9):
            leaf=system.http.products.pop(root+str(sid*1000+k))
            leaf.update(parent_id=parent['id'],type='variation',attributes=[{'option':f'flavor {k}'}])
            system.http.products[root+str(parent['id'])+'/variations/'+str(leaf['id'])]=leaf
            system.sql('UPDATE inv_site_sku_map SET wc_product_id=?,wc_variation_id=? WHERE id=?',(parent['id'],leaf['id'],leaf['id']))
        system.http.products[root+str(parent['id'])]=parent
    before=copy.deepcopy(system.http.products)
    preview=system.plan('reference_status',selection={'mode':'explicit','catalog_item_ids':[2,5]})
    assert preview['summary']['selected_products']==1 and preview['summary']['target_resources']==2
    assert system.execute(preview)['status']=='succeeded'
    changed={'https://target.test/wp-json/wc/v3/products/2000/variations/2002','https://target.test/wp-json/wc/v3/products/2000/variations/2005'}
    assert {u for u,p in system.http.puts}==changed
    assert all(system.http.products[k]==v for k,v in before.items() if k not in changed)


def test_duplicate_physical_resource_is_not_written_twice(system):
    system.seed(1,1)
    system.sql("UPDATE sites SET url='https://target.test' WHERE id=3")
    system.sql('UPDATE inv_site_sku_map SET wc_product_id=2001 WHERE id=3001')
    preview=system.plan('manual_hold',targets=[2,3])
    assert preview['summary']['conflict']==2
    assert not system.http.puts


def test_recreated_mapping_cannot_bypass_historical_hold(system):
    system.seed(1,1);system.execute(system.plan('manual_hold'))
    system.sql('UPDATE inv_site_sku_map SET id=9999 WHERE id=2001')
    preview=system.plan('reference_status')
    assert preview['items'][0]['reason']=='MAPPING_CHANGED'
    assert system.http.products['https://target.test/wp-json/wc/v3/products/2001']['stock_status']=='outofstock'


def test_endpoint_change_requires_explicit_rebind(system):
    system.seed(1,1);system.execute(system.plan())
    system.sql("UPDATE sites SET url='https://new-target.test' WHERE id=2")
    preview=system.plan()
    assert preview['items'][0]['reason']=='MAPPING_CHANGED'
