"""Observe durable progress from another connection while Woo reads are pending."""
from stock_sync_fixtures import system


def variable(s, pid):
    root=f'https://reference.test/wp-json/wc/v3/products/{pid}'
    parent=s.http.products[root]
    parent.update(type='variable',variations=[pid*10+1,pid*10+2])
    for vid in parent['variations']:
        s.http.products[root+f'/variations/{vid}']=dict(parent,id=vid,parent_id=pid,
            type='variation',attributes=[{'name':'flavor','option':f'Flavor {vid}'}])


def test_progress_visible_before_variations_in_first_page_finish(system):
    s=system;s.seed(3,3);variable(s,1002)
    id_=s.call('/catalog-scans','POST',{'source_site_id':1}).get_json()['id']
    queued=s.call('/catalog-scans/'+id_).get_json()
    assert queued['status']=='queued' and queued['read_progress']=={}
    observed=[]
    s.http.on_get=lambda url:observed.append(s.call('/catalog-scans/'+id_).get_json()) if url.endswith('/variations') else None
    s.work()
    assert len(observed)==1
    pending=observed[0]
    assert pending['status']=='scanning' and not pending['complete'] and pending['progress']==1
    detail=pending['read_progress']
    assert detail['products_read']==1 and detail['total_products']==3
    assert detail['phase']=='variations' and detail['current_product']=='Style 2'
    assert detail['total_variations']==2 and detail['variations_read']==0
    assert detail['updated_at'] and detail['started_at']
    finished=s.call('/catalog-scans/'+id_).get_json()
    assert finished['complete'] and finished['progress']==4
    assert finished['read_progress']['phase']=='complete' and finished['read_progress']['products_read']==3
    assert not s.http.puts


def test_variation_failure_keeps_only_complete_products_and_failed_progress(system):
    s=system;s.seed(2,2);variable(s,1002)
    s.http.fail_get=lambda url,kw:url.endswith('/variations')
    id_=s.call('/catalog-scans','POST',{'source_site_id':1}).get_json()['id'];s.work()
    status=s.call('/catalog-scans/'+id_).get_json()
    assert status['status']=='incomplete' and not status['complete']
    assert status['error']=='SOURCE_READ_FAILED' and status['progress']==1
    detail=status['read_progress']
    assert detail['phase']=='failed' and detail['products_read']==1 and detail['total_products']==2
    items=s.call('/catalog-snapshots/'+id_+'/products').get_json()['items']
    assert len(items)==1 and items[0]['product_id']==1001 and items[0]['complete']
    assert not s.http.puts
