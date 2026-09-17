import datetime
import json
import uuid
import concurrent.futures
import os

import pytest

from test_inventory_workflows import system, client, payload, post, value, detail, review
from inv_temporary_access_schema import migrate


def grant(system, uid=2, wid=2, **extra):
    data=dict(user_id=uid, warehouse_id=wid, reason='warehouse physical stock check',
              expires_at=(datetime.datetime.now(datetime.timezone.utc)+datetime.timedelta(hours=2)).isoformat(),
              request_key=str(uuid.uuid4()))
    data.update(extra)
    r=client(system,1).post('/api/inv/operations/temporary-access',json=data)
    assert r.status_code==200,r.get_json()
    return r.get_json()['id'],data


def count_payload(gid, **extra):
    return payload('stocktake',auto_approve=True,grant_id=gid,
                   items=[dict(sku_id=1,qty=12,baseline=10,baseline_reserved=2,baseline_movement=0)],**extra)


def revoke(system,gid):
    r=client(system,1).post(f'/api/inv/operations/temporary-access/{gid}/revoke',json={'reason':'counting complete'})
    assert r.status_code==200,r.get_json()
    return r.get_json()


def test_auto_count_records_operator_system_approval_grant_and_movement(system):
    gid,_=grant(system)
    data=count_payload(gid)
    first=post(system,data)
    assert first['status']=='approved'
    d=detail(system,first['id'])
    assert d['created_by']==2 and d['reviewed_by'] is None
    assert d['reviewed_name']=='系统自动审批' and d['reviewed_at']
    assert d['events'][-1]['action']=='auto_approved'
    assert d['events'][-1]['actor_id']==0
    assert json.loads(d['events'][-1]['detail'])['grant_id']==gid
    assert value(system,'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1')==12
    assert value(system,'SELECT reserved FROM inv_stock WHERE warehouse_id=2 AND sku_id=1')==2
    assert value(system,'SELECT operator_id FROM inv_movements')==2
    assert value(system,'SELECT qty_before FROM inv_movements')==10
    assert value(system,'SELECT qty_after FROM inv_movements')==12
    assert post(system,data)['replayed']
    revoke(system,gid)
    assert post(system,data)['replayed']
    assert value(system,'SELECT COUNT(*) FROM inv_movements')==1
    assert review(system,first['id'])['replayed']
    review(system,first['id'],decision='reject',expected=409)


def test_grant_and_revoke_admin_only_and_audited(system):
    gid,data=grant(system)
    url='/api/inv/operations/temporary-access'
    for uid in (2,3,4):
        assert client(system,uid).get(url).status_code==403
        assert client(system,uid).post(url,json=data).status_code==403
        assert client(system,uid).get('/settings/inventory-access').status_code==403
        assert client(system,uid).post(f'{url}/{gid}/revoke',json={'reason':'x'}).status_code==403
    assert client(system,1).post(url,json=data).get_json()['replayed']
    assert client(system,1).post(url,json=dict(data,reason='different')).status_code==409
    revoke(system,gid)
    assert revoke(system,gid)['replayed']
    assert value(system,"SELECT COUNT(*) FROM inv_document_events WHERE action='temporary_stock_granted'")==1
    assert value(system,"SELECT COUNT(*) FROM inv_document_events WHERE action='temporary_stock_revoked'")==1
    rows=client(system,1).get(url).get_json()['grants']
    assert not rows[0]['active'] and rows[0]['revoked_by']==1


def test_expired_revoked_and_foreign_grants_do_not_change_stock(system):
    gid,_=grant(system)
    post(system,count_payload(gid,warehouse_id=3),uid=4,expected=403)
    revoke(system,gid)
    post(system,count_payload(gid),expected=403)
    gid,_=grant(system)
    c=system[1]();c.execute("UPDATE inv_temporary_stock_grants SET expires_at='2000-01-01T00:00:00+00:00' WHERE id=?",(gid,));c.commit();c.close()
    post(system,count_payload(gid),expected=403)
    assert value(system,'SELECT COUNT(*) FROM inv_documents')==0
    assert value(system,'SELECT COUNT(*) FROM inv_movements')==0


def test_temporary_grant_does_not_grant_other_capabilities(system):
    gid,_=grant(system,uid=5)
    ctx=client(system,5).get('/api/inv/operations/context').get_json()
    assert [w['id'] for w in ctx['warehouses']]==[2]
    w=ctx['warehouses'][0]
    assert w['can_stocktake'] and w['auto_stocktake_grant']==gid
    assert not any(w[k] for k in ('can_receive','can_transfer','can_approve'))
    assert client(system,5).get('/api/inv/operations/catalog/1').status_code==403
    post(system,payload(),uid=5,expected=403)
    post(system,payload('transfer',source_warehouse_id=2,warehouse_id=1),uid=5,expected=400)
    post(system,payload(auto_approve=True,grant_id=gid),uid=5,expected=400)
    assert value(system,'SELECT can_manage_inventory FROM users WHERE id=5')==0
    revoke(system,gid)
    assert client(system,5).get('/api/inv/operations/context').status_code==403


@pytest.mark.parametrize('items',[
    [dict(sku_id=1,qty=1,baseline=10,baseline_reserved=2,baseline_movement=0)],
    [dict(sku_id=1,qty=12,baseline=9,baseline_reserved=2,baseline_movement=0)],
    [dict(sku_id=1,qty=12,baseline=10,baseline_reserved=1,baseline_movement=0)],
    [dict(sku_id=1,qty=12,baseline=10,baseline_reserved=2,baseline_movement=99)],
])
def test_automatic_count_preserves_reserved_and_stale_baselines(system,items):
    gid,_=grant(system)
    data=count_payload(gid);data['items']=items
    post(system,data,expected=409)
    assert value(system,'SELECT COUNT(*) FROM inv_documents')==0
    assert value(system,'SELECT COUNT(*) FROM inv_movements')==0


def test_auto_approval_failure_rolls_back_entire_document(system,monkeypatch):
    import inv_workflows
    gid,_=grant(system)
    original=inv_workflows.event
    def fail(conn,did,action,detail,**kw):
        if action=='auto_approved':
            raise RuntimeError('audit unavailable')
        return original(conn,did,action,detail,**kw)
    monkeypatch.setattr(inv_workflows,'event',fail)
    with pytest.raises(RuntimeError):
        post(system,count_payload(gid))
    assert value(system,'SELECT COUNT(*) FROM inv_documents')==0
    assert value(system,'SELECT COUNT(*) FROM inv_movements')==0
    assert value(system,'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1')==10


def test_ordinary_counts_remain_manual_and_cannot_self_approve(system):
    grant(system)
    data=count_payload(1);data.pop('auto_approve');data.pop('grant_id')
    did=post(system,data)['id']
    assert detail(system,did)['status']=='submitted'
    assert value(system,'SELECT COUNT(*) FROM inv_movements')==0
    review(system,did,uid=2,expected=403)


@pytest.mark.parametrize('field,value_', [('expires_at',''),('expires_at','2020-01-01T00:00:00Z'),
    ('expires_at','2030-01-01T00:00:00'),('reason',''),('warehouse_id',4),('warehouse_id',5)])
def test_invalid_grant_rejected(system,field,value_):
    data=dict(user_id=2,warehouse_id=2,reason='x',expires_at='2030-01-01T00:00:00Z',request_key=str(uuid.uuid4()))
    data[field]=value_
    assert client(system,1).post('/api/inv/operations/temporary-access',json=data).status_code==400
    assert value(system,'SELECT COUNT(*) FROM inv_temporary_stock_grants')==0


def test_migration_is_additive_and_grants_no_one(system):
    c=system[1]();migrate(c);migrate(c);c.close()
    assert value(system,'SELECT COUNT(*) FROM inv_temporary_stock_grants')==0
    assert value(system,'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1')==10


def test_simultaneous_duplicate_stocktake_is_applied_once(system):
    gid,_=grant(system)
    data=count_payload(gid)
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
        results=list(pool.map(lambda _:post(system,data),range(2)))
    assert results[0]['id']==results[1]['id']
    assert value(system,'SELECT COUNT(*) FROM inv_documents')==1
    assert value(system,'SELECT COUNT(*) FROM inv_movements')==1


def test_expiry_during_transaction_rolls_back(system,monkeypatch):
    import inv_workflows
    gid,_=grant(system)
    times=iter(['2000-01-01T00:00:00+00:00','2099-01-01T00:00:00+00:00'])
    monkeypatch.setattr(inv_workflows,'utc_now',lambda:next(times))
    post(system,count_payload(gid),expected=403)
    assert value(system,'SELECT COUNT(*) FROM inv_documents')==0
    assert value(system,'SELECT COUNT(*) FROM inv_movements')==0


@pytest.mark.skipif(not os.environ.get('INVENTORY_WORKFLOW_BROWSER'),reason='opt-in browser acceptance')
def test_browser_admin_grant_shipper_stocktake_and_revoke(system,tmp_path):
    import threading
    import requests
    from pathlib import Path
    from flask import Response
    from jinja2 import ChoiceLoader,DictLoader,FileSystemLoader
    from playwright.sync_api import sync_playwright,expect
    from werkzeug.serving import make_server
    app=system[0]
    assets={}
    for key,suffix in [('css','css/bootstrap.min.css'),('js','js/bootstrap.bundle.min.js')]:
        response=requests.get('https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/'+suffix,timeout=30)
        response.raise_for_status();assets[key]=response.content
    app.jinja_loader=ChoiceLoader([DictLoader({'base.html':'''<!doctype html><html><head><meta charset="UTF-8">
        <meta name="viewport" content="width=device-width,initial-scale=1"><link rel="stylesheet" href="/test-assets/css">
        </head><body class="bg-dark text-white">{% block content %}{% endblock %}<script src="/test-assets/js"></script></body></html>'''}),
        FileSystemLoader(str(Path(__file__).resolve().parents[1]/'templates'))])
    @app.get('/test-assets/<kind>')
    def assets_route(kind):
        return Response(assets[kind],mimetype='text/css' if kind=='css' else 'application/javascript')
    server=make_server('127.0.0.1',0,app,threaded=True)
    thread=threading.Thread(target=server.serve_forever,daemon=True);thread.start()
    base=f'http://127.0.0.1:{server.server_port}'
    try:
        with sync_playwright() as p:
            browser=p.chromium.launch(executable_path=os.environ['INVENTORY_WORKFLOW_BROWSER'],headless=True)
            errors=[]
            def page_for(uid):
                context=browser.new_context(viewport={'width':1440,'height':1000})
                context.add_cookies([{'name':'session','value':client(system,uid).get_cookie('session').value,'url':base}])
                page=context.new_page();page.on('pageerror',lambda e:errors.append(str(e)))
                page.on('dialog',lambda dialog:dialog.accept('counting complete') if dialog.type=='prompt' else dialog.accept())
                return page
            admin=page_for(1);admin.goto(base+'/settings/inventory-access')
            admin.locator('#accessUser').select_option('2');admin.locator('#accessWarehouse').select_option('2')
            admin.locator('#accessExpiry').fill('2030-01-01T12:30')
            admin.locator('#accessReason').fill('physical recount')
            admin.get_by_role('button',name='确认临时授权',exact=True).click()
            expect(admin.locator('#accessRows')).to_contain_text('有效')
            shipper=page_for(2);shipper.goto(base+'/inventory/operations')
            shipper.locator('#opTemporaryStock').click()
            expect(shipper.locator('#autoStockNotice')).to_contain_text('授权 #1')
            shipper.locator('#formReference').fill('browser-stocktake')
            shipper.locator('#formNote').fill('physical stock checked')
            shipper.locator('tr[data-sku="1"] .choose').check()
            shipper.locator('tr[data-sku="1"] .qty').fill('12')
            shipper.get_by_role('button',name='确认修改库存并自动生成盘点单',exact=True).click()
            expect(shipper.locator('#opTitle')).to_contain_text('已入账')
            expect(shipper.locator('#opBody')).to_contain_text('系统自动审批')
            assert value(system,'SELECT on_hand FROM inv_stock WHERE warehouse_id=2 AND sku_id=1')==12
            shipper.screenshot(path=str(tmp_path/'temporary-stocktake-approved.png'),full_page=True)
            admin.locator('[data-revoke="1"]').click()
            expect(admin.locator('#accessRows')).to_contain_text('已撤销')
            shipper.reload();expect(shipper.locator('#opTemporaryStock')).to_be_hidden()
            admin.set_viewport_size({'width':390,'height':844})
            assert admin.evaluate('document.documentElement.scrollWidth <= window.innerWidth')
            admin.screenshot(path=str(tmp_path/'temporary-grants-mobile.png'),full_page=True)
            assert errors==[]
            browser.close()
    finally:
        server.shutdown();thread.join(timeout=5)
