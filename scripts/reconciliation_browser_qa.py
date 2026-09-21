"""Browser acceptance against the local synthetic PostgreSQL application."""
import json
import sys
import os
from pathlib import Path
from playwright.sync_api import sync_playwright

BASE='http://127.0.0.1:5059'
ORDER=os.getenv('REC_QA_ORDER','1-1002')
SUFFIX=ORDER.split('-')[-1]
OUT=Path(__file__).resolve().parents[2]/'outputs'/'reconciliation-v2-dev'
OUT.mkdir(parents=True,exist_ok=True)

def run():
    with sync_playwright() as p:
        browser=p.chromium.launch(executable_path='C:/Program Files/Google/Chrome/Application/chrome.exe',headless=True)
        context=browser.new_context(viewport={'width':1440,'height':1100},accept_downloads=True)
        page=context.new_page();errors=[];failures=[]
        page.on('pageerror',lambda e:errors.append(str(e)))
        page.on('response',lambda r:failures.append({'url':r.url,'status':r.status}) if '/api/reconciliation-v2/' in r.url and r.status>=500 else None)
        def login(username):
            page.goto(BASE+'/login');page.locator('[name=username]').fill(username);page.locator('[name=password]').fill('local-reconciliation-demo')
            page.locator('button[type=submit]').click();page.wait_for_url(BASE+'/')
            page.goto(BASE+'/partner-reconciliation/modular');page.wait_for_selector('.rec-card')
        def tab(name):page.locator('[data-tab='+name+']').click()
        def act(name,ident=None):page.locator('[data-action='+name+']'+('[data-id="'+ident+'"]' if ident else '')).click()
        def fill(values):
            for key,value in values.items():
                node=page.locator('#rec-form [name="'+key+'"]')
                if node.evaluate('(n)=>n.tagName')=='SELECT':node.select_option(str(value))
                else:node.fill(str(value))
        def save():
            page.locator('#rec-form button[type=submit]').click()
            page.wait_for_function("!document.querySelector('#rec-dialog').open || !!document.querySelector('#rec-form-error').textContent")
            assert not page.locator('#rec-form-error').inner_text(),page.locator('#rec-form-error').inner_text()
            page.wait_for_function("!document.querySelector('#rec-dialog').open")
        login('admin')
        state=context.request.get(BASE+'/api/reconciliation-v2/state').json()
        partner=next(p['id'] for p in state['parties'] if p['name']=='演示 · 合作方')
        carrier=next(p['id'] for p in state['parties'] if p['name']=='演示 · 物流商')
        page.screenshot(path=str(OUT/'overview-final.png'),full_page=True)
        tab('cooperation');act('party');fill({'name':'浏览器验收主体','currency':'PLN'});save()
        tab('cooperation');act('fee');fill({'name':'浏览器验收费用版本','provider_id':partner,'amount':'20','currency':'CNY'});save()
        tab('inventory');page.screenshot(path=str(OUT/'inventory-final.png'),full_page=True)
        tab('orders');detail=context.request.get(BASE+'/api/reconciliation-v2/order/'+ORDER).json()
        if not detail['shipments']:
            act('preview',ORDER);page.wait_for_selector('#rec-detail[open]');page.locator('#rec-detail-close').click()
            act('plan',ORDER);page.wait_for_selector('#rec-detail[open]');page.locator('#rec-detail-close').click()
            fulfillment=context.request.get(BASE+'/api/fulfillment/order/'+ORDER).json()['fulfillments'][0]
            assert fulfillment['items'][0]['source_batches']
            response=context.request.post(BASE+'/api/fulfillment/'+fulfillment['id']+'/shipment',data={'tracking_number':'BROWSER-TRACK-'+SUFFIX})
            assert response.status==200,response.text()
            detail=context.request.get(BASE+'/api/reconciliation-v2/order/'+ORDER).json()
        shipment=detail['shipments'][0]['id']
        act('bill');fill({'shipment_id':shipment,'provider_id':carrier,'currency':'PLN','outbound':'8','reverse':'0','collection_fee':'2','collected':detail['expected_collection'],'remitted':str(float(detail['expected_collection'])-10),'collected_at':'2026-09-21','remitted_at':'2026-09-21','reference':'BROWSER-'+SUFFIX,'proof':'本地模拟物流账单'});save()
        act('economics',ORDER);page.wait_for_selector('#rec-detail[open]')
        assert '结算信息已齐全' in page.locator('#rec-detail-content').inner_text()
        page.screenshot(path=str(OUT/'economics-final.png'),full_page=True);page.locator('#rec-detail-close').click()
        act('recognize',ORDER);page.wait_for_selector('#rec-detail[open]');page.locator('#rec-detail-close').click()
        tab('statements');act('statement');fill({'party_id':partner,'currency':'CNY','start':'2026-09-01','end':'2026-10-01','bucket':'trade'});save()
        state=context.request.get(BASE+'/api/reconciliation-v2/state').json();statement=next(s['id'] for s in state['statements'] if s['party_id']==partner and s['currency']=='CNY')
        for status in ('verified','confirmed','locked'):
            act('transition',statement+'|'+status)
            page.wait_for_function('(status)=>document.querySelector("#rec-view").innerText.includes(status)',arg={'verified':'已核对','confirmed':'已确认','locked':'已锁定'}[status])
        act('cash');fill({'party_id':partner,'currency':'CNY','amount':'20','occurred_at':'2026-09-21','reference':'BROWSER-BANK-'+SUFFIX,'proof':'本地模拟银行流水'});save()
        state=context.request.get(BASE+'/api/reconciliation-v2/state').json();cash=next(c['id'] for c in state['cash'] if c['reference']=='BROWSER-BANK-'+SUFFIX)
        act('allocate');fill({'cash_id':cash,'statement_id':statement,'amount':'20'});save()
        page.wait_for_function('document.querySelector("#rec-view").innerText.includes("已结清")')
        page.screenshot(path=str(OUT/'paid-statement-final.png'),full_page=True)
        with page.expect_download() as download:
            page.locator('a[href="/api/reconciliation-v2/export/'+statement+'"]').click()
        download.value.save_as(str(OUT/'browser-statement.csv'))
        exported=(OUT/'browser-statement.csv').read_text(encoding='utf-8-sig')
        assert '20.00' in exported and 'unit_cost' not in exported
        page.set_viewport_size({'width':390,'height':844});page.screenshot(path=str(OUT/'mobile-final.png'),full_page=True)
        assert page.evaluate('document.documentElement.scrollWidth<=window.innerWidth+1')
        context.request.get(BASE+'/logout');page.goto(BASE+'/logout');login('partner_demo')
        tab('statements');page.wait_for_function('document.querySelector("#rec-view").innerText.includes("已结清")')
        page.screenshot(path=str(OUT/'partner-view-final.png'),full_page=True)
        scoped=context.request.get(BASE+'/api/reconciliation-v2/state').json()
        assert len(scoped['parties'])==1 and not scoped['can_edit'] and 'pools' not in scoped
        assert context.request.get(BASE+'/api/reconciliation-v2/order/'+ORDER).status==403
        assert not errors and not failures, (errors,failures)
        report={'result':'passed','browser_errors':errors,'server_errors':failures,'statement_status':'paid','fee':'20.00 CNY','partner_scope':'own party only','export':'verified','mobile':'390px no body overflow'}
        (OUT/'browser-acceptance.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        print(json.dumps(report,ensure_ascii=True))
        browser.close()

if __name__=='__main__':run()
