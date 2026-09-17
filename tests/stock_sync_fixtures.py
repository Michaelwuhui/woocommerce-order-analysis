"""Isolated synthetic business data; no production credentials or network."""
import copy
import json
import os
import sqlite3
import uuid
from pathlib import Path
from urllib.parse import urlsplit

import pytest
import requests
from flask import Flask
from flask_login import LoginManager, UserMixin
from jinja2 import ChoiceLoader, DictLoader, FileSystemLoader

import stock_sync_api as api
import stock_sync_guard as guard
from stock_sync_schema import migrate
from stock_sync_woo import Woo
from stock_sync_worker import run_once

FIXTURE_SQL='''
CREATE TABLE users(id INTEGER PRIMARY KEY,username TEXT,name TEXT,role TEXT,can_manage_products BOOLEAN,is_active BOOLEAN);
INSERT INTO users VALUES(1,'admin','Admin','admin',TRUE,TRUE),(2,'alice','Alice','admin',TRUE,TRUE),(3,'bob','Bob','user',TRUE,TRUE),(4,'viewer','Viewer','viewer',FALSE,TRUE);
CREATE TABLE sites(id INTEGER PRIMARY KEY,url TEXT,manager TEXT,country TEXT,consumer_key TEXT,consumer_secret TEXT,product_master_id INTEGER,is_active BOOLEAN);
INSERT INTO sites VALUES(1,'https://reference.test','Bob','PL','fixture-ck','fixture-cs',NULL,TRUE),(2,'https://target.test','Alice','PL','fixture-ck','fixture-cs',NULL,TRUE),(3,'https://outside.test','Bob','PL','fixture-ck','fixture-cs',NULL,TRUE);
CREATE TABLE product_masters(id INTEGER PRIMARY KEY,url TEXT);
CREATE TABLE warehouses(id INTEGER PRIMARY KEY,name TEXT,country TEXT,is_active BOOLEAN);
INSERT INTO warehouses VALUES(1,'Manual','PL',TRUE),(2,'Finite','PL',TRUE),(3,'External','PL',TRUE);
CREATE TABLE oms_warehouse_integrations(warehouse_id INTEGER PRIMARY KEY,inventory_authority TEXT,config_json TEXT);
INSERT INTO oms_warehouse_integrations VALUES(1,'manual_partner','{}'),(2,'local','{}'),(3,'external_wms','{}');
CREATE TABLE inv_market_warehouses(market_code TEXT,warehouse_id INTEGER,is_active BOOLEAN);
INSERT INTO inv_market_warehouses VALUES('PL',1,TRUE),('PL',2,TRUE),('PL',3,TRUE);
CREATE TABLE inv_skus(id INTEGER PRIMARY KEY,sku_code TEXT,name TEXT,is_active BOOLEAN);
CREATE TABLE inv_site_sku_map(id INTEGER PRIMARY KEY,site_id INTEGER,wc_product_id INTEGER,wc_variation_id INTEGER,sku_id INTEGER,qty_per_item INTEGER,is_active BOOLEAN,updated_at TEXT);
CREATE TABLE oms_sku_warehouses(sku_id INTEGER,warehouse_id INTEGER,is_enabled BOOLEAN);
CREATE TABLE inv_stock(warehouse_id INTEGER,sku_id INTEGER,on_hand INTEGER,reserved INTEGER);
CREATE TABLE oms_external_stock(warehouse_id INTEGER,sku_id INTEGER,available_quantity INTEGER,source_updated_at TEXT,synced_at TEXT);
CREATE TABLE inv_site_sync_config(site_id INTEGER PRIMARY KEY,mode TEXT,allocation_strategy TEXT,safety_stock INTEGER);
CREATE TABLE sync_runs(run_id TEXT,status TEXT,created_at TEXT);
CREATE TABLE sync_site_progress(run_id TEXT,site_id INTEGER,status TEXT);
CREATE TABLE inv_movements(id INTEGER PRIMARY KEY,sku_id INTEGER,qty_delta INTEGER);
'''


class Response:
    def __init__(self,data,code=200,headers=None):
        self.data,self.status_code,self.headers=data,code,headers or {}
    def json(self): return copy.deepcopy(self.data)


class FakeHTTP:
    def __init__(self):
        self.products={};self.calls=[];self.puts=[];self.fail_get=None
        self.fail_put=None;self.ignore_put=False;self.timeout_after_put=False
        self.bridge=False;self.on_put=None;self.on_get=None
    def request(self,method,url,**kw):
        self.calls.append((method,url,copy.deepcopy(kw.get('json'))))
        if self.on_get and method=='GET': self.on_get(url)
        if method=='GET' and self.fail_get and self.fail_get(url,kw):
            raise requests.Timeout('fixture timeout')
        parts=urlsplit(url);path=parts.path.split('/wp-json/wc/v3/')[1]
        if method=='GET' and path.endswith('products'):
            products=[v for k,v in self.products.items() if k.startswith(parts.scheme+'://'+parts.netloc+'/wp-json/wc/v3/products/') and '/variations/' not in k]
            products.sort(key=lambda p:p['id']);p=kw['params']['page'];size=kw['params']['per_page']
            return Response(products[(p-1)*size:p*size],headers={'X-WP-Total':str(len(products))})
        if method=='GET' and path.endswith('/variations'):
            products=[v for k,v in self.products.items() if k.startswith(url+'/')];p=kw['params']['page'];size=kw['params']['per_page']
            return Response(products[(p-1)*size:p*size],headers={'X-WP-Total':str(len(products))})
        if url not in self.products:return Response({},404)
        if method=='PUT':
            self.puts.append((url,copy.deepcopy(kw['json'])))
            failure=self.fail_put(url) if callable(self.fail_put) else self.fail_put
            if failure:return Response({},failure,{'Retry-After':'120'})
            if not self.ignore_put:
                self.products[url].update(copy.deepcopy(kw['json']))
                if self.bridge:
                    meta={m['key']:m['value'] for m in self.products[url].get('meta_data',[])}
                    if 'wcms_stock_manage' in meta:self.products[url]['manage_stock']=meta['wcms_stock_manage']=='yes'
                    if 'wcms_stock_qty' in meta:self.products[url]['stock_quantity']=meta['wcms_stock_qty']
                    if 'wcms_stock_status' in meta:self.products[url]['stock_status']=meta['wcms_stock_status']
            if self.on_put:self.on_put(url)
            if self.timeout_after_put:raise requests.Timeout('applied but client timed out')
        return Response(self.products[url])


@pytest.fixture
def system(tmp_path,monkeypatch):
    pg=os.getenv('STOCK_SYNC_TEST_POSTGRES')=='1'
    schema='stock_test_'+uuid.uuid4().hex
    path=tmp_path/'stock.db'
    if pg:
        import db_backend
        assert os.getenv('WOO_DB_NAME_OVERRIDE','').startswith('woo_stock_sync_test')
        c=db_backend.connect();c.execute('CREATE SCHEMA '+schema);c.commit();c.close()
    def connect():
        if pg:
            c=db_backend.connect();c._raw.execute('SET search_path TO '+schema);c.commit();return c
        c=sqlite3.connect(path,timeout=15);c.row_factory=sqlite3.Row;return c
    c=connect()
    for sql in FIXTURE_SQL.split(';'):
        if sql.strip():
            if pg:c._raw.execute(sql)
            else:c.execute(sql)
    migrate(c);c.close()
    monkeypatch.setenv('STOCK_SYNC_ENABLED','1')
    monkeypatch.setattr(api,'connect',connect);monkeypatch.setattr(guard,'connect',connect)
    monkeypatch.setenv('INV_DB_FILE',str(path));monkeypatch.setenv('WOO_SQLITE_PATH',str(path))
    http=FakeHTTP();woo=Woo(http)
    app=Flask(__name__);app.config.update(SECRET_KEY='stock-sync-test-only',TESTING=True)
    app.jinja_loader=ChoiceLoader([DictLoader({'base.html':'{% block content %}{% endblock %}'}),FileSystemLoader(str(Path(__file__).resolve().parents[1]/'templates'))])
    login=LoginManager(app)
    class User(UserMixin):
        def __init__(self,id_):self.id=id_
    @login.user_loader
    def load(id_):return User(id_)
    app.register_blueprint(api.bp)
    client=app.test_client()
    class System:
        def __init__(self):
            self.connect,self.http,self.woo,self.client,self.app=connect,http,woo,client,app
            self.worker_id='test-worker'
        def login(self,id_=1):
            with client.session_transaction() as s:s['_user_id']=str(id_);s['_fresh']=True
            self.token=client.get(api.PREFIX+'/options').get_json().get('csrf_token','')
        def call(self,path,method='GET',body=None):
            return client.open(api.PREFIX+path,method=method,json=body,headers={'X-CSRF-Token':self.token})
        def work(self):
            c=connect()
            try:return run_once(c,self.worker_id,woo)
            finally:c.close()
        def sql(self,sql,args=()):
            c=connect()
            try:
                r=c.execute(sql,args);result=[dict(x) for x in r.fetchall()] if r.description else []
                c.commit();return result
            finally:c.close()
        def seed(self,source=3,target=3):
            c=connect()
            for k in range(1,max(source,target)+1):
                c.execute('INSERT INTO inv_skus VALUES(?,?,?,TRUE)',(k,'SKU-'+str(k),'Style '+str(k)))
                c.execute('INSERT INTO oms_sku_warehouses VALUES(?,1,TRUE)',(k,))
            for sid,count in [(1,source),(2,target),(3,target)]:
                host={1:'reference',2:'target',3:'outside'}[sid]
                for k in range(1,count+1):
                    pid=sid*1000+k
                    c.execute('INSERT INTO inv_site_sku_map VALUES(?,?,?,0,?,1,TRUE,NULL)',(pid,sid,pid,k))
                    url=f'https://{host}.test/wp-json/wc/v3/products/{pid}'
                    http.products[url]={'id':pid,'parent_id':0,'name':'Style '+str(k),'type':'simple','status':'publish','sku':'SKU-'+str(k),'manage_stock':False,'stock_quantity':None,'stock_status':'instock' if sid==1 else 'outofstock','backorders':'no','regular_price':'10','description':'unchanged','images':[]}
            c.commit();c.close()
        def plan(self,operation='manual_hold',targets=None,selection=None,source_id=1,extra=None):
            ref=operation=='reference_status'
            r=self.call('/catalog-scans','POST',{'source_site_id':source_id} if ref else {'target_scope':{'mode':'all_authorized'}})
            assert r.status_code==202,r.get_json()
            scan=r.get_json()['id'];self.work()
            body={'operation':operation,'catalog_snapshot_id':scan,'selection':selection or {'mode':'all'},'target_scope':{'mode':'explicit_sites','site_ids':targets or [2]},'reason':'fixture operation'}
            if ref:body['source_site_id']=source_id
            body.update(extra or {})
            r=self.call('/plans','POST',body)
            assert r.status_code==202,r.get_json()
            id_=r.get_json()['id'];self.work()
            return self.call('/plans/'+id_).get_json()
        def execute(self,plan):
            r=self.call('/jobs','POST',{'plan_id':plan['id'],'plan_version':plan['version'],'idempotency_key':uuid.uuid4().hex,'accepted_item_ids':[i['id'] for i in plan['items'] if i['decision'] in ('change','unchanged')]})
            assert r.status_code==202,r.get_json()
            self.work();return self.call('/jobs/'+r.get_json()['id']).get_json()
    s=System();s.login();yield s
    if pg:
        import db_backend
        c=db_backend.connect();c.execute('DROP SCHEMA '+schema+' CASCADE');c.commit();c.close()
        db_backend.close_pools()
