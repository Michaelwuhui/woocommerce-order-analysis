"""Local HTTP end-to-end fixture with the real Flask app and a separate worker.

Requires an EMPTY production-schema clone named woo_stock_sync_test_browser on
127.0.0.1:55436. All remote catalog/stock HTTP is served on loopback port 8106.
No runtime configuration in the production application is changed by this file.
"""
from pathlib import Path
import json
import os
import subprocess
import sys
import threading

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
os.chdir(ROOT)
os.environ.update(WOO_DB_BACKEND='postgres',WOO_DB_HOST='127.0.0.1',WOO_DB_PORT='55436',
    WOO_DB_USER='stocktester',WOO_DB_PASSWORD='local-fixture',WOO_DB_NAME_OVERRIDE='woo_stock_sync_test_browser',
    WOO_DB_POOL_MAX='8',FLASK_SECRET_KEY='isolated-stock-sync-demo-only',STOCK_SYNC_ENABLED='1',
    WOO_SQLITE_PATH='woocommerce_orders.db',INV_DB_FILE='woocommerce_orders.db')

from flask import Flask,request,jsonify,redirect
from flask_login import login_user
from werkzeug.security import generate_password_hash
from werkzeug.serving import make_server
from stock_sync_common import connect,one
from stock_sync_schema import migrate

c=connect()
assert c.execute('SELECT current_database()').fetchone()[0]=='woo_stock_sync_test_browser'
migrate(c)
if not c.execute('SELECT 1 FROM users LIMIT 1').fetchone():
    for id_,username,name in [(1,'admin','演示管理员'),(2,'demo-owner','演示负责人')]:
        c.execute('INSERT INTO users(id,username,password_hash,name,role,can_manage_products) VALUES(?,?,?,?,?,?)',
            (id_,username,generate_password_hash('local-demo-only'),'演示管理员' if id_==1 else '演示负责人','admin' if id_==1 else 'user',True))
    for id_,name,authority in [(1,'人工供货仓','manual_partner'),(2,'数量仓','local')]:
        c.execute('INSERT INTO warehouses(id,name,code,country,is_active) VALUES(?,?,?,?,?)',(id_,name,'DEMO-'+str(id_),'PL',True))
        c.execute('INSERT INTO oms_warehouse_integrations(warehouse_id,inventory_authority,config_json) VALUES(?,?,?)',(id_,authority,'{}'))
        c.execute('INSERT INTO inv_market_warehouses(market_code,warehouse_id,is_active) VALUES(?,?,?)',('PL',id_,True))
    for sid,name in [(1,'参照站'),(2,'目标站 A'),(3,'目标站 B')]:
        c.execute('INSERT INTO sites(id,url,country,manager,consumer_key,consumer_secret) VALUES(?,?,?,?,?,?)',
            (sid,f'http://127.0.0.1:8106/site{sid}','PL','演示管理员' if sid==1 else '演示负责人','demo-key','demo-secret'))
    for k in range(1,21):
        quantity=k<=2
        c.execute('INSERT INTO inv_skus(id,sku_code,name,is_active) VALUES(?,?,?,?)',(k,'DEMO-'+str(k),['薄荷 / 独立口味','蓝莓 / 数量换算'][k-1] if k<=2 else '演示款式 '+str(k),True))
        c.execute('INSERT INTO oms_sku_warehouses(sku_id,warehouse_id,is_enabled) VALUES(?,?,?)',(k,2 if quantity else 1,True))
        if quantity:c.execute('INSERT INTO inv_stock(warehouse_id,sku_id,on_hand,reserved) VALUES(?,?,?,?)',(2,k,25 if k==2 else 20,0 if k==2 else 7))
        for sid in [1,2,3]:
            if sid==3 and k>3:continue
            c.execute('INSERT INTO inv_site_sku_map(id,site_id,wc_product_id,wc_variation_id,sku_id,qty_per_item,is_active) VALUES(?,?,?,?,?,?,?)',
                (sid*1000+k,sid,sid*1000+k,0,k,10 if k==2 else 1,True))
    c.commit()
c.close()

web=Flask('stock-sync-loopback-woo')
state_file=ROOT/'.test-stock-sync/woo-demo-state.json'
if state_file.exists():
    states=json.loads(state_file.read_text(encoding='utf-8'))
else:
    states={}
    for sid in [1,2,3]:
        states[str(sid)]={}
        for k in range(1,4 if sid==3 else 21):
            pid=sid*1000+k
            states[str(sid)][str(pid)]={'id':pid,'name':('蓝莓 / 数量换算' if k==2 else '演示款式 '+str(k)),
                'type':'simple','status':'publish','parent_id':0,'sku':'DEMO-'+str(k),
                'manage_stock':k<=2,'stock_quantity':1 if k<=2 else None,
                'stock_status':'instock' if sid==1 else 'outofstock','backorders':'no',
                'regular_price':'10.00','description':'本机测试商品','images':[]}

@web.route('/site<int:sid>/wp-json/wc/v3/products')
def catalog(sid):
    data=list(states[str(sid)].values());page=int(request.args.get('page',1));n=int(request.args.get('per_page',100))
    return jsonify(data[(page-1)*n:page*n]),200,{'X-WP-Total':str(len(data))}

@web.route('/site<int:sid>/wp-json/wc/v3/products/<int:pid>',methods=['GET','PUT'])
def product(sid,pid):
    p=states.get(str(sid),{}).get(str(pid))
    if p is None:return jsonify(error='not found'),404
    if request.method=='PUT':
        data=request.get_json()
        if set(data)-{'manage_stock','stock_status','stock_quantity','backorders','meta_data'}:return jsonify(error='non-stock field'),400
        p.update(data);state_file.write_text(json.dumps(states,ensure_ascii=False,indent=2),encoding='utf-8')
    return jsonify(p)

@web.route('/site<int:sid>/product/<int:pid>')
def public_product(sid,pid):
    p=states[str(sid)][str(pid)]
    status='有货' if p['stock_status']=='instock' else '售完'
    return f'<html lang="zh"><title>本机测试商品</title><h1>{p["name"]}</h1><p>{status}</p><p>库存：{p["stock_quantity"]}</p><p>价格：{p["regular_price"]}</p></html>'

from app import app,User
@app.route('/__stock_demo_login/<int:user_id>')
def demo_login(user_id):
    if request.remote_addr!='127.0.0.1' or user_id not in (1,2):return 'Forbidden',403
    c=connect();u=one(c,'SELECT * FROM users WHERE id=?',(user_id,));c.close()
    login_user(User(u['id'],u['username'],u['name'],u['role']))
    return redirect('/product-manager/stock-sync')

if __name__=='__main__':
    server=make_server('127.0.0.1',8106,web,threaded=True)
    threading.Thread(target=server.serve_forever,daemon=True).start()
    log=(ROOT/'.test-stock-sync/worker.log').open('a',encoding='utf-8')
    worker=subprocess.Popen([sys.executable,'-u','stock_sync_worker.py'],cwd=ROOT,env=os.environ,stdout=log,stderr=log,
        creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
    print('Demo app http://127.0.0.1:5106/__stock_demo_login/1 ; worker PID '+str(worker.pid),flush=True)
    try:app.run(host='127.0.0.1',port=5106,debug=False,use_reloader=False)
    finally:worker.terminate();worker.wait(10);server.shutdown();log.close()
