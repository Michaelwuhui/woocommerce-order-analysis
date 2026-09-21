"""Loopback-only synthetic development environment; never points at production."""
import argparse
import json
import os
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


def environment():
    import secrets
    os.environ.setdefault('FLASK_SECRET_KEY', secrets.token_hex(32))
    os.environ.update(WOO_DB_BACKEND='postgres', WOO_DB_HOST='127.0.0.1', WOO_DB_PORT='55439',
        WOO_DB_NAME_OVERRIDE='woo_reconciliation_test_dev', WOO_DB_USER='recdev', WOO_DB_PASSWORD='local-test-only',
        WOO_DB_POOL_MAX='12', CELERY_BROKER_URL='memory://', WOO_CELERY_BROKER_URL='memory://')


def seed():
    environment()
    from inv_common import get_conn
    from reconciliation_schema import migrate
    from reconciliation_core import save_object, configure_order
    from reconciliation_inventory import bind_batch
    from werkzeug.security import generate_password_hash
    conn=get_conn(); migrate(conn)
    if conn.execute('SELECT id FROM users LIMIT 1').fetchone():
        print('Synthetic database already seeded; preserving state.'); conn.close();return
    for username in ('admin','partner_demo'):
        conn.execute('INSERT INTO users(username,password_hash,name,can_view_reconciliation,can_view_costs) VALUES (?,?,?,?,?)',
                     (username,generate_password_hash('local-reconciliation-demo'),username,True,username=='admin'))
    conn.execute("INSERT INTO sites(id,url,consumer_key,consumer_secret,country) VALUES (1,'https://rec-demo.test','disabled','disabled','PL')")
    conn.execute("INSERT INTO warehouses(id,code,name,country) VALUES (1,'REC-DEMO','对账演示仓库','PL')")
    conn.execute("INSERT INTO inv_skus(id,sku_code,name,is_active) VALUES (1,'REC-SKU','演示商品',1)")
    conn.execute("INSERT INTO inv_site_sku_map(site_id,wc_product_id,wc_variation_id,wc_sku,raw_name,sku_id,qty_per_item,is_active) VALUES (1,101,0,'REC-SKU','演示商品',1,1,1)")
    conn.execute('INSERT INTO oms_sku_warehouses(sku_id,warehouse_id,is_primary,is_enabled) VALUES (1,1,1,1)')
    conn.execute("INSERT INTO inv_market_warehouses(market_code,warehouse_id,priority,is_active) VALUES ('PL',1,1,1)")
    conn.execute('INSERT INTO inv_stock(warehouse_id,sku_id,on_hand,reserved) VALUES (1,1,20,0)')
    for ident in (1,2):
        conn.execute('INSERT INTO inv_batches(id,warehouse_id,sku_id,batch_no,qty_received,qty_remaining,unit_cost) VALUES (?,1,1,?,10,10,10)',(ident,'DEMO-'+str(ident)))
    ours=save_object(conn,'party','演示 · 本团队',{'internal':True,'currency':'PLN'},1)
    partner=save_object(conn,'party','演示 · 合作方',{'currency':'CNY'},1)
    carrier=save_object(conn,'party','演示 · 物流商',{'currency':'PLN'},1)
    contracts={}
    for n,capital in enumerate(({ours:'1'},{ours:'.6',partner:'.4'})):
        pool=save_object(conn,'pool','演示货权池 '+str(n+1),{'ownership':capital,'capital':capital,'ownership_method':'physical'},1)
        contract=save_object(conn,'contract','演示合同 '+str(n+1),{'pool_id':pool,'template':'joint','profit':{ours:'.5',partner:'.5'},'loss':{ours:'.5',partner:'.5'},'shipping':{ours:'1'},'currency':'PLN','capital_status':'unpaid','recognition':'shipped','effective_from':'2020-01-01','effective_to':'2099-01-01'},1)
        contracts[pool]=contract
        bind_batch(conn,n+1,pool,'10','PLN',1)
    fee=save_object(conn,'fee','演示 · 每订单 20 CNY',{'provider_id':partner,'amount':'20','currency':'CNY','basis':'order','event':'shipped','return_policy':'keep','reship_policy':'free','effective_from':'2020-01-01','effective_to':'2099-01-01'},1)
    for number in (1001,1002,1003):
        order_id='1-'+str(number)
        lines=[{'id':501,'product_id':101,'variation_id':0,'sku':'REC-SKU','name':'演示商品','quantity':12 if number==1001 else 2,'total':'360' if number==1001 else '60','total_tax':'0'}]
        total='385' if number==1001 else '85'
        address=json.dumps({'country':'PL','first_name':'Demo','last_name':'Only','address_1':'Test 1','city':'Demo','postcode':'00-001','phone':'000000000'})
        conn.execute('INSERT INTO orders(id,number,status,source,line_items,billing,shipping,total,shipping_total,shipping_tax,currency,payment_method,date_created) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)',
            (order_id,str(number),'processing','https://rec-demo.test',json.dumps(lines),address,address,total,'25','0','PLN','cod','2026-09-21T10:00:00'))
        configure_order(conn,order_id,{'our_party_id':ours,'currency':'PLN','policy':'fefo','contracts':contracts,'fees':[fee],'provider_mode':'primary','primary_provider':partner,'fx':{'CNY/PLN':'.5'}},1)
    conn.execute('INSERT INTO rec_user_parties(user_id,party_id) VALUES (2,?)',(partner,))
    conn.commit();conn.close();print('Synthetic local-only PostgreSQL seed completed.')


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('command',choices=['seed','serve','migrate']);args=parser.parse_args()
    environment()
    if args.command=='seed':seed()
    elif args.command=='migrate':
        from inv_common import get_conn
        from reconciliation_schema import migrate
        conn=get_conn();migrate(conn);conn.close()
    else:
        from app import app
        app.config.update(TESTING=False)
        app.run(host='127.0.0.1',port=5059,debug=False,use_reloader=False)
