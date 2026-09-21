import os
import pytest
from test_reconciliation_postgres import pg, add_order
from test_reconciliation_v2 import setup
from reconciliation_core import save_object, rows
from reconciliation_history import save_rule, preview, create_draft
from fulfillment_service import plan_order, create_shipment


def test_history_postgres_uses_real_schema_and_does_not_replay_stock(pg):
    d,connect=pg;c=connect();add_order(d,c,'1-history',2)
    plan=plan_order(c,'1-history');shipment=create_shipment(c,plan['fulfillment_ids'][0],'HISTORY-TEST')
    c.execute('UPDATE oms_shipments SET shipped_at=? WHERE id=?',('2026-08-15 12:00:00',shipment['id']))
    c.execute("INSERT INTO shipping_logs(order_id,woo_order_id,source,tracking_number,shipped_at,status,is_reship) VALUES (?,?,?,?,?,?,?)",('1-history',123,'https://example.test','HISTORY-TEST','2026-08-15 12:00:00','shipped',False))
    c.execute("INSERT INTO exchange_rates(year_month,currency,rate_to_cny) VALUES (?,?,?)",('2026-08','PLN','1.81988'))
    c.execute("UPDATE orders SET total_tax=0 WHERE id='1-history'")
    pool=save_object(c,'pool','supplier',{'ownership':{d['partner']:'1'},'capital':{d['partner']:'1'},'ownership_method':'virtual'},1)
    contract=save_object(c,'contract','margin',{'pool_id':pool,'template':'margin','profit':{d['ours']:'1'},'loss':{d['ours']:'1'},'shipping':{d['ours']:'1'},'currency':'PLN','capital_status':'none','recognition':'shipped','effective_from':'2026-08-01','effective_to':'2099-01-01'},1)
    rule=save_rule(c,'history',{'warehouse_id':1,'pool_id':pool,'contract_id':contract,'our_party_id':d['ours'],'supplier_id':d['partner'],'provider_id':d['carrier'],'effective_from':'2026-08-01','effective_to':'2099-01-01','management_cny':'20','freight':{'PLN':{'outbound':'25','reverse':'0'}}},1)
    before=rows(c,'SELECT * FROM inv_stock ORDER BY warehouse_id,sku_id')
    snap=preview(c,rule,'2026-08','PLN')
    assert snap['counts']['pending_orders']==0,snap['rows']
    assert snap['totals']['revenue']=='85.00'
    assert snap['totals']['revenue_cny']=='154.69'
    ident=create_draft(c,{'rule_id':rule,'month':'2026-08','currency':'PLN','expected_digest':snap['digest'],'request_key':'postgres-history-0001'},1)
    c.commit()
    assert rows(c,'SELECT * FROM inv_stock ORDER BY warehouse_id,sku_id')==before
    assert c.execute('SELECT COUNT(*) FROM rec_entries').fetchone()[0]==0
    assert c.execute('SELECT kind FROM rec_objects WHERE id=?',(ident,)).fetchone()[0]=='history_draft'
    c.close()
