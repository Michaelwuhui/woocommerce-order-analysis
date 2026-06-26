"""
inv_selftest.py — 进销存模块整体端到端自测(在数据库副本上跑,自清理)。

覆盖:迁移状态 → 供应商/SKU/映射 → 采购入库(批次+流水) → 多仓库存 →
订单自动分仓拆单(预留→出库 FEFO) → 退货回补 → 下推可用量计算 →
通知扫描 → 报表总览 → 不变式校验。全程用合成数据(TEST-*/合成订单),跑完删除。

用法: venv/bin/python inv_selftest.py
非零退出码表示有断言失败。
"""

import json
import sys
import datetime

import app as appmod
from inv_common import get_conn
import inv_batches as B
import inv_migrations

FAIL = []


def check(name, cond):
    print(('  ✓ ' if cond else '  ✗ ') + name)
    if not cond:
        FAIL.append(name)


def main():
    print('== 迁移状态 ==')
    conn = get_conn()
    applied = inv_migrations._applied(conn)
    check('全部迁移已应用', all(v in applied for v, *_ in inv_migrations.MIGRATIONS))

    hu = conn.execute("SELECT id FROM warehouses WHERE code='HU'").fetchone()['id']
    pl = conn.execute("SELECT id FROM warehouses WHERE country='PL' ORDER BY id LIMIT 1").fetchone()['id']

    client = appmod.app.test_client()
    with client.session_transaction() as s:
        s['_user_id'] = '1'; s['_fresh'] = True

    # 清理可能的历史残留
    _cleanup(conn)

    print('== 供应商 / SKU / 映射 ==')
    sup = client.post('/api/inv/suppliers', json={'name': 'SELFTEST供应商'}).get_json()['id']
    sA = client.post('/api/inv/skus', json={'sku_code': 'TEST-ST-A', 'name': '自测A', 'shelf_life_days': 365}).get_json()['id']
    sB = client.post('/api/inv/skus', json={'sku_code': 'TEST-ST-B', 'name': '自测B'}).get_json()['id']
    check('SKU 创建', bool(sA and sB))

    print('== 采购入库(HU) ==')
    exp = (datetime.date.today() + datetime.timedelta(days=200)).isoformat()
    po = client.post('/api/inv/purchase-orders', json={'warehouse_id': hu, 'supplier_id': sup, 'currency': 'CNY',
        'items': [{'sku_id': sA, 'qty': 6, 'unit_cost': 5, 'batch_no': 'STA', 'expiry_date': exp},
                  {'sku_id': sB, 'qty': 4, 'unit_cost': 3, 'batch_no': 'STB', 'expiry_date': exp}]}).get_json()
    rcv = client.post(f"/api/inv/purchase-orders/{po['id']}/receive").get_json()
    check('收货成功', rcv.get('success'))
    huA = conn.execute('SELECT on_hand FROM inv_stock WHERE warehouse_id=? AND sku_id=?', (hu, sA)).fetchone()['on_hand']
    check('HU on_hand(A)=6', huA == 6)
    check('不变式 Σ批次==on_hand', B.total_remaining(conn, hu, sA) == 6)

    # PL 仓也补点货,逼出拆单(需 A=6,B=4;HU 有 A6/B4 可整单 → 为测拆单,把 HU B 调低)
    client.post('/api/inv/adjust', json={'warehouse_id': hu, 'sku_id': sB, 'qty_delta': -3, 'reason': '自测压低'})  # HU B=1
    conn.execute("INSERT INTO inv_stock (warehouse_id,sku_id,on_hand,reserved) VALUES (?,?,50,0)", (pl, sB))
    conn.execute("INSERT INTO inv_batches (warehouse_id,sku_id,batch_no,unit_cost,cost_currency,qty_received,qty_remaining) VALUES (?,?,?,?,?,?,?)", (pl, sB, 'PLB', 3, 'CNY', 50, 50))
    conn.commit()

    print('== 合成 CZ 订单 → 自动分仓拆单 ==')
    TID = 'TEST-ST-ORDER'
    conn.execute("DELETE FROM orders WHERE id=?", (TID,))
    site = conn.execute("SELECT id,url FROM sites WHERE country='PL' LIMIT 1").fetchone()
    li = [{'product_id': 970001, 'variation_id': 0, 'quantity': 6, 'name': 'A', 'sku': 'X'},
          {'product_id': 970002, 'variation_id': 0, 'quantity': 4, 'name': 'B', 'sku': 'Y'}]
    conn.execute("INSERT INTO orders (id,source,status,line_items,shipping,warehouse_id,number) VALUES (?,?,?,?,?,?,?)",
                 (TID, site['url'], 'processing', json.dumps(li), json.dumps({'country': 'CZ'}), pl, 'ST1'))
    conn.execute("INSERT INTO inv_site_sku_map (site_id,wc_product_id,wc_variation_id,sku_id,qty_per_item) VALUES (?,?,0,?,1)", (site['id'], 970001, sA))
    conn.execute("INSERT INTO inv_site_sku_map (site_id,wc_product_id,wc_variation_id,sku_id,qty_per_item) VALUES (?,?,0,?,1)", (site['id'], 970002, sB))
    conn.commit()

    res = client.post('/api/inv/process-order/' + TID).get_json()
    check('预留成功', res.get('action') == 'reserved')
    # A 整单从 HU(6),B 拆:HU1 + PL3
    rsvA_hu = conn.execute('SELECT reserved FROM inv_stock WHERE warehouse_id=? AND sku_id=?', (hu, sA)).fetchone()['reserved']
    check('A 预留在 HU=6', rsvA_hu == 6)

    conn.execute("UPDATE orders SET status='completed' WHERE id=?", (TID,)); conn.commit()
    res = client.post('/api/inv/process-order/' + TID).get_json()
    check('出库成功', res.get('action') == 'shipped')
    check('出库为拆单', res.get('is_split') is True)
    fuls = client.get('/api/inv/fulfillments/' + TID).get_json()
    check('生成多个分单', len(fuls) >= 2)
    # 出库后 HU A on_hand=0
    check('HU A 出库后=0', conn.execute('SELECT on_hand FROM inv_stock WHERE warehouse_id=? AND sku_id=?', (hu, sA)).fetchone()['on_hand'] == 0)
    check('不变式(HU,A)', B.total_remaining(conn, hu, sA) == 0)
    check('不变式(PL,B)', conn.execute('SELECT on_hand FROM inv_stock WHERE warehouse_id=? AND sku_id=?', (pl, sB)).fetchone()['on_hand'] == B.total_remaining(conn, pl, sB))

    print('== 退货回补 ==')
    conn.execute("UPDATE orders SET status='refunded' WHERE id=?", (TID,)); conn.commit()
    res = client.post('/api/inv/process-order/' + TID).get_json()
    check('退货成功', res.get('action') == 'returned')
    check('HU A 退货后回补=6', conn.execute('SELECT on_hand FROM inv_stock WHERE warehouse_id=? AND sku_id=?', (hu, sA)).fetchone()['on_hand'] == 6)

    print('== 下推可用量计算 ==')
    ss = client.get(f'/api/inv/site-stock/{site["id"]}').get_json()
    check('可发布库存计算返回', isinstance(ss, list))

    print('== 通知扫描 ==')
    scan = client.post('/api/inv/notifications/scan').get_json()
    check('扫描返回计数', 'generated' in scan)

    print('== 报表总览 ==')
    ov = client.get('/api/inv/reports/overview').get_json()
    check('总览含货值', 'valuation' in ov)
    val = client.get('/api/inv/reports/valuation').get_json()
    check('货值报表返回', 'by_warehouse' in val)

    print('== 页面可达 ==')
    for url in ['/inventory/warehouses', '/inventory/skus', '/inventory/stock', '/inventory/batches',
                '/inventory/orders', '/inventory/push', '/inventory/notifications', '/inventory/reports']:
        check(f'GET {url}', client.get(url).status_code == 200)

    print('== 清理 ==')
    _cleanup(conn, extra_skus=[sA, sB], sup=sup, tid=TID)
    conn.close()

    print()
    if FAIL:
        print(f'❌ 自测失败 {len(FAIL)} 项: {FAIL}')
        return 1
    print('✅ 整体自测全部通过')
    return 0


def _cleanup(conn, extra_skus=None, sup=None, tid='TEST-ST-ORDER'):
    ids = [r['id'] for r in conn.execute("SELECT id FROM inv_skus WHERE sku_code LIKE 'TEST-ST-%'")]
    ids = list(set(ids + (extra_skus or [])))
    conn.execute("DELETE FROM inv_order_state WHERE order_id=?", (tid,))
    conn.execute("DELETE FROM inv_fulfillments WHERE order_id=?", (tid,))
    conn.execute("DELETE FROM orders WHERE id=?", (tid,))
    for sid in ids:
        for t in ('inv_movements', 'inv_batches', 'inv_stock', 'inv_site_sku_map',
                  'inv_fulfillment_items', 'inv_notifications', 'inv_purchase_order_items'):
            conn.execute(f'DELETE FROM {t} WHERE sku_id=?', (sid,))
        conn.execute("DELETE FROM inv_skus WHERE id=?", (sid,))
    conn.execute("DELETE FROM inv_purchase_orders WHERE po_no LIKE 'PO%' AND id NOT IN (SELECT DISTINCT po_id FROM inv_purchase_order_items)")
    if sup:
        conn.execute("DELETE FROM inv_suppliers WHERE id=?", (sup,))
    conn.execute("DELETE FROM inv_suppliers WHERE name='SELFTEST供应商'")
    conn.commit()


if __name__ == '__main__':
    sys.exit(main())
