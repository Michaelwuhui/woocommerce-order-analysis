"""Owned stock integrated with OMS reservations and shipment movements."""
import json
from datetime import date
from decimal import Decimal
from reconciliation_core import (ReconciliationError, audit, dec, done, dump, integer,
    lock, now, object_, order_config, ready, rows, uid, convert)
from inv_common import record_movement


def lock_stock(conn, warehouse_id, sku_id):
    conn.execute('INSERT INTO inv_stock(warehouse_id,sku_id,on_hand,reserved) VALUES (?,?,0,0) ON CONFLICT(warehouse_id,sku_id) DO NOTHING', (warehouse_id,sku_id))
    return conn.execute('SELECT on_hand,reserved FROM inv_stock WHERE warehouse_id=? AND sku_id=?'+lock(conn),(warehouse_id,sku_id)).fetchone()


def bind_batch(conn, batch_id, pool_id, unit_cost, currency, actor):
    from reconciliation_core import currency as validate_currency
    object_(conn, pool_id, 'pool')
    batch = conn.execute('SELECT * FROM inv_batches WHERE id=?', (batch_id,)).fetchone()
    if not batch:
        raise ReconciliationError('实物批次不存在，请通过库存采购/入库流程建立')
    if conn.execute('SELECT batch_id FROM rec_batches WHERE batch_id=?', (batch_id,)).fetchone():
        raise ReconciliationError('批次货权已登记；转让必须通过货权转移流水')
    stock = conn.execute('SELECT * FROM inv_stock WHERE warehouse_id=? AND sku_id=?' + lock(conn),
                         (batch['warehouse_id'], batch['sku_id'])).fetchone()
    if not stock or stock['reserved']:
        raise ReconciliationError('批次登记要求库存已核对且没有旧预留')
    total = conn.execute('SELECT SUM(qty_remaining) AS n FROM inv_batches WHERE warehouse_id=? AND sku_id=?',
                         (batch['warehouse_id'], batch['sku_id'])).fetchone()['n']
    if total != stock['on_hand']:
        raise ReconciliationError('批次数量与实物库存不一致，请先完成库存盘点')
    if unit_cost not in (None, '') and dec(unit_cost) < 0:
        raise ReconciliationError('批次成本不能为负')
    conn.execute('INSERT INTO rec_batches(batch_id,pool_id,unit_cost,currency) VALUES (?,?,?,?)',
                 (batch_id, pool_id, None if unit_cost in (None, '') else str(dec(unit_cost)), validate_currency(currency)))
    audit(conn, uid(), 'bind_batch', None, {'batch_id': batch_id, 'pool_id': pool_id}, actor)


def available_batches(conn, order_id, warehouse_id, sku_id):
    cfg = order_config(conn, order_id)
    if not cfg:
        return []
    batches = rows(conn, '''SELECT b.*,r.pool_id,r.unit_cost AS owned_cost,r.currency AS owned_currency,r.reserved
        FROM inv_batches b JOIN rec_batches r ON r.batch_id=b.id
        WHERE b.warehouse_id=? AND b.sku_id=? ORDER BY b.id''', (warehouse_id, sku_id))
    today = date.today().isoformat()
    for b in batches:
        old = conn.execute('SELECT COALESCE(SUM(reserved),0) AS n FROM rec_sources WHERE order_id=? AND batch_id=?',
                           (order_id, b['id'])).fetchone()['n']
        b['available'] = max(0, int(b['qty_remaining']) - int(b['reserved']) + int(old))
    batches = [b for b in batches if b['available'] > 0 and b['pool_id'] in cfg['contracts']
               and (not b['expiry_date'] or str(b['expiry_date'])[:10] >= today)]
    if cfg['policy'] == 'manual':
        batches = [b for b in batches if b['id'] in cfg['preferred_batches']]
    def rank(b):
        base = (str(b['expiry_date'] or '9999-12-31'), b['id'])
        pool = object_(conn, b['pool_id'], 'pool')['data']
        if cfg['policy'] == 'own_first':
            return (-dec(pool['ownership'].get(cfg['our_party_id'], 0)), *base)
        if cfg['policy'] == 'manual':
            return (cfg['preferred_batches'].index(b['id']), *base)
        if cfg['policy'] == 'quota':
            shipped = rows(conn, 'SELECT snapshot,shipped FROM rec_sources WHERE batch_id IN (SELECT batch_id FROM rec_batches WHERE pool_id=?)', (b['pool_id'],))
            allocation = sum(int(s['shipped']) for s in shipped)
            # Pool ownership converts party quotas into a target share without duplicating stock.
            quota = sum(dec(cfg['quota'].get(p, 0)) * dec(w) for p, w in pool['ownership'].items())
            return (Decimal(allocation) / quota if quota else Decimal('Infinity'), *base)
        return base
    return sorted(batches, key=rank)


def candidates(conn, order_id, sku_id, values, required_qty):
    cfg = order_config(conn, order_id)
    if not cfg:
        if not ready(conn):
            return values
        # Owned inventory must never silently fall back to the legacy unlimited/warehouse ledger.
        return [v for v in values if not conn.execute(
            'SELECT r.batch_id FROM rec_batches r JOIN inv_batches b ON b.id=r.batch_id WHERE b.warehouse_id=? AND b.sku_id=? LIMIT 1',
            (v['warehouse_id'],sku_id)).fetchone()]
    result = []
    for candidate in values:
        candidate = dict(candidate)
        batches = available_batches(conn, order_id, candidate['warehouse_id'], sku_id)
        # Partner-reported unlimited stock cannot stand in for owned physical stock.
        stock = conn.execute('SELECT on_hand,reserved FROM inv_stock WHERE warehouse_id=? AND sku_id=?',
                             (candidate['warehouse_id'], sku_id)).fetchone()
        reusable = conn.execute('''SELECT COALESCE(SUM(s.reserved),0) AS n FROM rec_sources s
            JOIN inv_batches b ON b.id=s.batch_id WHERE s.order_id=? AND b.warehouse_id=? AND b.sku_id=?''',
                               (order_id, candidate['warehouse_id'], sku_id)).fetchone()['n']
        physical = max(0, int(stock['on_hand']) - int(stock['reserved']) + int(reusable)) if stock else 0
        candidate['available'] = min(physical, sum(b['available'] for b in batches))
        candidate['rec_batches'] = [{'id': b['id'], 'pool_id': b['pool_id'], 'available': b['available']} for b in batches]
        candidate['rec_config'] = cfg
        if not candidate['available']:
            continue
        if cfg['policy'] == 'cost':
            quote = cfg.get('quotes', {}).get(str(candidate['warehouse_id']))
            if not quote or any(b['owned_cost'] is None for b in batches):
                continue
            # Full quote is explicitly provided in order currency, with all four components.
            try:
                for batch in batches:
                    convert(batch['owned_cost'], batch['owned_currency'], cfg['currency'], cfg.get('fx', {}))
                components = [dec(quote[k]) for k in ('goods', 'service', 'outbound', 'collection')]
                if any(v < 0 for v in components):
                    raise ReconciliationError('报价不能为负')
                candidate['rec_rank'] = sum(components)
            except (KeyError, ReconciliationError):
                continue
        elif cfg['policy'] == 'fewest':
            candidate['rec_rank'] = -min(candidate['available'], required_qty)
        elif cfg['policy'] == 'own_first':
            candidate['rec_rank'] = -max(dec(object_(conn,b['pool_id'],'pool')['data']['ownership'].get(cfg['our_party_id'],0)) for b in batches)
        elif cfg['policy'] == 'fefo':
            candidate['rec_rank'] = min(str(b['expiry_date'] or '9999-12-31') for b in batches)
        elif cfg['policy'] == 'manual':
            candidate['rec_rank'] = min(cfg['preferred_batches'].index(b['id']) for b in batches)
        elif cfg['policy'] == 'quota':
            pool = object_(conn,batches[0]['pool_id'],'pool')['data']
            quota = sum(dec(cfg['quota'].get(p,0))*dec(w) for p,w in pool['ownership'].items())
            shipped = conn.execute('SELECT COALESCE(SUM(s.shipped),0) AS n FROM rec_sources s JOIN rec_batches b ON b.batch_id=s.batch_id WHERE b.pool_id=?',(batches[0]['pool_id'],)).fetchone()['n']
            candidate['rec_rank'] = Decimal(shipped)/quota if quota else Decimal('Infinity')
        else:
            candidate['rec_rank'] = 0
        result.append(candidate)
    return sorted(result, key=lambda c: c['rec_rank'])


def reserve(conn, order_id, item_id, warehouse_id, sku_id, quantity):
    cfg = order_config(conn, order_id)
    if not cfg:
        return False
    stock = conn.execute('SELECT on_hand,reserved FROM inv_stock WHERE warehouse_id=? AND sku_id=?' + lock(conn),
                         (warehouse_id, sku_id)).fetchone()
    if not stock or int(stock['on_hand']) - int(stock['reserved']) < quantity:
        raise ReconciliationError('实物库存已变化，请重新分仓')
    remaining = quantity
    for b in available_batches(conn, order_id, warehouse_id, sku_id):
        current = conn.execute('SELECT b.qty_remaining,r.reserved FROM inv_batches b JOIN rec_batches r ON r.batch_id=b.id WHERE b.id=?' + lock(conn), (b['id'],)).fetchone()
        take = min(remaining, int(current['qty_remaining']) - int(current['reserved']))
        if take <= 0:
            continue
        contract = object_(conn, cfg['contracts'][b['pool_id']], 'contract')
        if not contract['data']['effective_from'] <= now()[:10] < contract['data']['effective_to']:
            raise ReconciliationError('来源合同不在有效期')
        snapshot = {'pool': object_(conn, b['pool_id'], 'pool'), 'contract': contract,
                    'unit_cost': b['owned_cost'], 'currency': b['owned_currency'], 'config': cfg,
                    'reserved_at': now(), 'warehouse_id': warehouse_id, 'sku_id': sku_id}
        source_id = uid()
        conn.execute('INSERT INTO rec_sources(id,fulfillment_item_id,order_id,batch_id,quantity,reserved,snapshot) VALUES (?,?,?,?,?,?,?)',
                     (source_id, item_id, order_id, b['id'], take, take, dump(snapshot)))
        conn.execute('UPDATE rec_batches SET reserved=reserved+? WHERE batch_id=?', (take, b['id']))
        remaining -= take
        if not remaining:
            break
    if remaining:
        raise ReconciliationError('货权批次可用量不足，请重新分仓')
    record_movement(conn, warehouse_id=warehouse_id, sku_id=sku_id, movement_type='reserve',
                    reserved_delta=quantity, ref_type='rec_source', ref_id=str(item_id), order_id=order_id,
                    note='模块化对账：按货权批次预留')
    return True


def release(conn, fulfillment):
    if not order_config(conn, fulfillment['order_id']):
        return False
    for item in rows(conn,'SELECT sku_id FROM oms_fulfillment_items WHERE fulfillment_id=? ORDER BY sku_id',(fulfillment['id'],)):
        lock_stock(conn,fulfillment['warehouse_id'],item['sku_id'])
    sources = rows(conn, '''SELECT s.* FROM rec_sources s JOIN oms_fulfillment_items i ON i.id=s.fulfillment_item_id
        WHERE i.fulfillment_id=? AND s.reserved>0 ORDER BY s.batch_id,s.id''' + lock(conn), (fulfillment['id'],))
    for s in sources:
        snap = json.loads(s['snapshot'])
        qty = s['reserved']
        conn.execute('UPDATE rec_batches SET reserved=reserved-? WHERE batch_id=?', (qty, s['batch_id']))
        conn.execute('UPDATE rec_sources SET reserved=0 WHERE id=?', (s['id'],))
        record_movement(conn, warehouse_id=snap['warehouse_id'], sku_id=snap['sku_id'], movement_type='release',
                        reserved_delta=-qty, batch_id=s['batch_id'], ref_type='rec_source', ref_id=s['id'] + ':release',
                        order_id=s['order_id'], note='取消/重分仓释放货权预留')
    return True


def ship(conn, fulfillment, item, shipment_id, quantity):
    if not order_config(conn, fulfillment['order_id']):
        return False
    conn.execute('SELECT id FROM orders WHERE id=?' + lock(conn), (fulfillment["order_id"],)).fetchone()
    lock_stock(conn,fulfillment["warehouse_id"],item["sku_id"])
    key = f"stock:{item['id']}:{shipment_id}"
    if done(conn, key):
        return True
    remaining = quantity
    sources = rows(conn, 'SELECT * FROM rec_sources WHERE fulfillment_item_id=? ORDER BY batch_id,id' + lock(conn), (item['id'],))
    for s in sources:
        qty = min(remaining, int(s['reserved']))
        if qty <= 0:
            continue
        snap = json.loads(s['snapshot'])
        if dump(snap['config']) != dump(order_config(conn, fulfillment['order_id'])):
            raise ReconciliationError('来源配置已改变，请重新分仓后再发货')
        current = conn.execute('SELECT qty_remaining FROM inv_batches WHERE id=?' + lock(conn), (s['batch_id'],)).fetchone()
        if current['qty_remaining'] < qty:
            raise ReconciliationError('批次库存异常，停止出库')
        conn.execute('UPDATE inv_batches SET qty_remaining=qty_remaining-? WHERE id=?', (qty, s['batch_id']))
        conn.execute('UPDATE rec_batches SET reserved=reserved-? WHERE batch_id=?', (qty, s['batch_id']))
        conn.execute('UPDATE rec_sources SET reserved=reserved-?,shipped=shipped+? WHERE id=?', (qty, qty, s['id']))
        record_movement(conn, warehouse_id=snap['warehouse_id'], sku_id=snap['sku_id'], movement_type='sale_out',
                        qty_delta=-qty, reserved_delta=-qty, batch_id=s['batch_id'], ref_type='rec_source',
                        ref_id=s['id'] + ':' + shipment_id, order_id=s['order_id'], note='货权批次实际出库')
        audit(conn, s['id'] + ':' + shipment_id, 'source_shipped', s['order_id'],
              {'source_id': s['id'], 'shipment_id': shipment_id, 'quantity': qty})
        remaining -= qty
    if remaining:
        raise ReconciliationError('出库数量超出已预留货权')
    audit(conn, key, 'stock_shipped', fulfillment['order_id'], {'shipment_id': shipment_id, 'quantity': quantity})
    return True


def return_source(conn, source_id, quantity, salable, reason, event_key, actor):
    if not isinstance(salable, bool):
        raise ReconciliationError('是否可重新销售必须明确选择')
    if not event_key or not reason:
        raise ReconciliationError('退货必须有唯一凭证和原因')
    key = 'return:' + source_id + ':' + event_key
    original = conn.execute('SELECT * FROM rec_sources WHERE id=?', (source_id,)).fetchone()
    if not original:
        raise ReconciliationError('来源不存在')
    snap = json.loads(original['snapshot'])
    conn.execute('SELECT id FROM orders WHERE id=?'+lock(conn),(original['order_id'],)).fetchone()
    lock_stock(conn,snap['warehouse_id'],snap['sku_id'])
    s = conn.execute('SELECT * FROM rec_sources WHERE id=?' + lock(conn), (source_id,)).fetchone()
    if not s:
        raise ReconciliationError('来源不存在')
    qty = integer(quantity, 1)
    if done(conn, key):
        previous = json.loads(conn.execute('SELECT data FROM rec_actions WHERE id=?',(key,)).fetchone()['data'])
        if previous['quantity'] != qty or previous['salable'] != salable:
            raise ReconciliationError('同一退货凭证不能更改数量或验收结果')
        return
    if qty > int(s['shipped']) - int(s['returned']):
        raise ReconciliationError('退货数量超出实际出库量')
    snap = json.loads(s['snapshot'])
    conn.execute('UPDATE rec_sources SET returned=returned+? WHERE id=?', (qty, source_id))
    restock_batch = s['batch_id']
    if salable:
        owner = conn.execute('SELECT pool_id FROM rec_batches WHERE batch_id=?', (s['batch_id'],)).fetchone()
        if owner['pool_id'] != snap['pool']['id']:
            original_batch = conn.execute('SELECT * FROM inv_batches WHERE id=?',(s['batch_id'],)).fetchone()
            cur = conn.execute('INSERT INTO inv_batches(warehouse_id,sku_id,batch_no,expiry_date,qty_received,qty_remaining) VALUES (?,?,?,?,?,0)',
                (snap['warehouse_id'],snap['sku_id'],str(original_batch['batch_no'])+' / return '+event_key,original_batch['expiry_date'],qty))
            restock_batch = int(cur.lastrowid)
            conn.execute('INSERT INTO rec_batches(batch_id,pool_id,unit_cost,currency) VALUES (?,?,?,?)',
                (restock_batch,snap['pool']['id'],snap['unit_cost'],snap['currency']))
        conn.execute('UPDATE inv_batches SET qty_remaining=qty_remaining+? WHERE id=?', (qty, restock_batch))
        record_movement(conn, warehouse_id=snap['warehouse_id'], sku_id=snap['sku_id'], movement_type='return_in',
                        qty_delta=qty, batch_id=restock_batch, ref_type='rec_return', ref_id=event_key,
                        order_id=s['order_id'], note=reason)
    audit(conn, key, 'source_returned', s['order_id'], {'source_id': source_id, 'quantity': qty,
          'salable': bool(salable), 'restock_batch_id': restock_batch if salable else None, 'reason': reason}, actor)


def transfer_ownership(conn, batch_id, new_pool_id, reason, actor):
    if not reason:
        raise ReconciliationError('货权转移必须提供合同/凭证原因')
    object_(conn, new_pool_id, 'pool')
    batch = conn.execute('SELECT * FROM inv_batches WHERE id=?',(batch_id,)).fetchone()
    if not batch:
        raise ReconciliationError('批次不存在')
    lock_stock(conn,batch['warehouse_id'],batch['sku_id'])
    r = conn.execute('SELECT * FROM rec_batches WHERE batch_id=?' + lock(conn), (batch_id,)).fetchone()
    if not r or r['reserved']:
        raise ReconciliationError('货权未登记或仍有预留，不能转移')
    conn.execute('UPDATE rec_batches SET pool_id=? WHERE batch_id=?', (new_pool_id, batch_id))
    audit(conn, uid(), 'ownership_transfer', None, {'batch_id': batch_id, 'from': r['pool_id'],
          'to': new_pool_id, 'reason': reason}, actor)

def register_unclassified_batch(conn, warehouse_id, sku_id, quantity, batch_no, pool_id, unit_cost, currency, actor, expiry=None):
    """Describe already received physical stock; this never increases stock."""
    qty=integer(quantity,1)
    stock=conn.execute('SELECT * FROM inv_stock WHERE warehouse_id=? AND sku_id=?'+lock(conn),(warehouse_id,sku_id)).fetchone()
    if not stock or stock['reserved']:
        raise ReconciliationError('先完成入库/盘点，且不能有旧预留')
    classified=conn.execute('SELECT COALESCE(SUM(qty_remaining),0) AS n FROM inv_batches WHERE warehouse_id=? AND sku_id=?',(warehouse_id,sku_id)).fetchone()['n']
    if qty > int(stock['on_hand'])-int(classified):
        raise ReconciliationError('登记数量超过已经入库但尚未分批的数量')
    if expiry:
        date.fromisoformat(expiry)
    if not batch_no:
        raise ReconciliationError('批次编号必填')
    cur=conn.execute('INSERT INTO inv_batches(warehouse_id,sku_id,batch_no,expiry_date,qty_received,qty_remaining,unit_cost,cost_currency) VALUES (?,?,?,?,?,?,?,?)',
        (warehouse_id,sku_id,batch_no,expiry or None,qty,qty,unit_cost or 0,currency))
    ident=int(cur.lastrowid)
    # Several receipts may be classified in separate portions; only this portion is registered.
    object_(conn,pool_id,'pool')
    from reconciliation_core import currency as validate_currency
    validate_currency(currency)
    if unit_cost not in (None,'') and dec(unit_cost)<0:
        raise ReconciliationError('成本不能为负')
    conn.execute('INSERT INTO rec_batches(batch_id,pool_id,unit_cost,currency) VALUES (?,?,?,?)',
        (ident,pool_id,None if unit_cost in (None,'') else str(dec(unit_cost)),currency))
    audit(conn,uid(),'classify_received_stock',None,{'batch_id':ident,'warehouse_id':warehouse_id,'sku_id':sku_id,'quantity':qty},actor)
    return ident


def move_batch(conn, batch_id, target_warehouse, quantity, reference, reason, actor):
    if not reference or not reason:
        raise ReconciliationError('调拨需要唯一收货凭证和原因')
    key='owned-transfer:'+reference
    if done(conn,key):
        return
    qty=integer(quantity,1)
    initial=conn.execute('SELECT warehouse_id,sku_id FROM inv_batches WHERE id=?',(batch_id,)).fetchone()
    if not initial:
        raise ReconciliationError('批次不存在')
    for wid in sorted({initial['warehouse_id'],target_warehouse}):
        lock_stock(conn,wid,initial['sku_id'])
    if done(conn,key):
        return
    b=conn.execute('SELECT b.*,r.pool_id,r.unit_cost AS owned_cost,r.currency AS owned_currency,r.reserved FROM inv_batches b JOIN rec_batches r ON r.batch_id=b.id WHERE b.id=?'+lock(conn),(batch_id,)).fetchone()
    if not b or qty > int(b['qty_remaining'])-int(b['reserved']) or b['warehouse_id']==target_warehouse:
        raise ReconciliationError('可调拨数量不足或目标仓库无效')
    warehouse=conn.execute('SELECT id FROM warehouses WHERE id=? AND is_active=1',(target_warehouse,)).fetchone()
    if not warehouse:
        raise ReconciliationError('目标仓库不可用')
    conn.execute('UPDATE inv_batches SET qty_remaining=qty_remaining-? WHERE id=?',(qty,batch_id))
    cur=conn.execute('INSERT INTO inv_batches(warehouse_id,sku_id,batch_no,production_date,expiry_date,qty_received,qty_remaining,unit_cost,cost_currency) VALUES (?,?,?,?,?,?,?,?,?)',
        (target_warehouse,b['sku_id'],str(b['batch_no'])+' / '+reference,b['production_date'],b['expiry_date'],qty,qty,b['unit_cost'],b['cost_currency']))
    new_batch=int(cur.lastrowid)
    conn.execute('INSERT INTO rec_batches(batch_id,pool_id,unit_cost,currency) VALUES (?,?,?,?)',(new_batch,b['pool_id'],b['owned_cost'],b['owned_currency']))
    for wid,delta,kind,bid in ((b['warehouse_id'],-qty,'transfer_out',batch_id),(target_warehouse,qty,'transfer_in',new_batch)):
        record_movement(conn,warehouse_id=wid,sku_id=b['sku_id'],movement_type=kind,qty_delta=delta,batch_id=bid,
                        ref_type='rec_transfer',ref_id=reference,note=reason,operator_id=int(actor))
    audit(conn,key,'physical_transfer_received',None,{'from_batch':batch_id,'to_batch':new_batch,'quantity':qty,'reason':reason},actor)
    return new_batch

def fewest_warehouse_order(conn, order_id, items, market):
    """Find the smallest feasible warehouse set, after all existing routing constraints."""
    cfg=order_config(conn,order_id)
    if not cfg or cfg['policy']!='fewest':
        return []
    from itertools import combinations
    from fulfillment_service import _candidate_warehouses, managed_product_family
    demand={}; managed={}
    for item in items:
        sku=item['sku_id']
        if not sku:
            continue
        demand[sku]=demand.get(sku,0)+max(0,int(item['ordered_qty'])-int(item.get('cancelled_qty') or 0))
        managed[sku]=managed.get(sku,False) or bool(managed_product_family(conn,json.loads(item.get('raw_json') or '{}')))
    matrix={}
    for sku,qty in demand.items():
        for c in candidates(conn,order_id,sku,_candidate_warehouses(conn,market,sku,managed_family=managed[sku]),qty):
            matrix.setdefault(c['warehouse_id'],{})[sku]=c['available']
    warehouses=sorted(matrix)
    # Operational warehouse counts are small; cap exhaustive search and retain a deterministic greedy fallback.
    if len(warehouses)<=14:
        for size in range(1,len(warehouses)+1):
            for subset in combinations(warehouses,size):
                if all(sum(matrix[w].get(sku,0) for w in subset)>=qty for sku,qty in demand.items()):
                    return sorted(subset,key=lambda w:-sum(min(qty,matrix[w].get(sku,0)) for sku,qty in demand.items()))
    return sorted(warehouses,key=lambda w:-sum(min(qty,matrix[w].get(sku,0)) for sku,qty in demand.items()))
