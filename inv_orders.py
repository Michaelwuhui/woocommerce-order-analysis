"""
inv_orders.py — 模块 5:卖出扣减 + 订单联动(幂等处理器)。

约束:不能改拉单核心 sync_utils.py。因此库存联动做成**幂等处理器**:
给定订单当前状态,把库存对齐到应有效果,可手动 / 批量 / 独立 cron 触发,
重复执行结果一致(不会重复扣减)。

状态机(订单状态 → 目标库存效果):
  RESERVE(已付待发) : processing / on-hold        → 预留 reserved
  SHIP(已出库)      : shipped / completed /
                       delivered / partial-shipped → 出库 deduct(扣 on_hand + FEFO 批次)
  CLEAR(无效/退回)  : cancelled / failed /
                       refunded / checkout-draft /
                       cheat / offline             → 释放预留 或 退货入库(视当前状态)

每张订单在 inv_order_state 记录当前 inv_state 与 committed_json(已提交的 SKU/批次),
转换时按 committed_json 精确冲销旧效果、施加新效果,保证幂等与可逆。

选仓(模块5 单仓版;模块6 会升级为多仓自动拆单):
  优先用 orders.warehouse_id;否则取该订单市场(收货国)首选的、参与分仓的活跃仓。

不变式:Σ inv_batches.qty_remaining == inv_stock.on_hand(同仓同SKU)始终成立。
"""

import json
from flask import Blueprint, render_template, request, jsonify
from flask_login import login_required

from inv_common import (
    get_conn, inv_view_required, inv_manage_required,
    record_movement, current_operator,
)
import inv_resolver
import inv_batches

inv_ord_bp = Blueprint('inv_ord', __name__)

# 订单状态 → 目标效果
RESERVE_STATUSES = {'processing', 'on-hold'}
SHIP_STATUSES = {'shipped', 'completed', 'delivered', 'partial-shipped'}
CLEAR_STATUSES = {'cancelled', 'failed', 'refunded', 'checkout-draft', 'cheat', 'offline'}


def target_effect(status):
    """订单状态 → 目标库存效果: 'reserved' | 'shipped' | 'clear' | None(忽略)。"""
    if status in SHIP_STATUSES:
        return 'shipped'
    if status in RESERVE_STATUSES:
        return 'reserved'
    if status in CLEAR_STATUSES:
        return 'clear'
    return None


# ───────────────────────── 市场 / 选仓 ─────────────────────────

def order_market(conn, order):
    """订单市场(收货国):shipping.country → billing.country → 站点国家。"""
    for field in ('shipping', 'billing'):
        raw = order[field] if field in order.keys() else None
        if raw:
            try:
                c = (json.loads(raw) or {}).get('country')
                if c:
                    return c.upper()
            except Exception:
                pass
    s = conn.execute('SELECT country FROM sites WHERE url=?', (order['source'],)).fetchone()
    return (s['country'].upper() if s and s['country'] else None)


def pick_warehouse(conn, order, market):
    """选仓(单仓版)。返回 warehouse_id 或 None。

    1) orders.warehouse_id 若已指派,直接用;
    2) 否则按市场路由取首选(priority 最小)的、参与分仓的活跃仓;
    3) 再否则取该市场国家下任一活跃仓。
    """
    wid = order['warehouse_id'] if 'warehouse_id' in order.keys() else None
    if wid:
        return wid
    if market:
        r = conn.execute('''SELECT mw.warehouse_id FROM inv_market_warehouses mw
                            JOIN warehouses w ON w.id=mw.warehouse_id
                            LEFT JOIN inv_warehouse_ext we ON we.warehouse_id=w.id
                            WHERE mw.market_code=? AND mw.is_active=1 AND w.is_active=1
                              AND COALESCE(we.is_fulfillment,1)=1
                            ORDER BY mw.priority, mw.id LIMIT 1''', (market,)).fetchone()
        if r:
            return r['warehouse_id']
        r = conn.execute("SELECT id FROM warehouses WHERE country=? AND is_active=1 ORDER BY id LIMIT 1",
                         (market,)).fetchone()
        if r:
            return r['id']
    return None


# ───────────────────────── 效果施加/冲销 ─────────────────────────

def _aggregate_sku_lines(resolved):
    """把解析结果按 sku_id 聚合 sku 总数(同 SKU 多行合并)。"""
    agg = {}
    for ln in resolved['lines']:
        if ln['sku_id']:
            agg[ln['sku_id']] = agg.get(ln['sku_id'], 0) + ln['sku_qty']
    return agg


def _unreserve(conn, st, uid, uname):
    """冲销已有预留(committed mode=reserved)。"""
    c = json.loads(st['committed_json'] or '{}')
    wid = c.get('warehouse_id')
    for ln in c.get('lines', []):
        record_movement(conn, warehouse_id=wid, sku_id=ln['sku_id'], movement_type='release',
                        reserved_delta=-ln['qty'], ref_type='order', ref_id=st['order_id'],
                        order_id=st['order_id'], operator_id=uid, operator_name=uname,
                        note='释放预留(状态变更)')


def _unship(conn, st, uid, uname):
    """冲销已有出库(committed mode=shipped):退货入库 + 批次回补。"""
    c = json.loads(st['committed_json'] or '{}')
    wid = c.get('warehouse_id')
    # 批次回补(按当时分配)
    for b in c.get('batches', []):
        inv_batches.restock_batch(conn, b['batch_id'], b['qty'])
    for ln in c.get('lines', []):
        record_movement(conn, warehouse_id=wid, sku_id=ln['sku_id'], movement_type='return_in',
                        qty_delta=ln['qty'], ref_type='order', ref_id=st['order_id'],
                        order_id=st['order_id'], operator_id=uid, operator_name=uname,
                        note='退货入库(状态变更)')


def _apply_reserve(conn, order_id, wid, sku_agg, uid, uname):
    for sku_id, qty in sku_agg.items():
        record_movement(conn, warehouse_id=wid, sku_id=sku_id, movement_type='reserve',
                        reserved_delta=qty, ref_type='order', ref_id=order_id,
                        order_id=order_id, operator_id=uid, operator_name=uname, note='订单预留')
    return {'mode': 'reserved', 'warehouse_id': wid,
            'lines': [{'sku_id': k, 'qty': v} for k, v in sku_agg.items()], 'batches': []}


def _apply_ship(conn, order_id, wid, sku_agg, uid, uname, from_reserved):
    """出库:扣 on_hand + FEFO 批次。from_reserved=True 时同时释放对应预留。"""
    batches = []
    for sku_id, qty in sku_agg.items():
        allocs, shortage = inv_batches.plan_fefo(conn, wid, sku_id, qty)
        # 即使批次不足(shortage>0),也按 on_hand 扣减(负库存由调用前校验/允许超卖策略决定);
        # 这里记录实际可分配批次,剩余无批次部分不挂批次。
        inv_batches.consume_batches(conn, allocs)
        for a in allocs:
            batches.append({'batch_id': a['batch_id'], 'qty': a['qty']})
        reserved_delta = -qty if from_reserved else 0
        record_movement(conn, warehouse_id=wid, sku_id=sku_id, movement_type='sale_out',
                        qty_delta=-qty, reserved_delta=reserved_delta,
                        ref_type='order', ref_id=order_id, order_id=order_id,
                        operator_id=uid, operator_name=uname,
                        note='销售出库' + ('(原预留转出)' if from_reserved else ''))
    return {'mode': 'shipped', 'warehouse_id': wid,
            'lines': [{'sku_id': k, 'qty': v} for k, v in sku_agg.items()], 'batches': batches}


def _save_state(conn, order_id, source, inv_state, wid, committed, status, market, note=None):
    conn.execute('''INSERT INTO inv_order_state
        (order_id, source, inv_state, warehouse_id, committed_json, last_order_status, market_code, note, processed_at, updated_at)
        VALUES (?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)
        ON CONFLICT(order_id) DO UPDATE SET
            source=excluded.source, inv_state=excluded.inv_state, warehouse_id=excluded.warehouse_id,
            committed_json=excluded.committed_json, last_order_status=excluded.last_order_status,
            market_code=excluded.market_code, note=excluded.note,
            processed_at=CURRENT_TIMESTAMP, updated_at=CURRENT_TIMESTAMP''',
        (order_id, source, inv_state, wid, json.dumps(committed) if committed else None,
         status, market, note))


# ───────────────────────── 幂等处理器 ─────────────────────────

def process_order(conn, order_id, caches=None, commit=True):
    """把单张订单的库存对齐到其当前状态。幂等。返回处理摘要 dict。"""
    uid, uname = current_operator()
    order = conn.execute('SELECT * FROM orders WHERE id=?', (order_id,)).fetchone()
    if not order:
        return {'order_id': order_id, 'error': '订单不存在'}

    status = order['status']
    target = target_effect(status)
    st = conn.execute('SELECT * FROM inv_order_state WHERE order_id=?', (order_id,)).fetchone()
    cur_state = st['inv_state'] if st else 'none'

    market = order_market(conn, order)

    # 忽略的状态(如未知)——记录但不动库存
    if target is None:
        _save_state(conn, order_id, order['source'], cur_state if st else 'none',
                    (st['warehouse_id'] if st else None),
                    (json.loads(st['committed_json']) if st and st['committed_json'] else None),
                    status, market, note=f'状态 {status} 不参与库存联动')
        if commit:
            conn.commit()
        return {'order_id': order_id, 'status': status, 'action': 'ignored', 'inv_state': cur_state}

    # 解析 line_items → SKU
    if caches is None:
        caches = inv_resolver.build_caches(conn)
    resolved = inv_resolver.resolve_order(conn, order, caches)
    sku_agg = _aggregate_sku_lines(resolved)

    # 有未映射:无法安全扣减 → 先冲销旧效果(避免悬挂),标 blocked
    if resolved['unmatched'] > 0 or not sku_agg:
        if cur_state == 'reserved':
            _unreserve(conn, st, uid, uname)
        elif cur_state == 'shipped':
            _unship(conn, st, uid, uname)
        _save_state(conn, order_id, order['source'], 'blocked', (st['warehouse_id'] if st else None),
                    None, status, market,
                    note=f"{resolved['unmatched']} 行未映射,无法扣减")
        if commit:
            conn.commit()
        return {'order_id': order_id, 'status': status, 'action': 'blocked',
                'unmatched': resolved['unmatched'], 'inv_state': 'blocked'}

    wid = pick_warehouse(conn, order, market)
    if not wid:
        _save_state(conn, order_id, order['source'], 'blocked', None, None, status, market,
                    note='无可用仓库(未配置市场路由?)')
        if commit:
            conn.commit()
        return {'order_id': order_id, 'status': status, 'action': 'blocked', 'reason': 'no_warehouse'}

    # ---- 目标 = clear(释放/退货) ----
    if target == 'clear':
        if cur_state == 'reserved':
            _unreserve(conn, st, uid, uname)
            new_state = 'released'
        elif cur_state == 'shipped':
            _unship(conn, st, uid, uname)
            new_state = 'returned'
        else:
            new_state = cur_state if cur_state in ('released', 'returned') else 'released'
        _save_state(conn, order_id, order['source'], new_state, wid, None, status, market)
        if commit:
            conn.commit()
        return {'order_id': order_id, 'status': status, 'action': new_state, 'inv_state': new_state}

    # ---- 目标 = reserved ----
    if target == 'reserved':
        if cur_state == 'reserved':
            if commit:
                conn.commit()
            return {'order_id': order_id, 'status': status, 'action': 'noop', 'inv_state': 'reserved'}
        if cur_state == 'shipped':
            _unship(conn, st, uid, uname)  # 出库回退→再预留
        committed = _apply_reserve(conn, order_id, wid, sku_agg, uid, uname)
        _save_state(conn, order_id, order['source'], 'reserved', wid, committed, status, market)
        if commit:
            conn.commit()
        return {'order_id': order_id, 'status': status, 'action': 'reserved', 'inv_state': 'reserved',
                'warehouse_id': wid, 'skus': len(sku_agg)}

    # ---- 目标 = shipped ----
    if target == 'shipped':
        if cur_state == 'shipped':
            if commit:
                conn.commit()
            return {'order_id': order_id, 'status': status, 'action': 'noop', 'inv_state': 'shipped'}
        from_reserved = (cur_state == 'reserved')
        committed = _apply_ship(conn, order_id, wid, sku_agg, uid, uname, from_reserved)
        _save_state(conn, order_id, order['source'], 'shipped', wid, committed, status, market)
        if commit:
            conn.commit()
        return {'order_id': order_id, 'status': status, 'action': 'shipped', 'inv_state': 'shipped',
                'warehouse_id': wid, 'skus': len(sku_agg), 'batches': len(committed['batches'])}


# ───────────────────────────── API ─────────────────────────────

@inv_ord_bp.route('/inventory/orders')
@login_required
@inv_view_required
def orders_page():
    return render_template('inv_orders.html')


@inv_ord_bp.route('/api/inv/order-state/<path:order_id>', methods=['GET'])
@login_required
@inv_view_required
def get_order_state(order_id):
    """查看订单库存状态 + 解析预览(不写库)。"""
    conn = get_conn()
    try:
        order = conn.execute('SELECT * FROM orders WHERE id=?', (order_id,)).fetchone()
        if not order:
            return jsonify({'error': '订单不存在'}), 404
        st = conn.execute('SELECT * FROM inv_order_state WHERE order_id=?', (order_id,)).fetchone()
        resolved = inv_resolver.resolve_order(conn, order)
        return jsonify({
            'order_id': order_id, 'order_status': order['status'],
            'target_effect': target_effect(order['status']),
            'market': order_market(conn, order),
            'inv_state': (st['inv_state'] if st else 'none'),
            'committed': (json.loads(st['committed_json']) if st and st['committed_json'] else None),
            'resolved': resolved,
        })
    finally:
        conn.close()


@inv_ord_bp.route('/api/inv/process-order/<path:order_id>', methods=['POST'])
@login_required
@inv_manage_required
def api_process_order(order_id):
    conn = get_conn()
    try:
        res = process_order(conn, order_id)
        code = 200 if not res.get('error') else 404
        return jsonify(res), code
    finally:
        conn.close()


@inv_ord_bp.route('/api/inv/process-orders', methods=['POST'])
@login_required
@inv_manage_required
def api_process_orders():
    """批量处理最近订单。body: {limit, source, only_status:[...]}。返回各动作计数。"""
    d = request.get_json(silent=True) or {}
    try:
        limit = min(5000, int(d.get('limit') or 500))
    except (TypeError, ValueError):
        limit = 500
    conn = get_conn()
    try:
        sql = "SELECT id FROM orders WHERE 1=1"
        params = []
        if d.get('source'):
            sql += ' AND source=?'; params.append(d['source'])
        sql += ' ORDER BY id DESC LIMIT ?'; params.append(limit)
        ids = [r['id'] for r in conn.execute(sql, params).fetchall()]
        caches = inv_resolver.build_caches(conn)
        counts = {}
        for oid in ids:
            res = process_order(conn, oid, caches=caches, commit=False)
            act = res.get('action', 'error')
            counts[act] = counts.get(act, 0) + 1
        conn.commit()
        return jsonify({'processed': len(ids), 'counts': counts})
    finally:
        conn.close()


@inv_ord_bp.route('/api/inv/order-states', methods=['GET'])
@login_required
@inv_view_required
def list_order_states():
    """已处理订单的库存状态列表。?state= ?limit=。"""
    state = request.args.get('state')
    try:
        limit = min(2000, int(request.args.get('limit') or 200))
    except ValueError:
        limit = 200
    conn = get_conn()
    sql = '''SELECT os.*, o.status AS order_status, o.number
             FROM inv_order_state os LEFT JOIN orders o ON o.id=os.order_id WHERE 1=1'''
    params = []
    if state:
        sql += ' AND os.inv_state=?'; params.append(state)
    sql += ' ORDER BY os.updated_at DESC LIMIT ?'; params.append(limit)
    rows = conn.execute(sql, params).fetchall()
    summary = {r['inv_state']: r['n'] for r in conn.execute(
        'SELECT inv_state, COUNT(*) AS n FROM inv_order_state GROUP BY inv_state').fetchall()}
    conn.close()
    return jsonify({'summary': summary, 'items': [dict(r) for r in rows]})
