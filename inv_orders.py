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
import inv_allocator

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


# ───────────────────────── 效果施加/冲销(多仓) ─────────────────────────
# committed_json 形如:
#   {'allocations': [ {warehouse_id, mode:'reserved'|'shipped',
#                      lines:[{sku_id,qty}], batches:[{batch_id,qty}]}, ... ]}

def _aggregate_sku_lines(resolved):
    """把解析结果按 sku_id 聚合 sku 总数(同 SKU 多行合并)。"""
    agg = {}
    for ln in resolved['lines']:
        if ln['sku_id']:
            agg[ln['sku_id']] = agg.get(ln['sku_id'], 0) + ln['sku_qty']
    return agg


def _committed_allocs(st):
    if not st or not st['committed_json']:
        return []
    return (json.loads(st['committed_json']) or {}).get('allocations', [])


def _reserve_alloc(conn, order_id, alloc, uid, uname):
    """对一个仓的分配做预留。"""
    wid = alloc['warehouse_id']
    for ln in alloc['lines']:
        record_movement(conn, warehouse_id=wid, sku_id=ln['sku_id'], movement_type='reserve',
                        reserved_delta=ln['qty'], ref_type='order', ref_id=order_id,
                        order_id=order_id, operator_id=uid, operator_name=uname, note='订单预留')
    return {'warehouse_id': wid, 'mode': 'reserved',
            'lines': [dict(l) for l in alloc['lines']], 'batches': []}


def _ship_alloc(conn, order_id, alloc, uid, uname, from_reserved):
    """对一个仓的分配做出库:扣 on_hand + FEFO 批次;from_reserved 时释放预留。"""
    wid = alloc['warehouse_id']
    batches = []
    for ln in alloc['lines']:
        allocs, _short = inv_batches.plan_fefo(conn, wid, ln['sku_id'], ln['qty'])
        inv_batches.consume_batches(conn, allocs)
        for a in allocs:
            batches.append({'batch_id': a['batch_id'], 'qty': a['qty']})
        record_movement(conn, warehouse_id=wid, sku_id=ln['sku_id'], movement_type='sale_out',
                        qty_delta=-ln['qty'], reserved_delta=(-ln['qty'] if from_reserved else 0),
                        ref_type='order', ref_id=order_id, order_id=order_id,
                        operator_id=uid, operator_name=uname,
                        note='销售出库' + ('(原预留转出)' if from_reserved else ''))
    return {'warehouse_id': wid, 'mode': 'shipped',
            'lines': [dict(l) for l in alloc['lines']], 'batches': batches}


def _unreserve_allocs(conn, order_id, allocs, uid, uname):
    for a in allocs:
        for ln in a.get('lines', []):
            record_movement(conn, warehouse_id=a['warehouse_id'], sku_id=ln['sku_id'],
                            movement_type='release', reserved_delta=-ln['qty'],
                            ref_type='order', ref_id=order_id, order_id=order_id,
                            operator_id=uid, operator_name=uname, note='释放预留(状态变更)')


def _unship_allocs(conn, order_id, allocs, uid, uname):
    for a in allocs:
        for b in a.get('batches', []):
            inv_batches.restock_batch(conn, b['batch_id'], b['qty'])
        for ln in a.get('lines', []):
            record_movement(conn, warehouse_id=a['warehouse_id'], sku_id=ln['sku_id'],
                            movement_type='return_in', qty_delta=ln['qty'],
                            ref_type='order', ref_id=order_id, order_id=order_id,
                            operator_id=uid, operator_name=uname, note='退货入库(状态变更)')


def _undo_current(conn, order_id, st, uid, uname):
    """按当前状态冲销已提交效果(预留→释放;出库→退货回补)。"""
    if not st:
        return
    if st['inv_state'] == 'reserved':
        _unreserve_allocs(conn, order_id, _committed_allocs(st), uid, uname)
    elif st['inv_state'] == 'shipped':
        _unship_allocs(conn, order_id, _committed_allocs(st), uid, uname)


# ── inv_fulfillments:由 committed 派生,每次状态转换重建 ──

def _clear_fulfillments(conn, order_id):
    conn.execute('DELETE FROM inv_fulfillments WHERE order_id=?', (order_id,))  # 级联删 items


def _rebuild_fulfillments(conn, order_id, source, committed_allocs, status, uid, uname):
    _clear_fulfillments(conn, order_id)
    n = len(committed_allocs)
    for i, a in enumerate(committed_allocs, 1):
        cur = conn.execute('''INSERT INTO inv_fulfillments
            (order_id, source, warehouse_id, status, is_split, split_index, split_total,
             operator_id, operator_name)
            VALUES (?,?,?,?,?,?,?,?,?)''',
            (order_id, source, a['warehouse_id'], status, 1 if n > 1 else 0, i, n, uid, uname))
        fid = cur.lastrowid
        # 把批次按 sku 粗略对应到 item(单 sku 单 item;批次明细在流水里已可追溯)
        for ln in a.get('lines', []):
            conn.execute('INSERT INTO inv_fulfillment_items (fulfillment_id, sku_id, qty) VALUES (?,?,?)',
                         (fid, ln['sku_id'], ln['qty']))


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


def _plan_allocations(conn, order, market, sku_agg):
    """决定分配方案。

    设计要点:现有系统给每张订单都自动设了 orders.warehouse_id(按站点国家)。
    若直接尊重它,自动分仓/拆单就永远不会发生。因此:
      - 该市场**已配置**路由(inv_market_warehouses 有候选仓)→ 用分配引擎
        (单仓优先,缺货拆单)——这是本系统的核心目标行为;
      - 该市场**未配置**路由 → 回退到订单既有 warehouse_id(单仓),
        保证未接入自动分仓的市场(如 AU/AE)照旧工作。
    未来新市场只要在仓库管理里加路由,就自动切换到分仓引擎,无需改代码。

    返回 (allocations_for_apply, shortage, is_split, plan_reason)。
    """
    cands = inv_allocator.candidate_warehouses(conn, market)
    if cands:
        plan = inv_allocator.allocate(conn, market, sku_agg)
        allocs = [{'warehouse_id': a['warehouse_id'], 'lines': a['lines']} for a in plan['allocations']]
        return allocs, plan['shortage'], plan['is_split'], plan['reason']
    # 未配置该市场路由 → 回退订单既有仓库(单仓)
    forced = order['warehouse_id'] if 'warehouse_id' in order.keys() else None
    if forced:
        return ([{'warehouse_id': forced,
                  'lines': [{'sku_id': k, 'qty': v} for k, v in sku_agg.items()]}],
                {}, False, 'legacy_warehouse')
    return [], dict(sku_agg), False, 'no_warehouse'


# ───────────────────────── 幂等处理器(多仓) ─────────────────────────

def process_order(conn, order_id, caches=None, commit=True):
    """把单张订单的库存对齐到其当前状态。幂等。支持多仓自动拆单。"""
    uid, uname = current_operator()
    order = conn.execute('SELECT * FROM orders WHERE id=?', (order_id,)).fetchone()
    if not order:
        return {'order_id': order_id, 'error': '订单不存在'}

    status = order['status']
    target = target_effect(status)
    st = conn.execute('SELECT * FROM inv_order_state WHERE order_id=?', (order_id,)).fetchone()
    cur_state = st['inv_state'] if st else 'none'
    market = order_market(conn, order)

    # 不参与联动的状态
    if target is None:
        _save_state(conn, order_id, order['source'], cur_state if st else 'none',
                    (st['warehouse_id'] if st else None),
                    (json.loads(st['committed_json']) if st and st['committed_json'] else None),
                    status, market, note=f'状态 {status} 不参与库存联动')
        if commit:
            conn.commit()
        return {'order_id': order_id, 'status': status, 'action': 'ignored', 'inv_state': cur_state}

    if caches is None:
        caches = inv_resolver.build_caches(conn)
    resolved = inv_resolver.resolve_order(conn, order, caches)
    sku_agg = _aggregate_sku_lines(resolved)

    # 未映射:冲销旧效果并标 blocked
    if resolved['unmatched'] > 0 or not sku_agg:
        _undo_current(conn, order_id, st, uid, uname)
        _clear_fulfillments(conn, order_id)
        _save_state(conn, order_id, order['source'], 'blocked', None, None, status, market,
                    note=f"{resolved['unmatched']} 行未映射,无法扣减")
        if commit:
            conn.commit()
        return {'order_id': order_id, 'status': status, 'action': 'blocked',
                'unmatched': resolved['unmatched'], 'inv_state': 'blocked'}

    # ---- clear:释放/退货 ----
    if target == 'clear':
        if cur_state == 'reserved':
            _unreserve_allocs(conn, order_id, _committed_allocs(st), uid, uname); new_state = 'released'
        elif cur_state == 'shipped':
            _unship_allocs(conn, order_id, _committed_allocs(st), uid, uname); new_state = 'returned'
        else:
            new_state = cur_state if cur_state in ('released', 'returned') else 'released'
        _clear_fulfillments(conn, order_id)
        _save_state(conn, order_id, order['source'], new_state, None, None, status, market)
        if commit:
            conn.commit()
        return {'order_id': order_id, 'status': status, 'action': new_state, 'inv_state': new_state}

    # ---- reserved ----
    if target == 'reserved':
        if cur_state == 'reserved':
            if commit:
                conn.commit()
            return {'order_id': order_id, 'status': status, 'action': 'noop', 'inv_state': 'reserved'}
        if cur_state == 'shipped':
            _unship_allocs(conn, order_id, _committed_allocs(st), uid, uname)  # 回退出库
        allocs, shortage, is_split, reason = _plan_allocations(conn, order, market, sku_agg)
        if not allocs:
            _clear_fulfillments(conn, order_id)
            _save_state(conn, order_id, order['source'], 'blocked', None, None, status, market,
                        note='无可用仓库/库存(无法分配)')
            if commit:
                conn.commit()
            return {'order_id': order_id, 'status': status, 'action': 'blocked', 'reason': reason}
        committed_allocs = [_reserve_alloc(conn, order_id, a, uid, uname) for a in allocs]
        committed = {'allocations': committed_allocs}
        _rebuild_fulfillments(conn, order_id, order['source'], committed_allocs, 'reserved', uid, uname)
        note = ('部分缺货:' + json.dumps(shortage)) if shortage else (('拆 %d 仓' % len(allocs)) if is_split else None)
        _save_state(conn, order_id, order['source'], 'reserved',
                    (allocs[0]['warehouse_id'] if len(allocs) == 1 else None),
                    committed, status, market, note=note)
        if commit:
            conn.commit()
        return {'order_id': order_id, 'status': status, 'action': 'reserved', 'inv_state': 'reserved',
                'is_split': is_split, 'warehouses': [a['warehouse_id'] for a in allocs],
                'shortage': shortage}

    # ---- shipped ----
    if target == 'shipped':
        if cur_state == 'shipped':
            if commit:
                conn.commit()
            return {'order_id': order_id, 'status': status, 'action': 'noop', 'inv_state': 'shipped'}
        if cur_state == 'reserved':
            # 出库已预留的相同分配(保持与预留一致,不重新分配)
            committed_allocs = [_ship_alloc(conn, order_id, a, uid, uname, from_reserved=True)
                                for a in _committed_allocs(st)]
            shortage, is_split = {}, len(committed_allocs) > 1
        else:
            allocs, shortage, is_split, reason = _plan_allocations(conn, order, market, sku_agg)
            if not allocs:
                _clear_fulfillments(conn, order_id)
                _save_state(conn, order_id, order['source'], 'blocked', None, None, status, market,
                            note='无可用仓库/库存(无法分配)')
                if commit:
                    conn.commit()
                return {'order_id': order_id, 'status': status, 'action': 'blocked', 'reason': reason}
            committed_allocs = [_ship_alloc(conn, order_id, a, uid, uname, from_reserved=False)
                                for a in allocs]
        committed = {'allocations': committed_allocs}
        _rebuild_fulfillments(conn, order_id, order['source'], committed_allocs, 'shipped', uid, uname)
        note = ('部分缺货:' + json.dumps(shortage)) if shortage else (('拆 %d 仓' % len(committed_allocs)) if is_split else None)
        _save_state(conn, order_id, order['source'], 'shipped',
                    (committed_allocs[0]['warehouse_id'] if len(committed_allocs) == 1 else None),
                    committed, status, market, note=note)
        if commit:
            conn.commit()
        return {'order_id': order_id, 'status': status, 'action': 'shipped', 'inv_state': 'shipped',
                'is_split': is_split, 'warehouses': [a['warehouse_id'] for a in committed_allocs],
                'shortage': shortage}


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


@inv_ord_bp.route('/api/inv/allocate-preview/<path:order_id>', methods=['GET'])
@login_required
@inv_view_required
def allocate_preview(order_id):
    """分仓预览(不写库):展示该订单按市场路由+实时可用会怎么分/拆/缺货。"""
    conn = get_conn()
    try:
        order = conn.execute('SELECT * FROM orders WHERE id=?', (order_id,)).fetchone()
        if not order:
            return jsonify({'error': '订单不存在'}), 404
        resolved = inv_resolver.resolve_order(conn, order)
        sku_agg = _aggregate_sku_lines(resolved)
        market = order_market(conn, order)
        if resolved['unmatched'] > 0 or not sku_agg:
            return jsonify({'order_id': order_id, 'market': market, 'blocked': True,
                            'unmatched': resolved['unmatched'], 'resolved': resolved})
        forced = order['warehouse_id'] if 'warehouse_id' in order.keys() else None
        plan = inv_allocator.allocate(conn, market, sku_agg)
        # 附 SKU 名称便于展示
        names = {r['id']: r['sku_code'] for r in conn.execute(
            'SELECT id, sku_code FROM inv_skus WHERE id IN (%s)' %
            (','.join('?' * len(sku_agg)) or 'NULL'), list(sku_agg.keys())).fetchall()} if sku_agg else {}
        for a in plan['allocations']:
            for ln in a['lines']:
                ln['sku_code'] = names.get(ln['sku_id'])
        return jsonify({'order_id': order_id, 'market': market, 'forced_warehouse_id': forced,
                        'needs': [{'sku_id': k, 'sku_code': names.get(k), 'qty': v} for k, v in sku_agg.items()],
                        'plan': plan})
    finally:
        conn.close()


@inv_ord_bp.route('/api/inv/fulfillments/<path:order_id>', methods=['GET'])
@login_required
@inv_view_required
def list_fulfillments(order_id):
    """订单的分单(fulfillment)明细。"""
    conn = get_conn()
    try:
        fuls = conn.execute('''SELECT f.*, w.name AS warehouse_name,
                                COALESCE(we.ownership_type,'self') AS ownership_type, we.partner_name
                               FROM inv_fulfillments f
                               LEFT JOIN warehouses w ON w.id=f.warehouse_id
                               LEFT JOIN inv_warehouse_ext we ON we.warehouse_id=f.warehouse_id
                               WHERE f.order_id=? ORDER BY f.split_index''', (order_id,)).fetchall()
        out = []
        for f in fuls:
            items = conn.execute('''SELECT i.*, k.sku_code, k.name AS sku_name
                                    FROM inv_fulfillment_items i JOIN inv_skus k ON k.id=i.sku_id
                                    WHERE i.fulfillment_id=?''', (f['id'],)).fetchall()
            d = dict(f); d['items'] = [dict(x) for x in items]
            out.append(d)
        return jsonify(out)
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
