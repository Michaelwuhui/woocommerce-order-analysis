"""
inv_batches.py — 模块 4:批次 + 保质期(FEFO 先到期先发)。

批次由采购收货(模块3)生成。本模块提供:
  - FEFO 分配:按到期日升序(最早先发,无到期日的批次排最后)规划/执行消耗。
  - 退货回补:把退回的数量回补到批次(优先回原批次)。
  - 保质期判定:过期 / 临期(默认 90 天内到期)/ 正常。
  - 批次列表、临期过期报表 API + 页面。

设计:plan_fefo 是**纯规划**(不改库),consume_batches / restock_batch 才落库。
这样模块 5(卖出扣减)可以:plan_fefo → 按批次写 sale_out 流水(record_movement)
→ consume_batches,三步在同一事务里保持批次与台账一致。

注意:批次的 qty_remaining 之和 == inv_stock.on_hand(同仓同SKU)是一条不变式;
所有改动批次余额的路径都应同时经 record_movement 改台账(由调用方负责)。
"""

import datetime
from flask import Blueprint, render_template, request, jsonify
from flask_login import login_required

from inv_common import get_conn, inv_view_required

inv_batch_bp = Blueprint('inv_batch', __name__)

DEFAULT_NEAR_DAYS = 90  # 临期阈值默认 90 天


# ───────────────────────── 保质期判定 ─────────────────────────

def batch_status(expiry_date, near_days=DEFAULT_NEAR_DAYS, today=None):
    """返回 'expired' | 'near' | 'ok' | 'none'(无到期日)。"""
    if not expiry_date:
        return 'none'
    today = today or datetime.date.today()
    try:
        exp = datetime.date.fromisoformat(str(expiry_date)[:10])
    except ValueError:
        return 'none'
    if exp < today:
        return 'expired'
    if (exp - today).days <= near_days:
        return 'near'
    return 'ok'


# ───────────────────────── FEFO 分配 ─────────────────────────

def plan_fefo(conn, warehouse_id, sku_id, qty, include_expired=False):
    """规划从哪些批次取货以满足 qty(纯规划,不改库)。

    顺序:到期日升序(最早先发);无到期日的批次排最后;同到期日按入库先后(id)。
    默认跳过已过期批次(include_expired=False)——过期货不应自动发出。

    返回 (allocations, shortage):
      allocations = [{batch_id, qty, expiry_date, unit_cost, batch_no}, ...]
      shortage    = 仍未满足的数量(>0 表示可用批次不足)
    """
    if qty <= 0:
        return [], 0
    today = datetime.date.today().isoformat()
    rows = conn.execute('''
        SELECT id, batch_no, expiry_date, unit_cost, qty_remaining
        FROM inv_batches
        WHERE warehouse_id=? AND sku_id=? AND qty_remaining > 0
        ORDER BY (expiry_date IS NULL), expiry_date, id
    ''', (warehouse_id, sku_id)).fetchall()

    allocations = []
    need = qty
    for r in rows:
        if need <= 0:
            break
        if not include_expired and r['expiry_date'] and str(r['expiry_date'])[:10] < today:
            continue  # 跳过过期批次
        take = min(need, r['qty_remaining'])
        allocations.append({'batch_id': r['id'], 'qty': take, 'expiry_date': r['expiry_date'],
                            'unit_cost': r['unit_cost'], 'batch_no': r['batch_no']})
        need -= take
    return allocations, need


def consume_batches(conn, allocations):
    """执行分配:扣减各批次 qty_remaining。调用方需在同事务另写台账流水。"""
    for a in allocations:
        conn.execute('UPDATE inv_batches SET qty_remaining = qty_remaining - ?, updated_at=CURRENT_TIMESTAMP '
                     'WHERE id=?', (a['qty'], a['batch_id']))


def restock_batch(conn, batch_id, qty):
    """退货回补:把 qty 加回指定批次的 qty_remaining。"""
    conn.execute('UPDATE inv_batches SET qty_remaining = qty_remaining + ?, updated_at=CURRENT_TIMESTAMP '
                 'WHERE id=?', (qty, batch_id))


def total_remaining(conn, warehouse_id, sku_id):
    """某仓某 SKU 所有批次剩余之和(应等于 inv_stock.on_hand)。"""
    r = conn.execute('SELECT COALESCE(SUM(qty_remaining),0) AS s FROM inv_batches '
                     'WHERE warehouse_id=? AND sku_id=?', (warehouse_id, sku_id)).fetchone()
    return r['s']


# ───────────────────────────── API ─────────────────────────────

@inv_batch_bp.route('/inventory/batches')
@login_required
@inv_view_required
def batches_page():
    conn = get_conn()
    warehouses = conn.execute('SELECT id, name, code FROM warehouses WHERE is_active=1 ORDER BY name').fetchall()
    conn.close()
    return render_template('inv_batches.html', warehouses=[dict(w) for w in warehouses])


@inv_batch_bp.route('/api/inv/batches', methods=['GET'])
@login_required
@inv_view_required
def list_batches():
    """批次列表。?warehouse_id= ?sku_id= ?q= ?status=(expired/near/ok) ?near_days= ?remaining=1。"""
    wid = request.args.get('warehouse_id')
    sku_id = request.args.get('sku_id')
    q = (request.args.get('q') or '').strip()
    status = request.args.get('status')
    remaining = request.args.get('remaining')
    try:
        near_days = int(request.args.get('near_days') or DEFAULT_NEAR_DAYS)
    except ValueError:
        near_days = DEFAULT_NEAR_DAYS

    conn = get_conn()
    sql = '''SELECT b.*, w.name AS warehouse_name, k.sku_code, k.name AS sku_name
             FROM inv_batches b
             JOIN warehouses w ON w.id=b.warehouse_id
             JOIN inv_skus k ON k.id=b.sku_id WHERE 1=1'''
    params = []
    if wid:
        sql += ' AND b.warehouse_id=?'; params.append(wid)
    if sku_id:
        sql += ' AND b.sku_id=?'; params.append(sku_id)
    if q:
        sql += ' AND (k.sku_code LIKE ? OR k.name LIKE ? OR b.batch_no LIKE ?)'; params += [f'%{q}%', f'%{q}%', f'%{q}%']
    if remaining == '1':
        sql += ' AND b.qty_remaining > 0'
    sql += ' ORDER BY (b.expiry_date IS NULL), b.expiry_date, b.id'
    rows = conn.execute(sql, params).fetchall()
    conn.close()

    out = []
    for r in rows:
        d = dict(r)
        d['status'] = batch_status(r['expiry_date'], near_days)
        if r['expiry_date']:
            try:
                exp = datetime.date.fromisoformat(str(r['expiry_date'])[:10])
                d['days_to_expiry'] = (exp - datetime.date.today()).days
            except ValueError:
                d['days_to_expiry'] = None
        else:
            d['days_to_expiry'] = None
        if status and d['status'] != status:
            continue
        out.append(d)
    return jsonify(out)


@inv_batch_bp.route('/api/inv/expiry-report', methods=['GET'])
@login_required
@inv_view_required
def expiry_report():
    """临期 + 过期报表(仅含 qty_remaining>0 的批次)。?near_days= ?warehouse_id=。"""
    try:
        near_days = int(request.args.get('near_days') or DEFAULT_NEAR_DAYS)
    except ValueError:
        near_days = DEFAULT_NEAR_DAYS
    wid = request.args.get('warehouse_id')
    conn = get_conn()
    sql = '''SELECT b.*, w.name AS warehouse_name, k.sku_code, k.name AS sku_name
             FROM inv_batches b JOIN warehouses w ON w.id=b.warehouse_id JOIN inv_skus k ON k.id=b.sku_id
             WHERE b.qty_remaining > 0 AND b.expiry_date IS NOT NULL'''
    params = []
    if wid:
        sql += ' AND b.warehouse_id=?'; params.append(wid)
    sql += ' ORDER BY b.expiry_date, b.id'
    rows = conn.execute(sql, params).fetchall()
    conn.close()

    expired, near = [], []
    expired_qty = near_qty = 0
    for r in rows:
        st = batch_status(r['expiry_date'], near_days)
        d = dict(r)
        try:
            exp = datetime.date.fromisoformat(str(r['expiry_date'])[:10])
            d['days_to_expiry'] = (exp - datetime.date.today()).days
        except ValueError:
            d['days_to_expiry'] = None
        if st == 'expired':
            expired.append(d); expired_qty += r['qty_remaining']
        elif st == 'near':
            near.append(d); near_qty += r['qty_remaining']
    return jsonify({
        'near_days': near_days,
        'expired': expired, 'near': near,
        'expired_count': len(expired), 'near_count': len(near),
        'expired_qty': expired_qty, 'near_qty': near_qty,
    })
