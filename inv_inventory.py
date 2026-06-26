"""
inv_inventory.py — 模块 3:库存台账 + 出入库流水 + 采购入库。

职责:
  - 供应商(inv_suppliers)CRUD。
  - 采购单(inv_purchase_orders + items):建单(draft)→ 收货(received)。
    收货时:为每条明细生成批次(inv_batches),并通过 record_movement 写
    purchase_in 流水、同步更新 inv_stock。收货后单据锁定不可重复收。
  - 手工盘点调整:对某仓某 SKU 直接增减(adjust),必填原因,写审计。
  - 库存台账查询:仓 × SKU 的 on_hand / reserved / available(=on_hand-reserved)。
  - 出入库流水查询:按仓/SKU/类型/日期过滤。

所有改变库存的写操作一律经 inv_common.record_movement,保证:
  审计(操作人+时间+前后数量)与 inv_stock 物化值始终一致。
"""

import datetime
from flask import Blueprint, render_template, request, jsonify
from flask_login import login_required

from inv_common import (
    get_conn, inv_view_required, inv_manage_required,
    record_movement, current_operator, warehouse_scope_clause,
)

inv_inv_bp = Blueprint('inv_inv', __name__)


# ───────────────────────────── 页面 ─────────────────────────────

@inv_inv_bp.route('/inventory/stock')
@login_required
@inv_view_required
def stock_page():
    conn = get_conn()
    warehouses = conn.execute(
        'SELECT id, name, code, country FROM warehouses WHERE is_active=1 ORDER BY country, name'
    ).fetchall()
    conn.close()
    return render_template('inv_inventory.html', warehouses=[dict(w) for w in warehouses])


# ─────────────────────────── 供应商 API ───────────────────────────

@inv_inv_bp.route('/api/inv/suppliers', methods=['GET'])
@login_required
@inv_view_required
def list_suppliers():
    conn = get_conn()
    rows = conn.execute('SELECT * FROM inv_suppliers ORDER BY is_active DESC, name').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@inv_inv_bp.route('/api/inv/suppliers', methods=['POST'])
@login_required
@inv_manage_required
def create_supplier():
    d = request.get_json(force=True) or {}
    name = (d.get('name') or '').strip()
    if not name:
        return jsonify({'error': '供应商名称必填'}), 400
    conn = get_conn()
    try:
        cur = conn.execute('''INSERT INTO inv_suppliers (name, contact, phone, email, address, currency, notes)
                              VALUES (?,?,?,?,?,?,?)''',
                           (name, d.get('contact'), d.get('phone'), d.get('email'),
                            d.get('address'), (d.get('currency') or 'CNY').strip(), d.get('notes')))
        conn.commit()
        return jsonify({'success': True, 'id': cur.lastrowid})
    finally:
        conn.close()


@inv_inv_bp.route('/api/inv/suppliers/<int:sid>', methods=['PUT'])
@login_required
@inv_manage_required
def update_supplier(sid):
    d = request.get_json(force=True) or {}
    conn = get_conn()
    try:
        conn.execute('''UPDATE inv_suppliers SET name=?, contact=?, phone=?, email=?, address=?,
                        currency=?, notes=?, is_active=?, updated_at=CURRENT_TIMESTAMP WHERE id=?''',
                     (d.get('name'), d.get('contact'), d.get('phone'), d.get('email'),
                      d.get('address'), (d.get('currency') or 'CNY'), d.get('notes'),
                      1 if d.get('is_active', True) else 0, sid))
        conn.commit()
        return jsonify({'success': True})
    finally:
        conn.close()


# ─────────────────────────── 采购单 API ───────────────────────────

@inv_inv_bp.route('/api/inv/purchase-orders', methods=['GET'])
@login_required
@inv_view_required
def list_pos():
    """采购单列表(可 ?status= / ?warehouse_id= 过滤)。"""
    status = request.args.get('status')
    wid = request.args.get('warehouse_id')
    conn = get_conn()
    sql = '''SELECT po.*, w.name AS warehouse_name, sup.name AS supplier_name,
                    (SELECT COUNT(*) FROM inv_purchase_order_items i WHERE i.po_id=po.id) AS item_count,
                    (SELECT COALESCE(SUM(i.qty),0) FROM inv_purchase_order_items i WHERE i.po_id=po.id) AS total_qty
             FROM inv_purchase_orders po
             LEFT JOIN warehouses w ON w.id=po.warehouse_id
             LEFT JOIN inv_suppliers sup ON sup.id=po.supplier_id WHERE 1=1'''
    params = []
    if status:
        sql += ' AND po.status=?'; params.append(status)
    if wid:
        sql += ' AND po.warehouse_id=?'; params.append(wid)
    sql += ' ORDER BY po.id DESC LIMIT 500'
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@inv_inv_bp.route('/api/inv/purchase-orders/<int:pid>', methods=['GET'])
@login_required
@inv_view_required
def get_po(pid):
    conn = get_conn()
    try:
        po = conn.execute('''SELECT po.*, w.name AS warehouse_name, sup.name AS supplier_name
                             FROM inv_purchase_orders po
                             LEFT JOIN warehouses w ON w.id=po.warehouse_id
                             LEFT JOIN inv_suppliers sup ON sup.id=po.supplier_id WHERE po.id=?''', (pid,)).fetchone()
        if not po:
            return jsonify({'error': '采购单不存在'}), 404
        items = conn.execute('''SELECT i.*, k.sku_code, k.name AS sku_name
                                FROM inv_purchase_order_items i JOIN inv_skus k ON k.id=i.sku_id
                                WHERE i.po_id=?''', (pid,)).fetchall()
        out = dict(po)
        out['items'] = [dict(r) for r in items]
        return jsonify(out)
    finally:
        conn.close()


def _gen_po_no(conn):
    today = datetime.date.today().strftime('%Y%m%d')
    n = conn.execute("SELECT COUNT(*) FROM inv_purchase_orders WHERE po_no LIKE ?", (f'PO{today}%',)).fetchone()[0]
    return f'PO{today}-{n+1:03d}'


@inv_inv_bp.route('/api/inv/purchase-orders', methods=['POST'])
@login_required
@inv_manage_required
def create_po():
    """建采购单(draft)。items=[{sku_id, qty, unit_cost, batch_no, production_date, expiry_date}]。"""
    d = request.get_json(force=True) or {}
    warehouse_id = d.get('warehouse_id')
    items = d.get('items') or []
    if not warehouse_id:
        return jsonify({'error': '入库仓库必填'}), 400
    if not items:
        return jsonify({'error': '至少一条采购明细'}), 400
    uid, uname = current_operator()
    conn = get_conn()
    try:
        if not conn.execute('SELECT 1 FROM warehouses WHERE id=?', (warehouse_id,)).fetchone():
            return jsonify({'error': '仓库不存在'}), 404
        total = 0.0
        for it in items:
            if not it.get('sku_id') or not it.get('qty'):
                return jsonify({'error': '明细需含 sku_id 与 qty'}), 400
            total += float(it.get('qty') or 0) * float(it.get('unit_cost') or 0)
        po_no = _gen_po_no(conn)
        cur = conn.execute('''INSERT INTO inv_purchase_orders
            (po_no, supplier_id, warehouse_id, status, order_date, currency, total_amount,
             operator_id, operator_name, note)
            VALUES (?,?,?,'draft',?,?,?,?,?,?)''',
            (po_no, d.get('supplier_id') or None, warehouse_id,
             d.get('order_date') or datetime.date.today().isoformat(),
             (d.get('currency') or 'CNY'), total, uid, uname, d.get('note')))
        pid = cur.lastrowid
        for it in items:
            conn.execute('''INSERT INTO inv_purchase_order_items
                (po_id, sku_id, qty, unit_cost, batch_no, production_date, expiry_date, note)
                VALUES (?,?,?,?,?,?,?,?)''',
                (pid, it['sku_id'], int(it['qty']), float(it.get('unit_cost') or 0),
                 it.get('batch_no'), it.get('production_date'), it.get('expiry_date'), it.get('note')))
        conn.commit()
        return jsonify({'success': True, 'id': pid, 'po_no': po_no})
    finally:
        conn.close()


@inv_inv_bp.route('/api/inv/purchase-orders/<int:pid>/receive', methods=['POST'])
@login_required
@inv_manage_required
def receive_po(pid):
    """收货入库:为每条明细建批次 + 写 purchase_in 流水 + 更新台账。幂等防重复收货。"""
    uid, uname = current_operator()
    conn = get_conn()
    try:
        po = conn.execute('SELECT * FROM inv_purchase_orders WHERE id=?', (pid,)).fetchone()
        if not po:
            return jsonify({'error': '采购单不存在'}), 404
        if po['status'] == 'received':
            return jsonify({'error': '该采购单已收货,不能重复入库'}), 400
        if po['status'] == 'cancelled':
            return jsonify({'error': '该采购单已取消'}), 400
        items = conn.execute('SELECT * FROM inv_purchase_order_items WHERE po_id=?', (pid,)).fetchall()
        if not items:
            return jsonify({'error': '采购单无明细'}), 400

        wid = po['warehouse_id']
        received = []
        for it in items:
            qty = int(it['qty'])
            # 1) 建批次
            bcur = conn.execute('''INSERT INTO inv_batches
                (warehouse_id, sku_id, batch_no, production_date, expiry_date,
                 unit_cost, cost_currency, qty_received, qty_remaining, purchase_order_id)
                VALUES (?,?,?,?,?,?,?,?,?,?)''',
                (wid, it['sku_id'], it['batch_no'], it['production_date'], it['expiry_date'],
                 it['unit_cost'], po['currency'], qty, qty, pid))
            batch_id = bcur.lastrowid
            # 2) 写流水 + 更新台账(唯一入口)
            mid = record_movement(conn, warehouse_id=wid, sku_id=it['sku_id'],
                                  movement_type='purchase_in', qty_delta=qty, batch_id=batch_id,
                                  ref_type='po', ref_id=str(pid),
                                  operator_id=uid, operator_name=uname,
                                  note=f"采购入库 {po['po_no']}")
            # 3) 标记明细已收
            conn.execute('UPDATE inv_purchase_order_items SET received_qty=? WHERE id=?', (qty, it['id']))
            received.append({'sku_id': it['sku_id'], 'qty': qty, 'batch_id': batch_id, 'movement_id': mid})

        conn.execute('''UPDATE inv_purchase_orders SET status='received', received_date=?,
                        updated_at=CURRENT_TIMESTAMP WHERE id=?''',
                     (datetime.datetime.now().isoformat(timespec='seconds'), pid))
        conn.commit()
        return jsonify({'success': True, 'received': received})
    finally:
        conn.close()


@inv_inv_bp.route('/api/inv/purchase-orders/<int:pid>/cancel', methods=['POST'])
@login_required
@inv_manage_required
def cancel_po(pid):
    """取消采购单(仅 draft 可取消;已收货不可取消,需走退货/调整冲销)。"""
    conn = get_conn()
    try:
        po = conn.execute('SELECT status FROM inv_purchase_orders WHERE id=?', (pid,)).fetchone()
        if not po:
            return jsonify({'error': '采购单不存在'}), 404
        if po['status'] != 'draft':
            return jsonify({'error': '仅草稿状态可取消'}), 400
        conn.execute("UPDATE inv_purchase_orders SET status='cancelled', updated_at=CURRENT_TIMESTAMP WHERE id=?", (pid,))
        conn.commit()
        return jsonify({'success': True})
    finally:
        conn.close()


# ─────────────────────────── 手工调整 API ───────────────────────────

@inv_inv_bp.route('/api/inv/adjust', methods=['POST'])
@login_required
@inv_manage_required
def adjust_stock():
    """盘点/手工调整。body: {warehouse_id, sku_id, qty_delta(有符号), reason}。

    qty_delta 为对 on_hand 的有符号增量(盘盈正、盘亏负)。必填 reason。
    不允许把现存调成负数。
    """
    d = request.get_json(force=True) or {}
    wid = d.get('warehouse_id'); sku_id = d.get('sku_id')
    reason = (d.get('reason') or '').strip()
    try:
        qty_delta = int(d.get('qty_delta'))
    except (TypeError, ValueError):
        return jsonify({'error': 'qty_delta 必须是整数'}), 400
    if not wid or not sku_id:
        return jsonify({'error': '仓库与 SKU 必填'}), 400
    if qty_delta == 0:
        return jsonify({'error': '调整数量不能为 0'}), 400
    if not reason:
        return jsonify({'error': '必须填写调整原因'}), 400
    uid, uname = current_operator()
    conn = get_conn()
    try:
        cur = conn.execute('SELECT on_hand FROM inv_stock WHERE warehouse_id=? AND sku_id=?', (wid, sku_id)).fetchone()
        on_hand = cur['on_hand'] if cur else 0
        if on_hand + qty_delta < 0:
            return jsonify({'error': f'调整后现存为负(当前 {on_hand},调整 {qty_delta})'}), 400
        mid = record_movement(conn, warehouse_id=wid, sku_id=sku_id, movement_type='adjust',
                              qty_delta=qty_delta, ref_type='manual',
                              operator_id=uid, operator_name=uname, note=reason)
        conn.commit()
        return jsonify({'success': True, 'movement_id': mid})
    finally:
        conn.close()


# ─────────────────────────── 台账 / 流水 查询 ───────────────────────────

@inv_inv_bp.route('/api/inv/stock', methods=['GET'])
@login_required
@inv_view_required
def list_stock():
    """库存台账:仓×SKU 的现存/预留/可用。可 ?warehouse_id= / ?q=(SKU) / ?nonzero=1。"""
    wid = request.args.get('warehouse_id')
    q = (request.args.get('q') or '').strip()
    nonzero = request.args.get('nonzero')
    conn = get_conn()
    sql = '''SELECT st.warehouse_id, st.sku_id, st.on_hand, st.reserved,
                    (st.on_hand - st.reserved) AS available, st.updated_at,
                    w.name AS warehouse_name, w.code AS warehouse_code,
                    k.sku_code, k.name AS sku_name, k.unit
             FROM inv_stock st
             JOIN warehouses w ON w.id=st.warehouse_id
             JOIN inv_skus k ON k.id=st.sku_id WHERE 1=1'''
    params = []
    if wid:
        sql += ' AND st.warehouse_id=?'; params.append(wid)
    if q:
        sql += ' AND (k.sku_code LIKE ? OR k.name LIKE ?)'; params += [f'%{q}%', f'%{q}%']
    if nonzero == '1':
        sql += ' AND (st.on_hand != 0 OR st.reserved != 0)'
    # 角色仓库可见性(合伙人只看自己仓)
    sc, sp = warehouse_scope_clause('st.warehouse_id'); sql += sc; params += sp
    sql += ' ORDER BY w.name, k.sku_code'
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@inv_inv_bp.route('/api/inv/movements', methods=['GET'])
@login_required
@inv_view_required
def list_movements():
    """出入库流水查询。?warehouse_id= ?sku_id= ?type= ?date_from= ?date_to= ?limit=。"""
    conn = get_conn()
    sql = '''SELECT m.*, w.name AS warehouse_name, k.sku_code, k.name AS sku_name
             FROM inv_movements m
             LEFT JOIN warehouses w ON w.id=m.warehouse_id
             LEFT JOIN inv_skus k ON k.id=m.sku_id WHERE 1=1'''
    params = []
    for arg, col in (('warehouse_id', 'm.warehouse_id'), ('sku_id', 'm.sku_id'), ('type', 'm.movement_type')):
        v = request.args.get(arg)
        if v:
            sql += f' AND {col}=?'; params.append(v)
    if request.args.get('date_from'):
        sql += ' AND m.ts >= ?'; params.append(request.args.get('date_from'))
    if request.args.get('date_to'):
        sql += ' AND m.ts <= ?'; params.append(request.args.get('date_to') + ' 23:59:59')
    sc, sp = warehouse_scope_clause('m.warehouse_id'); sql += sc; params += sp
    try:
        limit = min(2000, int(request.args.get('limit') or 300))
    except ValueError:
        limit = 300
    sql += ' ORDER BY m.id DESC LIMIT ?'; params.append(limit)
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])
