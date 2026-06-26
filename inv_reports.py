"""
inv_reports.py — 模块 10:库存报表。

报表(均套用角色仓库可见性 warehouse_scope_clause,合伙人只看自己仓):
  - overview:总览看板(仓库/SKU 数、现存/预留/可用合计、库存货值、临期过期、
    订单库存状态分布、未读通知数)。
  - valuation:库存货值 = Σ(批次 qty_remaining × unit_cost),按仓 + 按 SKU。
  - movements:出入库汇总,按类型 + 期间。
  - fulfillments:分仓/拆单统计。
"""

import datetime
from flask import Blueprint, render_template, request, jsonify
from flask_login import login_required

from inv_common import get_conn, inv_view_required, warehouse_scope_clause

inv_report_bp = Blueprint('inv_report', __name__)


@inv_report_bp.route('/inventory/reports')
@login_required
@inv_view_required
def reports_page():
    conn = get_conn()
    warehouses = conn.execute('SELECT id, name, code FROM warehouses WHERE is_active=1 ORDER BY name').fetchall()
    conn.close()
    return render_template('inv_reports.html', warehouses=[dict(w) for w in warehouses])


@inv_report_bp.route('/api/inv/reports/overview', methods=['GET'])
@login_required
@inv_view_required
def overview():
    conn = get_conn()
    try:
        sc, sp = warehouse_scope_clause('warehouse_id')

        stock = conn.execute(
            f'''SELECT COUNT(DISTINCT sku_id) AS skus,
                       COALESCE(SUM(on_hand),0) AS on_hand,
                       COALESCE(SUM(reserved),0) AS reserved,
                       COALESCE(SUM(on_hand-reserved),0) AS available
                FROM inv_stock WHERE 1=1{sc}''', sp).fetchone()

        wh = conn.execute(
            f"SELECT COUNT(*) AS n FROM warehouses w WHERE is_active=1"
            + (sc.replace('warehouse_id', 'w.id') if sc else ''), sp).fetchone()

        # 库存货值(按币种)
        val_rows = conn.execute(
            f'''SELECT cost_currency, COALESCE(SUM(qty_remaining*unit_cost),0) AS val
                FROM inv_batches WHERE qty_remaining>0{sc} GROUP BY cost_currency''', sp).fetchall()
        valuation = {r['cost_currency'] or '?': round(r['val'], 2) for r in val_rows}

        # 临期/过期(有余量、有到期日)
        today = datetime.date.today().isoformat()
        near_cut = (datetime.date.today() + datetime.timedelta(days=90)).isoformat()
        exp = conn.execute(
            f'''SELECT
                  COALESCE(SUM(CASE WHEN expiry_date < ? THEN qty_remaining ELSE 0 END),0) AS expired,
                  COALESCE(SUM(CASE WHEN expiry_date >= ? AND expiry_date <= ? THEN qty_remaining ELSE 0 END),0) AS near
                FROM inv_batches WHERE qty_remaining>0 AND expiry_date IS NOT NULL{sc}''',
            [today, today, near_cut] + sp).fetchone()

        # 订单库存状态分布
        ostate = {r['inv_state']: r['n'] for r in conn.execute(
            'SELECT inv_state, COUNT(*) AS n FROM inv_order_state GROUP BY inv_state').fetchall()}

        # 未读通知
        sc2, sp2 = warehouse_scope_clause('warehouse_id')
        nscope = (' AND (warehouse_id IS NULL' + sc2.replace(' AND ', ' OR ', 1) + ')') if sc2 else ''
        unread = conn.execute(
            f"SELECT COUNT(*) FROM inv_notifications WHERE status='unread'{nscope}", sp2).fetchone()[0]

        return jsonify({
            'warehouses': wh['n'], 'skus': stock['skus'],
            'on_hand': stock['on_hand'], 'reserved': stock['reserved'], 'available': stock['available'],
            'valuation': valuation,
            'expired_qty': exp['expired'], 'near_qty': exp['near'],
            'order_states': ostate, 'unread_notifications': unread,
        })
    finally:
        conn.close()


@inv_report_bp.route('/api/inv/reports/valuation', methods=['GET'])
@login_required
@inv_view_required
def valuation():
    """库存货值:按仓汇总 + 按 SKU 明细(基于批次剩余 × 单位成本)。"""
    wid = request.args.get('warehouse_id')
    conn = get_conn()
    try:
        sc, sp = warehouse_scope_clause('b.warehouse_id')
        extra, params = '', []
        if wid:
            extra = ' AND b.warehouse_id=?'; params = [wid]
        by_wh = conn.execute(
            f'''SELECT b.warehouse_id, w.name AS warehouse_name, b.cost_currency,
                       SUM(b.qty_remaining) AS qty, ROUND(SUM(b.qty_remaining*b.unit_cost),2) AS value
                FROM inv_batches b JOIN warehouses w ON w.id=b.warehouse_id
                WHERE b.qty_remaining>0{sc}{extra}
                GROUP BY b.warehouse_id, b.cost_currency ORDER BY w.name''', sp + params).fetchall()
        by_sku = conn.execute(
            f'''SELECT b.sku_id, k.sku_code, k.name AS sku_name, b.warehouse_id, w.name AS warehouse_name,
                       b.cost_currency, SUM(b.qty_remaining) AS qty,
                       ROUND(SUM(b.qty_remaining*b.unit_cost),2) AS value
                FROM inv_batches b JOIN inv_skus k ON k.id=b.sku_id JOIN warehouses w ON w.id=b.warehouse_id
                WHERE b.qty_remaining>0{sc}{extra}
                GROUP BY b.sku_id, b.warehouse_id, b.cost_currency
                ORDER BY value DESC LIMIT 500''', sp + params).fetchall()
        return jsonify({'by_warehouse': [dict(r) for r in by_wh],
                        'by_sku': [dict(r) for r in by_sku]})
    finally:
        conn.close()


@inv_report_bp.route('/api/inv/reports/movements', methods=['GET'])
@login_required
@inv_view_required
def movements_summary():
    """出入库汇总:按类型统计笔数与数量变动。?date_from= ?date_to= ?warehouse_id=。"""
    conn = get_conn()
    try:
        sc, sp = warehouse_scope_clause('m.warehouse_id')
        sql = '''SELECT m.movement_type, COUNT(*) AS cnt,
                        COALESCE(SUM(m.qty_delta),0) AS qty_delta,
                        COALESCE(SUM(CASE WHEN m.qty_delta>0 THEN m.qty_delta ELSE 0 END),0) AS qty_in,
                        COALESCE(SUM(CASE WHEN m.qty_delta<0 THEN -m.qty_delta ELSE 0 END),0) AS qty_out
                 FROM inv_movements m WHERE 1=1''' + sc
        params = list(sp)
        if request.args.get('warehouse_id'):
            sql += ' AND m.warehouse_id=?'; params.append(request.args.get('warehouse_id'))
        if request.args.get('date_from'):
            sql += ' AND m.ts >= ?'; params.append(request.args.get('date_from'))
        if request.args.get('date_to'):
            sql += ' AND m.ts <= ?'; params.append(request.args.get('date_to') + ' 23:59:59')
        sql += ' GROUP BY m.movement_type ORDER BY cnt DESC'
        rows = conn.execute(sql, params).fetchall()
        return jsonify([dict(r) for r in rows])
    finally:
        conn.close()


@inv_report_bp.route('/api/inv/reports/fulfillments', methods=['GET'])
@login_required
@inv_view_required
def fulfillment_stats():
    """分仓/拆单统计:整单 vs 拆单数、按仓出货分单数。"""
    conn = get_conn()
    try:
        sc, sp = warehouse_scope_clause('warehouse_id')
        total_orders = conn.execute(
            'SELECT COUNT(DISTINCT order_id) AS n FROM inv_fulfillments WHERE 1=1' + sc, sp).fetchone()['n']
        split_orders = conn.execute(
            'SELECT COUNT(*) AS n FROM (SELECT order_id FROM inv_fulfillments WHERE is_split=1'
            + sc + ' GROUP BY order_id)', sp).fetchone()['n']
        by_wh = conn.execute(
            '''SELECT f.warehouse_id, w.name AS warehouse_name, COUNT(*) AS fulfillments,
                      SUM(CASE WHEN f.is_split=1 THEN 1 ELSE 0 END) AS split_parts
               FROM inv_fulfillments f JOIN warehouses w ON w.id=f.warehouse_id
               WHERE 1=1''' + sc.replace('warehouse_id', 'f.warehouse_id')
            + ' GROUP BY f.warehouse_id ORDER BY fulfillments DESC', sp).fetchall()
        return jsonify({'total_orders': total_orders, 'split_orders': split_orders,
                        'single_orders': total_orders - split_orders,
                        'by_warehouse': [dict(r) for r in by_wh]})
    finally:
        conn.close()
