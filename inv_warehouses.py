"""
inv_warehouses.py — 模块 1:仓库管理(后台页 + API)。

职责:
  - 仓库主数据管理:增 / 改 / 软删(停用),复用现有 warehouses 表,
    用 inv_warehouse_ext 挂载库存属性(自营/合伙人、合伙人名、是否参与分仓)。
  - 市场 → 仓库优先级路由(inv_market_warehouses)的增删改查,
    未来新市场只需在 UI 加数据,不改代码。

设计要点:
  - 不与 app.py 已有的 /api/warehouses 冲突 —— 全部走 /inventory 与 /api/inv/* 命名空间。
  - 不改 warehouses 现有列;ownership/partner 等只写 inv_warehouse_ext。
  - 仓库主数据属管理员级(inv_admin_required)。
  - 删除仓库走软删(is_active=0):仓库可能被 product_costs / orders / inv_stock 引用,
    硬删会破坏历史数据,违反"禁止手删数据"。仅当无任何引用时才允许硬删。
"""

from flask import Blueprint, render_template, request, jsonify
from flask_login import login_required

from inv_common import (
    get_conn, inv_view_required, inv_admin_required,
)

inv_wh_bp = Blueprint('inv_wh', __name__)


# ───────────────────────────── 页面 ─────────────────────────────

@inv_wh_bp.route('/inventory/warehouses')
@login_required
@inv_view_required
def warehouses_page():
    return render_template('inv_warehouses.html')


# ─────────────────────────── 仓库 API ───────────────────────────

@inv_wh_bp.route('/api/inv/warehouses', methods=['GET'])
@login_required
@inv_view_required
def list_warehouses():
    """列出所有仓库 + 库存扩展属性(left join,无扩展则给默认自营)。"""
    conn = get_conn()
    rows = conn.execute('''
        SELECT w.id, w.name, w.code, w.country, w.default_currency, w.is_active, w.notes,
               COALESCE(we.ownership_type, 'self') AS ownership_type,
               we.partner_name,
               we.partner_id,
               COALESCE(we.is_fulfillment, 1) AS is_fulfillment,
               we.region,
               (SELECT COUNT(*) FROM inv_market_warehouses mw WHERE mw.warehouse_id = w.id) AS route_count
        FROM warehouses w
        LEFT JOIN inv_warehouse_ext we ON we.warehouse_id = w.id
        ORDER BY w.is_active DESC, w.country, w.name
    ''').fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@inv_wh_bp.route('/api/inv/warehouses', methods=['POST'])
@login_required
@inv_admin_required
def create_warehouse():
    """新建仓库。同时写入 warehouses(主)+ inv_warehouse_ext(库存属性)。"""
    d = request.get_json(force=True) or {}
    name = (d.get('name') or '').strip()
    code = (d.get('code') or '').strip()
    country = (d.get('country') or '').strip().upper()
    currency = (d.get('default_currency') or 'PLN').strip().upper()
    notes = (d.get('notes') or '').strip()
    ownership = d.get('ownership_type') or 'self'
    partner_name = (d.get('partner_name') or '').strip() or None
    is_fulfillment = 1 if d.get('is_fulfillment', True) else 0

    if not name or not code or not country:
        return jsonify({'error': '名称、编码、国家为必填项'}), 400
    if ownership not in ('self', 'partner'):
        return jsonify({'error': 'ownership_type 必须是 self 或 partner'}), 400
    if ownership == 'partner' and not partner_name:
        return jsonify({'error': '合伙人仓必须填写合伙人名称'}), 400

    conn = get_conn()
    try:
        cur = conn.execute(
            'INSERT INTO warehouses (name, code, country, default_currency, notes) VALUES (?,?,?,?,?)',
            (name, code, country, currency, notes))
        wid = cur.lastrowid
        conn.execute('''INSERT INTO inv_warehouse_ext
                        (warehouse_id, ownership_type, partner_name, is_fulfillment)
                        VALUES (?,?,?,?)''',
                     (wid, ownership, partner_name, is_fulfillment))
        conn.commit()
        return jsonify({'success': True, 'id': wid})
    except Exception as e:
        if 'UNIQUE' in str(e):
            return jsonify({'error': '仓库编码已存在'}), 400
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()


@inv_wh_bp.route('/api/inv/warehouses/<int:wid>', methods=['PUT'])
@login_required
@inv_admin_required
def update_warehouse(wid):
    """更新仓库主数据 + 库存扩展属性(upsert inv_warehouse_ext)。"""
    d = request.get_json(force=True) or {}
    conn = get_conn()
    try:
        wh = conn.execute('SELECT id FROM warehouses WHERE id=?', (wid,)).fetchone()
        if not wh:
            return jsonify({'error': '仓库不存在'}), 404

        ownership = d.get('ownership_type', 'self')
        partner_name = (d.get('partner_name') or '').strip() or None
        if ownership not in ('self', 'partner'):
            return jsonify({'error': 'ownership_type 必须是 self 或 partner'}), 400
        if ownership == 'partner' and not partner_name:
            return jsonify({'error': '合伙人仓必须填写合伙人名称'}), 400
        if ownership == 'self':
            partner_name = None  # 自营仓清空合伙人名

        # warehouses 主数据(只更新这几列,不触碰其它列含义)
        conn.execute('''UPDATE warehouses SET name=?, code=?, country=?, default_currency=?,
                        notes=?, is_active=?, updated_at=CURRENT_TIMESTAMP WHERE id=?''',
                     (d.get('name'), d.get('code'), (d.get('country') or '').upper(),
                      (d.get('default_currency') or 'PLN').upper(), d.get('notes', ''),
                      1 if d.get('is_active', True) else 0, wid))

        is_fulfillment = 1 if d.get('is_fulfillment', True) else 0
        conn.execute('''INSERT INTO inv_warehouse_ext
                        (warehouse_id, ownership_type, partner_name, is_fulfillment)
                        VALUES (?,?,?,?)
                        ON CONFLICT(warehouse_id) DO UPDATE SET
                            ownership_type=excluded.ownership_type,
                            partner_name=excluded.partner_name,
                            is_fulfillment=excluded.is_fulfillment,
                            updated_at=CURRENT_TIMESTAMP''',
                     (wid, ownership, partner_name, is_fulfillment))
        conn.commit()
        return jsonify({'success': True})
    except Exception as e:
        if 'UNIQUE' in str(e):
            return jsonify({'error': '仓库编码已存在'}), 400
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()


@inv_wh_bp.route('/api/inv/warehouses/<int:wid>', methods=['DELETE'])
@login_required
@inv_admin_required
def delete_warehouse(wid):
    """删除仓库。默认软删(停用 is_active=0)。

    仅当无任何引用(product_costs / orders / inv_stock / inv_movements /
    市场路由)时,带 ?hard=1 才允许硬删。否则一律软删,保护历史数据。
    """
    hard = request.args.get('hard') == '1'
    conn = get_conn()
    try:
        wh = conn.execute('SELECT id FROM warehouses WHERE id=?', (wid,)).fetchone()
        if not wh:
            return jsonify({'error': '仓库不存在'}), 404

        refs = {
            'product_costs': conn.execute('SELECT COUNT(*) FROM product_costs WHERE warehouse_id=?', (wid,)).fetchone()[0],
            'orders': conn.execute('SELECT COUNT(*) FROM orders WHERE warehouse_id=?', (wid,)).fetchone()[0],
            'inv_stock': conn.execute('SELECT COUNT(*) FROM inv_stock WHERE warehouse_id=?', (wid,)).fetchone()[0],
            'inv_movements': conn.execute('SELECT COUNT(*) FROM inv_movements WHERE warehouse_id=?', (wid,)).fetchone()[0],
            'market_routes': conn.execute('SELECT COUNT(*) FROM inv_market_warehouses WHERE warehouse_id=?', (wid,)).fetchone()[0],
        }
        total_refs = sum(refs.values())

        if hard and total_refs == 0:
            conn.execute('DELETE FROM inv_warehouse_ext WHERE warehouse_id=?', (wid,))
            conn.execute('DELETE FROM warehouses WHERE id=?', (wid,))
            conn.commit()
            return jsonify({'success': True, 'mode': 'hard'})

        # 软删:停用
        conn.execute('UPDATE warehouses SET is_active=0, updated_at=CURRENT_TIMESTAMP WHERE id=?', (wid,))
        conn.commit()
        return jsonify({'success': True, 'mode': 'soft', 'refs': refs})
    finally:
        conn.close()


# ──────────────────────── 市场 → 仓 路由 API ────────────────────────

@inv_wh_bp.route('/api/inv/market-routes', methods=['GET'])
@login_required
@inv_view_required
def list_market_routes():
    """按市场分组返回路由(含仓库名/自营合伙人标记)。可选 ?market=CZ 过滤。"""
    market = request.args.get('market')
    conn = get_conn()
    sql = '''
        SELECT mw.id, mw.market_code, mw.warehouse_id, mw.priority, mw.is_active, mw.notes,
               w.name AS warehouse_name, w.code AS warehouse_code, w.country AS warehouse_country,
               COALESCE(we.ownership_type,'self') AS ownership_type, we.partner_name
        FROM inv_market_warehouses mw
        JOIN warehouses w ON w.id = mw.warehouse_id
        LEFT JOIN inv_warehouse_ext we ON we.warehouse_id = w.id
    '''
    params = ()
    if market:
        sql += ' WHERE mw.market_code = ?'
        params = (market.upper(),)
    sql += ' ORDER BY mw.market_code, mw.priority, mw.id'
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@inv_wh_bp.route('/api/inv/market-routes', methods=['POST'])
@login_required
@inv_admin_required
def create_market_route():
    d = request.get_json(force=True) or {}
    market = (d.get('market_code') or '').strip().upper()
    warehouse_id = d.get('warehouse_id')
    priority = d.get('priority')
    notes = (d.get('notes') or '').strip()
    if not market or not warehouse_id:
        return jsonify({'error': '市场代码和仓库为必填项'}), 400
    try:
        priority = int(priority) if priority is not None else 100
    except (TypeError, ValueError):
        return jsonify({'error': '优先级必须是整数'}), 400

    conn = get_conn()
    try:
        if not conn.execute('SELECT 1 FROM warehouses WHERE id=?', (warehouse_id,)).fetchone():
            return jsonify({'error': '仓库不存在'}), 404
        conn.execute('''INSERT INTO inv_market_warehouses (market_code, warehouse_id, priority, notes)
                        VALUES (?,?,?,?)''', (market, warehouse_id, priority, notes))
        conn.commit()
        return jsonify({'success': True, 'id': conn.execute('SELECT last_insert_rowid()').fetchone()[0]})
    except Exception as e:
        if 'UNIQUE' in str(e):
            return jsonify({'error': '该市场已配置此仓库,请直接调整其优先级'}), 400
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()


@inv_wh_bp.route('/api/inv/market-routes/<int:rid>', methods=['PUT'])
@login_required
@inv_admin_required
def update_market_route(rid):
    d = request.get_json(force=True) or {}
    conn = get_conn()
    try:
        r = conn.execute('SELECT id FROM inv_market_warehouses WHERE id=?', (rid,)).fetchone()
        if not r:
            return jsonify({'error': '路由不存在'}), 404
        priority = d.get('priority')
        try:
            priority = int(priority)
        except (TypeError, ValueError):
            return jsonify({'error': '优先级必须是整数'}), 400
        conn.execute('''UPDATE inv_market_warehouses SET priority=?, is_active=?, notes=?,
                        updated_at=CURRENT_TIMESTAMP WHERE id=?''',
                     (priority, 1 if d.get('is_active', True) else 0, d.get('notes', ''), rid))
        conn.commit()
        return jsonify({'success': True})
    finally:
        conn.close()


@inv_wh_bp.route('/api/inv/market-routes/<int:rid>', methods=['DELETE'])
@login_required
@inv_admin_required
def delete_market_route(rid):
    conn = get_conn()
    try:
        conn.execute('DELETE FROM inv_market_warehouses WHERE id=?', (rid,))
        conn.commit()
        return jsonify({'success': True})
    finally:
        conn.close()
