"""
inv_skus.py — 模块 2:SKU 主档 + 站点商品映射(后台页 + API)。

职责:
  - SKU 主档(inv_skus)增删改查;SKU 可携带品牌/系列/口数/口味 taxonomy,
    以便与现有商品名解析(brands/series/product_mappings)对接、做 taxonomy 兜底匹配。
  - 站点商品映射(inv_site_sku_map):WC 商品(product+variation 或 wc_sku)↔ SKU,
    支持 bundle(qty_per_item:一件 = N 个 SKU)。
  - 解析建议:输入商品名,返回解析出的 taxonomy + 候选 SKU(辅助建档)。
  - 未映射商品发现:扫描最近订单 line_items,找出尚不能解析到 SKU 的商品。
  - 订单解析预览:对单张订单展示每行命中哪个 SKU、经哪条路径、是否有未映射。

写操作权限:can_manage_inventory(inv_manage_required)。
"""

import json
from flask import Blueprint, render_template, request, jsonify
from flask_login import login_required

from inv_common import get_conn, inv_view_required, inv_manage_required
import inv_resolver

inv_sku_bp = Blueprint('inv_sku', __name__)


# ───────────────────────────── 页面 ─────────────────────────────

@inv_sku_bp.route('/inventory/skus')
@login_required
@inv_view_required
def skus_page():
    conn = get_conn()
    sites = conn.execute('SELECT id, url, country, manager FROM sites ORDER BY country, url').fetchall()
    brands = conn.execute('SELECT id, name FROM brands ORDER BY name').fetchall()
    conn.close()
    return render_template('inv_skus.html',
                           sites=[dict(s) for s in sites],
                           brands=[dict(b) for b in brands])


# ─────────────────────────── SKU 主档 API ───────────────────────────

@inv_sku_bp.route('/api/inv/skus', methods=['GET'])
@login_required
@inv_view_required
def list_skus():
    """列出 SKU(可 ?q= 模糊搜 code/name/barcode;?active=1 仅启用)。"""
    q = (request.args.get('q') or '').strip()
    active = request.args.get('active')
    conn = get_conn()
    sql = '''SELECT s.*, b.name AS brand_name, se.name AS series_name,
                    (SELECT COUNT(*) FROM inv_site_sku_map m WHERE m.sku_id=s.id) AS map_count
             FROM inv_skus s
             LEFT JOIN brands b ON b.id = s.brand_id
             LEFT JOIN series se ON se.id = s.series_id WHERE 1=1'''
    params = []
    if q:
        sql += ' AND (s.sku_code LIKE ? OR s.name LIKE ? OR s.barcode LIKE ?)'
        params += [f'%{q}%', f'%{q}%', f'%{q}%']
    if active == '1':
        sql += ' AND s.is_active = 1'
    sql += ' ORDER BY s.is_active DESC, s.sku_code'
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


def _sku_payload(d):
    return (
        (d.get('sku_code') or '').strip(),
        (d.get('name') or '').strip(),
        d.get('brand_id') or None,
        d.get('series_id') or None,
        d.get('puff_count') or None,
        (d.get('flavor') or '').strip() or None,
        (d.get('barcode') or '').strip() or None,
        (d.get('unit') or 'pcs').strip(),
        d.get('shelf_life_days') or None,
        (d.get('notes') or '').strip() or None,
    )


@inv_sku_bp.route('/api/inv/skus', methods=['POST'])
@login_required
@inv_manage_required
def create_sku():
    d = request.get_json(force=True) or {}
    vals = _sku_payload(d)
    if not vals[0] or not vals[1]:
        return jsonify({'error': 'SKU 编码与名称为必填项'}), 400
    conn = get_conn()
    try:
        cur = conn.execute('''INSERT INTO inv_skus
            (sku_code, name, brand_id, series_id, puff_count, flavor, barcode, unit, shelf_life_days, notes)
            VALUES (?,?,?,?,?,?,?,?,?,?)''', vals)
        conn.commit()
        return jsonify({'success': True, 'id': cur.lastrowid})
    except Exception as e:
        if 'UNIQUE' in str(e):
            return jsonify({'error': 'SKU 编码已存在'}), 400
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()


@inv_sku_bp.route('/api/inv/skus/<int:sid>', methods=['PUT'])
@login_required
@inv_manage_required
def update_sku(sid):
    d = request.get_json(force=True) or {}
    vals = _sku_payload(d)
    if not vals[0] or not vals[1]:
        return jsonify({'error': 'SKU 编码与名称为必填项'}), 400
    is_active = 1 if d.get('is_active', True) else 0
    conn = get_conn()
    try:
        if not conn.execute('SELECT 1 FROM inv_skus WHERE id=?', (sid,)).fetchone():
            return jsonify({'error': 'SKU 不存在'}), 404
        conn.execute('''UPDATE inv_skus SET sku_code=?, name=?, brand_id=?, series_id=?, puff_count=?,
                        flavor=?, barcode=?, unit=?, shelf_life_days=?, notes=?, is_active=?,
                        updated_at=CURRENT_TIMESTAMP WHERE id=?''', vals + (is_active, sid))
        conn.commit()
        return jsonify({'success': True})
    except Exception as e:
        if 'UNIQUE' in str(e):
            return jsonify({'error': 'SKU 编码已存在'}), 400
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()


@inv_sku_bp.route('/api/inv/skus/<int:sid>', methods=['DELETE'])
@login_required
@inv_manage_required
def delete_sku(sid):
    """删除 SKU。有库存/流水/映射引用时改为停用(软删),保护历史数据。"""
    conn = get_conn()
    try:
        if not conn.execute('SELECT 1 FROM inv_skus WHERE id=?', (sid,)).fetchone():
            return jsonify({'error': 'SKU 不存在'}), 404
        refs = (conn.execute('SELECT COUNT(*) FROM inv_stock WHERE sku_id=?', (sid,)).fetchone()[0]
                + conn.execute('SELECT COUNT(*) FROM inv_movements WHERE sku_id=?', (sid,)).fetchone()[0]
                + conn.execute('SELECT COUNT(*) FROM inv_site_sku_map WHERE sku_id=?', (sid,)).fetchone()[0])
        if refs == 0:
            conn.execute('DELETE FROM inv_skus WHERE id=?', (sid,))
            conn.commit()
            return jsonify({'success': True, 'mode': 'hard'})
        conn.execute('UPDATE inv_skus SET is_active=0, updated_at=CURRENT_TIMESTAMP WHERE id=?', (sid,))
        conn.commit()
        return jsonify({'success': True, 'mode': 'soft'})
    finally:
        conn.close()


# ─────────────────────────── 站点映射 API ───────────────────────────

@inv_sku_bp.route('/api/inv/sku-maps', methods=['GET'])
@login_required
@inv_view_required
def list_sku_maps():
    """列出站点映射。可 ?sku_id= 或 ?site_id= 过滤。"""
    sku_id = request.args.get('sku_id')
    site_id = request.args.get('site_id')
    conn = get_conn()
    sql = '''SELECT m.*, s.url AS site_url, s.country AS site_country,
                    k.sku_code, k.name AS sku_name
             FROM inv_site_sku_map m
             JOIN sites s ON s.id = m.site_id
             JOIN inv_skus k ON k.id = m.sku_id WHERE 1=1'''
    params = []
    if sku_id:
        sql += ' AND m.sku_id=?'; params.append(sku_id)
    if site_id:
        sql += ' AND m.site_id=?'; params.append(site_id)
    sql += ' ORDER BY s.country, s.url, m.wc_product_id'
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@inv_sku_bp.route('/api/inv/sku-maps', methods=['POST'])
@login_required
@inv_manage_required
def create_sku_map():
    d = request.get_json(force=True) or {}
    site_id = d.get('site_id')
    sku_id = d.get('sku_id')
    wc_product_id = d.get('wc_product_id') or None
    wc_variation_id = d.get('wc_variation_id') or 0
    wc_sku = (d.get('wc_sku') or '').strip() or None
    raw_name = (d.get('raw_name') or '').strip() or None
    qty_per_item = d.get('qty_per_item') or 1
    if not site_id or not sku_id:
        return jsonify({'error': '站点与 SKU 为必填项'}), 400
    if not wc_product_id and not wc_sku:
        return jsonify({'error': '须填写 WC 商品ID 或 WC SKU 字符串之一'}), 400
    try:
        qty_per_item = max(1, int(qty_per_item))
    except (TypeError, ValueError):
        return jsonify({'error': '每件折合数量必须是正整数'}), 400
    conn = get_conn()
    try:
        conn.execute('''INSERT INTO inv_site_sku_map
            (site_id, wc_product_id, wc_variation_id, wc_sku, raw_name, sku_id, qty_per_item)
            VALUES (?,?,?,?,?,?,?)''',
            (site_id, wc_product_id, wc_variation_id, wc_sku, raw_name, sku_id, qty_per_item))
        conn.commit()
        return jsonify({'success': True, 'id': conn.execute('SELECT last_insert_rowid()').fetchone()[0]})
    except Exception as e:
        if 'UNIQUE' in str(e):
            return jsonify({'error': '该站点的此商品(product+variation)已映射'}), 400
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()


@inv_sku_bp.route('/api/inv/sku-maps/<int:mid>', methods=['PUT'])
@login_required
@inv_manage_required
def update_sku_map(mid):
    d = request.get_json(force=True) or {}
    conn = get_conn()
    try:
        if not conn.execute('SELECT 1 FROM inv_site_sku_map WHERE id=?', (mid,)).fetchone():
            return jsonify({'error': '映射不存在'}), 404
        qty = max(1, int(d.get('qty_per_item') or 1))
        conn.execute('''UPDATE inv_site_sku_map SET sku_id=?, qty_per_item=?, wc_sku=?, raw_name=?,
                        is_active=?, updated_at=CURRENT_TIMESTAMP WHERE id=?''',
                     (d.get('sku_id'), qty, (d.get('wc_sku') or '').strip() or None,
                      (d.get('raw_name') or '').strip() or None,
                      1 if d.get('is_active', True) else 0, mid))
        conn.commit()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()


@inv_sku_bp.route('/api/inv/sku-maps/<int:mid>', methods=['DELETE'])
@login_required
@inv_manage_required
def delete_sku_map(mid):
    conn = get_conn()
    try:
        conn.execute('DELETE FROM inv_site_sku_map WHERE id=?', (mid,))
        conn.commit()
        return jsonify({'success': True})
    finally:
        conn.close()


# ─────────────────────────── 解析辅助 API ───────────────────────────

@inv_sku_bp.route('/api/inv/parse-suggest', methods=['GET'])
@login_required
@inv_view_required
def parse_suggest():
    """输入 ?name=&source= ,返回解析出的 taxonomy + 候选 SKU(辅助建档/映射)。"""
    name = (request.args.get('name') or '').strip()
    source = (request.args.get('source') or '').strip()
    if not name:
        return jsonify({'error': 'name 必填'}), 400
    conn = get_conn()
    try:
        brand_id, series_id, puff, flavor = inv_resolver.resolve_taxonomy(conn, name, source)
        cand = []
        if brand_id:
            rows = conn.execute('''SELECT id, sku_code, name FROM inv_skus
                WHERE is_active=1 AND brand_id=?
                  AND (? IS NULL OR series_id=? OR series_id IS NULL)
                  AND (? IS NULL OR puff_count=? OR puff_count IS NULL)
                ORDER BY sku_code LIMIT 20''',
                (brand_id, series_id, series_id, puff, puff)).fetchall()
            cand = [dict(r) for r in rows]
        brand_name = None
        if brand_id:
            b = conn.execute('SELECT name FROM brands WHERE id=?', (brand_id,)).fetchone()
            brand_name = b['name'] if b else None
        return jsonify({
            'taxonomy': {'brand_id': brand_id, 'brand_name': brand_name,
                         'series_id': series_id, 'puff_count': puff, 'flavor': flavor},
            'candidates': cand,
        })
    finally:
        conn.close()


@inv_sku_bp.route('/api/inv/resolve-order/<path:order_id>', methods=['GET'])
@login_required
@inv_view_required
def resolve_order_preview(order_id):
    """解析单张订单的 line_items → SKU,展示每行命中情况(预览,不写库存)。"""
    conn = get_conn()
    try:
        o = conn.execute('SELECT id, source, line_items FROM orders WHERE id=?', (order_id,)).fetchone()
        if not o:
            return jsonify({'error': '订单不存在'}), 404
        return jsonify(inv_resolver.resolve_order(conn, o))
    finally:
        conn.close()


@inv_sku_bp.route('/api/inv/unmapped-products', methods=['GET'])
@login_required
@inv_view_required
def unmapped_products():
    """扫描最近订单 line_items,聚合出无法解析到 SKU 的商品(供建档/映射)。

    参数:?limit= 扫描订单数(默认 800)、?site_id= 限定站点。
    返回按 (站点, 商品名) 聚合的未映射列表 + 出现次数。
    """
    try:
        limit = min(5000, int(request.args.get('limit') or 800))
    except ValueError:
        limit = 800
    site_id = request.args.get('site_id')
    conn = get_conn()
    try:
        caches = inv_resolver.build_caches(conn)
        sql = ("SELECT id, source, line_items FROM orders "
               "WHERE line_items IS NOT NULL AND line_items != '' AND line_items != '[]'")
        params = []
        if site_id:
            s = conn.execute('SELECT url FROM sites WHERE id=?', (site_id,)).fetchone()
            if s:
                sql += ' AND source=?'; params.append(s['url'])
        sql += ' ORDER BY id DESC LIMIT ?'; params.append(limit)
        orders = conn.execute(sql, params).fetchall()

        # 站点 url -> id 映射(避免每行查库)
        site_ids = {r['url']: r['id'] for r in conn.execute('SELECT id, url FROM sites').fetchall()}
        agg = {}
        scanned = 0
        for o in orders:
            scanned += 1
            sid = site_ids.get(o['source'])
            try:
                items = json.loads(o['line_items'] or '[]')
            except Exception:
                continue
            for it in items:
                res = inv_resolver.resolve_line_item(conn, sid, o['source'], it, caches)
                if res['sku_id']:
                    continue
                name = res['name'] or '(无名)'
                key = (o['source'], name)
                if key not in agg:
                    agg[key] = {'source': o['source'], 'site_id': sid, 'name': name,
                                'wc_product_id': it.get('product_id'),
                                'wc_sku': it.get('sku'), 'count': 0, 'qty': 0,
                                'example_order': o['id']}
                agg[key]['count'] += 1
                agg[key]['qty'] += int(it.get('quantity') or 0)
        out = sorted(agg.values(), key=lambda x: -x['count'])
        return jsonify({'scanned_orders': scanned, 'unmapped_count': len(out), 'items': out[:300]})
    finally:
        conn.close()
