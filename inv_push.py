"""
inv_push.py — 模块 7:库存可用量定时下推 WordPress + 对账。

模型(对齐锁定架构):库存真源=本系统。算出每仓"可用 = 现存 - 预留",
按市场把可用汇总下推到该市场各站点的 WC 商品库存。

某站点某 WC 商品(→SKU)的可发布数量:
    publishable = floor( Σ(serving warehouses) max(0, on_hand - reserved) / qty_per_item )
serving warehouses = 该站点市场(site.country)配置的路由仓;未配置则取该国家活跃仓。

共享库存说明:同一仓库被同国多个站点共享,各站会显示同一可用量。这是有意为之
(架构要求下推可用量),靠"高频下推 + 下单即预留(reserved 立减可用)"自我纠偏,
容忍瞬时超卖。

下推复用现有 Product Manager 的 PUT 白名单:
    app._build_product_update_payload(白名单校验) + app.get_product_api_endpoint(端点路由)
    + 同样的 req.put(/wp-json/wc/v3/products/<id>) + app._parse_wc_response。
绝不另起一套绕过白名单的写法。

注意:实际 PUT 会写到线上站点。所有写操作默认 dry_run=True(只算不推),
真正下推需显式 dry_run=False(UI 按钮 / cron 显式开启)。
"""

import json
from flask import Blueprint, render_template, request, jsonify
from flask_login import login_required

from inv_common import get_conn, inv_view_required, inv_admin_required, current_operator
import inv_allocator

inv_push_bp = Blueprint('inv_push', __name__)


# ───────────────────────── 计算可发布库存 ─────────────────────────

def _serving_warehouses(conn, market):
    """该市场的服务仓 id 列表。配置了路由→路由仓;否则→该国家活跃仓。"""
    cands = inv_allocator.candidate_warehouses(conn, market)
    if cands:
        return [c['warehouse_id'] for c in cands]
    rows = conn.execute("SELECT id FROM warehouses WHERE country=? AND is_active=1", (market,)).fetchall()
    return [r['id'] for r in rows]


def compute_site_stock(conn, site_id):
    """计算某站点所有已映射商品的可发布库存。返回 list(dict)。"""
    site = conn.execute('SELECT id, url, country FROM sites WHERE id=?', (site_id,)).fetchone()
    if not site:
        return []
    market = (site['country'] or '').upper()
    wh_ids = _serving_warehouses(conn, market)
    maps = conn.execute('''SELECT m.*, k.sku_code, k.name AS sku_name
                           FROM inv_site_sku_map m JOIN inv_skus k ON k.id=m.sku_id
                           WHERE m.site_id=? AND m.is_active=1''', (site_id,)).fetchall()
    out = []
    for m in maps:
        avail_sku = 0
        if wh_ids:
            q = ','.join('?' * len(wh_ids))
            r = conn.execute(
                f'SELECT COALESCE(SUM(MAX(on_hand - reserved, 0)),0) AS a '
                f'FROM inv_stock WHERE sku_id=? AND warehouse_id IN ({q})',
                [m['sku_id']] + wh_ids).fetchone()
            avail_sku = max(0, r['a'] or 0)
        qpi = m['qty_per_item'] or 1
        publishable = avail_sku // qpi
        out.append({
            'site_id': site_id, 'source': site['url'], 'market': market,
            'wc_product_id': m['wc_product_id'], 'wc_variation_id': m['wc_variation_id'] or 0,
            'sku_id': m['sku_id'], 'sku_code': m['sku_code'], 'sku_name': m['sku_name'],
            'qty_per_item': qpi, 'available_sku': avail_sku, 'publishable': publishable,
            'serving_warehouses': wh_ids,
        })
    return out


# ───────────────────────── 下推 / 对账 ─────────────────────────

def _put_stock(api_url, ck, cs, product_id, qty):
    """复用 Product Manager 的白名单 PUT 把单个商品库存写到 WC。返回 (ok, error)。"""
    import requests as req
    from app import _build_product_update_payload, _parse_wc_response
    payload, err = _build_product_update_payload({'manage_stock': True, 'stock_quantity': qty})
    if err:
        return False, err
    try:
        resp = req.put(f'{api_url}/wp-json/wc/v3/products/{product_id}',
                       auth=(ck, cs), json=payload, timeout=90,
                       headers={'User-Agent': 'WooCommerce API Client-Python/3.0.0',
                                'Content-Type': 'application/json', 'Accept': 'application/json'})
    except req.RequestException as e:
        return False, f'连接失败: {e}'
    _p, err = _parse_wc_response(resp)
    return (err is None), err


def push_site(conn, site_id, dry_run=True, only_changed=True):
    """把某站点的可发布库存下推到 WC。dry_run=True 只算不推。

    only_changed=True 时,仅推送与 WC 现值不同的(需要 GET 现值,对账模式);
    为避免大量 GET,默认走"全量推送当前可发布值"(only_changed 在此实现里仅控制
    是否记录 prev_qty,不做差异跳过——真实 PUT 幂等,WC 端相同值不变)。
    """
    uid, uname = current_operator()
    items = compute_site_stock(conn, site_id)
    site = conn.execute('SELECT id, url, consumer_key, consumer_secret, product_master_id FROM sites WHERE id=?',
                        (site_id,)).fetchone()
    result = {'site_id': site_id, 'source': (site['url'] if site else None),
              'dry_run': dry_run, 'total': len(items), 'ok': 0, 'error': 0, 'items': []}

    api_url = ck = cs = None
    if not dry_run:
        from app import get_product_api_endpoint
        api_url, ck, cs = get_product_api_endpoint(conn, site)
        if not (api_url and ck and cs):
            result['error'] = len(items)
            result['fatal'] = '站点缺少可用的商品 API 凭据'
            return result

    for it in items:
        status, err = 'dry', None
        if not dry_run:
            ok, err = _put_stock(api_url, ck, cs, it['wc_product_id'], it['publishable'])
            status = 'ok' if ok else 'error'
            if ok:
                result['ok'] += 1
            else:
                result['error'] += 1
        conn.execute('''INSERT INTO inv_push_logs
            (site_id, source, wc_product_id, wc_variation_id, sku_id, pushed_qty, status, error, operator_id, operator_name)
            VALUES (?,?,?,?,?,?,?,?,?,?)''',
            (site_id, it['source'], it['wc_product_id'], it['wc_variation_id'], it['sku_id'],
             it['publishable'], status, err, uid, uname))
        result['items'].append({**{k: it[k] for k in ('wc_product_id', 'sku_code', 'publishable')},
                                'status': status, 'error': err})
    conn.commit()
    return result


# ───────────────────────────── API ─────────────────────────────

@inv_push_bp.route('/inventory/push')
@login_required
@inv_view_required
def push_page():
    conn = get_conn()
    sites = conn.execute('SELECT id, url, country FROM sites ORDER BY country, url').fetchall()
    conn.close()
    return render_template('inv_push.html', sites=[dict(s) for s in sites])


@inv_push_bp.route('/api/inv/site-stock/<int:site_id>', methods=['GET'])
@login_required
@inv_view_required
def api_site_stock(site_id):
    """对账视图:某站点各商品的本系统可发布库存(不写库、不连 WC)。"""
    conn = get_conn()
    try:
        return jsonify(compute_site_stock(conn, site_id))
    finally:
        conn.close()


@inv_push_bp.route('/api/inv/push/<int:site_id>', methods=['POST'])
@login_required
@inv_admin_required
def api_push_site(site_id):
    """下推某站点库存。body: {dry_run: true|false}。默认 dry_run=true 防误推。"""
    d = request.get_json(silent=True) or {}
    dry = d.get('dry_run', True)
    conn = get_conn()
    try:
        return jsonify(push_site(conn, site_id, dry_run=bool(dry)))
    finally:
        conn.close()


@inv_push_bp.route('/api/inv/push-logs', methods=['GET'])
@login_required
@inv_view_required
def api_push_logs():
    site_id = request.args.get('site_id')
    try:
        limit = min(2000, int(request.args.get('limit') or 200))
    except ValueError:
        limit = 200
    conn = get_conn()
    sql = '''SELECT pl.*, k.sku_code FROM inv_push_logs pl
             LEFT JOIN inv_skus k ON k.id=pl.sku_id WHERE 1=1'''
    params = []
    if site_id:
        sql += ' AND pl.site_id=?'; params.append(site_id)
    sql += ' ORDER BY pl.id DESC LIMIT ?'; params.append(limit)
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])
