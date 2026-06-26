"""
inv_resolver.py — 把订单 line_items 解析到 SKU 的核心逻辑。

这是模块 5(卖出扣减)与模块 6(分仓引擎)依赖的底座:给定一个站点的
WooCommerce line_item,确定它对应本系统的哪个 inv_skus,以及折合多少个 SKU 单位。

解析优先级(从最确定到最模糊):
  1. 站点精确映射  (site_id, wc_product_id, wc_variation_id)  —— 人工/学习到的精确对应
  2. 站点产品级映射 (site_id, wc_product_id, 0)               —— 该商品无变体或变体不区分库存
  3. WC SKU 字符串  line_items.sku → inv_site_sku_map.wc_sku / inv_skus.sku_code / barcode
  4. 品牌+系列+口数+口味 taxonomy —— 复用现有商品名解析(parse_product_name /
     _resolve_product_to_brand),落到带相同 taxonomy 的 inv_skus。
  5. 无法解析 → 返回 None(上层据此报"未映射")。

为避免与 22k 行的 app.py 循环导入,对 app 内解析函数一律**懒加载**(函数内 import)。
"""

import json


# ───────────────────────── 缓存构建 ─────────────────────────

def _load_brands_cache(conn):
    """构造 parse_product_name 期望的 brands_cache 结构。"""
    rows = conn.execute('SELECT id, name, aliases FROM brands').fetchall()
    cache = []
    for r in rows:
        aliases = []
        if r['aliases']:
            try:
                aliases = json.loads(r['aliases'])
            except Exception:
                aliases = []
        cache.append({
            'id': r['id'], 'name': r['name'], 'aliases': aliases,
            'patterns': [r['name'].upper()] + [a.upper() for a in aliases],
        })
    return cache


def _load_pm_cache(conn):
    """product_mappings 缓存,键 = (normalized_raw_name, source)。复用 app.normalize_raw_name。"""
    from app import normalize_raw_name
    rows = conn.execute(
        'SELECT raw_name, source, brand_id, series_id, puff_count, flavor FROM product_mappings'
    ).fetchall()
    cache = {}
    for pm in rows:
        cache[(normalize_raw_name(pm['raw_name']), pm['source'])] = dict(pm)
    return cache


def _tax_key(brand_id, series_id, puff_count, flavor):
    """taxonomy 归一化键。flavor 大小写/空白不敏感;None 与空串等价。"""
    fl = (flavor or '').strip().upper()
    return (brand_id, series_id or None, puff_count or None, fl or None)


def _load_sku_indexes(conn):
    """构建 SKU 侧索引:按 sku_code、barcode、taxonomy 各建一份。"""
    rows = conn.execute(
        'SELECT id, sku_code, barcode, brand_id, series_id, puff_count, flavor, name '
        'FROM inv_skus WHERE is_active = 1'
    ).fetchall()
    by_code, by_barcode, by_tax = {}, {}, {}
    for r in rows:
        if r['sku_code']:
            by_code[r['sku_code'].strip().upper()] = r['id']
        if r['barcode']:
            by_barcode[r['barcode'].strip().upper()] = r['id']
        # taxonomy 索引只在 brand_id 存在时建,避免 (None,...) 误命中
        if r['brand_id']:
            by_tax.setdefault(_tax_key(r['brand_id'], r['series_id'], r['puff_count'], r['flavor']), r['id'])
    return {'by_code': by_code, 'by_barcode': by_barcode, 'by_tax': by_tax}


def _load_site_map(conn):
    """站点映射索引:键 (site_id, wc_product_id, wc_variation_id) 与 wc_sku。"""
    rows = conn.execute(
        'SELECT site_id, wc_product_id, wc_variation_id, wc_sku, sku_id, qty_per_item '
        'FROM inv_site_sku_map WHERE is_active = 1'
    ).fetchall()
    by_pv, by_sku = {}, {}
    for r in rows:
        by_pv[(r['site_id'], r['wc_product_id'], r['wc_variation_id'] or 0)] = r
        if r['wc_sku']:
            by_sku[(r['site_id'], r['wc_sku'].strip().upper())] = r
    return {'by_pv': by_pv, 'by_sku': by_sku}


def build_caches(conn):
    """一次性构建解析所需的全部缓存。批量解析(扫描/对账)时复用,避免 N 次查库。"""
    return {
        'brands': _load_brands_cache(conn),
        'pm': _load_pm_cache(conn),
        'sku': _load_sku_indexes(conn),
        'site_map': _load_site_map(conn),
    }


# ───────────────────────── 站点 id 解析 ─────────────────────────

def site_id_for_source(conn, source):
    """source(站点 url)→ sites.id。线上 orders 用 source 存站点 url。"""
    r = conn.execute('SELECT id FROM sites WHERE url = ?', (source,)).fetchone()
    return r['id'] if r else None


# ───────────────────────── 核心解析 ─────────────────────────

def resolve_line_item(conn, site_id, source, item, caches=None):
    """把单个 line_item 解析到 SKU。

    item: WooCommerce line_items 里的一条 dict(含 product_id/variation_id/sku/name/quantity)。
    返回 dict:
      {
        'sku_id': int|None, 'sku_code': str|None, 'qty_per_item': int,
        'qty': int,            # 该行 WC 数量
        'sku_qty': int,        # 折合的 SKU 总数 = qty * qty_per_item
        'match_via': str,      # exact|product|wc_sku|taxonomy|unmatched
        'name': str,
      }
    """
    if caches is None:
        caches = build_caches(conn)

    name = (item.get('name') or item.get('parent_name') or '').strip()
    wc_pid = item.get('product_id')
    wc_vid = item.get('variation_id') or 0
    wc_sku = (item.get('sku') or '').strip()
    qty = int(item.get('quantity') or 0)

    sku_id = None
    qty_per_item = 1
    match_via = 'unmatched'

    sm = caches['site_map']

    # 1. 精确(product + variation)
    hit = sm['by_pv'].get((site_id, wc_pid, wc_vid))
    # 2. 产品级(variation=0)
    if not hit and wc_vid:
        hit = sm['by_pv'].get((site_id, wc_pid, 0))
    if hit:
        sku_id = hit['sku_id']
        qty_per_item = hit['qty_per_item'] or 1
        match_via = 'exact' if (site_id, wc_pid, wc_vid) in sm['by_pv'] else 'product'

    # 3. WC SKU 字符串
    if not sku_id and wc_sku:
        hit = sm['by_sku'].get((site_id, wc_sku.upper()))
        if hit:
            sku_id, qty_per_item, match_via = hit['sku_id'], (hit['qty_per_item'] or 1), 'wc_sku'
        else:
            key = wc_sku.upper()
            sku_id = caches['sku']['by_code'].get(key) or caches['sku']['by_barcode'].get(key)
            if sku_id:
                match_via = 'wc_sku'

    # 4. taxonomy(品牌/系列/口数/口味)
    if not sku_id and name:
        brand_id, series_id, puff, flavor = resolve_taxonomy(conn, name, source, caches)
        if brand_id:
            sku_id = caches['sku']['by_tax'].get(_tax_key(brand_id, series_id, puff, flavor))
            if sku_id:
                match_via = 'taxonomy'

    sku_code = None
    if sku_id:
        # 仅在需要展示时回查 code(命中通常很少,代价可忽略)
        rc = conn.execute('SELECT sku_code FROM inv_skus WHERE id=?', (sku_id,)).fetchone()
        sku_code = rc['sku_code'] if rc else None

    return {
        'sku_id': sku_id, 'sku_code': sku_code, 'qty_per_item': qty_per_item,
        'qty': qty, 'sku_qty': qty * (qty_per_item or 1),
        'match_via': match_via, 'name': name,
    }


def resolve_taxonomy(conn, name, source, caches=None):
    """商品名 → (brand_id, series_id, puff_count, flavor)。复用 app 的解析逻辑。"""
    from app import _resolve_product_to_brand
    if caches is None:
        caches = {'brands': _load_brands_cache(conn), 'pm': _load_pm_cache(conn)}
    return _resolve_product_to_brand(name, source, caches['brands'], caches['pm'])


def resolve_order(conn, order_row, caches=None):
    """解析整张订单的 line_items。order_row 需含 id/source/line_items。

    返回 {'order_id', 'source', 'site_id', 'lines':[...], 'all_matched':bool,
          'matched':int, 'unmatched':int}。
    """
    if caches is None:
        caches = build_caches(conn)
    source = order_row['source']
    site_id = site_id_for_source(conn, source)
    try:
        items = json.loads(order_row['line_items'] or '[]')
    except Exception:
        items = []

    lines, matched, unmatched = [], 0, 0
    for it in items:
        res = resolve_line_item(conn, site_id, source, it, caches)
        lines.append(res)
        if res['sku_id']:
            matched += 1
        else:
            unmatched += 1
    return {
        'order_id': order_row['id'], 'source': source, 'site_id': site_id,
        'lines': lines, 'matched': matched, 'unmatched': unmatched,
        'all_matched': unmatched == 0 and matched > 0,
    }
