"""Product identities from catalogues and product-only order evidence.

Reading never creates products or stock. New master records are materialized
only inside the mapping confirmation transaction.
"""
from collections import defaultdict
import html
import json
import re
import unicodedata

from product_recognition import parse_product_name
from stock_sync_common import SyncError, digest, dumps, exists, one, rows


def key(value):
    return re.sub(r'\s+', ' ', unicodedata.normalize('NFKC', html.unescape(str(value or '')))).strip().casefold()


def words(value):
    value = key(value)
    value = re.sub(r'(\d)\s*[- ]?in\s*[- ]?(\d)', r'\1in\2', value)
    return ' '.join(re.findall(r'[^\W_]+', value, re.UNICODE))


def flavor_key(value):
    return words(re.sub(r'\bon[ _-]+ice\b', 'ice', key(value)))


def model_key(brand, model):
    value = words(model)
    return re.sub(r'^randm\s+', '', value) if key(brand) == 'fumot' else value


def is_flavor_attribute(value):
    label = ''.join(ch for ch in unicodedata.normalize('NFKD', key(value)) if not unicodedata.combining(ch)).replace('_', ' ')
    return bool(re.search(r'smak|flavo[u]?r|口味|taste|aroma|\b(?:iz|izcsoport|prichut)\b', label))


PUFFS = re.compile(r'\b(\d{1,3}(?:[\s.,]\d{3})+|\d+)\+?\s*(?:puffs?|zaciągnięć|口)\b|\b(\d+)\s*k\b', re.I)


def profile(name, sku, attributes, rules, tax=None, *, _allow_sku=True):
    """Require a named brand, model, puff count and full flavor, never a score."""
    name = html.unescape(str(name or ''))
    parsed = parse_product_name(name, rules['brands'], rules['series'])
    tax = tax or {}
    sku_profile = None
    sku_text = re.sub(r'[-_]+', ' ', str(sku or ''))
    sku_puff = PUFFS.search(sku_text) or re.search(r'\b(\d{3,6})\b', sku_text)
    if _allow_sku and sku_puff and sku_text[sku_puff.end():].strip():
        tail = re.sub(r'^dual smak\s+', '', sku_text[sku_puff.end():].strip(), flags=re.I)
        sku_profile = profile(sku_text[:sku_puff.end()]+' - '+tail,
                              '', [], rules, _allow_sku=False)
    def missing():
        if not sku_profile:
            return None
        brand = tax.get('brand_id') or parsed.get('brand_id')
        puffs = tax.get('puff_count') or parsed.get('puffs')
        flavors = {flavor_key(v) for n,v in attributes if is_flavor_attribute(n)}
        if any(not is_flavor_attribute(n) for n,v in attributes):
            return None
        if ((brand and brand != sku_profile['brand_id']) or (puffs and puffs != sku_profile['puff_count']) or
                (flavors and flavors != {sku_profile['flavor_key']})):
            return None
        return dict(sku_profile)
    brand_id = tax.get('brand_id') or parsed.get('brand_id')
    brand = next((b for b in rules['brands'] if b['id'] == brand_id), None)
    if not brand:
        return missing()
    puff = PUFFS.search(name)
    bare_puff = re.search(r'\b(\d{3,6})\b', name)
    puff_count = tax.get('puff_count') or parsed.get('puffs')
    if not puff_count and puff:
        puff_count = int(re.sub(r'[\s.,]', '', puff.group(1))) if puff.group(1) else int(puff.group(2))*1000
    if bare_puff and (not puff or (bare_puff.start() < puff.start() and int(bare_puff.group(1)) == puff_count)):
        puff = bare_puff
        puff_count = puff_count or int(bare_puff.group(1))
    if not puff_count:
        return missing()
    series = next((s for s in rules['series'] if s['id'] == tax.get('series_id') and s['brand_id'] == brand_id), None)
    prefix = name[:puff.start()] if puff else ''
    model = ''
    brand_matched = False
    for alias in dict.fromkeys([brand['name']] + sorted(brand.get('patterns', []), key=len)):
        pattern = r'[\s_-]*'.join(re.escape(part) for part in alias.split())
        match = re.search(r'(?<!\w)' + pattern + r'(?!\w)', prefix, re.I)
        if match:
            model = prefix[match.end():]
            brand_matched = True
            break
    # Product-type and marketing words are not part of the model.
    model = re.split(r'\b(?:jednorazowy|disposable|e[ -]?papieros|vape|dwusmakowe)\b', model, flags=re.I)[0].strip(' -–|/:')
    if series:
        model = series['name']
    model_from_brand = not model and brand_matched
    if model_from_brand:
        model = brand['name']
    if not model or len(model) > 65 or not re.search(r'[a-zA-Z]', model):
        return missing()
    model_identity = '@brand' if model_from_brand else model_key(brand['name'], model)
    flavors = {flavor_key(v): v for n, v in attributes if is_flavor_attribute(n)}
    if len(flavors) > 1:
        return None
    flavor = next(iter(flavors.values())) if flavors else tax.get('flavor') or parsed.get('flavor')
    if not flavor or len(str(flavor)) > 600:
        return missing()
    flavor = html.unescape(str(flavor)).strip()
    # Do not mistake a title's marketing suffix for a simple product flavor.
    if not flavors and re.search(r'\b(?:bestseller|ekran|screen|mesh|wykończenie|papieros|disposable)\b', flavor, re.I):
        return None
    quantity = 1
    quantity_matches = re.findall(r'\b(?:pack of|box of|opakowanie)\s*(\d+)\b|\b(\d+)\s*(?:pcs|szt\.?|pieces|支装)\b', name, re.I)
    quantities = {int(a or b) for a, b in quantity_matches}
    if len(quantities) > 1 or (quantities and not 1 <= next(iter(quantities)) <= 100000):
        return None
    if quantities:
        quantity = next(iter(quantities))
    strength = set()
    for number, unit in re.findall(r'\b(\d+(?:[.,]\d+)?)\s*(%|mg(?:/ml)?)', name, re.I):
        value = float(number.replace(',', '.'))
        strength.add(round(value * 10 if unit == '%' else value, 4))
    if len(strength) > 1:
        return None
    for attr, value in attributes:
        if is_flavor_attribute(attr):
            continue
        if re.search(r'nicot|nikot|尼古丁|strength', attr, re.I):
            values = re.findall(r'(\d+(?:[.,]\d+)?)\s*(%|mg(?:/ml)?)', value, re.I)
            if not values:
                return None
            for number, unit in values:
                n = float(number.replace(',', '.')); strength.add(round(n*10 if unit == '%' else n, 4))
        elif re.search(r'pack|ilość|quantity|包装|数量', attr, re.I):
            if not re.fullmatch(r'\s*\d+\s*(?:pcs|szt|支)?\s*', value, re.I):
                return None
            quantity = int(re.search(r'\d+', value)[0])
        else:
            # Color, resistance and other unknown variant dimensions matter.
            return None
    if len(strength) > 1 or not 1 <= quantity <= 100000:
        return None
    flavor = re.sub(r'\b(?:pack of|box of|opakowanie)\s*\d+\b|\b\d+\s*(?:pcs|szt\.?|pieces|支装)\b', '', flavor, flags=re.I).strip(' -()/')
    if not flavor:
        return None
    result = {'brand_id': brand_id, 'brand': brand['name'], 'model': model,
              'model_key': model_identity, 'model_from_brand': model_from_brand,
              'series_id': series['id'] if series else parsed.get('series_id'),
              'puff_count': int(puff_count), 'flavor': flavor, 'flavor_key': flavor_key(flavor),
              'nicotine': next(iter(strength)) if strength else None, 'quantity': quantity}
    result['base_key'] = digest([brand_id, model_identity, int(puff_count), flavor_key(flavor)])
    sku_same_flavor = sku_profile and re.sub(r'\W', '', result['flavor_key']) == re.sub(r'\W', '', sku_profile['flavor_key'])
    if sku_profile and (result['brand_id'] != sku_profile['brand_id'] or result['model_key'] != sku_profile['model_key'] or result['puff_count'] != sku_profile['puff_count'] or not sku_same_flavor):
        result['sku_conflict'] = True
    return result


def compatible(a, b):
    return bool(a and b and a['base_key'] == b['base_key'] and
                (a['nicotine'] is None or b['nicotine'] is None or a['nicotine'] == b['nicotine']))


def classification_rules(c):
    brands = rows(c, 'SELECT * FROM brands') if exists(c, 'brands') else []
    series = rows(c, 'SELECT * FROM series') if exists(c, 'series') else []
    for brand in brands:
        try:
            aliases = json.loads(brand.get('aliases') or '[]')
        except (ValueError, TypeError):
            aliases = []
        brand['patterns'] = [brand['name'].upper()] + [a.upper() for a in aliases if isinstance(a, str)] if isinstance(aliases, list) else [brand['name'].upper()]
    return {'brands': brands, 'series': series}


def order_products(c, sources):
    """Only item identity fields; no customer, address, price, or order IDs.

    Use all distinct product evidence in the most recent 50,000 authorized
    orders. The complete current Woo catalogue covers products without orders.
    """
    if not sources or not exists(c, 'orders'):
        return []
    marks = ','.join('?' for _ in sources)
    recent = f'SELECT source,line_items FROM orders WHERE source IN ({marks}) ORDER BY date_created DESC LIMIT 50000'
    if hasattr(c, '_raw'):
        sql = f'''SELECT DISTINCT o.source,li->>'name' AS name,li->>'sku' AS sku,
                  li->>'product_id' AS product_id,li->>'variation_id' AS variation_id
                  FROM ({recent}) o CROSS JOIN LATERAL jsonb_array_elements(
                    COALESCE(NULLIF(o.line_items,''),'[]')::jsonb) li'''
    else:
        sql = f'''SELECT DISTINCT o.source,json_extract(li.value,'$.name') AS name,
                  json_extract(li.value,'$.sku') AS sku,json_extract(li.value,'$.product_id') AS product_id,
                  json_extract(li.value,'$.variation_id') AS variation_id
                  FROM ({recent}) o,json_each(CASE WHEN json_valid(o.line_items) THEN o.line_items ELSE '[]' END) li'''
    result = []
    for item in rows(c, sql, tuple(sorted(sources))):
        try:
            item['product_id'] = int(item['product_id']); item['variation_id'] = int(item['variation_id'] or 0)
        except (TypeError, ValueError):
            continue
        if item['product_id'] > 0 and isinstance(item['name'], str):
            result.append(item)
    return result


def manual_sources(c):
    if not exists(c, 'oms_warehouse_integrations'):
        return []
    # SELECT wi.* also supports older read-only fixtures without is_enabled.
    result = rows(c, '''SELECT wi.*,w.name,w.country,w.is_active FROM oms_warehouse_integrations wi
                       JOIN warehouses w ON w.id=wi.warehouse_id WHERE wi.inventory_authority='manual_partner' ''')
    markets = rows(c, 'SELECT * FROM inv_market_warehouses WHERE is_active=1')
    for w in result:
        w['markets'] = sorted({m['market_code'] for m in markets if m['warehouse_id'] == w['warehouse_id']})
    return [w for w in result if w['is_active'] and w.get('is_enabled', True)]


def context(c, skus, rules, sites, authorized_sources):
    orders = order_products(c, authorized_sources)
    by_identity, by_sku = defaultdict(list), defaultdict(list)
    for order in orders:
        order['profile'] = profile(order['name'], order['sku'], [], rules)
        by_identity[(order['source'], order['product_id'], order['variation_id'])].append(order)
        if order['sku']:
            by_sku[key(order['sku'])].append(order)
    sku_profiles = {s['id']: profile(s['name'], s['sku_code'], [], rules, s.get('recognition')) for s in skus}
    manual = manual_sources(c)
    manual_ids = {w['warehouse_id'] for w in manual}
    links = rows(c, 'SELECT * FROM oms_sku_warehouses WHERE is_enabled=1')
    groups = defaultdict(list)
    blocked = set()
    for table in ('inv_stock', 'inv_movements', 'oms_external_stock'):
        if exists(c, table):
            blocked.update(r['sku_id'] for r in rows(c, f'SELECT DISTINCT sku_id FROM {table}'))
    for s in skus:
        p = sku_profiles[s['id']]
        supply = {l['warehouse_id'] for l in links if l['sku_id'] == s['id']}
        if (p and not p.get('sku_conflict') and p['quantity'] == 1 and s['sku_code'].startswith(('MANUAL-PL-', 'CAT-')) and supply and
                supply <= manual_ids and s['id'] not in blocked and (s.get('unit') or 'pcs') == 'pcs'):
            groups[(p['base_key'], p['nicotine'], tuple(sorted(supply)))].append(s['id'])
    canonical = {}
    for ids in groups.values():
        if len(ids) > 1:
            canonical.update({i: min(ids) for i in ids})
    protected = set()
    if exists(c, 'stock_sync_bindings'):
        protected.update(r['map_id'] for r in rows(c, 'SELECT map_id FROM stock_sync_bindings'))
        protected.update(r['map_id'] for r in rows(c, 'SELECT map_id FROM stock_sync_controls WHERE active=1'))
    return {'orders': orders, 'by_identity': by_identity, 'by_sku': by_sku,
            'sku_profiles': sku_profiles, 'canonical': canonical, 'manual_sources': manual,
            'protected': protected, 'sites': {s['id']: s for s in sites}}


def enrich(items, skus, rules, aliases, ctx):
    """Auto-select exact identities; preserve uncertain and protected records."""
    sku_by_id = {s['id']: s for s in skus}
    profiles = {}
    for item in items:
        site = ctx['sites'][item['site_id']]
        evidence = ctx['by_identity'].get((site['url'], item['product_id'], item['variation_id']), [])
        if not evidence and item['wc_sku']:
            evidence = ctx['by_sku'].get(key(item['wc_sku']), [])
        p = profile(item['name'], item['wc_sku'], item['attributes'], rules, item.get('recognition'))
        if p:
            item['product_identity'] = p
        known = [e['profile'] for e in evidence if e['profile']]
        if any(v.get('sku_conflict') for v in known):
            if item['state'] != 'mapped':
                item.update(state='review', certain=False, proposed_sku_id=None,
                            explanation='历史订单的名称与 SKU 有冲突，请核对')
            continue
        if p and any(not compatible(p, v) for v in known):
            # A reused generic WC SKU cannot override an independently complete
            # variation identity. Same product/variation disagreements do block.
            direct = ctx['by_identity'].get((site['url'], item['product_id'], item['variation_id']), [])
            if direct:
                if item['state'] != 'mapped':
                    item.update(state='review', certain=False, proposed_sku_id=None,
                                explanation='订单记录与当前商品的型号、口数、口味或规格不一致，请核对')
                continue
            evidence = []; known = []
        if not p and known and len({(v['base_key'],v['nicotine'],v['quantity']) for v in known}) == 1:
            # Apply current variant dimensions to the recovered title as well.
            # Unknown color/resistance attributes must not vanish in fallback.
            recovered = next(e for e in evidence if e['profile'])
            p = profile(recovered['name'], item['wc_sku'], item['attributes'], rules)
            if not p:
                continue
            # Only backfill title/model gaps, never overwrite a conflicting
            # current variation attribute or a different brand/puff count.
            current = item.get('recognition') or {}
            if ((current.get('brand_id') and current['brand_id'] != p['brand_id']) or
                (current.get('puff_count') and current['puff_count'] != p['puff_count']) or
                (current.get('flavor') and flavor_key(current['flavor']) != p['flavor_key'])):
                continue
        if not p:
            continue
        if p.get('sku_conflict') and item['state'] != 'mapped':
            item.update(state='review', certain=False, proposed_sku_id=None,
                        explanation='名称与 SKU 中的品牌、型号、口数或口味不一致，请核对')
            continue
        profiles[item['id']] = p
        item['product_identity'] = p
        item['order_evidence'] = [{k: e[k] for k in ('source','product_id','variation_id','name','sku')} for e in evidence]
        item['evidence_label'] = '订单商品与目录交叉识别' if known else '站点商品属性与名称识别'

    strengths = defaultdict(set)
    for p in list(profiles.values()) + [p for p in ctx['sku_profiles'].values() if p]:
        if p['nicotine'] is not None:
            strengths[p['base_key']].add(p['nicotine'])
    for item in items:
        p = profiles.get(item['id'])
        if not p or item['state'] in ('conflict', 'unsupported'):
            continue
        if len(strengths[p['base_key']]) > 1 and p['nicotine'] is None:
            if item['state'] != 'mapped':
                item.update(state='review', certain=False, proposed_sku_id=None, explanation='同款存在不同尼古丁规格，当前商品未注明规格')
            continue
        existing = aliases.get((item['site_id'], item['product_id'], item['variation_id']), [])
        if existing:
            m = existing[0]; canonical = ctx['canonical'].get(m['sku_id'], m['sku_id'])
            if (len(existing) == 1 and m['is_active'] and canonical != m['sku_id'] and
                    m['id'] not in ctx['protected'] and compatible(p, ctx['sku_profiles'].get(canonical))):
                item.update(state='align', certain=True, proposed_sku_id=canonical, expected_mapping=dict(m),
                            explanation='同一商品存在多个人工供货临时 SKU；统一当前映射，保留原 SKU 和历史履约记录')
            continue
        # Do not override an explicit classification/attribute conflict.
        if item['state'] == 'review' and not item.get('candidates'):
            continue
        matching = [s for s in skus if compatible(p, ctx['sku_profiles'][s['id']])]
        if any(ctx['sku_profiles'][s['id']].get('sku_conflict') for s in matching):
            item.update(state='review', certain=False, proposed_sku_id=None,
                        candidates=[s['id'] for s in matching], explanation='已有 SKU 主档的名称和编码不一致，请核对')
            continue
        if any((s.get('unit') or 'pcs') != 'pcs' or ctx['sku_profiles'][s['id']]['quantity'] != 1 for s in matching):
            item.update(state='review', certain=False, proposed_sku_id=None, candidates=[s['id'] for s in matching],
                        explanation='已有主档包含整盒或其他包装单位，请核对每件折合数量')
            continue
        candidates = [s['id'] for s in matching]
        if len(strengths[p['base_key']]) > 1 and any(ctx['sku_profiles'][i]['nicotine'] is None for i in candidates):
            item.update(state='review', certain=False, proposed_sku_id=None, candidates=candidates,
                        explanation='同款存在多种尼古丁规格，已有 SKU 未注明规格，请核对主档')
            continue
        candidates = sorted({ctx['canonical'].get(s, s) for s in candidates})
        if item.get('proposed_sku_id') and item['proposed_sku_id'] not in candidates:
            # Exact opaque identifiers still need to agree with any readable
            # identity. Keep legacy suggestions without a parseable SKU name.
            other = ctx['sku_profiles'].get(item['proposed_sku_id'])
            if other and not compatible(p, other):
                item.update(state='review', certain=False, proposed_sku_id=None, explanation='SKU 标识与商品身份不一致，请核对')
                continue
            if not other:
                continue
        if len(candidates) == 1:
            item.update(state='suggested', certain=True, proposed_sku_id=candidates[0], candidates=candidates,
                        qty_per_item=p['quantity'], explanation=item['evidence_label']+'：品牌、型号、口数和口味一致，已自动选择')
        elif len(candidates) > 1:
            item.update(state='review', certain=False, proposed_sku_id=None, candidates=candidates,
                        explanation='存在多个产品主档或不同供货来源，请人工选择；有库存或业务记录的主档不会自动合并')
        else:
            possible_aliases = []
            for s in skus:
                other = ctx['sku_profiles'][s['id']]
                if not other or other['puff_count'] != p['puff_count'] or other['flavor_key'] != p['flavor_key']:
                    continue
                a = words(p['brand']+' '+p['model']).replace(' ', '')
                b = words(other['brand']+' '+other['model']).replace(' ', '')
                if (other['brand_id'] != p['brand_id'] and
                    ((other['model_key'] != '@brand' and other['model_key'].replace(' ', '') in a) or
                     (p['model_key'] != '@brand' and p['model_key'].replace(' ', '') in b))):
                    possible_aliases.append(s['id'])
            if possible_aliases:
                item.update(state='review', certain=False, proposed_sku_id=None, candidates=possible_aliases,
                            explanation='品牌称呼与现有 SKU 可能是别名，请核对，避免将数量仓商品重复建为人工供货商品')
                continue
            site = ctx['sites'][item['site_id']]
            supply = [w for w in ctx['manual_sources'] if site.get('country') in w['markets']]
            if len(supply) != 1:
                item.update(state='review', certain=False, proposed_sku_id=None,
                            explanation='产品已识别，尚无唯一的人工供货仓可用于补建，请先配置供货来源或选择已有 SKU')
                continue
            nicotine = p['nicotine']
            if nicotine is None and strengths[p['base_key']]:
                nicotine = next(iter(strengths[p['base_key']]))
            identity_key = digest([p['base_key'], nicotine])
            family = p['brand'] if p.get('model_from_brand') else p['brand']+' '+p['model_key'].title()
            name = f"{family} {p['puff_count']} Puffs - {p['flavor_key']}"
            if nicotine is not None:
                name += f' ({nicotine:g}mg/ml)'
            proposal = dict(p, nicotine=nicotine, identity_key=identity_key, sku_code='CAT-'+identity_key[:24].upper(),
                            name=name, warehouse_id=supply[0]['warehouse_id'], warehouse_name=supply[0]['name'])
            item.update(state='new', certain=True, proposed_sku_id=None, new_sku=proposal,
                        qty_per_item=p['quantity'], explanation=item['evidence_label']+'：已识别，保存时自动补建统一 SKU；人工供货，不创建库存数量')
    return items


def create_master(c, proposal, actor_id, confirmation_id):
    """Called with the master lock held; no stock or order writes."""
    previous = one(c, 'SELECT * FROM inv_skus WHERE sku_code=?', (proposal['sku_code'],))
    if previous:
        if not previous['is_active'] or key(previous['name']) != key(proposal['name']):
            raise SyncError('MAPPING_INPUT_CHANGED', '自动产品编码已有不同记录，请重新读取')
        if not one(c, 'SELECT sku_id FROM oms_sku_warehouses WHERE sku_id=? AND warehouse_id=?', (previous['id'], proposal['warehouse_id'])):
            c.execute('INSERT INTO oms_sku_warehouses(sku_id,warehouse_id,is_enabled) VALUES(?,?,1)', (previous['id'], proposal['warehouse_id']))
        return previous['id']
    series_id = proposal.get('series_id')
    if not series_id and not proposal.get('model_from_brand'):
        same = [s for s in rows(c, 'SELECT id,name FROM series WHERE brand_id=?', (proposal['brand_id'],)) if model_key(proposal['brand'], s['name']) == proposal['model_key']]
        if len(same) > 1:
            raise SyncError('MAPPING_INPUT_CHANGED', '系列主档存在重复，请先处理')
        if same:
            series_id = same[0]['id']
        else:
            series_id = c.execute('INSERT INTO series(brand_id,name) VALUES(?,?) RETURNING id',
                                  (proposal['brand_id'], proposal['model'])).fetchone()['id']
    notes = dumps({'product_identity': proposal['identity_key'], 'model': proposal['model'],
                   'source': 'stock_sync_mapping', 'confirmation_id': confirmation_id, 'actor_id': actor_id})
    sku_id = c.execute('''INSERT INTO inv_skus(sku_code,name,brand_id,series_id,puff_count,flavor,unit,is_active,notes)
                         VALUES(?,?,?,?,?,?,'pcs',1,?) RETURNING id''',
                       (proposal['sku_code'], proposal['name'], proposal['brand_id'], series_id,
                        proposal['puff_count'], proposal['flavor'], notes)).fetchone()['id']
    c.execute('INSERT INTO oms_sku_warehouses(sku_id,warehouse_id,is_enabled) VALUES(?,?,1)',
              (sku_id, proposal['warehouse_id']))
    return sku_id
