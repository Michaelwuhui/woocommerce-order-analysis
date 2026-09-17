"""Reviewed catalogue-to-SKU bindings. All WooCommerce operations are GETs.

Uses the existing durable snapshot/work/event tables. Suggestions never become
bindings until explicitly selected; confirmations reread identity in the worker.
"""
from collections import Counter
import html
import logging
import re
import unicodedata

from stock_sync_common import (SyncError, begin, digest, dumps, event, exists,
    loads, lock_clause, now, one, parse_time, rows, stamp, uid)
from stock_sync_permissions import actor, can_write, reference_site, require_object, target_sites
from stock_sync_catalog import enqueue
from stock_sync_woo import Woo, site_identity
from product_recognition import parse_product_name


def permitted(u):
    return bool(u['superadmin'] or u.get('can_manage_inventory'))


def require_manager(u):
    if not permitted(u):
        raise SyncError('MAPPING_PERMISSION_REQUIRED', '补齐映射需要库存管理权限', 403)


def key(value):
    # Preserve punctuation and non-Latin characters: SKU-A is not SKU/A.
    return re.sub(r'\s+', ' ', unicodedata.normalize('NFKC', html.unescape(str(value or '')))).strip().casefold()


def sku_recognition(sku, rules):
    """Recognize missing legacy SKU metadata for suggestions, without saving it."""
    value = {f: sku.get(f) for f in ('brand_id', 'series_id', 'puff_count', 'flavor')}
    parsed = parse_product_name(sku['name'], rules['brands'], rules['series'])
    if value['brand_id'] and parsed.get('brand_id') and value['brand_id'] != parsed['brand_id']:
        return value
    for field, source in (('brand_id', 'brand_id'), ('series_id', 'series_id'),
                          ('puff_count', 'puffs'), ('flavor', 'flavor')):
        if value[field] in (None, ''):
            value[field] = parsed.get(source)
    return value


def master_data(c):
    skus = rows(c, 'SELECT * FROM inv_skus WHERE is_active=1 ORDER BY id')
    rules = rows(c, 'SELECT * FROM product_mappings ORDER BY id') if exists(c, 'product_mappings') else []
    fields = ('id','sku_code','name','barcode','brand_id','series_id','puff_count','flavor','unit','is_active')
    skus = [{f: s.get(f) for f in fields} for s in skus]
    rules = [{f: r.get(f) for f in ('id','raw_name','source','brand_id','series_id','puff_count','flavor')} for r in rules]
    brands = rows(c, 'SELECT id,name,aliases FROM brands ORDER BY id') if exists(c, 'brands') else []
    series = rows(c, 'SELECT id,brand_id,name FROM series ORDER BY id') if exists(c, 'series') else []
    for brand in brands:
        try:
            aliases = loads(brand['aliases'], [])
        except (TypeError, ValueError):
            aliases = []
        if not isinstance(aliases, list):
            aliases = []
        brand['patterns'] = [str(brand['name']).upper()] + [a.upper() for a in aliases if isinstance(a, str)]
    rules = {'mappings':rules, 'brands':brands, 'series':series}
    for sku in skus:
        sku['recognition'] = sku_recognition(sku, rules)
    return skus, rules, digest([skus, rules])


def identity(site_id, parent, leaf):
    variation = leaf['id'] if leaf['id'] != parent['id'] else 0
    attrs = sorted([(str(a.get('name') or a.get('slug') or ''), str(a.get('option') or ''))
                    for a in leaf.get('attributes', []) if a.get('option')])
    parent_name = html.unescape(str(parent.get('name') or ''))
    name = parent_name + (' / ' + ', '.join(v for _, v in attrs) if variation else '')
    result = {'id': f'{site_id}:{parent["id"]}:{variation}', 'site_id': site_id,
        'product_id': parent['id'], 'variation_id': variation, 'parent_name': parent_name,
        'name': name, 'wc_sku': str(leaf.get('sku') or ''),
        'barcode': str(leaf.get('global_unique_id') or ''), 'attributes': attrs,
        'status': leaf.get('status'), 'parent_status': parent.get('status'),
        'type': parent.get('type')}
    result['identity_hash'] = digest(result)
    return result


def recognition(item, site, rules):
    names = [item['name']]
    if item['variation_id']:
        names.append(item['parent_name'] + ' - ' + ' - '.join(v for _, v in item['attributes']))
    names.append(item['parent_name'])
    found = []
    for name in dict.fromkeys(names):
        exact = [r for r in rules['mappings'] if key(r['raw_name']) == key(name)]
        found = [r for r in exact if r['source'] == site['url']]
        if not found:
            found = [r for r in exact if not r['source']]
        if found:
            break
    values = [{f: r.get(f) for f in ('brand_id','series_id','puff_count','flavor')} for r in found]
    if len({dumps(v) for v in values}) > 1:
        return {}, '已有产品识别规则相互冲突'
    if values:
        value = values[0]
    else:
        parsed = parse_product_name(item['parent_name'] if item['variation_id'] else item['name'], rules['brands'], rules['series'])
        value = {'brand_id':parsed.get('brand_id'), 'series_id':parsed.get('series_id'),
                 'puff_count':parsed.get('puffs'), 'flavor':parsed.get('flavor')}
    flavors = {key(v): v for n, v in item['attributes']
               if any(t in key(n) for t in ('smak','flavor','flavour','口味','taste','aroma'))}
    if len(flavors) > 1:
        return {}, '站点包含多个不同的口味属性，请人工核对'
    if len(flavors) == 1:
        flavor = next(iter(flavors.values()))
        # Stored classification is explicit evidence; parsed title suffixes may
        # be marketing copy. Variation attributes take precedence over parsing.
        if values and value.get('flavor') and key(value['flavor']) != key(flavor):
            return {}, '已有识别口味与站点变体口味不同'
        value = dict(value, flavor=flavor)
    return value, ''


def suggest(item, site, skus, rules, aliases, sku_counts):
    existing = aliases.get((item['site_id'], item['product_id'], item['variation_id']), [])
    result = dict(item, candidates=[], proposed_sku_id=None, qty_per_item=1,
                  certain=False, state='unmatched', explanation='未找到现有 SKU，请手动选择；没有主档的商品需先建立 SKU')
    if existing:
        m = existing[0]
        result.update(state='mapped' if len(existing) == 1 and m['is_active'] else 'conflict',
            proposed_sku_id=m['sku_id'], qty_per_item=m['qty_per_item'],
            explanation='已有映射，保持原记录' if len(existing) == 1 and m['is_active'] else '已有停用或冲突映射，请在映射管理中处理')
        return result
    if item['status'] != 'publish' or item['parent_status'] != 'publish' or item['type'] not in ('simple','variable'):
        result.update(state='unsupported', explanation='仅为已发布的简单商品和变体生成新映射')
        return result
    tax, warning = recognition(item, site, rules)
    result['recognition'] = tax
    if warning:
        result.update(state='review', explanation=warning)
        return result
    identifiers = {key(v) for v in (item['wc_sku'], item['barcode']) if key(v)}
    direct = [s for s in skus if identifiers.intersection({key(s['sku_code']), key(s['barcode'])})]
    named = [s for s in skus if key(s['name']) in {key(item['name']), key(item['parent_name'] + ' - ' + ' - '.join(v for _, v in item['attributes']))}]
    classified = []
    if tax.get('brand_id') and tax.get('puff_count') and tax.get('flavor'):
        classified = [s for s in skus if s['recognition']['brand_id'] == tax['brand_id'] and s['recognition']['puff_count'] == tax['puff_count']
                      and key(s['recognition']['flavor']) == key(tax['flavor'])
                      and (not tax.get('series_id') or s['recognition']['series_id'] == tax['series_id'])]
    chosen = direct or named or classified
    method = 'SKU / 条码精确匹配' if direct else ('完整名称精确匹配' if named else '复用已有品牌、系列、口数和口味识别')
    # Conflicting independent evidence must never be silently resolved by priority.
    other = named or classified
    disagreement = bool(direct and other and {s['id'] for s in direct} != {s['id'] for s in other})
    if disagreement:
        chosen = list({s['id']: s for s in direct + other}.values())
        method = 'SKU 标识与完整名称或已有产品识别结果不一致'
    flavor_mismatch = bool(direct and tax.get('flavor') and any(s['recognition'].get('flavor') and key(s['recognition']['flavor']) != key(tax['flavor']) for s in direct))
    if flavor_mismatch:
        method = 'SKU 标识与口味属性不一致，请人工核对'
    result['candidates'] = [s['id'] for s in chosen]
    if chosen:
        unique = len(chosen) == 1 and not flavor_mismatch
        duplicate = bool(item['wc_sku'] and sku_counts[(item['site_id'], key(item['wc_sku']))] > 1)
        result.update(state='suggested' if unique else 'review', proposed_sku_id=chosen[0]['id'] if unique else None,
            certain=bool(unique and (direct or named) and not duplicate),
            explanation=method + ('；该站多个商品共用此 WC SKU，请逐项核对口味' if duplicate else '')
            + ('；存在多个候选，请人工选择' if not unique else '；请核对规格和每件折合数量'))
    return result


def create(c, u, data):
    require_manager(u)
    targets = target_sites(c, u, data.get('target_scope') or {})
    source = data.get('source_site_id')
    if source is not None and (type(source) is not int or source < 1):
        raise SyncError('INVALID_INPUT', status=400)
    if source:
        reference_site(c, u, source)
        targets = [s for s in targets if s['id'] != source]
    if not targets:
        raise SyncError('EMPTY_TARGET', '请先勾选目标站点', 400)
    scope = {'purpose': 'mapping_assist', 'mode': 'explicit_sites', 'site_ids': [s['id'] for s in targets]}
    id_ = uid()
    c.execute('''INSERT INTO stock_sync_catalog_snapshots(id,actor_id,site_id,scope_json,status,created_at,expires_at)
                 VALUES(?,?,?,?,'queued',?,?)''', (id_, u['id'], source, dumps(scope), stamp(), stamp(3600)))
    enqueue(c, 'mapping', id_)
    event(c, 'mapping_scan_requested', u['id'], id_, {'targets':scope['site_ids'], 'source':source})
    c.commit()
    return id_


def accessible(c, u, id_, purpose='mapping_assist'):
    require_manager(u)
    snap = one(c, 'SELECT * FROM stock_sync_catalog_snapshots WHERE id=?', (id_,))
    if not snap or loads(snap['scope_json']).get('purpose') != purpose:
        raise SyncError('NOT_FOUND', status=404)
    scope = loads(snap['scope_json'])
    require_object(c, u, snap, scope['site_ids'], snap['site_id'])
    return snap, scope


def status(c, u, id_, confirmation=False):
    snap, scope = accessible(c, u, id_, 'mapping_confirmation' if confirmation else 'mapping_assist')
    items = loads(snap['items_json'], [])
    sites = {s['id']: s for s in rows(c, 'SELECT * FROM sites')}
    for item in items:
        item.pop('identity_hash', None)
        item['writable'] = can_write(u, sites[item['site_id']])
        item['site_url'] = sites[item['site_id']]['url']
    result = {k: snap[k] for k in ('id','status','complete','progress','error','expires_at')}
    result.update(items=items, counts=dict(Counter(i['state'] for i in items)),
                  site_ids=scope['site_ids'], source_site_id=snap['site_id'])
    if not confirmation:
        result['skus'] = master_data(c)[0]
    return result


def scan(c, id_, woo=None):
    woo = woo or Woo(); woo.connection = c
    snap = one(c, 'SELECT * FROM stock_sync_catalog_snapshots WHERE id=?', (id_,))
    scope = loads(snap['scope_json'])
    items = []
    try:
        u = actor(c, snap['actor_id']); accessible(c, u, id_)
        sites = target_sites(c, u, scope)
        if snap['site_id']:
            sites.insert(0, reference_site(c, u, snap['site_id']))
        skus, rules, fingerprint = master_data(c)
        scope.update(master_hash=fingerprint, site_hashes={str(s['id']): site_identity(s) for s in sites})
        for site in sites:
            for page in woo.pages(site, 'products'):
                accessible(c, actor(c, snap['actor_id']), id_)
                for product in page:
                    if product.get('type') == 'variable':
                        leaves = [leaf for batch in woo.pages(site, f'products/{product["id"]}/variations') for leaf in batch]
                        if set(product.get('variations', [])) != {v['id'] for v in leaves}:
                            raise SyncError('SOURCE_INCOMPLETE', '商品口味列表在读取期间发生变化')
                        items.extend(identity(site['id'], product, leaf) for leaf in leaves)
                    else:
                        items.append(identity(site['id'], product, product))
                c.execute("UPDATE stock_sync_catalog_snapshots SET status='scanning',progress=? WHERE id=?", (len(items), id_)); c.commit()
        if len({i['id'] for i in items}) != len(items):
            raise SyncError('SOURCE_INCOMPLETE', '商品目录包含重复记录')
        aliases = {}
        for m in rows(c, 'SELECT * FROM inv_site_sku_map'):
            aliases.setdefault((m['site_id'], m['wc_product_id'], m.get('wc_variation_id') or 0), []).append(m)
        counts = Counter((i['site_id'], key(i['wc_sku'])) for i in items if i['wc_sku'])
        by_site = {s['id']: s for s in sites}
        items = [suggest(i, by_site[i['site_id']], skus, rules, aliases, counts) for i in items]
        if master_data(c)[2] != fingerprint:
            raise SyncError('MAPPING_INPUT_CHANGED', 'SKU 主档或产品识别规则已改变，请重新读取')
        c.execute("UPDATE stock_sync_catalog_snapshots SET status='complete',complete=1,items_json=?,scope_json=?,progress=?,observed_at=? WHERE id=?",
                  (dumps(items), dumps(scope), len(items), stamp(), id_)); c.commit()
    except Exception as exc:
        c.rollback()
        logging.getLogger(__name__).warning('Mapping scan %s failed: %s', id_, type(exc).__name__)
        c.execute("UPDATE stock_sync_catalog_snapshots SET status='failed',error=?,complete=0 WHERE id=?",
                  (str(exc) if isinstance(exc, SyncError) else '读取映射建议失败，请重新读取', id_)); c.commit()


def confirm(c, u, id_, data):
    snap, scope = accessible(c, u, id_)
    if not snap['complete'] or snap['status'] != 'complete':
        raise SyncError('SOURCE_INCOMPLETE', '目录未完整读取，不能保存映射')
    if parse_time(snap['expires_at']) <= now():
        raise SyncError('CATALOG_EXPIRED', '映射建议已过期，请重新读取')
    if data.get('reviewed') is not True:
        raise SyncError('MAPPING_REVIEW_REQUIRED', '请确认已核对所选商品的口味、规格和每件折合数量', 400)
    reason = data.get('reason')
    if not isinstance(reason, str) or not reason.strip() or len(reason) > 500:
        raise SyncError('REASON_REQUIRED', '请填写操作原因（最多 500 字）', 400)
    choices = data.get('items')
    if not isinstance(choices, list) or not choices or len(choices) > 1000:
        raise SyncError('INVALID_SELECTION', '请选择 1 至 1000 项映射', 400)
    skus, _, fingerprint = master_data(c)
    if fingerprint != scope['master_hash']:
        raise SyncError('MAPPING_INPUT_CHANGED', 'SKU 主档或产品识别规则已改变，请重新读取')
    sku_ids = {s['id'] for s in skus}
    original = {i['id']: i for i in loads(snap['items_json'], [])}
    approved, seen = [], set()
    for choice in choices:
        if not isinstance(choice, dict) or not isinstance(choice.get('id'), str):
            raise SyncError('INVALID_SELECTION', status=400)
        item = original.get(choice['id'])
        if not item or item['id'] in seen or item['state'] not in ('unmatched','suggested','review'):
            raise SyncError('INVALID_SELECTION', '选择项不属于待处理映射', 400)
        target_sites(c, u, {'mode':'explicit_sites', 'site_ids':[item['site_id']]})
        sku, quantity = choice.get('sku_id'), choice.get('qty_per_item')
        if type(sku) is not int or sku not in sku_ids or type(quantity) is not int or not 1 <= quantity <= 100000:
            raise SyncError('INVALID_INPUT', '请选择有效 SKU，并填写正整数的每件折合数量', 400)
        approved.append(dict(item, chosen_sku_id=sku, chosen_quantity=quantity, state='pending'))
        seen.add(item['id'])
    # Same immutable selection/reason is idempotent even after a response is lost.
    selection_hash = digest([id_, sorted([(i['id'], i['chosen_sku_id'], i['chosen_quantity']) for i in approved]), reason.strip()])
    confirmation_id = digest(['mapping_confirmation', u['id'], selection_hash])[:32]
    begin(c)
    c.execute('SELECT id FROM stock_sync_catalog_snapshots WHERE id=?'+lock_clause(c), (id_,)).fetchone()
    previous = one(c, 'SELECT id FROM stock_sync_catalog_snapshots WHERE id=?', (confirmation_id,))
    if previous:
        c.commit()
        return confirmation_id
    new_scope = dict(scope, purpose='mapping_confirmation', parent_id=id_, reason=reason.strip())
    c.execute('''INSERT INTO stock_sync_catalog_snapshots(id,actor_id,site_id,scope_json,status,items_json,created_at,expires_at)
                 VALUES(?,?,?,?,'queued',?,?,?)''',
        (confirmation_id, u['id'], snap['site_id'], dumps(new_scope), dumps(approved), stamp(), snap['expires_at']))
    enqueue(c, 'mapping_confirm', confirmation_id)
    event(c, 'mapping_confirmation_requested', u['id'], confirmation_id, {'scan_id':id_, 'items':choices, 'reason':reason.strip()})
    c.commit()
    return confirmation_id


def apply(c, id_, woo=None):
    woo = woo or Woo(); woo.connection = c
    snap = one(c, 'SELECT * FROM stock_sync_catalog_snapshots WHERE id=?', (id_,))
    scope = loads(snap['scope_json']); items = loads(snap['items_json'], [])
    try:
        u = actor(c, snap['actor_id']); accessible(c, u, id_, 'mapping_confirmation')
        if parse_time(snap['expires_at']) <= now():
            raise SyncError('CATALOG_EXPIRED', '映射建议已过期，请重新读取')
        for index, item in enumerate(items):
            u = actor(c, snap['actor_id']); require_manager(u)
            site = target_sites(c, u, {'mode':'explicit_sites','site_ids':[item['site_id']]})[0]
            if site_identity(site) != scope['site_hashes'][str(site['id'])]:
                raise SyncError('MAPPING_INPUT_CHANGED', '站点配置已改变，请重新读取')
            parent, _ = woo.request('GET', site, f'products/{item["product_id"]}')
            leaf = parent
            if item['variation_id']:
                leaf, _ = woo.request('GET', site, f'products/{item["product_id"]}/variations/{item["variation_id"]}')
            if (parent.get('id') != item['product_id'] or leaf.get('id') != (item['variation_id'] or item['product_id'])
                    or identity(site['id'], parent, leaf)['identity_hash'] != item['identity_hash']):
                raise SyncError('MAPPING_INPUT_CHANGED', '站点商品身份已改变，请重新读取映射建议')
            c.execute("UPDATE stock_sync_catalog_snapshots SET status='checking',progress=? WHERE id=?", (index + 1, id_)); c.commit()
        begin(c)
        u = actor(c, snap['actor_id']); accessible(c, u, id_, 'mapping_confirmation')
        if parse_time(snap['expires_at']) <= now():
            raise SyncError('CATALOG_EXPIRED', '映射建议已过期，请重新读取')
        if master_data(c)[2] != scope['master_hash']:
            raise SyncError('MAPPING_INPUT_CHANGED', 'SKU 主档或产品识别规则已改变，请重新读取')
        # Cooperating requests serialize by site; the table's unique key also
        # protects against mappings created by the existing inventory screens.
        for sid in sorted({i['site_id'] for i in items}):
            site = target_sites(c, u, {'mode':'explicit_sites','site_ids':[sid]})[0]
            if site_identity(site) != scope['site_hashes'][str(sid)]:
                raise SyncError('MAPPING_INPUT_CHANGED', '站点配置已改变')
            if hasattr(c, '_raw'):
                c._raw.execute('SELECT pg_advisory_xact_lock(%s,%s)', (194761, sid))
        for item in items:
            previous = rows(c, 'SELECT * FROM inv_site_sku_map WHERE site_id=? AND wc_product_id=? AND COALESCE(wc_variation_id,0)=?'+lock_clause(c),
                            (item['site_id'], item['product_id'], item['variation_id']))
            if previous:
                if len(previous) != 1 or not previous[0]['is_active'] or previous[0]['sku_id'] != item['chosen_sku_id'] or previous[0]['qty_per_item'] != item['chosen_quantity']:
                    raise SyncError('MAPPING_CHANGED', '部分商品已被其他操作绑定，本批未保存，请重新读取')
                item.update(state='unchanged', map_id=previous[0]['id'])
                continue
            # Exact product/variation binding only. Never add an ambiguous
            # WC-SKU fallback or overwrite another fulfilment identity.
            inserted = c.execute('''INSERT INTO inv_site_sku_map(site_id,wc_product_id,wc_variation_id,sku_id,qty_per_item,is_active)
                                   VALUES(?,?,?,?,?,1) RETURNING id''',
                (item['site_id'], item['product_id'], item['variation_id'], item['chosen_sku_id'], item['chosen_quantity'])).fetchone()
            item.update(state='saved', map_id=inserted['id'])
            event(c, 'mapping_created', u['id'], str(inserted['id']), {'confirmation_id':id_,
                'site_id':item['site_id'], 'product_id':item['product_id'], 'variation_id':item['variation_id'],
                'sku_id':item['chosen_sku_id'], 'qty_per_item':item['chosen_quantity'],
                'recognition':item.get('recognition'), 'reason':scope['reason']})
        c.execute("UPDATE stock_sync_catalog_snapshots SET status='complete',complete=1,items_json=?,progress=?,observed_at=? WHERE id=?",
                  (dumps(items), len(items), stamp(), id_)); c.commit()
    except Exception as exc:
        c.rollback()
        logging.getLogger(__name__).warning('Mapping confirmation %s failed: %s', id_, type(exc).__name__)
        c.execute("UPDATE stock_sync_catalog_snapshots SET status='failed',error=?,complete=0 WHERE id=?",
                  (str(exc) if isinstance(exc, SyncError) else '映射未保存，可能存在并发变更，请重新读取', id_)); c.commit()
