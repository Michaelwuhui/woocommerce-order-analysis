"""Exact-ID WooCommerce adapter. No search/name fallback or master routing."""
from urllib.parse import urlsplit
from email.utils import parsedate_to_datetime
import math
import requests

from stock_sync_common import SyncError, one, rows, stamp, digest, now, parse_time, loads
from stock_sync_policy import STOCK_FIELDS, matches

STATE_FIELDS = ('id','parent_id','type','status','manage_stock','stock_quantity','stock_status','backorders')


def stock_state(item):
    return {k: item.get(k) for k in STATE_FIELDS}


def resource_key(site, product_id, variation_id=0):
    p = urlsplit(site['url'].rstrip('/'))
    # Scheme is not part of identity: http/https aliases cannot acquire two locks.
    return f'{p.netloc.lower()}{p.path.rstrip("/")}/products/{product_id}/variations/{variation_id or 0}'


def site_identity(site):
    return digest({k:site.get(k) for k in ('id','url','product_master_id','consumer_key','consumer_secret','is_active')})


def managed_resource_keys(c,site,mapping):
    keys=[resource_key(site,mapping['wc_product_id'],mapping.get('wc_variation_id'))]
    if site.get('product_master_id') and _masters_exist(c):
        master=one(c,'SELECT * FROM product_masters WHERE id=?',(site['product_master_id'],))
        url=(master or {}).get('url') or (master or {}).get('api_url')
        if url:keys.append(resource_key({'url':url},mapping['wc_product_id'],mapping.get('wc_variation_id')))
    return sorted(set(keys))


def physical_scope(c, site, mapping, target_ids, source_id=None):
    key = resource_key(site, mapping['wc_product_id'], mapping.get('wc_variation_id'))
    same = [s for s in rows(c, 'SELECT id,url FROM sites') if resource_key(s, mapping['wc_product_id'], mapping.get('wc_variation_id')) == key]
    affected = {s['id'] for s in same}
    cap = one(c, 'SELECT * FROM stock_sync_site_capabilities WHERE site_id=?', (site['id'],))
    try:
        proof=loads(cap['evidence']) if cap else {}
    except (ValueError,TypeError):
        proof={}
    if site.get('product_master_id') and not (cap and cap['independent_write'] == 1 and proof.get('site_identity')==site_identity(site) and proof.get('evidence')):
        raise SyncError('PHYSICAL_SCOPE_CONFLICT', '子站独立写入能力尚未核实，不能自动改写主站')
    # A configured product master may distribute writes to its children.
    masters = rows(c, 'SELECT * FROM product_masters') if _masters_exist(c) else []
    for master in masters:
        master_url = master.get('url') or master.get('api_url')
        if master_url and resource_key({'url':master_url}, mapping['wc_product_id'], mapping.get('wc_variation_id')) == key:
            if c.execute('SELECT 1 FROM sites WHERE product_master_id=? LIMIT 1', (master['id'],)).fetchone():
                raise SyncError('PHYSICAL_SCOPE_CONFLICT', '主站传播范围不能隔离，请使用已核实的子站接口')
    if not affected.issubset(set(target_ids)) or source_id in affected:
        raise SyncError('PHYSICAL_SCOPE_CONFLICT', '实际资源涉及未选站点或只读参照站')
    return {'resource_key': key, 'affected_site_ids': sorted(affected), 'capability': cap,
            'site_identity': site_identity(site)}


def _masters_exist(c):
    from stock_sync_common import exists
    return exists(c, 'product_masters')


class Woo:
    def __init__(self, session=None):
        self.http = session or requests.Session()
        self.connection = None

    def request(self, method, site, path, **kwargs):
        if not site.get('consumer_key') or not site.get('consumer_secret') or site.get('is_active',1) == 0:
            raise SyncError('SITE_UNAVAILABLE', '站点禁用或凭据不完整')
        endpoint=resource_key(site,0).split('/products/')[0]
        c=self.connection
        if c is not None:
            cooldown=one(c,'SELECT retry_at FROM stock_sync_rate_limits WHERE endpoint_key=?',(endpoint,))
            c.commit()
            if cooldown and parse_time(cooldown['retry_at'])>now():
                raise SyncError('REMOTE_RATE_LIMITED','站点仍在退避期间，请在 '+str(cooldown['retry_at'])+' 后重新预览')
        try:
            r = self.http.request(method, site['url'].rstrip('/') + '/wp-json/wc/v3/' + path,
                auth=(site['consumer_key'], site['consumer_secret']), timeout=(5,20),
                headers={'Accept':'application/json','Cache-Control':'no-cache','User-Agent':'WooStockSync/1.0'},
                allow_redirects=False, **kwargs)
        except requests.RequestException:
            raise SyncError('WRITE_RESULT_UNKNOWN' if method == 'PUT' else 'SOURCE_READ_FAILED', '远程请求未能确认结果') from None
        if c is not None and (r.status_code==429 or r.status_code>=500):
            value=r.headers.get('Retry-After')
            try:
                delay=max(1,math.ceil(float(value)))
            except (ValueError,TypeError):
                try:
                    delay=max(1,math.ceil((parsedate_to_datetime(value)-now()).total_seconds()))
                except (ValueError,TypeError,OverflowError):
                    delay=60 if r.status_code==429 else 10
            retry_at=stamp(delay)
            c.execute('''INSERT INTO stock_sync_rate_limits(endpoint_key,retry_at,reason) VALUES(?,?,?)
                ON CONFLICT(endpoint_key) DO UPDATE SET retry_at=CASE
                WHEN stock_sync_rate_limits.retry_at>excluded.retry_at THEN stock_sync_rate_limits.retry_at
                ELSE excluded.retry_at END,reason=excluded.reason''',(endpoint,retry_at,'HTTP_'+str(r.status_code)))
            c.commit()
        if r.status_code == 429:
            error = SyncError('REMOTE_RATE_LIMITED', '站点限流，请稍后重新预览')
            # Never silently shorten a server requested delay.
            error.retry_after = r.headers.get('Retry-After')
            raise error
        if r.status_code not in (200,201):
            code = 'REMOTE_AUTH_FAILED' if r.status_code in (401,403) else ('WRITE_RESULT_UNKNOWN' if method == 'PUT' and r.status_code >= 500 else 'REMOTE_HTTP_ERROR')
            raise SyncError(code, f'站点返回 HTTP {r.status_code}')
        try:
            return r.json(), r.headers
        except ValueError:
            raise SyncError('REMOTE_INVALID_JSON', '站点未返回有效 JSON') from None

    def path(self, m):
        p = f'products/{int(m["wc_product_id"])}'
        return p + f'/variations/{int(m["wc_variation_id"])}' if m.get('wc_variation_id') else p

    def read(self, site, m):
        item, _ = self.request('GET', site, self.path(m))
        if not isinstance(item, dict) or item.get('id') != (m.get('wc_variation_id') or m['wc_product_id']):
            raise SyncError('RESOURCE_ID_MISMATCH')
        if m.get('wc_variation_id'):
            parent, _ = self.request('GET', site, f'products/{m["wc_product_id"]}')
            if parent.get('id') != m['wc_product_id']:
                raise SyncError('RESOURCE_ID_MISMATCH')
            item['_parent_manage_stock'] = parent.get('manage_stock')
            item['_parent_status'] = parent.get('status')
        return item

    def validate(self, item, m):
        if item.get('status') != 'publish' or item.get('_parent_status', 'publish') != 'publish' or item.get('type') in ('grouped','external','variable'):
            raise SyncError('UNSUPPORTED_PRODUCT', '只处理已发布的简单商品或独立口味变体')
        manage = item.get('manage_stock')
        if manage == 'parent' or (m.get('wc_variation_id') and manage is False and item.get('_parent_manage_stock') is True):
            raise SyncError('PARENT_STOCK_SHARED', '该口味共享父级库存，不能隔离写入')
        if type(manage) is not bool or item.get('stock_status') not in ('instock','outofstock','onbackorder'):
            raise SyncError('INVENTORY_UNKNOWN', '库存字段不完整')
        if item.get('backorders') not in ('no','notify','yes'):
            raise SyncError('BACKORDER_POLICY_CONFLICT')

    def pages(self, site, path, on_page=None):
        page, count, expected = 1, 0, None
        while page <= 10000:
            data, headers = self.request('GET', site, path, params={'per_page':100,'page':page,'status':'any','orderby':'id','order':'asc'})
            if not isinstance(data,list):
                raise SyncError('SOURCE_INCOMPLETE')
            total = headers.get('X-WP-Total')
            if total is not None:
                try:
                    total = int(total)
                except ValueError:
                    raise SyncError('SOURCE_INCOMPLETE') from None
                if expected is not None and total != expected:
                    raise SyncError('SOURCE_INCOMPLETE', '目录扫描期间商品总数变化，请重新扫描')
                expected = total
            count += len(data)
            if on_page:
                on_page({'page': page, 'total': expected, 'read': count})
            yield data
            if len(data) < 100 or (expected is not None and count >= expected):
                if expected is not None and count != expected:
                    raise SyncError('SOURCE_INCOMPLETE')
                return
            page += 1
        raise SyncError('SOURCE_INCOMPLETE', '目录超过扫描上限')

    def publish(self, site, m, before, intended, check):
        if set(intended) - STOCK_FIELDS:
            raise SyncError('INVALID_PAYLOAD')
        phases = [dict(intended)]
        if intended.get('manage_stock') is False and intended.get('stock_status') == 'outofstock' and before.get('manage_stock') is True:
            # Close sales before disabling numerical management. This website-only
            # quantity is a stop switch; physical stock is never changed.
            phases.insert(0, {'manage_stock':True,'stock_quantity':0,'stock_status':'outofstock','backorders':'no'})
        meta = {x.get('key'):x.get('value') for x in before.get('meta_data',[]) if isinstance(x,dict)}
        for phase in phases:
            check()
            if 'wcms_stock_manage' in meta:
                names = {'manage_stock':'wcms_stock_manage','stock_quantity':'wcms_stock_qty','stock_status':'wcms_stock_status'}
                phase['meta_data'] = [{'key':names[k], 'value':('yes' if v else 'no') if k=='manage_stock' else v} for k,v in phase.items() if k in names]
            self.request('PUT', site, self.path(m), json=phase)
