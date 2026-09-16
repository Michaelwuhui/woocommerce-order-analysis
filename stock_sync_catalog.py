"""Durable immutable catalog snapshots; a failed page never means sold out."""
import logging
from stock_sync_common import SyncError, uid, rows, one, dumps, loads, stamp, event
from stock_sync_permissions import actor, target_sites, reference_site
from stock_sync_supply import mappings, mapping_hash
from stock_sync_woo import Woo, stock_state


def enqueue(c, kind, object_id):
    c.execute("INSERT INTO stock_sync_work(id,kind,object_id,status,created_at) VALUES(?,?,?,'queued',?)", (uid(),kind,object_id,stamp()))


def create_scan(c, u, request):
    site_id = request.get('source_site_id')
    scope = request.get('target_scope') or {'mode':'all_authorized'}
    if site_id:
        site_id = int(site_id)
        reference_site(c,u,site_id)
        scope = {'mode':'explicit_sites','site_ids':[]}
    else:
        scope = {'mode':'explicit_sites','site_ids':[s['id'] for s in target_sites(c,u,scope)]}
    id_ = uid()
    c.execute('''INSERT INTO stock_sync_catalog_snapshots(id,actor_id,site_id,scope_json,status,created_at,expires_at)
        VALUES(?,?,?,?,'queued',?,?)''', (id_,u['id'],site_id,dumps(scope),stamp(),stamp(3600)))
    enqueue(c,'catalog',id_)
    c.commit()
    return id_


def scan(c, id_, woo=None):
    woo = woo or Woo()
    woo.connection = c
    snap = one(c,'SELECT * FROM stock_sync_catalog_snapshots WHERE id=?',(id_,))
    items, error = [], None
    try:
        u = actor(c,snap['actor_id'])
        scope = loads(snap['scope_json'])
        site_id = snap['site_id']
        if not site_id:
            sites = target_sites(c,u,scope)
            for m in mappings(c,[s['id'] for s in sites]):
                items.append({'id':m['id'],'site_id':m['site_id'],'product_id':m['wc_product_id'],
                    'variation_id':m['wc_variation_id'] or 0,'name':m['sku_name'],'sku_code':m['sku_code'],
                    'map_ids':[m['id']],'mapping_hashes':{str(m['id']):mapping_hash(m)},'sku_id':m['sku_id'],
                    'qty_per_item':m['qty_per_item'],'complete':True})
        else:
            site = reference_site(c,u,site_id)
            maps = mappings(c,[site_id])
            c.commit()
            seen = set()
            def append(product, leaf, variation_id):
                key = (product['id'],variation_id)
                if key in seen:
                    raise SyncError('SOURCE_INCOMPLETE','目录出现重复资源')
                seen.add(key)
                ms = [m for m in maps if m['wc_product_id']==product['id'] and (m['wc_variation_id'] or 0)==variation_id]
                name = product.get('name','')
                if variation_id:
                    name += ' / ' + ', '.join(str(a.get('option','')) for a in leaf.get('attributes',[]))
                items.append({'id':len(items)+1,'site_id':site_id,'product_id':product['id'],
                    'variation_id':variation_id,'name':name,'sku_code':ms[0]['sku_code'] if len(ms)==1 else '',
                    'sku_id':ms[0]['sku_id'] if len(ms)==1 else None,'qty_per_item':ms[0]['qty_per_item'] if len(ms)==1 else None,
                    'map_ids':[m['id'] for m in ms],'mapping_hashes':{str(m['id']):mapping_hash(m) for m in ms},
                    'state':stock_state(leaf),'complete':True})
            for page in woo.pages(site,'products'):
                reference_site(c,actor(c,snap['actor_id']),site_id)
                c.commit()
                for product in page:
                    if product.get('type') == 'variable':
                        leaves = []
                        for variations in woo.pages(site,f'products/{product["id"]}/variations'):
                            leaves.extend(variations)
                        if set(product.get('variations',[])) != {v['id'] for v in leaves}:
                            raise SyncError('SOURCE_INCOMPLETE','口味列表在扫描期间发生变化')
                        for leaf in leaves:
                            append(product,leaf,leaf['id'])
                    else:
                        append(product,product,0)
                c.execute("UPDATE stock_sync_catalog_snapshots SET status='scanning',progress=?,items_json=? WHERE id=?", (len(items),dumps(items),id_))
                c.commit()
    except SyncError as exc:
        error = exc.code
    except Exception:
        c.rollback()
        logging.getLogger(__name__).exception('Catalog snapshot %s failed', id_)
        error = 'CATALOG_READ_FAILED'
    c.execute('''UPDATE stock_sync_catalog_snapshots SET status=?,complete=?,items_json=?,progress=?,error=?,observed_at=? WHERE id=?''',
        ('incomplete' if error else 'complete',0 if error else 1,dumps(items),len(items),error,stamp(),id_))
    c.commit()


def select_items(snap, selection):
    items = loads(snap['items_json'],[])
    mode = selection.get('mode')
    if mode in ('all','filtered_all'):
        if not snap['complete']:
            raise SyncError('SOURCE_INCOMPLETE','目录未完整读取，不能全选')
        chosen = items
        if mode=='filtered_all':
            term = str((selection.get('filter') or {}).get('search','')).strip().casefold()
            chosen = [i for i in chosen if term in (i['name']+' '+i['sku_code']).casefold()]
        excluded = selection.get('excluded_catalog_item_ids',[])
        if not isinstance(excluded,list) or any(x not in {i['id'] for i in items} for x in excluded):
            raise SyncError('INVALID_SELECTION','排除项不属于本次目录',400)
        chosen = [i for i in chosen if i['id'] not in excluded]
    elif mode=='explicit':
        ids = selection.get('catalog_item_ids',[])
        if not isinstance(ids,list) or any(x not in {i['id'] for i in items} for x in ids):
            raise SyncError('INVALID_SELECTION','选择项不属于本次目录',400)
        chosen = [i for i in items if i['id'] in ids and i['complete']]
    else:
        raise SyncError('INVALID_SELECTION','请选择全选或明确商品集合',400)
    if not chosen:
        raise SyncError('EMPTY_SELECTION','没有选中商品',400)
    return chosen
