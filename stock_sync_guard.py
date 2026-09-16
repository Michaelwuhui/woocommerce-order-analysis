"""Serialize legacy writers, and hand enrolled resources to the stock-sync path."""
from contextlib import contextmanager
from pathlib import Path
from urllib.parse import urlsplit
import os
import re

from stock_sync_common import SyncError, connect, exists, rows, uid, loads
from stock_sync_jobs import acquire, release
from stock_sync_policy import STOCK_FIELDS
from stock_sync_woo import resource_key


@contextmanager
def legacy_write(resource_url, payload):
    if not STOCK_FIELDS.intersection(payload):
        yield
        return
    import db_backend
    path=os.getenv('INV_DB_FILE',os.getenv('WOO_SQLITE_PATH','woocommerce_orders.db'))
    if not db_backend.is_postgres_backend() and not Path(path).exists():
        yield
        return
    c=connect()
    owner='legacy:'+uid()
    leased=False
    try:
        if not exists(c,'stock_sync_bindings'):
            yield
            return
        parts=urlsplit(resource_url)
        match=re.fullmatch(r'(.*)/wp-json/wc/v3/products/(\d+)(?:/variations/(\d+))?',parts.path)
        if not match:
            raise SyncError('STOCK_SYNC_RESOURCE_UNRESOLVED')
        endpoint=f'{parts.scheme}://{parts.netloc}{match[1]}'
        pid,vid=int(match[2]),int(match[3] or 0)
        key=resource_key({'url':endpoint},pid,vid)
        acquire(c,[key,'endpoint:'+key.split('/products/')[0]],owner)
        leased=True
        # Keep the original physical resource protected even if its mapping was
        # deleted or changed outside this module.
        for b in rows(c,'SELECT baseline_json FROM stock_sync_bindings'):
            for managed in loads(b['baseline_json']).get('managed_resource_keys',[]):
                if managed==key or (not vid and managed.split('/variations/')[0]==key.split('/variations/')[0]):
                    raise SyncError('STOCK_SYNC_MANAGED','该资源已接入库存同步，请通过差异预览修改库存')
        bound=rows(c,'''SELECT b.map_id,m.wc_product_id,m.wc_variation_id,s.url,s.product_master_id
            FROM stock_sync_bindings b JOIN inv_site_sku_map m ON m.id=b.map_id
            JOIN sites s ON s.id=b.site_id''')
        # The parent resource can indirectly change all inherited variations.
        for b in bound:
            urls=[b['url']]
            if b.get('product_master_id') and exists(c,'product_masters'):
                masters=rows(c,'SELECT * FROM product_masters WHERE id=?',(b['product_master_id'],))
                urls.extend(x.get('url') or x.get('api_url') or '' for x in masters)
            if any(resource_key({'url':url},pid,0).split('/products/')[0]==key.split('/products/')[0] for url in urls) and b['wc_product_id']==pid and (not vid or (b['wc_variation_id'] or 0)==vid):
                raise SyncError('STOCK_SYNC_MANAGED','该商品已接入库存同步，请通过差异预览修改库存；价格仍可单独修改')
        c.commit()
        yield
    finally:
        if leased:
            release(c,owner)
        c.close()
