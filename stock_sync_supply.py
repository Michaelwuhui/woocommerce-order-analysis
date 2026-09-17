"""Read existing SKU/warehouse links and reservation-aware quantities."""
from stock_sync_common import one, rows, exists, loads, digest, now, parse_time
from stock_sync_policy import supply_result


def mapping_hash(m):
    return digest({k: m.get(k) for k in ('id','site_id','sku_id','wc_product_id','wc_variation_id','qty_per_item','is_active','updated_at')})


def mappings(c, site_ids=None):
    args, where = [], ''
    if site_ids is not None:
        if not site_ids:
            return []
        args = list(site_ids)
        where = ' AND m.site_id IN (' + ','.join('?' for _ in args) + ')'
    return rows(c, '''SELECT m.*,k.sku_code,k.name AS sku_name FROM inv_site_sku_map m
        JOIN inv_skus k ON k.id=m.sku_id WHERE m.is_active=1 AND k.is_active=1''' + where + ' ORDER BY m.id', args)


def sources_for(c, sku_id, market):
    # No country fallback: an absent serving relationship is unknown, not zero.
    return rows(c, '''SELECT DISTINCT w.id,w.name,wi.inventory_authority,wi.config_json
        FROM oms_sku_warehouses sw JOIN warehouses w ON w.id=sw.warehouse_id
        JOIN inv_market_warehouses mw ON mw.warehouse_id=w.id
        LEFT JOIN oms_warehouse_integrations wi ON wi.warehouse_id=w.id
        WHERE sw.sku_id=? AND sw.is_enabled=1 AND w.is_active=1 AND mw.is_active=1
          AND mw.market_code=? ORDER BY w.id''', (sku_id, market))


def snapshot(c, mapping, site):
    sources = []
    warehouses = sources_for(c, mapping['sku_id'], site.get('country'))
    for w in warehouses:
        config = loads(w.get('config_json'))
        authority = w.get('inventory_authority') or 'local'
        if config.get('requires_quantity_inventory') is False:
            authority = 'manual_partner'
        source = {'pool_id': str(config.get('physical_stock_pool') or w['id']),
                  'warehouse_id': w['id'], 'authority': authority}
        if authority == 'local':
            stock = one(c, 'SELECT on_hand,reserved FROM inv_stock WHERE warehouse_id=? AND sku_id=?', (w['id'], mapping['sku_id']))
            if not stock or any(type(stock[k]) is not int or stock[k] < 0 for k in ('on_hand','reserved')):
                source['error'] = 'INVENTORY_UNKNOWN'
            else:
                source.update(stock)
                source['available'] = max(0, stock['on_hand'] - stock['reserved'])
        elif authority == 'external_wms':
            stock = rows(c, 'SELECT available_quantity,source_updated_at,synced_at FROM oms_external_stock WHERE warehouse_id=? AND sku_id=?', (w['id'], mapping['sku_id']))
            # The adapter's existing available value must already include OMS
            # reservations; explicitly reconciled sources opt into this contract.
            if config.get('stock_sync_available_includes_oms_reservations') is not True or len(stock) != 1:
                source['error'] = 'EXTERNAL_RESERVATION_UNCONFIRMED'
            else:
                source.update(stock[0])
                try:
                    observed = parse_time(stock[0]['source_updated_at'] or stock[0]['synced_at'])
                    age = (now()-observed).total_seconds() if observed else float('inf')
                except (ValueError, TypeError):
                    age = float('inf')
                if age < -60 or age > int(config.get('stock_sync_freshness_seconds', 10800)):
                    source['error'] = 'EXTERNAL_STOCK_STALE'
                else:
                    source['available'] = stock[0]['available_quantity']
        elif authority == 'manual_partner':
            source['status'] = 'unknown'
        else:
            source['error'] = 'INVENTORY_UNKNOWN'
        sources.append(source)
    # Different aliases and markets sharing a physical pool use the same safety
    # rule: the maximum existing policy once per aggregate, never per alias.
    participants = rows(c, '''SELECT DISTINCT s.id,s.country FROM inv_site_sku_map m
        JOIN sites s ON s.id=m.site_id WHERE m.sku_id=? AND m.is_active=1''', (mapping['sku_id'],))
    safety, policies, backlog = 0, [], False
    physical = {s['pool_id'] for s in sources}
    for p in participants:
        linked = sources_for(c, mapping['sku_id'], p['country'])
        pools = {str(loads(w.get('config_json')).get('physical_stock_pool') or w['id']) for w in linked}
        if not pools.intersection(physical):
            continue
        if exists(c, 'inv_site_sync_config'):
            cfg = one(c, 'SELECT mode,allocation_strategy,safety_stock FROM inv_site_sync_config WHERE site_id=?', (p['id'],))
            if cfg:
                safety = max(safety, int(cfg.get('safety_stock') or 0))
                policies.append({'site_id': p['id'], **cfg})
        if exists(c, 'sync_site_progress'):
            latest = one(c, '''SELECT sp.status FROM sync_site_progress sp
                JOIN sync_runs r ON r.run_id=sp.run_id WHERE sp.site_id=?
                ORDER BY r.created_at DESC,r.run_id DESC LIMIT 1''', (p['id'],))
            backlog = backlog or bool(latest and latest['status'] != 'success')
            # A second, older run still executing can also change reservations.
            backlog = backlog or bool(c.execute("SELECT 1 FROM sync_site_progress WHERE site_id=? AND status IN ('queued','fetching','writing','recovering') LIMIT 1", (p['id'],)).fetchone())
        if exists(c, 'sync_page_receipts'):
            backlog = backlog or bool(c.execute("SELECT 1 FROM sync_page_receipts WHERE site_id=? AND post_commit_status IN ('pending','processing','error') LIMIT 1", (p['id'],)).fetchone())
    result = supply_result(sources, safety, mapping['qty_per_item'])
    if result['mode'] == 'quantity' and backlog:
        result.update(quantity=None, status='unknown', errors=['ORDER_SYNC_BACKLOG'])
    return {'sources': sources, 'safety_stock': safety, 'policies': policies, 'result': result}


def strategy_conflict(supply):
    return any(p.get('mode') in ('live','observe') and p.get('allocation_strategy') == 'quota' for p in supply['policies'])
