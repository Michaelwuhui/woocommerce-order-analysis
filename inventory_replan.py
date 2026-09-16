"""Transactional restock jobs; allocation stays in the fulfillment domain."""

from fulfillment_service import (
    _candidate_warehouses, enqueue_job, json_load, managed_product_family,
    plan_order, record_event, utcnow,
)
from inv_common import _table_exists

SHORTAGE_REASON = '订单存在缺货或未映射商品，需联系客户处理'
WAREHOUSE_JOB = 'REPLAN_RESTOCKED_WAREHOUSE'
ORDER_JOB = 'REPLAN_ORDER_AFTER_RESTOCK'
ACTOR = {'type': 'system', 'name': '补货后自动重算'}


def _enabled(conn):
    if not all(_table_exists(conn, name) for name in (
        'settings', 'oms_order_fulfillment_state', 'oms_integration_jobs',
    )):
        return False
    row = conn.execute("SELECT value FROM settings WHERE key='oms_fulfillment_enabled'").fetchone()
    return bool(row and str(row['value']).lower() in ('1', 'true', 'yes', 'on'))


def _local_warehouse(conn, warehouse_id):
    return bool(conn.execute('''SELECT w.id FROM warehouses w
        JOIN oms_warehouse_integrations wi ON wi.warehouse_id=w.id
        WHERE w.id=? AND w.is_active=1 AND wi.is_enabled=1
          AND wi.inventory_authority='local' ''', (warehouse_id,)).fetchone())


def queue_restock_replan(conn, *, warehouse_id, sku_id, movement_id,
                        ref_type=None, ref_id=None):
    """Called after a positive physical-stock movement, before its commit.

    One inventory document can update many SKUs. Its warehouse job runs only
    after the entire document commits, then queues affected orders oldest first.
    Non-document movements each have their own durable trigger.
    """
    if not _enabled(conn) or not _local_warehouse(conn, warehouse_id):
        return None
    if not conn.execute('''SELECT oi.id FROM oms_order_items oi
        JOIN orders o ON o.id=oi.order_id
        JOIN oms_order_fulfillment_state s ON s.order_id=o.id
        WHERE oi.sku_id=? AND oi.shortage_qty>0 AND s.has_shortage=1
          AND o.status IN ('processing','offline','partial-shipped') LIMIT 1''',
        (sku_id,)).fetchone():
        return None
    event = (f'document:{ref_id}' if ref_type == 'inventory_document' and ref_id is not None
             else f'movement:{movement_id}')
    return enqueue_job(
        conn, WAREHOUSE_JOB, 'warehouse', str(warehouse_id),
        f'restock:{warehouse_id}:{event}',
        {'warehouse_id': warehouse_id, 'movement_id': movement_id,
         'ref_type': ref_type, 'ref_id': ref_id}, available_at=utcnow(),
    )


def eligible_restock_orders(conn, warehouse_id, order_id=None):
    """Only shortages with currently usable stock and unstarted fulfillment."""
    if not _enabled(conn) or not _local_warehouse(conn, warehouse_id):
        return []
    sql = '''SELECT DISTINCT o.id,o.date_created,site.country
        FROM orders o JOIN sites site ON site.url=o.source
        JOIN oms_order_fulfillment_state s ON s.order_id=o.id
        JOIN oms_order_items oi ON oi.order_id=o.id
        JOIN inv_stock st ON st.sku_id=oi.sku_id AND st.warehouse_id=?
        WHERE o.status IN ('processing','offline','partial-shipped')
          AND s.has_shortage=1 AND s.aggregate_status='stock_shortage'
          AND (s.manual_review=0 OR s.manual_reason=?)
          AND oi.shortage_qty>0 AND st.on_hand>st.reserved
          AND NOT EXISTS (SELECT 1 FROM oms_fulfillments f
              WHERE f.order_id=o.id AND f.revision=s.revision
                AND f.status NOT IN ('planned','ready_to_pick','ready_to_submit','stock_shortage'))
          AND NOT EXISTS (SELECT 1 FROM oms_shipments sp
              JOIN oms_fulfillments sf ON sf.id=sp.fulfillment_id
              WHERE sf.order_id=o.id AND sp.status!='cancelled')'''
    params = [warehouse_id, SHORTAGE_REASON]
    if order_id is not None:
        sql += ' AND o.id=?'
        params.append(order_id)
    sql += ' ORDER BY o.date_created,o.id'
    result = []
    for row in conn.execute(sql, params).fetchall():
        for item in conn.execute('''SELECT sku_id,raw_json FROM oms_order_items
            WHERE order_id=? AND shortage_qty>0 AND sku_id IS NOT NULL''', (row['id'],)).fetchall():
            family = managed_product_family(conn, json_load(item['raw_json'], {}) or {})
            candidates = _candidate_warehouses(conn, row['country'], item['sku_id'], managed_family=bool(family))
            if any(candidate['warehouse_id'] == warehouse_id and candidate['available'] > 0
                   for candidate in candidates):
                result.append(dict(row))
                break
    return result


def handle_restocked_warehouse(conn, job, payload):
    warehouse_id = int(payload['warehouse_id'])
    orders = eligible_restock_orders(conn, warehouse_id)
    available_at = utcnow()
    job_ids = []
    for order in orders:
        job_ids.append(enqueue_job(
            conn, ORDER_JOB, 'order', order['id'],
            f"restock-order:{job['id']}:{order['id']}",
            {'order_id': order['id'], 'warehouse_id': warehouse_id,
             'restock_job_id': job['id'], 'movement_id': payload.get('movement_id')},
            available_at=available_at,
        ))
    # The worker commits these jobs and the parent success together.
    return {'order_ids': [order['id'] for order in orders], 'job_ids': job_ids}


def handle_restocked_order(conn, job, payload):
    order_id = payload['order_id']
    warehouse_id = int(payload['warehouse_id'])
    if hasattr(conn, '_raw'):
        conn.execute('SELECT id FROM orders WHERE id=? FOR UPDATE', (order_id,)).fetchall()
        conn.execute('SELECT order_id FROM oms_order_fulfillment_state WHERE order_id=? FOR UPDATE', (order_id,)).fetchall()
    # State/routing can change between the stock receipt and worker execution.
    if not eligible_restock_orders(conn, warehouse_id, order_id):
        return {'order_id': order_id, 'action': 'skipped', 'reason': 'no_eligible_shortage'}
    result = plan_order(conn, order_id, actor=ACTOR, commit=False)
    if result['action'] == 'planned':
        record_event(conn, 'order', order_id, 'restock_replanned', actor=ACTOR,
                     reason='库存入账后重新分配待发货缺货商品',
                     payload={**payload, 'job_id': job['id'], 'revision': result['revision']})
    # Reservation updates and job success are committed atomically by the worker.
    return result
