"""Pure stock decision rules: no networking, no writes, no invented stock."""
from stock_sync_common import SyncError

STOCK_FIELDS = frozenset({'manage_stock', 'stock_quantity', 'stock_status', 'backorders'})


def supply_result(sources, safety, qty_per_item):
    if type(qty_per_item) is not int or qty_per_item <= 0:
        raise SyncError('INVALID_UNIT', '销售单位必须是正整数')
    if type(safety) is not int or safety < 0:
        raise SyncError('INVENTORY_UNKNOWN', '安全库存口径无效')
    pools = {}
    for source in sources:
        key = source['pool_id']
        comparable=lambda s:{k:v for k,v in s.items() if k not in ('warehouse_id','source_updated_at','synced_at')}
        if key in pools and comparable(pools[key]) != comparable(source):
            raise SyncError('INVENTORY_UNKNOWN', '同一物理库存存在矛盾来源')
        pools[key] = source
    sources = list(pools.values())
    mode = 'quantity' if sources and all(s['authority'] != 'manual_partner' for s in sources) else 'status'
    total, states, errors = 0, [], []
    for s in sources:
        if s.get('error'):
            states.append('unknown'); errors.append(s['error']); continue
        if s['authority'] == 'manual_partner':
            states.append(s.get('status', 'unknown')); continue
        available = s.get('available')
        if type(available) is not int or available < 0:
            states.append('unknown'); errors.append('INVENTORY_UNKNOWN'); continue
        total += available
        states.append('instock' if available >= qty_per_item else 'outofstock')
    if mode == 'quantity':
        if errors or 'unknown' in states:
            return {'mode': mode, 'status': 'unknown', 'quantity': None, 'errors': errors}
        quantity = max(0, total - safety) // qty_per_item
        return {'mode': mode, 'status': 'instock' if quantity else 'outofstock', 'quantity': quantity, 'errors': []}
    # Mixed sources do not publish an incomplete numerical sum.
    quantity_available = max(0, total - safety) >= qty_per_item
    manual_available = any(s.get('status') == 'instock' and not s.get('error') for s in sources if s['authority'] == 'manual_partner')
    status = 'instock' if quantity_available or manual_available else (
        'outofstock' if states and all(s == 'outofstock' for s in states) else 'unknown')
    return {'mode': mode, 'status': status, 'quantity': None, 'errors': errors}


def resolve_target_stock(binding, supply, controls):
    holds = [x for x in controls if x['kind'] == 'manual_hold']
    refs = [x for x in controls if x['kind'] == 'reference_snapshot']
    if len(refs) > 1:
        raise SyncError('SOURCE_STATUS_CONFLICT')
    ref = refs[0]['stock_status'] if refs else None
    if ref == 'onbackorder':
        raise SyncError('BACKORDER_POLICY_CONFLICT')
    if holds or ref == 'outofstock':
        return {'manage_stock': False, 'stock_status': 'outofstock', 'backorders': 'no'}, (
            'MANUAL_HOLD_ACTIVE' if holds else 'REFERENCE_HOLD_ACTIVE')
    if supply['mode'] == 'quantity':
        if supply['quantity'] is None:
            raise SyncError((supply.get('errors') or ['INVENTORY_UNKNOWN'])[0])
        q = supply['quantity']
        return {'manage_stock': True, 'stock_quantity': q, 'stock_status': 'instock' if q else 'outofstock', 'backorders': 'no'}, 'ACTUAL_STOCK_ZERO' if q == 0 else 'ACTUAL_STOCK'
    # A genuinely quantity-managed resource cannot be restored without a complete
    # authority, even if a reference shop happens to say it is in stock.
    if binding.get('mode') == 'quantity':
        raise SyncError('INVENTORY_UNKNOWN', '原数量管理商品缺少可靠数量，不能关闭数量保护恢复')
    confirmations = [x for x in controls if x['kind'] == 'manual_availability']
    state = ref or (confirmations[-1]['stock_status'] if confirmations else supply['status'])
    if state not in ('instock', 'outofstock'):
        raise SyncError('INVENTORY_UNKNOWN', '没有可靠供货依据，请明确确认供货或选择参照站')
    return {'manage_stock': False, 'stock_status': state, 'backorders': 'no'}, 'STATUS_SOURCE'


def matches(actual, intended):
    # bool('parent') and int(None) are deliberately not accepted.
    return all(type(actual.get(k)) is type(v) and actual.get(k) == v for k, v in intended.items() if k in STOCK_FIELDS)
