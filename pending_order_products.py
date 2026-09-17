"""Pending-list presentation only; never grants warehouse or shipping access."""


def include_market_shortages(original_items, allocations, shortages):
    """Include unallocated shortage units once, without exposing other allocations."""
    by_line = {}
    for allocated in allocations:
        key = str(allocated['id'])
        if key not in by_line:
            by_line[key] = dict(allocated)
        else:
            by_line[key]['quantity'] += allocated['quantity']
            by_line[key]['total'] += allocated['total']
    for original in original_items:
        key = str(original.get('id'))
        qty = min(max(0, int(shortages.get(key, 0))), int(original.get('quantity') or 0))
        if not qty:
            continue
        if key not in by_line:
            by_line[key] = dict(original, quantity=0, total=0,
                                name=original.get('name', '') + '（缺货待分仓，仅供核对）')
        item = by_line[key]
        item['quantity'] += qty
        item['total'] = float(item.get('total') or 0) + float(original.get('total') or 0) * qty / max(1, int(original.get('quantity') or 1))
        item['shortage_quantity'] = qty
    return [by_line[str(item.get('id'))] for item in original_items if str(item.get('id')) in by_line]
