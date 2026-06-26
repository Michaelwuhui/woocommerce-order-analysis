"""
inv_allocator.py — 模块 6:多仓自动分配 / 拆单引擎(纯算法)。

目标(对齐锁定架构):
  订单 line_items 解析到 SKU 后,按"该市场可用仓优先级 + 各仓实时库存":
    1. 优先用单仓凑齐整单(少拆包);
    2. 单仓都凑不齐才拆单,按优先级贪心分配;
    3. 仍不足的记为 shortage(缺货/待补)。

实时可用 available = inv_stock.on_hand - reserved。
候选仓 = 该市场 inv_market_warehouses 中 is_active 且仓库 is_active 且参与分仓
(inv_warehouse_ext.is_fulfillment=1),按 priority 升序。

本模块不写库,只产出分配方案;落库(预留/出库/生成 fulfillment)由 inv_orders 负责。
"""


def candidate_warehouses(conn, market):
    """返回该市场的候选仓(按优先级),含自营/合伙人标记。"""
    rows = conn.execute('''
        SELECT mw.warehouse_id, mw.priority, w.name, w.code, w.country,
               COALESCE(we.ownership_type,'self') AS ownership_type, we.partner_name
        FROM inv_market_warehouses mw
        JOIN warehouses w ON w.id = mw.warehouse_id
        LEFT JOIN inv_warehouse_ext we ON we.warehouse_id = w.id
        WHERE mw.market_code = ? AND mw.is_active = 1 AND w.is_active = 1
          AND COALESCE(we.is_fulfillment, 1) = 1
        ORDER BY mw.priority, mw.id
    ''', (market,)).fetchall()
    return [dict(r) for r in rows]


def _available_map(conn, warehouse_ids, sku_ids):
    """{(warehouse_id, sku_id): available}。available = on_hand - reserved。"""
    if not warehouse_ids or not sku_ids:
        return {}
    whq = ','.join('?' * len(warehouse_ids))
    skq = ','.join('?' * len(sku_ids))
    rows = conn.execute(
        f'''SELECT warehouse_id, sku_id, (on_hand - reserved) AS avail
            FROM inv_stock WHERE warehouse_id IN ({whq}) AND sku_id IN ({skq})''',
        list(warehouse_ids) + list(sku_ids)).fetchall()
    return {(r['warehouse_id'], r['sku_id']): r['avail'] for r in rows}


def allocate(conn, market, needs, prefer_single=True):
    """计算分配方案。

    needs: {sku_id: qty}(已聚合的整单 SKU 需求)。
    返回 dict:
      {
        'market': str,
        'is_split': bool,
        'allocations': [ {warehouse_id, name, ownership_type, partner_name,
                          lines: [{sku_id, qty}]}, ... ],
        'shortage': {sku_id: qty},      # 所有候选仓都凑不齐的部分
        'candidates': [...],            # 候选仓(调试/展示用)
        'reason': 'single'|'split'|'no_candidates'|'empty',
      }
    """
    needs = {k: v for k, v in (needs or {}).items() if v and v > 0}
    cands = candidate_warehouses(conn, market)
    base = {'market': market, 'candidates': cands, 'shortage': {}}
    if not needs:
        return {**base, 'is_split': False, 'allocations': [], 'reason': 'empty'}
    if not cands:
        return {**base, 'is_split': False, 'allocations': [], 'shortage': dict(needs), 'reason': 'no_candidates'}

    wh_ids = [c['warehouse_id'] for c in cands]
    avail = _available_map(conn, wh_ids, list(needs.keys()))

    def av(wh, sku):
        return max(0, avail.get((wh, sku), 0))

    # 1) 单仓优先:第一个能整单凑齐的仓
    if prefer_single:
        for c in cands:
            wh = c['warehouse_id']
            if all(av(wh, sku) >= qty for sku, qty in needs.items()):
                return {**base, 'is_split': False, 'reason': 'single',
                        'allocations': [{
                            'warehouse_id': wh, 'name': c['name'],
                            'ownership_type': c['ownership_type'], 'partner_name': c['partner_name'],
                            'lines': [{'sku_id': sku, 'qty': qty} for sku, qty in needs.items()],
                        }]}

    # 2) 拆单:按优先级贪心
    remaining = dict(needs)
    local_av = {(c['warehouse_id'], sku): av(c['warehouse_id'], sku)
                for c in cands for sku in needs}
    allocations = []
    for c in cands:
        wh = c['warehouse_id']
        lines = []
        for sku in list(remaining.keys()):
            if remaining[sku] <= 0:
                continue
            take = min(remaining[sku], local_av.get((wh, sku), 0))
            if take > 0:
                lines.append({'sku_id': sku, 'qty': take})
                remaining[sku] -= take
        if lines:
            allocations.append({'warehouse_id': wh, 'name': c['name'],
                                'ownership_type': c['ownership_type'], 'partner_name': c['partner_name'],
                                'lines': lines})
        if all(v <= 0 for v in remaining.values()):
            break

    shortage = {sku: q for sku, q in remaining.items() if q > 0}
    is_split = len(allocations) > 1
    reason = 'split' if is_split else ('single' if allocations else 'no_candidates')
    return {**base, 'is_split': is_split, 'allocations': allocations,
            'shortage': shortage, 'reason': reason}
