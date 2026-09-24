"""Non-posting historical statements: commercial snapshots, never stock replay.

Rules and immutable drafts use versioned rec_objects. They cannot enter cash
allocation or the final ledger while supplier cost/collection evidence is absent.
"""
import hashlib
import html
import json
import re
from collections import defaultdict
from datetime import date
from decimal import Decimal

from reconciliation_core import (ReconciliationError, audit, currency, dec, dump,
                                 money, now, object_, objects, rows, uid)
from product_recognition import parse_product_name


def month_bounds(month):
    if not isinstance(month, str) or not re.fullmatch(r'\d{4}-\d{2}', month):
        raise ReconciliationError('月份格式应为 YYYY-MM')
    try:
        y, m = map(int, month.split('-'))
        start = date(y, m, 1)
        end = date(y + (m == 12), 1 if m == 12 else m + 1, 1)
    except ValueError as exc:
        raise ReconciliationError('月份无效') from exc
    return start.isoformat(), end.isoformat()


def rate_at(conn, unit, month):
    month_bounds(month)
    if unit == 'CNY':
        return {'currency': unit, 'month': month, 'requested_month': month, 'rate': '1', 'fallback': False}
    row = conn.execute('''SELECT year_month,CAST(rate_to_cny AS TEXT) AS rate
        FROM exchange_rates WHERE currency=? AND year_month<=?
        ORDER BY year_month DESC LIMIT 1''', (currency(unit), month)).fetchone()
    if not row or dec(row['rate']) <= 0:
        return None
    return {'currency': unit, 'month': row['year_month'], 'requested_month': month,
            'rate': str(dec(row['rate'])), 'fallback': row['year_month'] != month}


def save_rule(conn, name, data, actor):
    if not isinstance(name, str) or not name.strip() or len(name) > 150:
        raise ReconciliationError('规则名称必填，最多150字')
    d = {k: data.get(k) for k in ('warehouse_id', 'pool_id', 'contract_id', 'our_party_id', 'supplier_id', 'provider_id', 'effective_from', 'effective_to')}
    try:
        d['warehouse_id'] = int(d['warehouse_id'])
        start, end = date.fromisoformat(d['effective_from']), date.fromisoformat(d['effective_to'])
    except (ValueError, TypeError) as exc:
        raise ReconciliationError('仓库及生效日期必填') from exc
    if start >= end:
        raise ReconciliationError('结束日期必须晚于开始日期（结束日不含）')
    warehouse = conn.execute('SELECT id,name FROM warehouses WHERE id=?', (d['warehouse_id'],)).fetchone()
    if not warehouse:
        raise ReconciliationError('仓库不存在')
    parties = {k: object_(conn, d[k], 'party') for k in ('our_party_id', 'supplier_id', 'provider_id')}
    if len({d[k] for k in parties}) != 3 or not parties['our_party_id']['data']['internal']:
        raise ReconciliationError('本团队、供货方、发货服务商须为三个不同主体')
    pool = object_(conn, d['pool_id'], 'pool')
    contract = object_(conn, d['contract_id'], 'contract')
    cd = contract['data']
    if pool['data']['ownership'] != {d['supplier_id']: '1'} and {p: dec(v) for p, v in pool['data']['ownership'].items()} != {d['supplier_id']: Decimal(1)}:
        raise ReconciliationError('此历史规则要求供货方100%货权')
    if cd['pool_id'] != pool['id'] or cd['template'] not in ('margin', 'self') or cd['capital_status'] != 'none':
        raise ReconciliationError('请选择实际毛利、货值另账（不产生回本应付）的合同')
    for field in ('profit', 'loss', 'shipping'):
        if {p: dec(v) for p, v in cd[field].items()} != {d['our_party_id']: Decimal(1)}:
            raise ReconciliationError('当前历史规则要求商品和运费盈亏100%归本团队')
    if cd['recognition'] != 'shipped' or d['effective_from'] < cd['effective_from'] or d['effective_to'] > cd['effective_to']:
        raise ReconciliationError('需使用出库确认合同，且规则生效期间不得超出合同')
    rates = data.get('freight', {})
    if not isinstance(rates, dict) or not rates or set(rates) - {'PLN', 'CZK', 'HUF'}:
        raise ReconciliationError('请配置 PLN/CZK/HUF 市场运费')
    d['freight'] = {}
    for unit, value in rates.items():
        if set(value) != {'outbound', 'reverse'}:
            raise ReconciliationError('每个市场须明确去程和回程费用')
        if any(dec(value[k]) < 0 for k in value):
            raise ReconciliationError('运费不得为负')
        d['freight'][unit] = {k: str(money(value[k])) for k in value}
    fee = dec(data.get('management_cny'))
    if fee < 0:
        raise ReconciliationError('管理费不得为负')
    d.update(management_cny=str(money(fee)), reship_policy='charge_except_warehouse_omission',
             freight_basis='original_order', returns_keep_fees=True, supplier_cost='pending',
             recognition='shipped_return_zero', fx_policy='month_then_previous', display_currency='CNY',
             date_basis='stored_shipped_date', warehouse_name=warehouse['name'],
             party_names={key: obj['name'] for key, obj in parties.items()},
             pool_snapshot=pool, contract_snapshot=contract)
    ident = uid()
    conn.execute('INSERT INTO rec_objects(id,kind,name,data,created_at,actor) VALUES (?,?,?,?,?,?)',
                 (ident, 'history_rule', name.strip(), dump(d), now(), str(actor)))
    audit(conn, 'history-rule:' + ident, 'history_rule_created', None, {'rule_id': ident}, actor)
    return ident


def _json(value):
    return json.loads(value or '[]', parse_float=Decimal)


def _cost_context(conn, warehouse_id):
    """Only costs assigned to this physical warehouse can value its shipments."""
    brands = []
    for row in rows(conn, 'SELECT id,name,aliases FROM brands'):
        try:
            aliases = json.loads(row['aliases'] or '[]')
        except (ValueError, TypeError):
            aliases = []
        brands.append({'id': row['id'], 'name': row['name'],
                       'patterns': [row['name'].upper()] + [str(a).upper() for a in aliases]})
    mappings = {}
    for row in rows(conn, 'SELECT raw_name,source,brand_id,series_id,puff_count,flavor FROM product_mappings'):
        mappings[(html.unescape(row['raw_name'] or ''), row['source'])] = row
    return {'brands': brands, 'series': rows(conn, 'SELECT id,brand_id,name FROM series'),
            'mappings': mappings, 'costs': rows(conn, '''SELECT id,brand_id,series_id,puff_count,flavor,
                CAST(cost_price AS TEXT) AS price,cost_currency,effective_date FROM product_costs
                WHERE warehouse_id=? ORDER BY effective_date DESC,id DESC''', (warehouse_id,))}


def _line_cost(item, source, shipped_date, context, rates):
    name = html.unescape(str(item.get('name') or ''))
    mapping = next((context['mappings'].get((name, scope)) for scope in (source, '', None)
                    if context['mappings'].get((name, scope))), None)
    if mapping:
        brand, series, puffs, flavor = (mapping[k] for k in ('brand_id','series_id','puff_count','flavor'))
    else:
        parsed = parse_product_name(name, context['brands'], context['series'])
        brand, series, puffs, flavor = (parsed.get(k) for k in ('brand_id','series_id','puffs','flavor'))
        if not puffs:
            shorthand = re.search(r'\b(\d{1,3})\s*[kK]\b', name)
            puffs = int(shorthand.group(1)) * 1000 if shorthand else None
    if not brand or not puffs:
        return None, '商品品牌或口数未识别：' + name
    matches = [r for r in context['costs'] if r['brand_id'] == brand and r['puff_count'] == puffs
               and (r['series_id'] is None or r['series_id'] == series)
               and (not r['flavor'] or str(r['flavor']).casefold() == str(flavor or '').casefold())
               and r['effective_date'] <= shipped_date]
    if not matches:
        return None, '本仓库出库日无对应成本：' + name
    matches.sort(key=lambda r: (r['series_id'] is not None, bool(r['flavor']), r['effective_date'], r['id']), reverse=True)
    chosen = matches[0]
    unit = currency(chosen['cost_currency'])
    fx = rates.get(unit)
    if not fx:
        return None, '成本币种缺少当月及历史汇率：' + unit
    qty = dec(item['quantity'])
    return {'name': name, 'quantity': str(qty), 'cost_id': chosen['id'],
            'cost_effective_date': chosen['effective_date'], 'unit_cost': chosen['price'],
            'cost_currency': unit, 'cost_fx_month': fx['month'],
            'value_cny': str(money(dec(chosen['price']) * qty * dec(fx['rate'])))}, None


def _calculate(order, rule, rate, start, end, decisions, costs=None, cost_rates=None):
    """Single original order; unsupported partial/mixed cases stay visibly pending."""
    pending = []
    all_parcels = [p for p in order['parcels'] if p.get('shipped_at') and p.get('status') not in ('cancelled', 'label_pending')]
    ours = [p for p in all_parcels if p['warehouse_id'] == rule['warehouse_id']]
    month_parcels = [p for p in ours if start <= p['shipped_at'][:10] < end]
    legacy = defaultdict(list)
    for p in order['legacy_parcels']:
        legacy[str(p['tracking_number'])].append(p)
    original, reships, seen = [], [], set()
    for p in ours:
        tracking = str(p.get('tracking_number') or '')
        if not tracking:
            pending.append('运单号缺失，无法核对补发')
        if tracking and tracking in seen:
            pending.append('重复运单记录待核实')
            continue
        seen.add(tracking)
        logs = legacy.get(tracking, [])
        if len(logs) != 1:
            pending.append('运单补发标记缺失或冲突：' + p['id'])
        if any(l.get('is_reship') for l in logs):
            reships.append(p)
        else:
            original.append(p)
    first = min((p['shipped_at'][:10] for p in original), default='')
    base = bool(first and start <= first < end)
    if not original:
        pending.append('缺少原始出库记录')
    if any(p['warehouse_id'] != rule['warehouse_id'] for p in all_parcels):
        pending.append('多仓订单：收入和费用分摊待核实')
    lines = _json(order['line_items'])
    line_qty = {str(i['id']): dec(i['quantity']) for i in lines}
    quantities = defaultdict(Decimal)
    for item in order['shipped_items']:
        if any(p['id'] == item['shipment_id'] for p in original if start <= p['shipped_at'][:10] < end):
            quantities[str(item['woo_line_item_id'])] += dec(item['quantity'])
    if base and dict(quantities) != line_qty:
        pending.append('本期原始出库商品与整单数量不一致，收入待核实')
    if not base and any(p in original for p in month_parcels):
        pending.append('跨月分批出库，收入归属待核实')
    returned = bool(order['is_undelivered']) or bool(original and all(p['status'] == 'returned' for p in original))
    if not returned and any(p['status'] == 'returned' for p in original):
        pending.append('部分包裹退回，收入待核实')
    if order['is_problem_return']:
        pending.append('问题退货/货损待核实')
    if _json(order['refunds']) and not returned:
        pending.append('退款金额及口径待核实')
    goods = sum((dec(i['total']) for i in lines), Decimal(0))
    shipping = dec(order['shipping_total'])
    other = sum((dec(i.get('total', 0)) for i in _json(order['fee_lines'])), Decimal(0))
    tax = dec(order['total_tax'])
    if abs(goods + shipping + other + tax - dec(order['total'])) > Decimal('.02'):
        pending.append('订单金额守恒待核实')
    if not base or returned:
        goods = shipping = other = tax = Decimal(0)
    freight_rule = rule['freight'][order['currency']]
    freight = (dec(freight_rule['outbound']) + (dec(freight_rule['reverse']) if returned else 0)) if base else Decimal(0)
    fee = dec(rule['management_cny']) * int(base)
    reship_details = []
    for p in reships:
        if not start <= p['shipped_at'][:10] < end:
            continue
        decision = decisions.get(p['id'], {})
        reason = decision.get('reason', 'unknown')
        if reason not in ('ordinary', 'warehouse_omission', 'unknown'):
            raise ReconciliationError('补发原因无效')
        evidence = str(decision.get('evidence', '')).strip()
        if reason != 'unknown' and not evidence:
            raise ReconciliationError('补发分类需要填写依据')
        if reason == 'ordinary':
            fee += dec(rule['management_cny'])
        if reason == 'unknown':
            pending.append('补发原因待核实：' + p['id'])
        reship_details.append({'shipment_id': p['id'], 'tracking_number': p['tracking_number'],
                               'reason': reason, 'evidence': evidence, 'fee_cny': str(money(rule['management_cny'])) if reason == 'ordinary' else ('0.00' if reason == 'warehouse_omission' else None)})
    if not rate:
        pending.append('当月及历史月份汇率缺失')
    revenue = money(goods + shipping + other + tax)
    # Any ambiguous source must not silently feed an apparently complete total.
    complete = not pending
    convert = lambda value: str(money(value * dec(rate['rate']))) if rate and complete else None
    cost_lines, cost_missing = [], []
    if base and not returned and complete:
        for item in lines:
            matched, reason = _line_cost(item, order['source'], first, costs or {'brands': [], 'series': [], 'mappings': {}, 'costs': []}, cost_rates or {})
            if reason:
                cost_missing.append(reason)
            else:
                cost_lines.append(matched)
    known_cny = money(sum((dec(line['value_cny']) for line in cost_lines), Decimal(0)))
    known_native = money(known_cny / dec(rate['rate'])) if rate and complete else None
    cost_ready = complete and (returned or not base or not cost_missing) and bool(rate)
    goods_value = known_native if cost_ready else None
    goods_value_cny = known_cny if cost_ready else None
    product_profit = money(goods - goods_value) if goods_value is not None else None
    contribution = money(revenue - goods_value - freight - fee / dec(rate['rate'])) if goods_value is not None else None
    return {'order_id': order['id'], 'number': order['number'], 'site': order['source'], 'currency': order['currency'],
            'shipped_at': min(p['shipped_at'] for p in month_parcels), 'quantity': str(sum(quantities.values(), Decimal(0))),
            'state': '拒收退回' if returned else ('已签收' if all(p['status'] == 'delivered' for p in original) else '已出库，未确认签收'),
            'returned': returned, 'base_order': base, 'woo_status': order['status'], 'rate': rate,
            'goods_income': str(money(goods)) if complete else None, 'shipping_income': str(money(shipping)) if complete else None,
            'tax': str(money(tax)) if complete else None, 'other_income': str(money(other)) if complete else None,
            'revenue': str(revenue) if complete else None, 'revenue_cny': convert(revenue),
            'freight': str(money(freight)) if complete else None, 'freight_cny': convert(freight),
            'management_cny': str(money(fee)) if complete else None,
            'shipping_net': str(money(shipping - freight)) if complete else None,
            'supplier_goods_value': str(goods_value) if goods_value is not None else None,
            'supplier_goods_value_cny': str(goods_value_cny) if goods_value_cny is not None else None,
            'known_goods_value': str(known_native) if known_native is not None else None,
            'known_goods_value_cny': str(known_cny) if complete else None,
            'product_profit': str(product_profit) if product_profit is not None else None,
            'contribution_profit': str(contribution) if contribution is not None else None,
            'cost_matched_quantity': str(sum((dec(line['quantity']) for line in cost_lines), Decimal(0))),
            'cost_missing_quantity': str(sum((dec(item['quantity']) for item in lines), Decimal(0)) - sum((dec(line['quantity']) for line in cost_lines), Decimal(0))) if base and not returned and complete else '0',
            'cost_lines': cost_lines, 'cost_missing': cost_missing,
            'cost_status': ('出库待核实' if not complete else '退回不结商品货值' if returned else '补发不重复计货值' if not base else '部分成本待补' if cost_missing else '本仓成本已匹配'),
            'collection_status': '回款待凭证核对', 'pending': list(dict.fromkeys(pending)), 'reships': reship_details}


def preview(conn, rule_id, month, unit, decisions=None):
    rule = object_(conn, rule_id, 'history_rule')
    cfg = rule['data']
    start, end = month_bounds(month)
    if start < cfg['effective_from'] or end > cfg['effective_to']:
        raise ReconciliationError('月份须完整处于规则有效期内')
    if unit not in cfg['freight']:
        raise ReconciliationError('所选市场未配置物流费')
    decisions = decisions or {}
    rate = rate_at(conn, unit, month)
    orders = rows(conn, '''SELECT DISTINCT o.id,o.number,o.source,o.status,o.currency,
        o.is_undelivered,o.is_problem_return,o.line_items,o.fee_lines,o.refunds,
        CAST(o.total AS TEXT) AS total,CAST(o.shipping_total AS TEXT) AS shipping_total,
        CAST(o.total_tax AS TEXT) AS total_tax
        FROM orders o JOIN oms_fulfillments f ON f.order_id=o.id
        JOIN oms_shipments s ON s.fulfillment_id=f.id
        WHERE f.warehouse_id=? AND substr(s.shipped_at,1,10)>=? AND substr(s.shipped_at,1,10)<?
        AND s.status NOT IN ('cancelled','label_pending') AND o.currency=? ORDER BY o.id''',
        (cfg['warehouse_id'], start, end, unit))
    if len(orders) > 2000:
        raise ReconciliationError('本次超过2000单，请缩小仓库范围')
    cost_context = _cost_context(conn, cfg['warehouse_id'])
    cost_rates = {unit: rate_at(conn, unit, month) for unit in
                  {currency(row['cost_currency']) for row in cost_context['costs']}}
    details, known_ids = [], set()
    for order in orders:
        order['parcels'] = rows(conn, '''SELECT s.id,s.tracking_number,s.shipped_at,s.status,f.warehouse_id
            FROM oms_shipments s JOIN oms_fulfillments f ON f.id=s.fulfillment_id WHERE f.order_id=?''', (order['id'],))
        order['legacy_parcels'] = rows(conn, '''SELECT tracking_number,shipped_at,is_reship,reship_reason
            FROM shipping_logs WHERE order_id=? AND status<>'pending_sync' ''', (order['id'],))
        order['shipped_items'] = rows(conn, '''SELECT s.id AS shipment_id,oi.woo_line_item_id,si.quantity
            FROM oms_shipment_items si JOIN oms_shipments s ON s.id=si.shipment_id
            JOIN oms_fulfillment_items fi ON fi.id=si.fulfillment_item_id
            JOIN oms_order_items oi ON oi.id=fi.order_item_id
            JOIN oms_fulfillments f ON f.id=fi.fulfillment_id WHERE f.order_id=? AND f.warehouse_id=?''', (order['id'], cfg['warehouse_id']))
        detail = _calculate(order, cfg, rate, start, end, decisions, cost_context, cost_rates)
        details.append(detail)
        known_ids.update(r['shipment_id'] for r in detail['reships'])
    if set(decisions) - known_ids:
        raise ReconciliationError('补发分类包含本期范围外的运单')
    details.sort(key=lambda r: (r['shipped_at'], r['order_id']))
    fields = ('goods_income','shipping_income','tax','other_income','revenue','revenue_cny','freight','freight_cny','management_cny','shipping_net')
    totals = {k: str(money(sum((dec(r[k]) for r in details if r[k] is not None), Decimal(0)))) for k in fields}
    for key in ('known_goods_value', 'known_goods_value_cny'):
        totals[key] = str(money(sum((dec(r[key]) for r in details if r[key] is not None), Decimal(0))))
    all_costs_known = all(r['supplier_goods_value'] is not None for r in details)
    for key in ('supplier_goods_value', 'supplier_goods_value_cny', 'product_profit', 'contribution_profit'):
        totals[key] = str(money(sum((dec(r[key]) for r in details), Decimal(0)))) if all_costs_known else None
    incomplete = sum(bool(r['pending']) for r in details)
    cost_pending = sum(r['supplier_goods_value'] is None for r in details if r['base_order'] and not r['returned'])
    result = {'version': 2, 'rule_id': rule_id, 'rule_snapshot': rule, 'month': month, 'currency': unit,
            'generated_at': now(), 'status': 'draft', 'date_basis': '保存的出库日期字段，不变更历史时区口径',
            'state_basis': '读取时最新退回状态，非月末快照', 'rows': details, 'rate': rate,
            'counts': {'orders': len(details), 'returns': sum(r['returned'] for r in details),
                       'income_orders': sum(r['base_order'] and not r['returned'] for r in details),
                       'pending_orders': incomplete, 'cost_pending_orders': cost_pending,
                       'cost_matched_quantity': str(sum((dec(r['cost_matched_quantity']) for r in details), Decimal(0))),
                       'cost_missing_quantity': str(sum((dec(r['cost_missing_quantity']) for r in details), Decimal(0)))},
            'totals': totals, 'totals_label': '已核实行合计（待核实行不计入）' if incomplete else '本期合计',
            'supplier_statement': {'party_id': cfg['supplier_id'], 'goods_value': totals['supplier_goods_value'],
                                   'goods_value_cny': totals['supplier_goods_value_cny'],
                                   'known_subtotal_cny': totals['known_goods_value_cny'],
                                   'status': '部分成本待补，已匹配小计不可当应付总额' if cost_pending else '系统成本参考值，待供货方凭证确认'},
            'provider_statement': {'party_id': cfg['provider_id'], 'freight': totals['freight'], 'currency': unit,
                                   'management_cny': totals['management_cny'], 'status': '待核对草稿'},
            'can_lock': False, 'blocking': (['商品成本未完全匹配'] if cost_pending else []) +
                                    ['供货方价格与货值待凭证确认', '回款未核销', '历史草稿不向正式权益账重复记账']}
    stable = {k: v for k, v in result.items() if k != 'generated_at'}
    result['digest'] = hashlib.sha256(dump(stable).encode()).hexdigest()
    return result


def create_draft(conn, request, actor):
    key = request.get('request_key', '')
    if not re.fullmatch(r'[A-Za-z0-9_-]{16,100}', key):
        raise ReconciliationError('生成草稿需要稳定的请求编号')
    ident = 'history-draft-' + hashlib.sha256(key.encode()).hexdigest()[:32]
    inputs = {k: request.get(k) for k in ('rule_id', 'month', 'currency', 'reship_decisions', 'expected_digest')}
    existing = conn.execute('SELECT data FROM rec_objects WHERE id=?', (ident,)).fetchone()
    if existing:
        old = json.loads(existing['data'])
        if old['request'] != inputs:
            raise ReconciliationError('请求编号已用于另一份草稿')
        return ident
    snap = preview(conn, inputs['rule_id'], inputs['month'], inputs['currency'], inputs['reship_decisions'])
    if inputs['expected_digest'] != snap['digest']:
        raise ReconciliationError('出库、退回、费用分类或汇率已变化，请重新预览核对')
    if not snap['rows']:
        raise ReconciliationError('本期没有出库记录，不生成空草稿')
    payload = {'request': inputs, 'snapshot': snap}
    name = snap['month'] + ' · ' + snap['rule_snapshot']['name'] + ' · ' + snap['currency']
    conn.execute('INSERT INTO rec_objects(id,kind,name,data,created_at,actor) VALUES (?,?,?,?,?,?) ON CONFLICT(id) DO NOTHING',
                 (ident, 'history_draft', name, dump(payload), now(), str(actor)))
    stored = object_(conn, ident, 'history_draft')['data']
    if stored['request'] != inputs:
        raise ReconciliationError('并发请求编号冲突')
    return ident
