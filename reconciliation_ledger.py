"""Order economics, carrier evidence, immutable differences and cash clearing."""
import hashlib
import json
from collections import defaultdict
from decimal import Decimal
from datetime import datetime
from reconciliation_core import *


def entry(conn, key, party, order, category, amount, unit, occurred, data):
    amount = money(amount)
    object_(conn, party, 'party')
    ident = uid()
    conn.execute('''INSERT INTO rec_entries(id,event_key,party_id,order_id,site,category,amount,currency,occurred_at,data)
        VALUES (?,?,?,?,?,?,?,?,?,?)''', (ident, key, party, order['id'] if order else None,
        order['source'] if order else None, category, str(amount), currency(unit), occurred, dump(data)))
    return ident


def shipments(conn, order_id):
    return rows(conn, '''SELECT DISTINCT s.* FROM oms_shipments s
        JOIN oms_shipment_items si ON si.shipment_id=s.id
        JOIN oms_fulfillment_items fi ON fi.id=si.fulfillment_item_id
        JOIN oms_fulfillments f ON f.id=fi.fulfillment_id WHERE f.order_id=? ORDER BY s.id''', (order_id,))


def record_bill(conn, shipment_id, data, actor):
    shipment = conn.execute('SELECT * FROM oms_shipments WHERE id=?' + lock(conn), (shipment_id,)).fetchone()
    if not shipment:
        raise ReconciliationError('运单不存在')
    orders = rows(conn, '''SELECT DISTINCT f.order_id FROM oms_fulfillment_items i
        JOIN oms_shipment_items si ON si.fulfillment_item_id=i.id JOIN oms_fulfillments f ON f.id=i.fulfillment_id
        WHERE si.shipment_id=?''', (shipment_id,))
    if len(orders) != 1 or not order_config(conn, orders[0]['order_id']):
        raise ReconciliationError('账单必须匹配一笔已启用货权的站点订单')
    object_(conn, data.get('provider_id'), 'party')
    currency(data.get('currency'))
    for field in ('outbound', 'reverse', 'collection_fee', 'collected', 'remitted'):
        if dec(data.get(field)) < 0:
            raise ReconciliationError('物流账单金额不能为负')
        data[field] = str(money(data[field]))
    if not data.get('reference') or not data.get('proof'):
        raise ReconciliationError('账单编号和凭证必填')
    for field, amount in [('collected_at', 'collected'), ('remitted_at', 'remitted')]:
        if dec(data[amount]) and not data.get(field):
            raise ReconciliationError('收款/回款必须记录实际日期')
        if data.get(field):
            datetime.fromisoformat(data[field])
    key = 'bill:' + shipment_id + ':' + str(data['reference'])
    if done(conn, key):
        old = conn.execute('SELECT data FROM rec_actions WHERE id=?', (key,)).fetchone()
        if json.loads(old['data']) != {**data, 'shipment_id': shipment_id}:
            raise ReconciliationError('相同账单编号不能覆盖原凭证，请创建修订凭证')
        return
    data = {**data, 'shipment_id': shipment_id}
    audit(conn, key, 'bill_verified', orders[0]['order_id'], data, actor)


def latest_bills(conn, order_id):
    bills = {}
    for r in rows(conn, "SELECT * FROM rec_actions WHERE order_id=? AND kind='bill_verified' ORDER BY created_at,id", (order_id,)):
        data = json.loads(r['data'])
        bills[data['shipment_id']] = data
    return bills


def economics(conn, order_id):
    cfg = order_config(conn, order_id)
    if not cfg:
        raise ReconciliationError('订单未启用模块化对账')
    order = dict(conn.execute('SELECT * FROM orders WHERE id=?', (order_id,)).fetchone())
    monetary_columns = [k for k in ('total','shipping_total','shipping_tax') if k in order]
    exact = conn.execute('SELECT '+','.join('CAST('+k+' AS TEXT) AS '+k for k in monetary_columns)+' FROM orders WHERE id=?',(order_id,)).fetchone()
    order.update(dict(exact))
    source_rows = rows(conn, '''SELECT s.*,i.order_item_id,o.raw_json,o.ordered_qty,o.sku_id
        FROM rec_sources s JOIN oms_fulfillment_items i ON i.id=s.fulfillment_item_id
        JOIN oms_order_items o ON o.id=i.order_item_id WHERE s.order_id=? AND s.shipped>0 ORDER BY s.id''', (order_id,))
    pending = []
    if not source_rows:
        pending.append('尚无已出库来源')
    unit = currency(cfg.get('currency', order['currency']))
    corrections = conn.execute("SELECT data FROM rec_actions WHERE order_id=? AND kind='financial_correction' ORDER BY created_at DESC,id DESC LIMIT 1", (order_id,)).fetchone()
    correction = json.loads(corrections['data']) if corrections else {}
    cfg = {**cfg, **correction.get('values', {})}
    fx = cfg.get('fx', {})
    totals = defaultdict(Decimal)
    entry_dates = {}
    source_details = []
    split_signature = None
    parcel_rows = shipments(conn, order_id)
    bills = latest_bills(conn, order_id)
    source_dates = {}
    bill_dates = {}
    def timestamp(value):
        parsed = datetime.fromisoformat(value)
        return (parsed.replace(tzinfo=timezone.utc) if parsed.tzinfo is None else parsed).astimezone(timezone.utc).isoformat()
    for event in rows(conn,"SELECT kind,data,created_at FROM rec_actions WHERE order_id=? AND kind IN ('source_shipped','bill_verified') ORDER BY created_at,id",(order_id,)):
        payload=json.loads(event['data'])
        if event['kind']=='source_shipped':
            parcel=next((p for p in parcel_rows if p['id']==payload['shipment_id']),None)
            shipped_at=timestamp(parcel['shipped_at'] if parcel and parcel.get('shipped_at') else event['created_at'])
            source_dates[payload['source_id']]=max(source_dates.get(payload['source_id'],shipped_at),shipped_at)
        else:
            bill_dates[payload['shipment_id']]=timestamp(event['created_at'])
    business_date = max(source_dates.values(), default=now())
    if any(p['id'] not in bills for p in parcel_rows):
        pending.append('实际物流账单缺失（零费用也需确认）')
    active_f = rows(conn, "SELECT status FROM oms_fulfillments WHERE order_id=? AND status NOT IN ('cancelled','superseded')", (order_id,))
    if any(r['status'] not in {'shipped', 'delivered', 'returned'} for r in active_f):
        pending.append('订单仍有未完成出库来源')
    if conn.execute('SELECT id FROM oms_order_items WHERE order_id=? AND shortage_qty>0', (order_id,)).fetchone():
        pending.append('订单仍有缺货')
    def cv(value, native):
        try:
            return convert(value, native, unit, fx)
        except ReconciliationError:
            pending.append('汇率缺失：' + native + '/' + unit)
            return Decimal(0)
    def add(party, category, amount, curr=unit, occurred=None):
        key=(party, category, curr)
        totals[key] += money(amount)
        date_at=timestamp(occurred or business_date)
        entry_dates[key]=max(entry_dates.get(key,date_at),date_at)
    raw_lines = json.loads(order.get('line_items') or '[]')
    net_goods = sum((dec(x.get('total')) for x in raw_lines), Decimal(0))
    gross_goods = sum((dec(x.get('total')) + dec(x.get('total_tax', 0)) for x in raw_lines), Decimal(0))
    customer_shipping = dec(order.get('shipping_total'))
    shipping_tax = dec(order.get('shipping_tax') or 0)
    fees_raw = json.loads(order.get('fee_lines') or '[]')
    shipping_fees = set(str(x) for x in cfg.get('customer_shipping_fee_ids', []))
    customer_fee_net = sum((dec(f.get('total')) for f in fees_raw if str(f.get('id')) in shipping_fees), Decimal(0))
    fee_gross = sum((dec(f.get('total')) + dec(f.get('total_tax', 0)) for f in fees_raw), Decimal(0))
    if money(gross_goods + customer_shipping + shipping_tax + fee_gross) != money(order['total']):
        pending.append('订单商品、运费、税和附加费与总额不守恒，请核对退款/折扣映射')
    if any(str(f.get('id')) not in shipping_fees and dec(f.get('total')) != 0 for f in fees_raw):
        pending.append('存在未分类的订单附加费')
    refunds = cfg.get('refunds', {})
    refunds_already_netted = cfg.get('refunds_in_order_totals') is True
    # Explicit net refunds, not a second deduction from a previously netted total.
    if (order.get('status') == 'refunded' or json.loads(order.get('refunds') or '[]')) and not refunds:
        pending.append('退款订单缺少已核实的净收入映射')
    for s in source_rows:
        snap = json.loads(s['snapshot'])
        terms = snap['contract']['data']
        raw = json.loads(s['raw_json'])
        recognition = terms['recognition']
        if recognition == 'delivered' and any(p['status'] != 'delivered' for p in parcel_rows):
            pending.append('合同要求签收后确认利润')
        if recognition == 'bill_verified' and any(p['id'] not in bills for p in parcel_rows):
            pending.append('合同要求账单核实后确认利润')
        if s['id'] in cfg.get('source_costs', {}):
            snap['unit_cost'] = cfg['source_costs'][s['id']]
        if snap['unit_cost'] is None:
            pending.append('批次成本缺失：' + str(s['batch_id']))
        refunded = dec(refunds.get(str(raw['id']), '0'))
        if refunded < 0 or (not refunds_already_netted and refunded > dec(raw['total'])):
            pending.append('退款金额超出原商品金额')
        revenue = cv((dec(raw['total']) - (Decimal(0) if refunds_already_netted else refunded)) * Decimal(s['shipped']) / Decimal(s['ordered_qty']), order['currency'])
        salable = 0
        for r in rows(conn, "SELECT data FROM rec_actions WHERE order_id=? AND kind='source_returned'", (order_id,)):
            ret = json.loads(r['data'])
            if ret['source_id'] == s['id'] and ret['salable']:
                salable += ret['quantity']
        cost = cv(dec(snap['unit_cost'] or 0) * (s['shipped'] - salable), snap['currency'])
        weight = dec(raw.get('weight', 0)) * s['shipped']
        recognition_date = source_dates.get(s['id'],business_date)
        if recognition == 'delivered':
            recognition_date = max((timestamp(p['delivered_at']) for p in parcel_rows if p['delivered_at']),default=recognition_date)
        elif recognition == 'bill_verified':
            recognition_date = max(bill_dates.values(),default=recognition_date)
        source_details.append({'recognized_at': recognition_date, 'id': s['id'], 'quantity': s['shipped'], 'weight': weight,
             'revenue': revenue, 'cost': cost, 'snapshot': snap, 'terms': terms})
    by_line = defaultdict(list)
    for source, detail in zip(source_rows, source_details):
        by_line[source['order_item_id']].append((source, detail))
    for group in by_line.values():
        raw = json.loads(group[0][0]['raw_json'])
        sold = sum(s['shipped'] for s, _ in group)
        ordered = group[0][0]['ordered_qty']
        revenue = cv((dec(raw['total']) - (Decimal(0) if refunds_already_netted else dec(refunds.get(str(raw['id']), '0')))) * Decimal(sold) / Decimal(ordered), order['currency'])
        parts = allocate(revenue, {s['id']: s['shipped'] for s,_ in group})
        for source, detail in group:
            detail['revenue'] = parts[source['id']]
    # Freeze fee rules selected at reservation; a new version never creates another charge.
    fee_cost = Decimal(0)
    fee_details = []
    seen_fee = set()
    warehouses = {json.loads(s['snapshot'])['warehouse_id'] for s in source_rows}
    pools = {json.loads(s['snapshot'])['pool']['id'] for s in source_rows}
    site_country = conn.execute('SELECT country FROM sites WHERE url=?', (order['source'],)).fetchone()
    for fee_id in cfg.get('fees', []):
        rule = object_(conn, fee_id, 'fee')
        fee = rule['data']; provider = fee['provider_id']
        scope = fee['scope']
        if scope.get('site') and scope['site'] != order['source']:
            continue
        if scope.get('country') and (not site_country or scope['country'] != site_country['country']):
            continue
        if scope.get('warehouse_id') and int(scope['warehouse_id']) not in warehouses:
            continue
        if scope.get('pool_id') and scope['pool_id'] not in pools:
            continue
        if scope.get('channel') and scope['channel'] != cfg.get('channel'):
            continue
        date_at = min((json.loads(s['snapshot'])['reserved_at'][:10] for s in source_rows), default=now()[:10])
        if not fee['effective_from'] <= date_at < fee['effective_to']:
            continue
        if cfg['provider_mode'] == 'primary' and provider != cfg['primary_provider']:
            continue
        identity = provider
        if identity in seen_fee:
            pending.append('同一服务商存在重叠收费规则，请明确版本')
            continue
        seen_fee.add(identity)
        if cfg['provider_mode'] == 'split':
            signature = (money(fee['amount']), fee['currency'], fee['basis'], fee['event'])
            if split_signature is not None and signature != split_signature:
                pending.append('总额分摊模式必须使用相同金额、币种、单位和时点')
            split_signature = signature
        eligible = []
        for parcel in parcel_rows:
            parts = rows(conn, '''SELECT f.warehouse_id,s.snapshot FROM oms_shipment_items si
                JOIN oms_fulfillment_items i ON i.id=si.fulfillment_item_id
                JOIN oms_fulfillments f ON f.id=i.fulfillment_id
                JOIN rec_sources s ON s.fulfillment_item_id=i.id WHERE si.shipment_id=?''', (parcel['id'],))
            if any((not scope.get('warehouse_id') or int(scope['warehouse_id']) == r['warehouse_id'])
                   and (not scope.get('pool_id') or json.loads(r['snapshot'])['pool']['id'] == scope['pool_id']) for r in parts):
                eligible.append(parcel)
        if fee['event'] == 'delivered':
            eligible = [p for p in eligible if p['status'] == 'delivered']
        elif fee['event'] == 'bill_verified':
            eligible = [p for p in eligible if p['id'] in bills]
        count = len(eligible) if fee['basis'] == 'parcel' else int(bool(eligible))
        reships = set(cfg.get('reshipment_ids', []))
        if fee['reship_policy'] == 'free' and fee['basis'] == 'parcel':
            count -= sum(p['id'] in reships for p in eligible)
        if fee['return_policy'] == 'reverse' and source_rows and all(s['returned'] == s['shipped'] for s in source_rows):
            count = 0
        amount = money(dec(fee['amount']) * count)
        if cfg['provider_mode'] == 'split':
            amount = allocate(amount, cfg['provider_weights']).get(provider, Decimal(0))
        fee_cost += cv(amount, fee['currency'])
        if count and not object_(conn, provider, 'party')['data']['internal']:
            dates = [bill_dates[p['id']] if fee['event']=='bill_verified' else timestamp(p.get('delivered_at') if fee['event']=='delivered' else p.get('shipped_at') or business_date) for p in eligible]
            fee_date = (min(dates) if fee['basis']=='order' else max(dates)) if dates else business_date
            add(provider, 'service', amount, fee['currency'], fee_date)
        fee_details.append({'provider': provider, 'rule_id': fee_id, 'amount': str(amount), 'currency': fee['currency']})
    if source_details and cfg.get('fees') and not seen_fee:
        pending.append('没有匹配本单的有效管理费规则')
    direct_goods = Decimal(0)
    direct_shipping = Decimal(0)
    for expense_ref, expense in cfg.get('expenses', {}).items():
        converted = cv(expense['amount'], expense['currency'])
        if expense['kind'] == 'shipping':
            direct_shipping += converted
        else:
            direct_goods += converted
        if not object_(conn,expense['party_id'],'party')['data']['internal']:
            add(expense['party_id'],'other_expense',expense['amount'],expense['currency'])
    if source_details:
        allocation_modes = {s['terms'].get('allocation', 'goods') for s in source_details}
        if len(allocation_modes) != 1:
            pending.append('同一订单的共同费用分摊方式必须一致')
        mode = sorted(allocation_modes)[0]
        weights = {s['id']: max(Decimal(0), s[{'goods': 'revenue', 'quantity': 'quantity', 'weight': 'weight'}[mode]]) for s in source_details}
        if not sum(weights.values()):
            if mode == 'weight':
                pending.append('重量分摊缺少重量')
            weights = {s['id']: s['quantity'] for s in source_details}
        charges = allocate(fee_cost, weights)
        direct_charges = allocate(direct_goods, weights)
        for s in source_details:
            terms, snap = s['terms'], s['snapshot']
            business_date = s['recognized_at']
            cost = s['cost']
            capital = cost
            if terms['template'] == 'fixed':
                capital = cv(dec(terms['supply_price']) * s['quantity'], terms['currency'])
            profit = s['revenue'] - capital - charges[s['id']] - direct_charges[s['id']]
            if terms['template'] == 'legacy':
                # Legacy ratio is a replacement entitlement, never stacked with capital repayment.
                entitlement = money(s['revenue'] * dec(terms['legacy_rate']))
                for p, amount in allocate(entitlement, terms['profit']).items():
                    add(p, 'legacy', amount)
                add(cfg['our_party_id'], 'profit', profit - entitlement)
            else:
                if terms['capital_status'] == 'unpaid':
                    for p, amount in allocate(capital, snap['pool']['data']['capital']).items():
                        add(p, 'capital', amount)
                for p, amount in allocate(profit, terms['loss'] if profit < 0 else terms['profit']).items():
                    add(p, 'profit', amount)
            s['service'] = charges[s['id']]
            s['other_cost'] = direct_charges[s['id']]
            s['profit'] = profit
        business_date = max((s['recognized_at'] for s in source_details),default=business_date)
        shipping_cost = sum((cv(dec(b['outbound']) + dec(b['reverse']) + dec(b['collection_fee']), b['currency']) for b in bills.values()), Decimal(0))
        ship_net = cv(customer_shipping + customer_fee_net - (Decimal(0) if refunds_already_netted else dec(refunds.get('shipping', 0))), order['currency']) - shipping_cost - direct_shipping
        shipping_rules = {dump(s['terms']['shipping']) for s in source_details}
        if len(shipping_rules) != 1:
            pending.append('共同包裹运费归属规则冲突')
        else:
            for p, amount in allocate(ship_net, json.loads(next(iter(shipping_rules)))).items():
                add(p, 'shipping_profit', amount)
    for b in bills.values():
        # Carrier cost is payable; collections and remittances remain separate from service/profit.
        actual_cost = dec(b['outbound']) + dec(b['reverse']) + dec(b['collection_fee'])
        bill_date=bill_dates.get(b['shipment_id'],business_date)
        add(b['provider_id'], 'carrier', actual_cost, b['currency'],bill_date)
        add(b['provider_id'], 'collection_due', -dec(b['collected']), b['currency'],b.get('collected_at') or bill_date)
        if dec(b['remitted']):
            add(b['provider_id'], 'collection_remitted', dec(b['remitted']), b['currency'],b.get('remitted_at') or bill_date)
    return {'order_id': order_id, 'site': order['source'], 'currency': unit,
            'pending': sorted(set(pending)), 'totals': [{'party_id': p, 'category': c, 'currency': u, 'amount': str(money(a)), 'occurred_at': entry_dates[(p,c,u)]}
                for (p,c,u),a in sorted(totals.items())], 'sources': source_details,
            'fees': fee_details, 'bills': list(bills.values()), 'customer_shipping': str(customer_shipping),
            'expected_collection': str(money(order['total'])) if order.get('payment_method') == 'cod' else '0.00',
            'collection_variance': str(money(sum((convert(b['collected'],b['currency'],order['currency'],fx) for b in bills.values()), Decimal(0)) - (dec(order['total']) if order.get('payment_method') == 'cod' else Decimal(0)))) if all(b['currency']==order['currency'] or b['currency']+'/'+order['currency'] in fx for b in bills.values()) else None,
            'profit_type': 'operating_confirmed' if cfg.get('operating_costs_complete') else 'contribution', 'fx': fx, 'expenses': cfg.get('expenses', {})}


def serializable(value):
    return json.loads(json.dumps(value, ensure_ascii=False, default=str))


def recognize(conn, order_id, actor):
    conn.execute('SELECT id FROM orders WHERE id=?' + lock(conn), (order_id,)).fetchone()
    result = serializable(economics(conn, order_id))
    if result['pending']:
        raise ReconciliationError('待补全：' + '；'.join(result['pending']))
    digest = hashlib.sha256(dump(result).encode()).hexdigest()
    previous = conn.execute("SELECT id,data FROM rec_actions WHERE order_id=? AND kind='economic_recognized' ORDER BY created_at DESC,id DESC LIMIT 1", (order_id,)).fetchone()
    if previous and dump(json.loads(previous['data'])) == dump(result):
        return result
    chain = hashlib.sha256((digest + (previous['id'] if previous else '')).encode()).hexdigest()
    key = 'economics:' + order_id + ':' + chain
    prior = defaultdict(Decimal)
    for row in rows(conn, 'SELECT * FROM rec_entries WHERE order_id=?', (order_id,)):
        if json.loads(row['data']).get('economic_snapshot'):
            prior[(row['party_id'], row['category'], row['currency'])] += dec(row['amount'])
    current = {(r['party_id'], r['category'], r['currency']): dec(r['amount']) for r in result['totals']}
    dates = {(r['party_id'], r['category'], r['currency']): r['occurred_at'] for r in result['totals']}
    order = dict(conn.execute('SELECT * FROM orders WHERE id=?', (order_id,)).fetchone())
    for p,c,u in sorted(set(prior) | set(current)):
        delta = current.get((p,c,u), Decimal(0)) - prior.get((p,c,u), Decimal(0))
        if delta or (p,c,u) not in prior:
            entry(conn, key, p, order, c, delta, u, now() if (p,c,u) in prior else dates.get((p,c,u),now()), {'economic_snapshot': digest, 'fx': result['fx'],
                  'adjustment': bool(prior), 'source_totals': result['totals']})
    audit(conn, key, 'economic_recognized', order_id, result, actor)
    return result


def adjustment(conn, party, amount, unit, category, reference, reason, actor):
    if not reference or not reason or category not in {'capital', 'profit', 'service', 'carrier', 'collection_due', 'advance', 'opening'}:
        raise ReconciliationError('调整需有效类别、唯一凭证和原因')
    key = 'adjustment:' + reference
    if done(conn, key):
        raise ReconciliationError('该调整凭证已经登记')
    entry(conn, key, party, None, category, amount, unit, now(), {'reason': reason, 'reference': reference})
    audit(conn, key, 'adjustment', None, {'party_id': party, 'amount': str(money(amount)), 'reason': reason}, actor)


def create_statement(conn, party, unit, start, end, bucket, actor, timezone_name="UTC"):
    object_(conn, party, 'party')
    currency(unit)
    datetime.fromisoformat(start); datetime.fromisoformat(end)
    if start >= end or bucket not in {'trade', 'collection'}:
        raise ReconciliationError('期间或对账类别无效，使用 UTC 含起日不含止日')
    from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
    try:
        zone = ZoneInfo(timezone_name)
    except ZoneInfoNotFoundError as exc:
        raise ReconciliationError('不支持的时区') from exc
    def cutoff(value):
        value = datetime.fromisoformat(value)
        return (value.replace(tzinfo=zone) if value.tzinfo is None else value).astimezone(timezone.utc).isoformat()
    utc_start, utc_end = cutoff(start), cutoff(end)
    if utc_start >= utc_end:
        raise ReconciliationError('结算截止时间必须晚于开始时间')
    # Lock party to serialize statements and prevent the same entitlement appearing twice.
    conn.execute('SELECT id FROM rec_objects WHERE id=?' + lock(conn), (party,)).fetchone()
    candidates_ = rows(conn, '''SELECT * FROM rec_entries WHERE party_id=? AND currency=?
        AND occurred_at>=? AND occurred_at<? AND statement_id IS NULL ORDER BY occurred_at,id''' + lock(conn), (party, unit, utc_start, utc_end))
    entries = [e for e in candidates_ if (e['category'].startswith('collection_')) == (bucket == 'collection')]
    if not entries:
        raise ReconciliationError('该期间没有可对账流水')
    opening = sum((dec(e['amount']) for e in rows(conn,'SELECT category,amount FROM rec_entries WHERE party_id=? AND currency=? AND occurred_at<?',(party,unit,utc_start)) if e['category'].startswith('collection_') == (bucket=='collection')),Decimal(0))
    for payment in rows(conn,'SELECT a.amount,s.snapshot FROM rec_cash_allocations a JOIN rec_statements s ON s.id=a.statement_id JOIN rec_cash c ON c.id=a.cash_id WHERE s.party_id=? AND s.currency=? AND c.occurred_at<?',(party,unit,utc_start)):
        if json.loads(payment['snapshot'])['bucket']==bucket:
            opening -= dec(payment['amount'])
    period_activity = sum((dec(e['amount']) for e in rows(conn,'SELECT category,amount FROM rec_entries WHERE party_id=? AND currency=? AND occurred_at>=? AND occurred_at<?',(party,unit,utc_start,utc_end)) if e['category'].startswith('collection_') == (bucket=='collection')),Decimal(0))
    statement_id = uid()
    snapshot = {'entries': [{**e, 'data': json.loads(e['data'])} for e in entries],
                'total': str(sum((dec(e['amount']) for e in entries), Decimal(0))), 'bucket': bucket,
                'timezone': timezone_name, 'utc_start': utc_start, 'utc_end': utc_end,
                'opening_outstanding': str(money(opening)),
                'closing_before_payment': str(money(opening + period_activity)),
                'actor': str(actor)}
    conn.execute('INSERT INTO rec_statements(id,party_id,currency,period_start,period_end,status,snapshot,created_at) VALUES (?,?,?,?,?,?,?,?)',
                 (statement_id, party, unit, start, end, 'draft', dump(snapshot), now()))
    for e in entries:
        conn.execute('UPDATE rec_entries SET statement_id=? WHERE id=? AND statement_id IS NULL', (statement_id, e['id']))
    return statement_id


def statement_transition(conn, ident, target, actor, reason=''):
    row = conn.execute('SELECT * FROM rec_statements WHERE id=?' + lock(conn), (ident,)).fetchone()
    if not row:
        raise ReconciliationError('对账单不存在')
    allowed = {'draft': {'verified'}, 'verified': {'confirmed', 'disputed'}, 'disputed': {'verified'}, 'confirmed': {'locked', 'disputed'}}
    if target not in allowed.get(row['status'], set()):
        raise ReconciliationError('不能执行该状态变更；锁单后只能追加差额')
    if target in {'verified','confirmed','locked'}:
        snap = json.loads(row['snapshot'])
        unresolved = pending_statement_orders(conn,row['party_id'],snap.get('utc_start',row['period_start']),snap.get('utc_end',row['period_end']))
        if unresolved:
            raise ReconciliationError('期间仍有待补全/待记账订单：' + '、'.join(r['order_id'] for r in unresolved))
    if target == 'disputed' and not reason:
        raise ReconciliationError('争议原因必填')
    final_status = 'paid' if target == 'locked' and money(json.loads(row['snapshot'])['total']) == 0 else target
    conn.execute('UPDATE rec_statements SET status=? WHERE id=?', (final_status, ident))
    audit(conn, uid(), 'statement_' + target, None, {'statement_id': ident, 'reason': reason}, actor)


def record_cash(conn, party, unit, amount, occurred, reference, proof, actor):
    object_(conn, party, 'party'); currency(unit)
    if not reference or not proof or not money(amount):
        raise ReconciliationError('资金记录需要非零金额、凭证编号及凭证')
    occurred_value = datetime.fromisoformat(occurred)
    occurred = (occurred_value.replace(tzinfo=timezone.utc) if occurred_value.tzinfo is None else occurred_value).astimezone(timezone.utc).isoformat()
    conn.execute('SELECT id FROM rec_objects WHERE id=?'+lock(conn),(party,)).fetchone()
    existing = conn.execute('SELECT * FROM rec_cash WHERE party_id=? AND currency=? AND reference=?',(party,unit,reference)).fetchone()
    if existing:
        if money(existing['amount']) == money(amount) and existing['occurred_at'] == occurred and json.loads(existing['data']).get('proof') == proof:
            return existing['id']
        raise ReconciliationError('相同资金凭证不能覆盖原金额或日期')
    ident = uid()
    conn.execute('INSERT INTO rec_cash(id,party_id,currency,amount,occurred_at,reference,data) VALUES (?,?,?,?,?,?,?)',
                 (ident, party, unit, str(money(amount)), occurred, reference, dump({'proof': proof, 'actor': str(actor)})))
    return ident


def allocate_cash(conn, cash_id, statement_id, amount, actor, request_key=None):
    cash = conn.execute('SELECT * FROM rec_cash WHERE id=?' + lock(conn), (cash_id,)).fetchone()
    statement = conn.execute('SELECT * FROM rec_statements WHERE id=?' + lock(conn), (statement_id,)).fetchone()
    allocation_key = 'cash-allocation:' + str(request_key) if request_key else uid()
    if done(conn,allocation_key):
        previous=json.loads(conn.execute('SELECT data FROM rec_actions WHERE id=?',(allocation_key,)).fetchone()['data'])
        if previous['cash_id']!=cash_id or previous['statement_id']!=statement_id or money(previous['amount'])!=money(amount):
            raise ReconciliationError('核销请求编号与原请求不一致')
        return
    if not cash or not statement or statement['status'] not in {'locked', 'partial_paid'}:
        raise ReconciliationError('仅已锁定或部分核销的对账单可核销')
    if cash['party_id'] != statement['party_id'] or cash['currency'] != statement['currency']:
        raise ReconciliationError('禁止跨主体或币种自动核销')
    amount = money(amount)
    total = dec(json.loads(statement['snapshot'])['total'])
    used_cash = sum((dec(r['amount']) for r in rows(conn, 'SELECT amount FROM rec_cash_allocations WHERE cash_id=?', (cash_id,))), Decimal(0))
    used_statement = sum((dec(r['amount']) for r in rows(conn, 'SELECT amount FROM rec_cash_allocations WHERE statement_id=?', (statement_id,))), Decimal(0))
    if not amount or amount * total <= 0 or amount * dec(cash['amount']) <= 0:
        raise ReconciliationError('核销方向与资金/账单不一致；付出为正，收到为负')
    if abs(amount) > abs(dec(cash['amount']) - used_cash) or abs(amount) > abs(total - used_statement):
        raise ReconciliationError('核销金额超过资金余额或账单未结金额')
    existing = conn.execute('SELECT amount FROM rec_cash_allocations WHERE cash_id=? AND statement_id=?', (cash_id, statement_id)).fetchone()
    cumulative = money(amount + (dec(existing['amount']) if existing else Decimal(0)))
    conn.execute('INSERT INTO rec_cash_allocations(cash_id,statement_id,amount) VALUES (?,?,?) ON CONFLICT(cash_id,statement_id) DO UPDATE SET amount=excluded.amount', (cash_id, statement_id, str(cumulative)))
    conn.execute('UPDATE rec_statements SET status=? WHERE id=?', ('paid' if money(used_statement + amount) == money(total) else 'partial_paid', statement_id))
    audit(conn, allocation_key, 'cash_allocated', None, {'cash_id': cash_id, 'statement_id': statement_id, 'amount': str(amount)}, actor)

def financial_correction(conn, order_id, values, reference, reason, actor):
    conn.execute('SELECT id FROM orders WHERE id=?' + lock(conn), (order_id,)).fetchone()
    cfg = order_config(conn, order_id)
    if not cfg or not reference or not reason:
        raise ReconciliationError('财务补全需要已启用订单、唯一凭证和原因')
    if set(values) - {'refunds', 'fx', 'source_costs', 'reshipment_ids', 'expenses', 'operating_costs_complete', 'refunds_in_order_totals'}:
        raise ReconciliationError('不能修改已冻结的来源、合同或计费规则')
    for field in ('refunds','source_costs'):
        for ident, value in values.get(field, {}).items():
            if dec(value) < 0:
                raise ReconciliationError('退款和成本不能为负')
            if field == 'source_costs' and not conn.execute('SELECT id FROM rec_sources WHERE id=? AND order_id=?', (ident,order_id)).fetchone():
                raise ReconciliationError('成本补全来源不属于该订单')
    for value in values.get('fx', {}).values():
        if dec(value) <= 0:
            raise ReconciliationError('汇率必须为正')
    if 'operating_costs_complete' in values and not isinstance(values['operating_costs_complete'], bool):
        raise ReconciliationError('完整成本确认必须为布尔值')
    for expense_ref, expense in values.get('expenses', {}).items():
        if not expense_ref or not expense.get('proof') or expense.get('kind') not in {'goods','shipping','overhead'} or dec(expense.get('amount')) < 0:
            raise ReconciliationError('费用需要唯一凭证、合法类别和非负金额')
        currency(expense.get('currency'));object_(conn,expense.get('party_id'),'party')
    parcel_ids={p['id'] for p in shipments(conn,order_id)}
    if set(values.get('reshipment_ids', [])) - parcel_ids:
        raise ReconciliationError('补发包裹不属于该订单')
    previous=conn.execute("SELECT data FROM rec_actions WHERE order_id=? AND kind='financial_correction' ORDER BY created_at DESC,id DESC LIMIT 1",(order_id,)).fetchone()
    combined=json.loads(previous['data'])['values'] if previous else {}
    for field,value in values.items():
        combined[field]={**combined.get(field,{}),**value} if isinstance(value,dict) else value
    key='financial-correction:'+order_id+':'+reference
    if done(conn,key):
        raise ReconciliationError('财务补全凭证已登记')
    audit(conn,key,'financial_correction',order_id,{'values':combined,'reason':reason,'reference':reference},actor)

def balances(conn, allowed=None):
    amounts=defaultdict(Decimal)
    for e in rows(conn,'SELECT party_id,currency,category,amount FROM rec_entries'):
        if allowed is None or e['party_id'] in allowed:
            amounts[(e['party_id'],e['currency'],'collection' if e['category'].startswith('collection_') else 'trade')]+=dec(e['amount'])
    for r in rows(conn,'SELECT s.party_id,s.currency,s.snapshot,a.amount FROM rec_cash_allocations a JOIN rec_statements s ON s.id=a.statement_id'):
        if allowed is None or r['party_id'] in allowed:
            amounts[(r['party_id'],r['currency'],json.loads(r['snapshot'])['bucket'])]-=dec(r['amount'])
    return [{'party_id':p,'currency':u,'bucket':b,'outstanding':str(money(a))} for (p,u,b),a in sorted(amounts.items())]


def profit_report(conn):
    result=defaultdict(lambda:defaultdict(Decimal))
    latest=rows(conn,"""SELECT data FROM (SELECT data,ROW_NUMBER() OVER(PARTITION BY order_id ORDER BY created_at DESC,id DESC) AS n
        FROM rec_actions WHERE kind='economic_recognized') ranked WHERE n=1""")
    for row in latest:
        data=json.loads(row['data'])
        for source in data['sources']:
            snap=source['snapshot'];key=(data['site'],snap['warehouse_id'],snap['sku_id'],snap['pool']['id'],data['currency'],data['profit_type'])
            for field in ('revenue','cost','service','other_cost','profit'):
                result[key][field]+=dec(source.get(field,0))
    return [{'site':site,'warehouse_id':w,'sku_id':sku,'pool_id':pool,'currency':unit,'profit_type':profit_type,**{k:str(money(v)) for k,v in values.items()}}
            for (site,w,sku,pool,unit,profit_type),values in sorted(result.items())]

def authorized_offset(conn, positive_statement, negative_statement, amount, reference, reason, actor):
    if positive_statement == negative_statement or not reference or not reason:
        raise ReconciliationError('抵扣需要两张不同账单、唯一授权凭证和说明')
    for ident in sorted([positive_statement,negative_statement]):
        conn.execute('SELECT id FROM rec_statements WHERE id=?'+lock(conn),(ident,)).fetchone()
    positive=conn.execute('SELECT * FROM rec_statements WHERE id=?',(positive_statement,)).fetchone()
    negative=conn.execute('SELECT * FROM rec_statements WHERE id=?',(negative_statement,)).fetchone()
    amount=money(amount)
    if not positive or not negative or amount<=0 or positive['party_id']!=negative['party_id'] or positive['currency']!=negative['currency']:
        raise ReconciliationError('抵扣必须是同一主体、同一币种，且金额为正')
    key='authorized-offset:'+reference
    if done(conn,key):
        raise ReconciliationError('授权抵扣凭证已登记')
    for statement,sign in ((positive,1),(negative,-1)):
        cash=record_cash(conn,statement['party_id'],statement['currency'],amount*sign,now(),key+':'+str(sign),reason,actor)
        conn.execute('UPDATE rec_cash SET data=? WHERE id=?',(dump({'kind':'authorized_offset','proof':reason,'reference':reference,'actor':str(actor)}),cash))
        allocate_cash(conn,cash,statement['id'],amount*sign,actor)
    audit(conn,key,'authorized_offset',None,{'positive_statement':positive_statement,'negative_statement':negative_statement,
        'amount':str(amount),'reason':reason},actor)

def pending_statement_orders(conn, party, start, end):
    pending=[]
    for row in rows(conn,'SELECT DISTINCT order_id FROM rec_sources WHERE shipped>0'):
        events=rows(conn,"SELECT created_at FROM rec_actions WHERE order_id=? AND kind='source_shipped'",(row['order_id'],))
        shipped_in_period = any(start<=e['created_at']<end for e in events)
        cfg=order_config(conn,row['order_id']);parties={cfg['our_party_id']}
        for source in rows(conn,'SELECT snapshot FROM rec_sources WHERE order_id=? AND shipped>0',(row['order_id'],)):
            snap=json.loads(source['snapshot']);terms=snap['contract']['data']
            for field in ('profit','loss','shipping'):parties.update(terms[field])
            parties.update(snap['pool']['data']['capital'])
        for fee in cfg.get('fees',[]):parties.add(object_(conn,fee,'fee')['data']['provider_id'])
        parties.update(b['provider_id'] for b in latest_bills(conn,row['order_id']).values())
        if party not in parties:continue
        report=serializable(economics(conn,row['order_id']))
        if not shipped_in_period and not any(start<=r['occurred_at']<end for r in report['totals'] if r['party_id']==party):
            continue
        latest=conn.execute("SELECT data FROM rec_actions WHERE order_id=? AND kind='economic_recognized' ORDER BY created_at DESC,id DESC LIMIT 1",(row['order_id'],)).fetchone()
        def amounts(totals):
            return sorted((r['party_id'],r['category'],r['currency'],money(r['amount'])) for r in totals if money(r['amount']) != 0)
        if report['pending'] or not latest or amounts(report['totals'])!=amounts(json.loads(latest['data'])['totals']):
            pending.append({'order_id':row['order_id'],'reasons':report['pending'] or ['权益尚未确认或存在未记账变更']})
    return pending
