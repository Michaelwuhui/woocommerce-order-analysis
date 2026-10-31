"""Product-source settlement and immutable versions of legacy partner bills.

Product income, supplier capital and warehouse service fees are separate rights.
No stock, WooCommerce, cash or modular-ledger writes are performed here.
"""
from collections import defaultdict
from datetime import datetime, timezone
from decimal import Decimal, ROUND_HALF_UP
import json

TRANSIT_CODE = 'PL-JYJG-TRANSIT'
LABELS = {'partner_stock': '合伙人自备货', 'transit_service': '我方备货 · 中转服务', 'pending': '来源待核实'}
CENT = Decimal('.01')


def dec(value):
    number = Decimal(str(value if value not in (None, '') else 0))
    if not number.is_finite():
        raise ValueError('金额无效')
    return number


def money(value):
    return dec(value).quantize(CENT, rounding=ROUND_HALF_UP)


def num(value):
    return float(money(value))


def allocate(amount, weights):
    """Stable cent allocation, including negative adjustments and zero weights."""
    if not weights:
        return []
    weights = [abs(dec(value)) for value in weights]
    if not sum(weights):
        weights = [Decimal(1)] * len(weights)
    amount = money(amount)
    sign = -1 if amount < 0 else 1
    units = int(abs(amount)/CENT)
    exact = [Decimal(units)*weight/sum(weights) for weight in weights]
    whole = [int(value) for value in exact]
    order = sorted(range(len(weights)),key=lambda index:(-(exact[index]-whole[index]),index))
    for index in order[:units-sum(whole)]:
        whole[index] += 1
    return [Decimal(value)*CENT*sign for value in whole]


def rows(conn, sql, args=()):
    return [dict(row) for row in conn.execute(sql, args).fetchall()]


def exists(conn, table):
    from inv_common import _table_exists
    return _table_exists(conn, table)


def init_sqlite_schema(conn):
    conn.execute('''CREATE TABLE IF NOT EXISTS reconciliation_statement_versions (
        statement_id INTEGER NOT NULL, revision INTEGER NOT NULL,
        summary_json TEXT NOT NULL, lines_json TEXT NOT NULL,
        fees_json TEXT NOT NULL, previous_json TEXT NOT NULL,
        created_at TEXT NOT NULL, actor_id INTEGER,
        settlement_ready INTEGER NOT NULL DEFAULT 1,
        packing_fee_cny TEXT NOT NULL DEFAULT '0.00', pending_orders INTEGER NOT NULL DEFAULT 0,
        PRIMARY KEY(statement_id, revision))''')


def read_version(conn, statement_id, *, details=False, revision=None):
    if not exists(conn, 'reconciliation_statement_versions'):
        return None
    columns = 'revision,summary_json,created_at'
    if details is True or details=='lines':
        columns += ',lines_json'
    if details is True or details=='fees':
        columns += ',fees_json'
    where = 'statement_id=?'
    args = [statement_id]
    if revision is not None:
        where += ' AND revision=?'
        args.append(revision)
    row = conn.execute(f'SELECT {columns} FROM reconciliation_statement_versions WHERE {where} ORDER BY revision DESC LIMIT 1', args).fetchone()
    if not row:
        return None
    result = json.loads(row['summary_json'])
    result.update(revision=row['revision'], snapshot_created_at=row['created_at'])
    if details is True or details=='lines':
        result['source_lines'] = json.loads(row['lines_json'])
    if details is True or details=='fees':
        result['service_lines'] = json.loads(row['fees_json'])
    return result


def save_version(conn, statement_id, snapshot, actor, previous=None):
    if not exists(conn, 'reconciliation_statement_versions'):
        raise ValueError('请先安装商品来源对账快照迁移')
    latest = conn.execute('SELECT MAX(revision) FROM reconciliation_statement_versions WHERE statement_id=?', (statement_id,)).fetchone()[0]
    revision = (latest or 0) + 1
    # Preserve the complete old statement and its order snapshot before refresh.
    previous = previous or {}
    summary = {key: value for key, value in snapshot.items() if key not in ('source_lines', 'service_lines')}
    stamp = datetime.now(timezone.utc).isoformat(timespec='seconds')
    dump = lambda value: json.dumps(value, ensure_ascii=False, separators=(',', ':'), default=str)
    conn.execute('''INSERT INTO reconciliation_statement_versions
        (statement_id,revision,summary_json,lines_json,fees_json,previous_json,created_at,actor_id,
         settlement_ready,packing_fee_cny,pending_orders)
        VALUES (?,?,?,?,?,?,?,?,?,?,?)''', (statement_id, revision, dump(summary),
        dump(snapshot['source_lines']), dump(snapshot['service_lines']), dump(previous), stamp, actor,
        int(snapshot.get('settlement',{}).get('ready',True)),
        str(money(snapshot.get('settlement',{}).get('packing_fee_cny',0))),
        snapshot.get('settlement',{}).get('pending_orders',0)))
    return revision


def _scope(conn, partner, year, month, site_filter=None, manager_filter=None):
    sites = rows(conn, '''SELECT s.id,s.url,s.country,s.manager,s.cod_on_hold_is_shipped
        FROM effective_partner_sites ps JOIN sites s ON s.id=ps.site_id WHERE ps.partner_id=?''', (partner['id'],))
    sites = [site for site in sites if (not site_filter or site['url'] == site_filter)
             and (not manager_filter or site['manager'] == manager_filter)]
    start = f'{year:04d}-{month:02d}-01'
    end = f'{year + (month == 12):04d}-{1 if month == 12 else month + 1:02d}-01'
    if not sites:
        return sites, start, end, '1=0', []
    placeholders = ','.join('?' for _ in sites)
    return sites, start, end, f'o.source IN ({placeholders}) AND o.currency=?', [site['url'] for site in sites] + [partner['currency'] or 'PLN']


def _success(order, site):
    status = order['status'] or ''
    pm = order['payment_method'] or 'cod'
    return (status not in ('failed', 'cancelled', 'checkout-draft', 'trash', 'cheat', 'refunded')
            and not order['is_undelivered']
            and not (status == 'pending' and pm != 'cod')
            and not (status == 'on-hold' and (pm == 'bacs' or not site.get('cod_on_hold_is_shipped'))))


def _source_parts(order, item, index, evidence, warehouses, default_warehouse, confirmation=None):
    quantity = dec(item.get('quantity'))
    candidates = [part for part in evidence if
                  (part['woo_line_item_id'] is not None and str(part['woo_line_item_id']) == str(item.get('id')))
                  or (part['woo_line_item_id'] in (None, '') and part['line_index'] == index)]
    parts = []
    for part in candidates:
        qty = max(Decimal(0), dec(part['allocated_qty']) - dec(part['cancelled_qty']))
        if part['status'] == 'cancelled':
            qty = dec(part['fulfilled_qty'])
        if qty <= 0:
            continue
        warehouse = warehouses.get(part['warehouse_id'], {})
        parts.append({'warehouse_id': part['warehouse_id'], 'warehouse_code': warehouse.get('code', ''),
                      'warehouse_name': warehouse.get('name', ''), 'quantity': qty,
                      'fulfillment_id': part['fulfillment_id'], 'fulfillment_item_id': part['id'],
                      'fulfilled_quantity': part['fulfilled_qty'], 'source_basis': 'OMS 商品分仓',
                      'mode': 'transit_service' if warehouse.get('code') == TRANSIT_CODE else 'partner_stock'})
    assigned = sum((part['quantity'] for part in parts), Decimal(0))
    if assigned > quantity:
        return [{'mode': 'pending', 'quantity': quantity, 'warehouse_id': None, 'warehouse_name': '',
                 'source_basis': '分仓数量超过订单商品数量', 'warehouse_code': ''}]
    if assigned < quantity:
        if evidence:
            mode, warehouse_id, basis = 'pending', None, 'OMS 商品尚未完整分仓'
        else:
            warehouse_id = order['warehouse_id'] or default_warehouse
            warehouse = warehouses.get(warehouse_id, {})
            mode = 'transit_service' if warehouse.get('code') == TRANSIT_CODE else 'partner_stock'
            basis = '订单指定仓库' if order['warehouse_id'] else '历史合伙人供货范围（无 OMS 分仓）'
            if not warehouse_id:
                mode, basis = 'pending', '未找到供货仓库'
        warehouse = warehouses.get(warehouse_id, {})
        parts.append({'mode': mode, 'quantity': quantity-assigned, 'warehouse_id': warehouse_id,
                      'warehouse_code': warehouse.get('code', ''), 'warehouse_name': warehouse.get('name', ''),
                      'source_basis': basis, 'fulfilled_quantity': 0})
    if confirmation:
        valid = (str(confirmation.get('product_id') or 0)==str(item.get('product_id') or 0)
                 and str(confirmation.get('variation_id') or 0)==str(item.get('variation_id') or 0)
                 and dec(confirmation.get('quantity'))==quantity
                 and confirmation.get('mode') in ('partner_stock','transit_service')
                 and str(confirmation.get('evidence') or '').strip())
        if not valid:
            return [{'mode':'pending','quantity':quantity,'warehouse_id':None,'warehouse_name':'',
                     'warehouse_code':'','source_basis':'供货来源确认与当前商品身份或数量不一致'}]
        for part in parts:
            part.update(mode=confirmation['mode'],source_override_id=confirmation['id'],
                        source_basis='供货来源已明确确认；'+part['source_basis'],
                        ownership_evidence=confirmation['evidence'])
    return parts


def _service_lines(conn, partner, start, end, scope_sql, scope_args, warehouses, decisions):
    transit_ids = [ident for ident, warehouse in warehouses.items() if warehouse['code'] == TRANSIT_CODE]
    if not transit_ids or not exists(conn, 'oms_shipments'):
        return []
    transit_id = transit_ids[0]
    # Physical parcel primary warehouse may be the ordinary partner warehouse.
    # Attribution comes from the shipped product items, including joint parcels.
    parcels = rows(conn, f'''SELECT DISTINCT s.id,s.tracking_number,s.shipped_at,s.status,
        f.order_id,o.number,o.source,o.currency,o.date_created
        FROM oms_shipments s JOIN oms_fulfillments f ON f.id=s.fulfillment_id
        JOIN orders o ON o.id=f.order_id
        WHERE {scope_sql} AND s.shipped_at<? AND s.shipped_at<>''
          AND s.status NOT IN ('cancelled','label_pending') AND (
            EXISTS (SELECT 1 FROM oms_shipment_items si JOIN oms_fulfillment_items fi ON fi.id=si.fulfillment_item_id
              JOIN oms_fulfillments item_f ON item_f.id=fi.fulfillment_id
              WHERE si.shipment_id=s.id AND item_f.warehouse_id=?)
            OR (NOT EXISTS (SELECT 1 FROM oms_shipment_items si WHERE si.shipment_id=s.id)
              AND f.warehouse_id=?)) ORDER BY s.shipped_at,s.id''', scope_args+[end, transit_id, transit_id])
    if not any(start <= parcel['shipped_at'][:10] < end for parcel in parcels):
        return []
    rules = []
    if exists(conn, 'rec_objects'):
        for row in rows(conn, "SELECT id,name,data,created_at FROM rec_objects WHERE kind='history_rule' ORDER BY created_at,id"):
            rule = json.loads(row['data'])
            if rule.get('warehouse_id') == transit_id:
                rules.append({**rule, 'id': row['id'], 'name': row['name']})
    logs = rows(conn, f'''SELECT l.order_id,l.tracking_number,l.is_reship,l.reship_reason
        FROM shipping_logs l JOIN orders o ON o.id=l.order_id WHERE {scope_sql}
        AND l.status<>'pending_sync' ''', scope_args)
    log_index = defaultdict(list)
    for log in logs:
        log_index[(log['order_id'], (log['tracking_number'] or '').strip())].append(log)
    classified = {}
    if exists(conn,'rec_objects'):
        for record in rows(conn,"SELECT id,data FROM rec_objects WHERE kind='legacy_shipment_classification' ORDER BY created_at,id"):
            value = json.loads(record['data'])
            classified[str(value.get('shipment_id'))] = {**value,'classification_id':record['id']}
    first_original = {}
    prior_unknown = set()
    result = []
    for parcel in parcels:
        stamp = parcel['shipped_at'][:10]
        matches = log_index[(parcel['order_id'], (parcel['tracking_number'] or '').strip())]
        flags = {bool(log['is_reship']) for log in matches if log['is_reship'] is not None}
        is_reship = flags.pop() if len(flags) == 1 else None
        stored = classified.get(parcel['id'],{})
        if stored.get('order_id') != parcel['order_id']:
            stored = {}
        decision = decisions[parcel['id']] if parcel['id'] in decisions else stored
        category = decision.get('reason', 'unknown')
        evidence_text = str(decision.get('evidence') or '').strip()
        if is_reship is None and category != 'unknown':
            if category not in ('original','ordinary','warehouse_omission') or not evidence_text:
                raise ValueError('原发/补发分类需要有效类型和核对依据')
            is_reship = category != 'original'
        rule = next((rule for rule in reversed(rules) if rule['effective_from'] <= stamp < rule['effective_to']), None)
        price = dec(rule['management_cny']) if rule and rule.get('management_cny') is not None else None
        pending, amount, reason = [], None, ''
        kind = 'reship' if is_reship else 'original'
        if is_reship is False:
            if parcel['order_id'] in first_original:
                amount, reason = Decimal(0), '同一原订单拆包不重复收费'
            else:
                first_original[parcel['order_id']] = parcel['id']
                amount, reason = price, '原订单首次中转出库'
                if parcel['order_id'] in prior_unknown:
                    pending.append('更早包裹原发/补发身份待核实，不能重复收原订单费')
        elif is_reship is True:
            category = decision.get('reason', 'unknown')
            evidence = str(decision.get('evidence') or '').strip()
            if category not in ('ordinary', 'warehouse_omission', 'unknown'):
                raise ValueError('补发分类无效')
            if category != 'unknown' and not evidence:
                raise ValueError('补发分类需要填写依据')
            if category == 'ordinary':
                amount, reason = price, '普通补发收费'
            elif category == 'warehouse_omission':
                amount, reason = Decimal(0), '仓库漏发补发免收'
            else:
                pending.append('补发原因待核实')
        else:
            pending.append('运单原发/补发标记缺失或冲突')
            prior_unknown.add(parcel['order_id'])
        if not rule:
            pending.append('缺少出库日生效的中转打包费规则')
        elif price is None:
            pending.append('打包费金额缺失')
        elif price < 0:
            raise ValueError('打包费规则金额不能为负')
        if not start <= stamp < end:
            continue
        result.append({'event_key': 'transit:'+parcel['id'], 'order_id': parcel['order_id'],
            'order_number': parcel['number'], 'source': parcel['source'], 'order_date': parcel['date_created'],
            'market_currency': parcel['currency'], 'shipment_id': parcel['id'],
            'tracking_number': parcel['tracking_number'], 'shipped_at': parcel['shipped_at'], 'kind': kind,
            'warehouse_id': transit_id, 'warehouse_name': warehouses[transit_id]['name'],
            'rule_id': rule['id'] if rule else None, 'rule_name': rule['name'] if rule else None,
            'provider_id': rule.get('provider_id') if rule else None,
            'unit_fee_cny': num(price) if price is not None else None,
            'amount_cny': num(amount) if amount is not None and not pending else None,
            'reason': reason, 'decision':{key:decision[key] for key in ('reason','evidence') if key in decision},
            'classification_id':stored.get('classification_id'), 'pending': pending,
            'reship_policy': 'charge_except_warehouse_omission', 'return_policy': 'keep'})
    used = {row['shipment_id'] for row in result}
    if set(decisions)-used:
        raise ValueError('补发分类包含本期范围外的运单')
    return result


def build_snapshot(conn, partner, year, month, mode, business, cost_lookup, *, decisions=None,
                   site_filter=None, manager_filter=None):
    if not 1 <= month <= 12 or not 2000 <= year <= 2200:
        raise ValueError('对账期间无效')
    sites, start, end, scope_sql, scope_args = _scope(conn, partner, year, month, site_filter, manager_filter)
    site_index = {site['url']: site for site in sites}
    orders = rows(conn, f'''SELECT o.id,o.number,o.source,o.status,o.payment_method,o.total,o.shipping_total,
        o.date_created,o.currency,o.warehouse_id,o.line_items,o.is_undelivered,o.is_problem_return,o.shipping_loss_amount
        FROM orders o WHERE {scope_sql} AND o.date_created>=? AND o.date_created<? ORDER BY o.date_created,o.id''', scope_args+[start,end])
    warehouses = {row['id']: row for row in rows(conn, 'SELECT id,code,name,country FROM warehouses ORDER BY id')}
    defaults = {}
    for ident, warehouse in warehouses.items():
        defaults.setdefault(warehouse['country'], ident)
    confirmations = {}
    if exists(conn,'rec_objects'):
        for record in rows(conn,"SELECT id,data FROM rec_objects WHERE kind='legacy_product_source' ORDER BY created_at,id"):
            data = json.loads(record['data'])
            confirmations[(str(data.get('order_id')),str(data.get('line_item_id')))] = {**data,'id':record['id']}
    evidence = defaultdict(list)
    if exists(conn, 'oms_fulfillments'):
        for row in rows(conn, f'''SELECT fi.id,fi.allocated_qty,fi.fulfilled_qty,fi.cancelled_qty,
            f.id AS fulfillment_id,f.order_id,f.warehouse_id,f.status,oi.woo_line_item_id,oi.line_index
            FROM oms_fulfillments f JOIN orders o ON o.id=f.order_id
            JOIN oms_fulfillment_items fi ON fi.fulfillment_id=f.id JOIN oms_order_items oi ON oi.id=fi.order_item_id
            WHERE {scope_sql} AND o.date_created>=? AND o.date_created<? AND f.status<>'superseded'
            ORDER BY f.order_id,fi.id''', scope_args+[start,end]):
            evidence[row['order_id']].append(row)
    cr = dec(partner['cost_ratio'] if partner['cost_ratio'] is not None else '.5')
    pp = dec(partner['partner_profit_ratio'] if partner['partner_profit_ratio'] is not None else '.25')
    op = dec(partner['our_profit_ratio'] if partner['our_profit_ratio'] is not None else '.25')
    pp_share = pp/(pp+op) if pp+op else Decimal('.5')
    source_lines, source_products = [], {}
    group = {kind: {'mode': kind, 'label': label, 'quantity': 0, 'revenue': 0., 'shipping_loss': 0.,
                   'cost_recovery': 0., 'partner_profit': 0., 'our_receivable': 0., 'pending_count': 0,
                   'known_revenue': 0., 'order_ids': set()} for kind,label in LABELS.items()}
    pending_orders = set()
    for order in orders:
        site = site_index[order['source']]
        success = _success(order, site)
        try:
            items = json.loads(order['line_items'] or '[]')
        except (TypeError, ValueError):
            items = []
        items = items if isinstance(items, list) else []
        valid_items = [(index,item) for index,item in enumerate(items) if dec(item.get('quantity')) > 0]
        residual = money(dec(order['total'])-dec(order['shipping_total'])-sum((dec(item.get('total')) for _,item in valid_items),Decimal(0)))
        adjustments = allocate(residual, [item.get('total',0) for _,item in valid_items])
        order_rows = []
        for (index,item), adjustment in zip(valid_items, adjustments):
            confirmation = confirmations.get((order['id'],str(item.get('id',index))))
            parts = _source_parts(order,item,index,evidence[order['id']],warehouses,defaults.get(site['country']),confirmation)
            amounts = allocate(dec(item.get('total'))+adjustment,[part['quantity'] for part in parts])
            for part, amount in zip(parts,amounts):
                kind = part['mode']
                issues = []
                if kind == 'pending':
                    issues.append(part['source_basis'])
                cost = None
                cost_info = {}
                if success and kind == 'partner_stock':
                    cost_info = cost_lookup(item.get('name') or '',order['source'],part['warehouse_id'],order['date_created'][:10]) or {}
                    if cost_info.get('unit_cost') is not None:
                        cost = money(dec(cost_info['unit_cost'])*part['quantity'])
                    elif mode == 'actual':
                        issues.append('合伙人自备货成本待补全')
                revenue = amount if success else Decimal(0)
                recovery = profit = receivable = None
                if not issues:
                    if kind == 'transit_service':
                        recovery, profit, receivable = Decimal(0),Decimal(0),revenue
                    elif kind == 'partner_stock':
                        recovery = (cost if mode == 'actual' else money(revenue*cr)) if success else Decimal(0)
                        profit = money((revenue-recovery)*pp_share) if mode == 'actual' else money(revenue*pp)
                        receivable = revenue-recovery-profit
                row = {'order_id':order['id'],'order_number':order['number'],'source':order['source'],
                    'manager':site['manager'],'order_date':order['date_created'],'order_status':order['status'],
                    'currency':order['currency'],'line_item_id':str(item.get('id',index)),
                    'product_id':item.get('product_id'),'variation_id':item.get('variation_id'),
                    'product_name':item.get('name') or '', 'sku':item.get('sku') or '',
                    **{key:value for key,value in part.items() if key!='quantity'},
                    'mode_label':LABELS[kind], 'quantity':float(part['quantity']),
                    'nominal_revenue':num(amount),'revenue':num(revenue),'shipping_loss':0.,
                    'actual_partner_cost':num(cost) if cost is not None else None,
                    'cost_recovery':num(recovery) if recovery is not None else None,
                    'partner_profit':num(profit) if profit is not None else None,
                    'our_receivable':num(receivable) if receivable is not None else None,
                    'cost_evidence':cost_info, 'pending':issues,'is_revenue':success,
                    'returned':bool(order['is_undelivered']), 'row_kind':'product'}
                order_rows.append(row)
        loss = dec(order['shipping_loss_amount']) if order['is_undelivered'] else Decimal(0)
        losses = allocate(loss,[row['nominal_revenue'] or row['quantity'] for row in order_rows])
        if not order_rows and (success or loss):
            pending_orders.add(order['id'])
            order_rows = [{'order_id':order['id'],'order_number':order['number'],'source':order['source'],
                'order_date':order['date_created'],'order_status':order['status'],'currency':order['currency'],
                'product_name':'订单商品缺失','line_item_id':'','mode':'pending','mode_label':LABELS['pending'],
                'warehouse_id':None,'warehouse_name':'','quantity':0,'nominal_revenue':num(dec(order['total'])-dec(order['shipping_total'])),
                'revenue':num(dec(order['total'])-dec(order['shipping_total'])) if success else 0,
                'shipping_loss':num(loss),'cost_recovery':None,'partner_profit':None,'our_receivable':None,
                'pending':['订单商品明细缺失'],'source_basis':'无法核对商品来源','row_kind':'product','is_revenue':success}]
            losses = [Decimal(0)]
        for row, allocated_loss in zip(order_rows,losses):
            kind = row['mode']
            row['shipping_loss'] += num(allocated_loss)
            if not row['pending'] and kind == 'partner_stock' and allocated_loss:
                partner_loss = money(allocated_loss*pp_share) if mode=='actual' else money(allocated_loss/2)
                row['partner_profit'] = num(dec(row['partner_profit'])-partner_loss)
                row['our_receivable'] = num(dec(row['our_receivable'])-(allocated_loss-partner_loss))
            bucket = group[kind]
            if success or allocated_loss:
                bucket['order_ids'].add(order['id'])
            bucket['quantity'] += row['quantity'] if success else 0
            bucket['revenue'] += row['revenue']
            bucket['shipping_loss'] += row['shipping_loss']
            if row['pending'] and (success or allocated_loss):
                bucket['pending_count'] += 1
                pending_orders.add(order['id'])
            elif not row['pending']:
                bucket['known_revenue'] += row['revenue']
                for key in ('cost_recovery','partner_profit','our_receivable'):
                    bucket[key] += row[key] or 0
            key = (kind,row.get('warehouse_id'),row['product_name'],row.get('product_id'),row.get('variation_id'))
            product = source_products.setdefault(key, {'mode':kind,'mode_label':LABELS[kind],
                'warehouse_id':row.get('warehouse_id'),'warehouse_name':row.get('warehouse_name',''),
                'product_name':row['product_name'],'product_id':row.get('product_id'),'variation_id':row.get('variation_id'),
                'quantity':0,'revenue':0.,'cost_recovery':0.,'partner_profit':0.,'our_receivable':0.,
                'shipping_loss':0.,'pending_count':0,'order_ids':set()})
            if success or allocated_loss:
                product['order_ids'].add(order['id'])
            product['quantity'] += row['quantity'] if success else 0
            for field in ('revenue','cost_recovery','partner_profit','our_receivable','shipping_loss'):
                product[field] += row.get(field) or 0
            product['pending_count'] += bool(row['pending'] and (success or allocated_loss))
            source_lines.append(row)
    # Balance fractional cents across the period instead of always giving the
    # half-cent to one party on every item. Row sums still equal the bill.
    eligible = [row for row in source_lines if row['mode']=='partner_stock' and not row['pending']]
    exact_profit = lambda row: ((dec(row['revenue'])-dec(row['cost_recovery'])-dec(row['shipping_loss']))*pp_share
        if mode=='actual' else dec(row['revenue'])*pp-dec(row['shipping_loss'])/2)
    target_profit = money(sum((exact_profit(row) for row in eligible),Decimal(0)))
    difference = money(target_profit-sum((dec(row['partner_profit']) for row in eligible),Decimal(0)))
    direction = 1 if difference>0 else -1
    ranked = sorted(eligible,key=lambda row:(-direction*(exact_profit(row)-dec(row['partner_profit'])),
                    row['order_id'],row['line_item_id'],row.get('warehouse_id') or 0))
    for row in ranked[:int(abs(difference)/CENT)]:
        change = CENT*direction
        row['partner_profit'] = num(dec(row['partner_profit'])+change)
        row['our_receivable'] = num(dec(row['our_receivable'])-change)
        group['partner_stock']['partner_profit'] += float(change)
        group['partner_stock']['our_receivable'] -= float(change)
        key = ('partner_stock',row.get('warehouse_id'),row['product_name'],row.get('product_id'),row.get('variation_id'))
        source_products[key]['partner_profit'] += float(change)
        source_products[key]['our_receivable'] -= float(change)
    fees = _service_lines(conn,partner,start,end,scope_sql,scope_args,warehouses,decisions or {})
    for bucket in list(group.values())+list(source_products.values()):
        bucket['orders'] = len(bucket.pop('order_ids'))
        for field in ('revenue','known_revenue','cost_recovery','partner_profit','our_receivable','shipping_loss'):
            if field in bucket:
                bucket[field] = num(bucket[field])
    fee_total = num(sum((dec(row['amount_cny']) for row in fees if row['amount_cny'] is not None),Decimal(0)))
    fee_pending = sum(bool(row['pending']) for row in fees)
    cost = num(sum(dec(bucket['cost_recovery']) for bucket in group.values()))
    profit = num(sum(dec(bucket['partner_profit']) for bucket in group.values()))
    receivable = num(sum(dec(bucket['our_receivable']) for bucket in group.values()))
    ready = not pending_orders and not fee_pending
    return {'schema_version':1,'calc_mode':mode,'currency':partner['currency'] or 'PLN',
        'period_year':year,'period_month':month,'partner_id':partner['id'],
        'business_detail':business,'groups':list(group.values()),
        'source_products':sorted(source_products.values(),key=lambda row:(row['mode'],-row['revenue'],row['product_name'])),
        'source_lines':source_lines,'service_lines':fees,
        'settlement':{'cost_recovery':cost,'partner_profit':profit,'our_receivable':receivable,
            'net':num(dec(cost)+dec(profit)+dec(receivable)),
            'transit_goods_receivable':group['transit_service']['our_receivable'],
            'transit_shipping_loss_separate':group['transit_service']['shipping_loss'],
            'packing_fee_cny':fee_total,'packing_charges':sum(dec(row['amount_cny'] or 0)>0 for row in fees),
            'packing_pending':fee_pending,'pending_orders':len(pending_orders),'ready':ready,
            'amounts_label':'完整核算' if ready else '已核实行小计 · 待核实行未计入',
            'service_fee_offset':False,'sales_date_basis':'下单月份','service_date_basis':'实际出库月份',
            'reship_policy':'普通补发收费，仓库漏发免收',
            'note':'中转商品不返还合伙人备货成本、不分商品利润；中转退件运费损失归我方另账。打包费以 CNY 单列，不自动抵扣商品货款。代收及实收回款另核凭证。'},
        'ratios':{'cost':str(cr),'partner_profit':str(pp),'our_profit':str(op)},
        'scope_sites':[site['url'] for site in sites]}
