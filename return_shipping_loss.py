"""Shared, currency-safe return freight policy for manual and automatic returns.

Rules are stored in settings, so no schema migration or import-time write is
needed. Amounts are per returned order. Multi-warehouse orders require a human
to confirm the actual amount instead of guessing which route to charge.
"""
import hashlib
import json
import re
from decimal import Decimal, InvalidOperation


SETTING_KEY = 'return_shipping_loss_rules_v1'


class PolicyConflict(ValueError):
    pass


def money(value):
    if isinstance(value, bool) or value is None or str(value).strip() == '':
        raise ValueError('运费金额不能为空')
    try:
        amount = Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError):
        raise ValueError('运费金额格式不正确') from None
    if not amount.is_finite() or amount < 0 or amount > 1000000:
        raise ValueError('运费金额须为 0 至 1000000 之间的有限数值')
    if amount != amount.quantize(Decimal('0.01')):
        raise ValueError('运费金额最多保留两位小数')
    return amount


def _version(raw):
    return hashlib.sha256((raw or '').encode('utf-8')).hexdigest()


def load_policy(conn):
    row = conn.execute('SELECT value FROM settings WHERE key=?', (SETTING_KEY,)).fetchone()
    raw = row[0] if row else None
    rules = json.loads(raw) if raw else []
    if not isinstance(rules, list):
        raise ValueError('退件运费规则格式异常，请在系统设置中检查')
    return {'rules': rules, 'version': _version(raw)}


def validate_rules(conn, rules):
    if not isinstance(rules, list) or len(rules) > 100:
        raise ValueError('最多可设置 100 条退件运费规则')
    warehouses = {int(r[0]) for r in conn.execute('SELECT id FROM warehouses').fetchall()}
    result, seen = [], set()
    for rule in rules:
        if not isinstance(rule, dict):
            raise ValueError('规则格式不正确')
        warehouse_id = rule.get('warehouse_id')
        if warehouse_id is not None:
            if isinstance(warehouse_id, bool) or not isinstance(warehouse_id, int) or warehouse_id not in warehouses:
                raise ValueError('请选择有效的发货仓库')
        country = str(rule.get('destination_country') or '').strip().upper()
        currency = str(rule.get('currency') or '').strip().upper()
        if not re.fullmatch('[A-Z]{2}', country):
            raise ValueError('收货区域须为两位国家代码，例如 CZ、PL')
        if not re.fullmatch('[A-Z]{3}', currency):
            raise ValueError('币种须为三位代码，例如 CZK、PLN')
        key = (warehouse_id, country)
        if key in seen:
            raise ValueError('同一仓库和收货区域只能设置一条规则')
        seen.add(key)
        outbound = money(rule.get('outbound_amount'))
        inbound = money(rule.get('return_amount'))
        money(outbound + inbound)
        result.append({
            'warehouse_id': warehouse_id, 'destination_country': country,
            'currency': currency, 'outbound_amount': float(outbound),
            'return_amount': float(inbound),
        })
    return result


def save_policy(conn, rules, expected_version):
    rules = validate_rules(conn, rules)
    row = conn.execute('SELECT value FROM settings WHERE key=?', (SETTING_KEY,)).fetchone()
    before = row[0] if row else None
    if expected_version != _version(before):
        raise PolicyConflict('规则已被其他人修改，请重新加载后再保存')
    raw = json.dumps(rules, ensure_ascii=False, separators=(',', ':'))
    if row:
        changed = conn.execute('UPDATE settings SET value=? WHERE key=? AND value=?',
                               (raw, SETTING_KEY, before)).rowcount
    else:
        changed = conn.execute('INSERT OR IGNORE INTO settings (key,value) VALUES (?,?)',
                               (SETTING_KEY, raw)).rowcount
    if changed != 1:
        raise PolicyConflict('规则已被其他人修改，请重新加载后再保存')
    return {'rules': rules, 'version': _version(raw)}


def _address(value):
    if isinstance(value, dict):
        return value
    try:
        data = json.loads(value or '{}')
        return data if isinstance(data, dict) else {}
    except (TypeError, ValueError):
        return {}


def quote_loss(conn, order, policy=None):
    """Return an amount in the ORDER currency, or an explicit manual-review error."""
    order = dict(order)
    policy = policy if policy is not None else load_policy(conn)
    currency = str(order.get('currency') or '').strip().upper()
    result = {'amount': None, 'currency': currency, 'matched': False,
              'warehouse_id': None, 'destination_country': '', 'message': ''}
    rules = policy['rules']
    if not rules:
        result.update(amount=float(money(order.get('shipping_total') or 0)),
                      message='未设置退件规则，沿用订单运费；可按实际损失修改')
        return result
    # Older callers select only the small candidate projection.
    if 'shipping' not in order or 'warehouse_id' not in order:
        full = conn.execute('SELECT * FROM orders WHERE id=?', (order['id'],)).fetchone()
        if full:
            order = dict(full)
    shipping, billing = _address(order.get('shipping')), _address(order.get('billing'))
    country = str(shipping.get('country') or billing.get('country') or '').strip().upper()
    if not country:
        site = conn.execute('SELECT country FROM sites WHERE url=?', (order.get('source'),)).fetchone()
        country = str(site[0] or '').strip().upper() if site else ''
    result['destination_country'] = country
    # Current, actually dispatched fulfillment beats the legacy order warehouse.
    rows = conn.execute('''
        SELECT DISTINCT f.warehouse_id
        FROM oms_fulfillments f
        JOIN oms_order_fulfillment_state st ON st.order_id=f.order_id AND st.revision=f.revision
        WHERE f.order_id=? AND f.status NOT IN ('cancelled','failed','superseded')
          AND (f.shipped_at IS NOT NULL OR EXISTS (
              SELECT 1 FROM oms_shipments s WHERE s.fulfillment_id=f.id
                AND s.status IN ('shipped','in_transit','delivered','returned')
          ))
    ''', (order['id'],)).fetchall()
    warehouse_ids = {int(r[0]) for r in rows if r[0] is not None}
    if len(warehouse_ids) > 1:
        result['message'] = '该订单由多个仓库发货，请人工填写实际退件运费损失'
        return result
    warehouse_id = next(iter(warehouse_ids), order.get('warehouse_id'))
    result['warehouse_id'] = warehouse_id
    matching = [r for r in rules if r['destination_country'] == country]
    specific = next((r for r in matching if warehouse_id is not None and r['warehouse_id'] == warehouse_id), None)
    rule = specific or next((r for r in matching if r['warehouse_id'] is None), None)
    if not rule:
        if matching and warehouse_id is None:
            result['message'] = '无法识别实际发货仓库，请补充区域通用规则或人工填写损失'
            return result
        result.update(amount=float(money(order.get('shipping_total') or 0)),
                      message='未匹配退件规则，沿用订单运费；可按实际损失修改')
        return result
    if rule['currency'] != currency:
        result['message'] = f"规则币种 {rule['currency']} 与订单币种 {currency} 不同，请按订单币种人工填写损失"
        return result
    outbound, inbound = money(rule['outbound_amount']), money(rule['return_amount'])
    total = money(outbound + inbound)
    scope = '区域通用规则' if rule['warehouse_id'] is None else '仓库专用规则'
    result.update(amount=float(total), matched=True, outbound_amount=float(outbound),
                  return_amount=float(inbound),
                  message=f'{scope}：正向 {outbound:.2f} + 逆向 {inbound:.2f} = {total:.2f} {currency}')
    return result


def required_loss(conn, order, policy=None):
    quote = quote_loss(conn, order, policy)
    if quote['amount'] is None:
        raise ValueError(quote['message'])
    return quote['amount']
