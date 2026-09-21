"""Versioned commercial terms and exact, auditable money primitives.

Money is decimal text deliberately: the legacy PostgreSQL adapter converts NUMERIC
results to binary float. All arithmetic and validation stay in Decimal here.
"""
import json
import uuid
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

from inv_common import _table_exists


class ReconciliationError(ValueError):
    pass


def now():
    return datetime.now(timezone.utc).isoformat(timespec='microseconds')


def uid():
    return str(uuid.uuid4())


def dump(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(',', ':'))


def dec(value):
    if value is None or value == '' or isinstance(value, bool):
        raise ReconciliationError('金额/汇率缺失，不能视为零')
    try:
        result = Decimal(str(value))
    except InvalidOperation as exc:
        raise ReconciliationError('无效数字') from exc
    if not result.is_finite() or abs(result) > Decimal('1000000000000'):
        raise ReconciliationError('数字超出允许范围')
    return result


def money(value):
    return dec(value).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)


def currency(value):
    value = str(value).upper()
    if value not in {'CNY', 'PLN', 'CZK', 'HUF', 'EUR', 'USD', 'GBP', 'AUD', 'CAD'}:
        raise ReconciliationError('不支持的币种')
    return value


def integer(value, minimum=0):
    number = dec(value)
    if number != int(number) or number < minimum:
        raise ReconciliationError('数量必须为符合范围的整数')
    return int(number)


def lock(conn):
    return ' FOR UPDATE' if hasattr(conn, '_raw') else ''


def ready(conn):
    return _table_exists(conn, 'rec_objects')


def rows(conn, sql, args=()):
    return [dict(r) for r in conn.execute(sql, args).fetchall()]


def object_(conn, ident, kind=None):
    row = conn.execute('SELECT * FROM rec_objects WHERE id=?', (ident,)).fetchone()
    if not row or (kind and row['kind'] != kind):
        raise ReconciliationError('配置不存在或类型错误')
    return {**dict(row), 'data': json.loads(row['data'])}


def objects(conn, kind):
    return [{**r, 'data': json.loads(r['data'])} for r in rows(
        conn, 'SELECT * FROM rec_objects WHERE kind=? ORDER BY created_at,id', (kind,))]


def shares(conn, values, label):
    if not isinstance(values, dict) or not values:
        raise ReconciliationError(label + '未配置')
    result = {}
    for party, weight in values.items():
        object_(conn, party, 'party')
        amount = dec(weight)
        if amount < 0:
            raise ReconciliationError(label + '不能为负')
        result[party] = str(amount)
    if sum(map(dec, result.values())) != 1:
        raise ReconciliationError(label + '之和必须为 1')
    return result


def allocate(amount, weights):
    """Largest remainder in stable key order; preserve negative totals exactly."""
    total = money(amount)
    weights = {str(k): dec(v) for k, v in weights.items()}
    denominator = sum(weights.values())
    if not weights or denominator <= 0 or any(w < 0 for w in weights.values()):
        raise ReconciliationError('分摊权重无效')
    sign = -1 if total < 0 else 1
    cents = int(abs(total) * 100)
    exact = {k: Decimal(cents) * w / denominator for k, w in weights.items()}
    result = {k: int(v) for k, v in exact.items()}
    left = cents - sum(result.values())
    for k in sorted(weights, key=lambda k: (-(exact[k] - result[k]), k))[:left]:
        result[k] += 1
    return {k: Decimal(sign * v) / 100 for k, v in result.items()}


def convert(amount, source, target, fx):
    if source == target:
        return money(amount)
    rate = dec(fx.get(source + '/' + target))
    if rate <= 0:
        raise ReconciliationError('汇率必须大于零')
    return money(dec(amount) * rate)


def save_object(conn, kind, name, data, actor):
    data = dict(data)
    if not name or len(name) > 150:
        raise ReconciliationError('名称必填且不超过 150 字')
    if kind == 'party':
        if 'internal' in data and not isinstance(data['internal'], bool):
            raise ReconciliationError('主体类型必须为布尔值')
        data = {'internal': bool(data.get('internal')), 'currency': currency(data.get('currency', 'CNY')),
                'roles': data.get('roles', []), 'legacy_partner_id': data.get('legacy_partner_id')}
        if data['legacy_partner_id'] is not None and not conn.execute(
                'SELECT id FROM partners WHERE id=?', (data['legacy_partner_id'],)).fetchone():
            raise ReconciliationError('旧合伙人不存在')
    elif kind == 'pool':
        data['ownership'] = shares(conn, data.get('ownership'), '货权比例')
        data['capital'] = shares(conn, data.get('capital'), '回本比例')
        if data.get('ownership_method') not in {'physical', 'virtual'}:
            raise ReconciliationError('必须明确实物拣货或虚拟货权台账')
    elif kind == 'contract':
        if data.get('template') not in {'self', 'margin', 'joint', 'fixed', 'legacy'}:
            raise ReconciliationError('请选择合作模板')
        object_(conn, data.get('pool_id'), 'pool')
        data['profit'] = shares(conn, data.get('profit'), '分润比例')
        data['loss'] = shares(conn, data.get('loss'), '亏损承担比例')
        data['shipping'] = shares(conn, data.get('shipping'), '运费收益及亏损比例')
        currency(data.get('currency'))
        if data.get('allocation', 'goods') not in {'goods', 'quantity', 'weight'}:
            raise ReconciliationError('分摊方式无效')
        if data.get('template') == 'fixed':
            if dec(data.get('supply_price')) < 0:
                raise ReconciliationError('供货价不能为负')
        if data.get('template') == 'legacy':
            if not 0 <= dec(data.get('legacy_rate')) <= 1:
                raise ReconciliationError('销售比例必须在 0 至 1 之间')
        if data.get('capital_status') not in {'unpaid', 'paid', 'none'}:
            raise ReconciliationError('必须明确是否已经支付货款')
        if data.get('recognition') not in {'shipped', 'delivered', 'bill_verified'}:
            raise ReconciliationError('必须配置确认时点')
    elif kind == 'profile':
        if not conn.execute('SELECT id FROM sites WHERE url=?',(data.get('site'),)).fetchone():
            raise ReconciliationError('站点不存在')
        if not isinstance(data.get('enabled'),bool) or not data.get('starts_on'):
            raise ReconciliationError('站点规则必须明确开关和开始日期')
        datetime.fromisoformat(data['starts_on'])
        config = data.get('config',{})
        if config.get('policy') in {'cost','manual'}:
            raise ReconciliationError('站点默认规则不能复用某笔订单的报价或指定批次')
        if not config.get('contracts') or not config.get('our_party_id'):
            raise ReconciliationError('请先配置一笔样板订单')
        for pool, contract in config['contracts'].items():
            if object_(conn,contract,'contract')['data']['pool_id'] != pool:
                raise ReconciliationError('合同货权池不匹配')
        for key in ('refunds','source_costs','reshipment_ids','quotes','preferred_batches','actor'):
            config.pop(key,None)
        data['config'] = config
    elif kind == 'fee':
        object_(conn, data.get('provider_id'), 'party')
        if data.get('basis') not in {'order', 'parcel'}:
            raise ReconciliationError('计费单位必须为订单或包裹')
        if data.get('event') not in {'shipped', 'delivered', 'bill_verified'}:
            raise ReconciliationError('收费时点无效')
        if dec(data.get('amount')) < 0:
            raise ReconciliationError('管理费不能为负')
        currency(data.get('currency'))
        if data.get('return_policy') not in {'keep', 'reverse'}:
            raise ReconciliationError('退货管理费处理方式必填')
        if data.get('reship_policy') not in {'free', 'charge'}:
            raise ReconciliationError('补发管理费处理方式必填')
        data.setdefault('scope', {})
        if set(data['scope']) - {'site', 'country', 'warehouse_id', 'channel', 'pool_id'}:
            raise ReconciliationError('不支持的费率适用范围')
    else:
        raise ReconciliationError('不支持的配置类型')
    if kind in {'contract', 'fee'}:
        start, end = data.get('effective_from'), data.get('effective_to')
        if not start or not end or start >= end:
            raise ReconciliationError('生效日期区间无效，结束日期不包含当日')
        datetime.fromisoformat(start)
        datetime.fromisoformat(end)
    ident = uid()
    conn.execute('INSERT INTO rec_objects(id,kind,name,data,created_at,actor) VALUES (?,?,?,?,?,?)',
                 (ident, kind, name, dump(data), now(), str(actor)))
    return ident


def order_config(conn, order_id):
    if not ready(conn):
        return None
    row = conn.execute('SELECT data FROM rec_orders WHERE order_id=?', (order_id,)).fetchone()
    return json.loads(row['data']) if row else None


def configure_order(conn, order_id, data, actor):
    order = conn.execute('SELECT * FROM orders WHERE id=?' + lock(conn), (order_id,)).fetchone()
    if not order:
        raise ReconciliationError('订单不存在')
    if conn.execute('SELECT id FROM rec_sources WHERE order_id=? AND shipped>0', (order_id,)).fetchone():
        raise ReconciliationError('已出库来源已冻结；请使用差额调整')
    if _table_exists(conn, 'inv_order_state'):
        legacy = conn.execute('SELECT * FROM inv_order_state WHERE order_id=?', (order_id,)).fetchone()
        if legacy and legacy['inv_state'] in {'reserved', 'shipped'}:
            raise ReconciliationError('旧库存处理器已有流水，请先完成库存迁移核对')
    existing = rows(conn, "SELECT status FROM oms_fulfillments WHERE order_id=? AND status NOT IN ('cancelled','superseded')", (order_id,))
    if not order_config(conn, order_id) and existing:
        raise ReconciliationError('已有履约不能直接切换货权；请取消未发履约后启用')
    data = dict(data)
    data['currency'] = currency(data.get('currency', order['currency']))
    policy = data.get('policy', 'fefo')
    if policy not in {'fefo', 'own_first', 'fewest', 'cost', 'quota', 'manual'}:
        raise ReconciliationError('来源策略无效')
    object_(conn, data.get('our_party_id'), 'party')
    if not object_(conn, data['our_party_id'])['data']['internal']:
        raise ReconciliationError('本团队主体必须为内部主体')
    contracts = data.get('contracts', {})
    if not contracts:
        raise ReconciliationError('至少配置一个货权池合同')
    for pool, contract in contracts.items():
        if object_(conn, contract, 'contract')['data']['pool_id'] != pool:
            raise ReconciliationError('合同货权池不匹配')
    if not data.get('fees') and data.get('no_service_fee') is not True:
        raise ReconciliationError('请选择管理费规则，或明确确认无需管理费')
    for fee_id in data.get('fees', []):
        object_(conn, fee_id, 'fee')
    if data.get('provider_mode') not in {'primary', 'split', 'contracts'}:
        raise ReconciliationError('必须明确多服务商计费模式')
    if data['provider_mode'] == 'primary':
        object_(conn, data.get('primary_provider'), 'party')
    elif data['provider_mode'] == 'split':
        data['provider_weights'] = shares(conn, data.get('provider_weights'), '服务费分摊比例')
    for pair in data.get('fx', {}):
        parts = pair.split('/')
        if len(parts) != 2:
            raise ReconciliationError('汇率方向必须为 原币/结算币')
        currency(parts[0]); currency(parts[1])
    for rate in data.get('fx', {}).values():
        if dec(rate) <= 0:
            raise ReconciliationError('汇率必须为正')
    if policy == 'manual' and (not data.get('reason') or not data.get('preferred_batches')):
        raise ReconciliationError('手动选批次必须填写原因')
    if policy == 'quota':
        data['quota'] = shares(conn, data.get('quota'), '轮转配额')
    data['policy'] = policy
    data['actor'] = str(actor)
    conn.execute('INSERT INTO rec_orders(order_id,data,created_at) VALUES (?,?,?) ON CONFLICT(order_id) DO UPDATE SET data=excluded.data',
                 (order_id, dump(data), now()))
    audit(conn, uid(), 'configure_order', order_id, data, actor)


def audit(conn, key, kind, order_id, data, actor='system'):
    conn.execute('INSERT INTO rec_actions(id,kind,order_id,data,created_at,actor) VALUES (?,?,?,?,?,?)',
                 (key, kind, order_id, dump(data), now(), str(actor)))


def done(conn, key):
    return conn.execute('SELECT id FROM rec_actions WHERE id=?', (key,)).fetchone() is not None

def apply_site_profile(conn, order_id):
    """Opt-in rules for newly planned orders; never rewrite legacy fulfillment/stock."""
    if not ready(conn) or order_config(conn,order_id):
        return
    order=conn.execute('SELECT * FROM orders WHERE id=?',(order_id,)).fetchone()
    if not order:
        return
    profiles=[p for p in objects(conn,'profile') if p['data']['site']==order['source']]
    if not profiles:
        return
    profile=profiles[-1];data=profile['data']
    if not data['enabled'] or not order['date_created'] or str(order['date_created'])[:10]<data['starts_on']:
        return
    if conn.execute("SELECT id FROM oms_fulfillments WHERE order_id=? AND status NOT IN ('cancelled','superseded')",(order_id,)).fetchone():
        return
    if _table_exists(conn,'inv_order_state'):
        old=conn.execute('SELECT inv_state FROM inv_order_state WHERE order_id=?',(order_id,)).fetchone()
        if old and old['inv_state'] in {'reserved','shipped'}:
            return
    configure_order(conn,order_id,{**data['config'],'site_profile_id':profile['id']},'site-profile')
