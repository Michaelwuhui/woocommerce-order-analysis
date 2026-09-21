"""Verify original parcels, restore eligible missing tracking, then reconcile.

Never creates another parcel or calls email/inventory endpoints. Conflicting
evidence remains unresolved; every restoration requires a separate source GET.
"""
from __future__ import annotations

import json
import logging
from datetime import datetime, timedelta, timezone

import requests

import db_backend as db
from external_operations import SHIPMENT_LOCK_NAMESPACE, canonical_hash, transition_operation
from oid_utils import make_oid, woo_post_id
from order_shipments import extract_tracking_candidates
from shipment_split import order_products
from shipment_repair import RepairNotAllowed, restore_payload
from sync_service import get_connection

LOG = logging.getLogger(__name__)
LOCK_NAMESPACE = SHIPMENT_LOCK_NAMESPACE
ACTIVE = ('pending', 'external_success', 'reconciliation_required')
GRACE_SECONDS = 300
EVIDENCE_KEY = 'shipment_reconciliation'


def parsed(value, default=None):
    if isinstance(value, str):
        return json.loads(value) if value else default
    return value if value is not None else default


def products(value):
    rows = parsed(value, [])
    result = sorted((str(r['item_id']), str(r['product']), int(r['qty'])) for r in rows)
    if not result or any(q <= 0 or not i or not p for i, p, q in result):
        raise ValueError('商品明细不完整')
    if len({i for i, _, _ in result}) != len(result):
        raise ValueError('商品行重复')
    return result


def carrier(value):
    key = str(value or '').strip().lower().replace(' ', '').replace('-', '').replace('_', '')
    return {'inpostpaczkomaty': 'inpost', 'inpostpl': 'inpost',
            'dpdpl': 'dpd', 'dpdpoland': 'dpd',
            'zasilkovna': 'packeta'}.get(key, key)


class ReviewRequired(ValueError):
    pass


def verify_remote(remote, payload):
    """Require the same order, entire original parcel and item-level coverage."""
    if not isinstance(remote, dict) or str(remote.get('id')) != woo_post_id(payload['order_id']):
        raise ReviewRequired('来源站点未返回对应订单')
    expected = products(payload['items'])
    if products(order_products(remote.get('line_items'))) != expected:
        raise ReviewRequired('来源订单商品、规格或数量与原发货不一致')
    number = str(payload['tracking_number']).strip()
    provider = carrier(payload['carrier_slug'])
    candidates = extract_tracking_candidates(
        remote.get('meta_data'), remote.get('line_items'), remote.get('shipping_lines'), [])
    numbers = {r['tracking_number'] for r in candidates}
    if not numbers:
        if remote.get('status') in {'pending', 'processing', 'offline', 'on-hold'}:
            return 'remote_absent'
        raise ReviewRequired('来源订单已变更状态，但缺少可核验的原运单')
    if numbers != {number}:
        raise ReviewRequired('来源站点存在其他运单，需要核对包裹')
    if remote.get('status') not in {'on-hold', 'shipped', 'completed'}:
        raise ReviewRequired('来源站点有运单，但订单状态仍需核实')
    if any(r['provider'] and carrier(r['provider']) != provider for r in candidates):
        raise ReviewRequired('来源站点物流商与原发货不一致')

    ast_rows = []
    for meta in remote.get('meta_data') or []:
        if meta.get('key') == '_wc_shipment_tracking_items':
            ast_rows.extend(parsed(meta.get('value'), []))
    if ast_rows:
        if len(ast_rows) != 1:
            raise ReviewRequired('来源站点存在多条包裹记录')
        row = ast_rows[0]
        if carrier(row.get('tracking_provider')) != provider:
            raise ReviewRequired('来源包裹缺少匹配的物流商')
        if products(row.get('products_list')) != expected:
            raise ReviewRequired('来源包裹的商品数量与原发货不一致')
        return 'verified'

    # VillaTheme assigns one full line to its tracking record. Custom line-item
    # integrations likewise store carrier_slug beside tracking_number.
    for line in remote.get('line_items') or []:
        if str(line.get('id')) not in {i for i, _, _ in expected}:
            continue
        meta = {r.get('key'): r.get('value') for r in line.get('meta_data') or []}
        if '_vi_wot_order_item_tracking_data' in meta:
            rows = parsed(meta['_vi_wot_order_item_tracking_data'], [])
            if len(rows) != 1 or str(rows[0].get('tracking_number', '')).strip() != number:
                raise ReviewRequired('来源商品行未完整绑定原运单')
            row = rows[0]
            if carrier(row.get('carrier_slug') or row.get('carrier_name')) != provider:
                raise ReviewRequired('来源商品行物流商不一致')
            for key in ('quantity', 'qty'):
                if key in row and int(row[key]) != int(line['quantity']):
                    raise ReviewRequired('来源商品行运单数量不一致')
        elif (str(meta.get('tracking_number', '')).strip() != number
              or carrier(meta.get('carrier_slug')) != provider):
            raise ReviewRequired('来源商品行缺少完整的运单及物流商证据')
    return 'verified'


def _snapshot(connection, operation_id, *, lock=False):
    suffix = ' FOR UPDATE' if lock else ''
    row = connection.execute('SELECT * FROM external_operations WHERE operation_id=?' + suffix,
                             (operation_id,)).fetchone()
    if not row:
        return None
    op = dict(row)
    order = connection.execute(
        'SELECT id,number,source,status,line_items,meta_data,shipping_lines,date_modified,'
        'delivery_confirmed,is_undelivered,is_problem_return FROM orders WHERE id=?' + suffix,
        (op['order_id'],)).fetchone()
    logs = connection.execute('SELECT * FROM shipping_logs WHERE order_id=? ORDER BY id' + suffix,
                              (op['order_id'],)).fetchall()
    site = connection.execute('SELECT id,url,consumer_key,consumer_secret FROM sites WHERE id=?',
                              (op['site_id'],)).fetchone()
    oms = connection.execute('SELECT 1 FROM oms_order_fulfillment_state WHERE order_id=?',
                             (op['order_id'],)).fetchone()
    payload = parsed(op['request_payload'], {})
    provider = connection.execute('SELECT name,tracking_url FROM shipping_carriers WHERE slug=?',
                                  (payload.get('carrier_slug'),)).fetchone()
    return {'op': op, 'order': dict(order) if order else None,
            'logs': [dict(r) for r in logs], 'site': dict(site) if site else None,
            'has_oms': bool(oms), 'carrier': dict(provider) if provider else None}


def validate_local(snapshot):
    op, order, logs, site = (snapshot[k] for k in ('op', 'order', 'logs', 'site'))
    payload = parsed(op['request_payload'], {})
    if not order or not site or order['source'].rstrip('/') != site['url'].rstrip('/'):
        raise ReviewRequired('订单或来源站点配置不一致')
    if (payload.get('order_id') != op['order_id'] or payload.get('site_id') != op['site_id']
            or make_oid(site['id'], woo_post_id(op['order_id'])) != op['order_id']
            or canonical_hash(payload) != op['request_hash']):
        raise ReviewRequired('原发货操作身份或请求记录不一致')
    if snapshot['has_oms'] or any(payload.get(k) for k in
                                 ('new_parcel', 'more_batches', 'is_reship', 'reship_reason')):
        raise ReviewRequired('多仓、分批或补发包裹需要人工核对')
    if order['is_undelivered'] or order['is_problem_return']:
        raise ReviewRequired('订单已经进入拒收或退回处理')
    if order['status'] not in {'pending', 'processing', 'offline', 'on-hold', 'shipped', 'completed'}:
        raise ReviewRequired('本地订单状态已变化，需要核实')
    expected = products(payload['items'])
    if products(order_products(parsed(order['line_items'], []))) != expected:
        raise ReviewRequired('本地商品、规格或数量与原发货不一致')
    if not str(payload.get('tracking_number') or '').strip() or not carrier(payload.get('carrier_slug')):
        raise ReviewRequired('原发货缺少运单或物流商')
    if len(logs) > 1:
        raise ReviewRequired('本地存在多个包裹，需要核对')
    if logs:
        row = logs[0]
        if (row['tracking_number'] != payload['tracking_number']
                or carrier(row['carrier_slug']) != carrier(payload['carrier_slug'])
                or row['is_partial'] or row['is_reship']
                or products(row['items_json']) != expected
                or row['status'] not in {'pending_sync', 'shipped', 'delivered'}):
            raise ReviewRequired('本地发货记录与原操作不一致')
    return payload


def due_operations(order_ids=None, limit=30, *, manual=False):
    if not db.is_postgres_backend():
        return []
    connection = get_connection()
    try:
        clause, args = '', []
        if order_ids is not None:
            if not order_ids:
                return []
            clause = ' AND order_id IN (' + ','.join('?' for _ in order_ids) + ')'
            args.extend(str(oid) for oid in order_ids)
        rows = connection.execute("""
            SELECT operation_id FROM external_operations
            WHERE operation_type='ship_order'
              AND status IN ('pending','external_success','reconciliation_required')
              AND updated_at <= CURRENT_TIMESTAMP - interval '300 seconds'
              AND (? OR COALESCE(NULLIF(external_evidence->'shipment_reconciliation'->>'next_check_at','')
                           ::timestamptz, '-infinity'::timestamptz) <= CURRENT_TIMESTAMP
              )
            """ + clause + ' ORDER BY updated_at LIMIT ?', (bool(manual), *args, max(1, min(limit, 100)))).fetchall()
        return [str(r['operation_id']) for r in rows]
    finally:
        connection.close()


def _duplicate_tracking(connection, order_id, payload):
    return connection.execute(
        'SELECT 1 FROM shipping_logs WHERE order_id<>? AND lower(trim(tracking_number))=lower(trim(?)) '
        'AND lower(trim(carrier_slug))=lower(trim(?)) LIMIT 1',
        (order_id, payload['tracking_number'], payload['carrier_slug'])).fetchone()


def reconcile_operation(operation_id, *, fetch=None, put=None, manual=False):
    """Network reads occur outside a transaction; local completion is atomic."""
    if not db.is_postgres_backend():
        return {'outcome': 'disabled'}
    fetch = fetch or requests.get
    put = put or requests.put
    connection = get_connection()
    locked, order_id = False, None
    try:
        row = connection.execute('SELECT order_id FROM external_operations WHERE operation_id=?',
                                 (operation_id,)).fetchone()
        if not row:
            return {'outcome': 'missing'}
        order_id = row['order_id']
        locked = bool(connection.execute('SELECT pg_try_advisory_lock(?,hashtext(?))',
                                         (LOCK_NAMESPACE, order_id)).fetchone()[0])
        if not locked:
            return {'outcome': 'busy'}
        before = _snapshot(connection, operation_id)
        op = before['op']
        now = datetime.now(timezone.utc)
        age = (now - datetime.fromisoformat(str(op['updated_at']))).total_seconds()
        previous = parsed(op['external_evidence'], {}).get(EVIDENCE_KEY, {})
        if op['operation_type'] != 'ship_order' or op['status'] not in ACTIVE:
            return {'outcome': 'already_closed'}
        if age < GRACE_SECONDS or (not manual and previous.get('next_check_at') and
                datetime.fromisoformat(previous['next_check_at']) > now):
            return {'outcome': 'not_due'}
        connection.commit()
        outcome, reason, remote = 'needs_review', '', None
        evidence = parsed(op['external_evidence'], {})
        try:
            payload = validate_local(before)
            site = before['site']
            url = site['url'].rstrip('/') + '/wp-json/wc/v3/orders/' + woo_post_id(order_id)
            options = {'auth': (site['consumer_key'], site['consumer_secret']),
                       'headers': {'User-Agent': 'WooCommerce API Client-Python/3.0.0',
                                   'Cache-Control': 'no-cache'},
                       'timeout': (5, 30), 'allow_redirects': False}
            response = fetch(url, **options)
            if response.status_code != 200:
                outcome, reason = 'read_failed', f'来源站点读取失败（HTTP {response.status_code}），稍后重查'
            else:
                remote = response.json()
                outcome = verify_remote(remote, payload)
                if outcome == 'remote_absent':
                    reason = '来源站点尚无原运单；等待再次核验后自动补同步'
                    patch = restore_payload(before, payload, remote, evidence, now)
                    if patch:
                        current = _snapshot(connection, operation_id, lock=True)
                        if current != before:
                            connection.rollback()
                            return {'outcome': 'changed_during_check'}
                        if _duplicate_tracking(connection, order_id, payload):
                            raise ReviewRequired('原运单同时绑定其他订单，需要核实')
                        if (current['order']['date_modified'] and remote.get('date_modified') and
                                str(current['order']['date_modified']).replace(' ', 'T') >
                                str(remote['date_modified']).replace(' ', 'T')):
                            raise ReviewRequired('来源站点返回的版本早于本地订单，不能补写')
                        # Persist BEFORE the PUT. A killed worker must leave a
                        # durable attempt and cooldown, not a free blind replay.
                        evidence = parsed(op['external_evidence'], {})
                        repair = evidence.get('shipment_repair', {})
                        repair = {**repair, 'attempts': int(repair.get('attempts', 0)) + 1,
                                  'started_at': datetime.now(timezone.utc).isoformat(),
                                  'http_status': None, 'confirmation': 'awaiting_readback'}
                        evidence['shipment_repair'] = repair
                        connection.execute('UPDATE external_operations SET external_evidence=?::jsonb, '
                                           'attempts=attempts+1,updated_at=CURRENT_TIMESTAMP WHERE operation_id=?',
                                           (json.dumps(evidence, ensure_ascii=False), operation_id))
                        connection.commit()
                        updated = _snapshot(connection, operation_id)
                        if any(updated[k] != before[k] for k in before if k != 'op'):
                            connection.rollback()
                            return {'outcome': 'changed_during_check'}
                        before = updated
                        op = before['op']
                        connection.commit()
                        try:
                            written = put(url, json=patch, **options)
                            repair['http_status'] = written.status_code
                        except requests.RequestException:
                            pass  # A timeout is ambiguous: always read, never replay here.
                        # Even HTTP 200 with a plausible body is insufficient.
                        response = fetch(url, **options)
                        if response.status_code != 200:
                            outcome, reason = 'read_failed', '原运单已尝试补同步，回读暂时失败，稍后核验'
                        else:
                            remote = response.json()
                            outcome = verify_remote(remote, payload)
                            reason = '' if outcome == 'verified' else '原运单补同步后尚未确认，稍后重新核验'
                        repair['confirmation'] = 'verified_get' if outcome == 'verified' else 'not_confirmed'
                        # Keep post-write evidence in memory until the atomic
                        # final update; before remains the actual DB snapshot.
                        evidence['shipment_repair'] = repair
        except requests.RequestException:
            outcome, reason = 'read_failed', '来源站点查询超时或连接失败，稍后重查'
        except (ValueError, TypeError, KeyError) as exc:
            outcome = 'needs_review'
            reason = str(exc) if isinstance(exc, (ReviewRequired, RepairNotAllowed)) else '包裹证据格式不完整，需要核实'

        current = _snapshot(connection, operation_id, lock=True)
        if current != before:
            connection.rollback()
            return {'outcome': 'changed_during_check'}
        # Prevent a legacy request from ever binding one tracking to two orders.
        if outcome == 'verified':
            duplicate = _duplicate_tracking(connection, order_id, payload)
            if duplicate:
                outcome, reason = 'needs_review', '原运单同时绑定其他订单，需要核实'
        if outcome == 'verified' and current['order']['date_modified'] and remote.get('date_modified'):
            if str(current['order']['date_modified']).replace(' ', 'T') > str(remote['date_modified']).replace(' ', 'T'):
                outcome, reason = 'needs_review', '来源站点返回的版本早于本地订单，稍后重新核验'
        checked = datetime.now(timezone.utc)
        checks = int(previous.get('checks', 0)) + 1
        detail = {'outcome': outcome, 'reason': reason, 'checks': checks,
                  'checked_at': checked.isoformat(), 'remote_status': remote.get('status') if isinstance(remote, dict) else None,
                  'next_check_at': None if outcome == 'verified' else
                  (checked + timedelta(seconds=min(3600, 300 * 2 ** min(checks - 1, 4)))).isoformat()}
        evidence[EVIDENCE_KEY] = detail
        if outcome == 'verified':
            if evidence.get('shipment_repair'):
                evidence['shipment_repair']['confirmation'] = 'verified_get'
            order = current['order']
            # Preserve terminal outcomes. No inventory, fulfilment, or email hooks.
            status = 'completed' if order['status'] == 'completed' or order['delivery_confirmed'] else remote['status']
            connection.execute('UPDATE orders SET status=?,line_items=?,meta_data=?,shipping_lines=? WHERE id=?',
                               (status, json.dumps(remote['line_items'], ensure_ascii=False),
                                json.dumps(remote.get('meta_data') or [], ensure_ascii=False),
                                json.dumps(remote.get('shipping_lines') or [], ensure_ascii=False), order_id))
            if current['logs']:
                connection.execute("UPDATE shipping_logs SET status='shipped' WHERE id=? AND status='pending_sync'",
                                   (current['logs'][0]['id'],))
            else:
                user = connection.execute('SELECT id FROM users WHERE id::text=?', (str(op['created_by']),)).fetchone()
                connection.execute("""INSERT INTO shipping_logs
                    (order_id,woo_order_id,source,tracking_number,carrier_slug,shipped_by,shipped_at,items_json,is_partial,status)
                    VALUES (?,?,?,?,?,?,?,?,0,'shipped')""",
                    (order_id, order['number'], order['source'], payload['tracking_number'], payload['carrier_slug'],
                     user['id'] if user else None, datetime.fromisoformat(str(op['created_at'])).astimezone(timezone.utc).strftime('%Y-%m-%d %H:%M:%S'),
                     json.dumps(payload['items'], ensure_ascii=False)))
            if op['status'] != 'external_success':
                transition_operation(connection, operation_id, 'external_success', evidence=evidence,
                                     external_reference=woo_post_id(order_id))
            transition_operation(connection, operation_id, 'local_committed', evidence=evidence)
            connection.execute("""INSERT INTO order_notes
                (order_id,note,date_created,customer_note,author,added_by_user)
                VALUES (?,?,datetime('now'),0,'系统自动对账',0)""",
                (order_id, ('原运单已补同步至来源站点并回读确认；' if evidence.get('shipment_repair') else '') +
                 '发货同步对账已补齐：来源站点原运单、物流商及商品数量核验一致；保留原发货时间。'))
        else:
            connection.execute("""UPDATE external_operations SET status='reconciliation_required',
                external_evidence=?::jsonb,last_error=?,updated_at=CURRENT_TIMESTAMP WHERE operation_id=?""",
                (json.dumps(evidence, ensure_ascii=False), reason, operation_id))
        connection.commit()
        LOG.info('Shipment reconciliation order=%s outcome=%s', order_id, outcome)
        return {'outcome': outcome, 'order_id': order_id}
    except Exception:
        connection.rollback()
        raise
    finally:
        if locked:
            connection.rollback()
            connection.execute('SELECT pg_advisory_unlock(?,hashtext(?))', (LOCK_NAMESPACE, order_id))
            connection.commit()
        connection.close()


def public_statuses(connection, order_ids):
    if not db.is_postgres_backend() or not order_ids:
        return {}
    rows = connection.execute('SELECT order_id,status,external_evidence,request_payload,created_at FROM external_operations '
                              "WHERE operation_type='ship_order' AND order_id IN (" +
                              ','.join('?' for _ in order_ids) + ') ORDER BY created_at DESC', tuple(order_ids)).fetchall()
    result = {}
    for row in rows:
        oid = row['order_id']
        if oid in result:
            continue
        detail = parsed(row['external_evidence'], {}).get(EVIDENCE_KEY, {})
        outcome = detail.get('outcome', 'waiting')
        if row['status'] in {'local_committed', 'notified'} and outcome != 'verified':
            outcome = 'completed'
        if row['status'] == 'failed':
            outcome = 'rejected'
        label = {'verified': '已补齐', 'remote_absent': '需处理', 'needs_review': '需处理',
                 'rejected': '可重试', 'completed': '已同步'}.get(outcome, '待核验')
        payload = parsed(row['request_payload'], {})
        result[oid] = {'outcome': outcome, 'label': label, 'reason': detail.get('reason') or
                       ('来源站点原包裹已核验，本地记录已补齐' if outcome == 'verified' else
                        '发货同步已完成' if outcome == 'completed' else
                        '站点明确拒绝了上次发货，可核对原运单后重试' if outcome == 'rejected' else
                        '系统每分钟检查，发货操作超过 5 分钟后自动核对站点'),
                       'checked_at': detail.get('checked_at'), 'next_check_at': detail.get('next_check_at'),
                       'pending': row['status'] in ACTIVE, 'tracking_number': payload.get('tracking_number'),
                       'carrier_slug': payload.get('carrier_slug'), 'shipped_at': str(row['created_at']),
                       'retry_allowed': row['status'] == 'failed'}
    return result
