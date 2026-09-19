"""Warehouse-scoped receipt, transfer and stocktake documents; no external writes."""
import datetime
import hashlib
import json
import re
from functools import wraps

from flask import Blueprint, jsonify, render_template, request
from flask_login import current_user, login_required
import db_backend
from inv_temporary_access import active_grants, utc_now

from inv_common import (get_conn, record_movement, can_manage_inventory,
                        visible_warehouse_ids, inv_view_required, _table_exists)

workflow_bp = Blueprint('inv_workflows', __name__)
CAPABILITIES = ('can_receive', 'can_transfer', 'can_stocktake', 'can_approve')
FINAL = {'approved', 'rejected', 'cancelled', 'completed', 'completed_with_difference'}


class WorkflowError(Exception):
    def __init__(self, message, status=400):
        super().__init__(message)
        self.status = status


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat(timespec='microseconds')


def actor():
    return int(current_user.id), (current_user.name or current_user.username)


def require(condition, message, status=400):
    if not condition:
        raise WorkflowError(message, status)


def integer(value, minimum=0):
    require(not isinstance(value, bool) and re.fullmatch(r'\d+', str(value or 0)) is not None,
            '数量和编号必须是整数')
    result = int(value or 0)
    require(minimum <= result <= 100000000, '数量或编号超出允许范围')
    return result


def permissions(conn, wid):
    if can_manage_inventory():
        return dict.fromkeys(CAPABILITIES, True)
    row = conn.execute('SELECT * FROM inv_operation_permissions WHERE user_id=? AND warehouse_id=?',
                       (current_user.id, wid)).fetchone()
    result = {key: bool(row and row[key]) for key in CAPABILITIES}
    if active_grants(conn, current_user.id, wid):
        result['can_stocktake'] = True
    return result


def visible(wid):
    ids = visible_warehouse_ids()
    return ids is None or wid in ids


def allowed(conn, wid, capability):
    require(visible(wid) and permissions(conn, wid)[capability], '没有该仓库的操作权限', 403)


def warehouse(conn, wid, destination=True):
    row = conn.execute('''SELECT w.*,COALESCE(wi.inventory_authority,'local') AS authority
        FROM warehouses w LEFT JOIN oms_warehouse_integrations wi ON wi.warehouse_id=w.id
        WHERE w.id=? AND w.is_active=1''', (wid,)).fetchone()
    require(row is not None, '仓库不存在或已停用')
    require(row['authority'] in (('local',) if destination else ('local', 'external_wms', 'manual_partner')),
            '该仓不采用本地数量库存，不能在此入库或盘点')
    return row


def eligible_sku(conn, wid, sid):
    return conn.execute('''SELECT k.id FROM inv_skus k WHERE k.id=? AND k.is_active=1 AND (
        EXISTS(SELECT 1 FROM inv_stock s WHERE s.warehouse_id=? AND s.sku_id=k.id)
        OR EXISTS(SELECT 1 FROM oms_sku_warehouses sw WHERE sw.warehouse_id=?
                  AND sw.sku_id=k.id AND sw.is_enabled=1))''', (sid, wid, wid)).fetchone()


def lock_suffix(conn):
    return ' FOR UPDATE' if hasattr(conn, '_raw') else ''


def stock(conn, wid, sid, lock=False):
    if lock:
        conn.execute('''INSERT INTO inv_stock(warehouse_id,sku_id,on_hand,reserved)
            VALUES (?,?,0,0) ON CONFLICT(warehouse_id,sku_id) DO NOTHING''', (wid, sid))
    row = conn.execute('SELECT on_hand,reserved FROM inv_stock WHERE warehouse_id=? AND sku_id=?'
                       + (lock_suffix(conn) if lock else ''), (wid, sid)).fetchone()
    movement = conn.execute('SELECT COALESCE(MAX(id),0) FROM inv_movements WHERE warehouse_id=? AND sku_id=?',
                            (wid, sid)).fetchone()[0]
    return (int(row['on_hand']) if row else 0, int(row['reserved']) if row else 0, int(movement))


def event(conn, did, action, detail, system=False):
    uid, name = (0, '系统自动审批') if system else actor()
    conn.execute('''INSERT INTO inv_document_events
        (document_id,action,actor_id,actor_name,detail,created_at) VALUES (?,?,?,?,?,?)''',
        (did, action, uid, name, json.dumps(detail, ensure_ascii=False), now()))


def document(conn, did, lock=False):
    row = conn.execute('SELECT * FROM inv_documents WHERE id=?' + (lock_suffix(conn) if lock else ''),
                       (did,)).fetchone()
    require(row is not None, '单据不存在', 404)
    # A transfer is shared with its two warehouses; unrelated documents remain hidden.
    require(visible(row['warehouse_id']) or (row['source_warehouse_id'] and visible(row['source_warehouse_id'])),
            '没有该单据的查看权限', 403)
    return dict(row)


def lines(conn, did):
    return [dict(r) for r in conn.execute('''SELECT l.*,k.sku_code,k.name FROM inv_document_lines l
        JOIN inv_skus k ON k.id=l.sku_id WHERE l.document_id=? ORDER BY l.sku_id''', (did,)).fetchall()]


def received(conn, did):
    return {r['sku_id']: dict(r) for r in conn.execute('''SELECT l.sku_id,
        SUM(l.qty) AS good, SUM(l.damaged_qty) AS damaged, SUM(l.short_qty) AS short
        FROM inv_document_lines l JOIN inv_documents d ON d.id=l.document_id
        WHERE d.parent_id=? AND d.status='approved' GROUP BY l.sku_id''', (did,)).fetchall()}


def remaining(conn, parent):
    totals = received(conn, parent['id'])
    return {line['sku_id']: line['qty'] - sum(int(totals.get(line['sku_id'], {}).get(k, 0))
            for k in ('good', 'damaged', 'short')) for line in lines(conn, parent['id'])}


def transact(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        conn = get_conn()
        try:
            require(_table_exists(conn, 'inv_documents'), '入库盘点模块尚未完成数据库升级', 503)
            if request.method != 'GET':
                conn.commit()  # finish the schema probe before SQLite's write transaction
                conn.execute('BEGIN IMMEDIATE')
            result = func(conn, *args, **kwargs)
            conn.commit()
            return jsonify(result)
        except WorkflowError as exc:
            conn.rollback()
            return jsonify(error=str(exc)), exc.status
        except db_backend.IntegrityError:
            conn.rollback()
            return jsonify(error='该凭证已存在或数据关联已变化，请检查原单据后重试'), 409
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
    return wrapper


@workflow_bp.route('/inventory/operations')
@login_required
@inv_view_required
def page():
    return render_template('inv_operations.html', permission_admin=can_manage_inventory())


@workflow_bp.route('/api/inv/operations/context')
@login_required
@inv_view_required
@transact
def context(conn):
    result = []
    for row in conn.execute('''SELECT w.id,w.name,COALESCE(wi.inventory_authority,'local') AS authority
        FROM warehouses w LEFT JOIN oms_warehouse_integrations wi ON wi.warehouse_id=w.id
        WHERE w.is_active=1 ORDER BY w.name''').fetchall():
        if visible(row['id']) and row['authority'] in ('local', 'external_wms', 'manual_partner'):
            grants = active_grants(conn, current_user.id, row['id'])
            result.append(dict(row, **permissions(conn, row['id']),
                               auto_stocktake_grant=grants[0]['id'] if grants else None,
                               auto_stocktake_expires_at=grants[0]['expires_at'] if grants else None))
    return {'warehouses': result, 'user_id': int(current_user.id), 'can_admin': can_manage_inventory()}


@workflow_bp.route('/api/inv/operations/catalog/<int:wid>')
@login_required
@inv_view_required
@transact
def catalog(conn, wid):
    require(visible(wid), '无权查看该仓', 403)
    warehouse(conn, wid, destination=False)
    rows = conn.execute('''SELECT k.id,k.sku_code,k.name,COALESCE(s.on_hand,0) AS on_hand,
        COALESCE(s.reserved,0) AS reserved,
        COALESCE((SELECT MAX(m.id) FROM inv_movements m WHERE m.warehouse_id=? AND m.sku_id=k.id),0) AS movement
        FROM inv_skus k LEFT JOIN inv_stock s ON s.sku_id=k.id AND s.warehouse_id=?
        WHERE k.is_active=1 AND (s.sku_id IS NOT NULL OR EXISTS(
            SELECT 1 FROM oms_sku_warehouses sw WHERE sw.sku_id=k.id AND sw.warehouse_id=? AND sw.is_enabled=1))
        ORDER BY k.sku_code''', (wid, wid, wid)).fetchall()
    return [dict(r) for r in rows]


@workflow_bp.route('/api/inv/operations', methods=['GET'])
@login_required
@inv_view_required
@transact
def listing(conn):
    ids = visible_warehouse_ids()
    sql, params = 'SELECT d.*,w.name AS warehouse_name,src.name AS source_name FROM inv_documents d JOIN warehouses w ON w.id=d.warehouse_id LEFT JOIN warehouses src ON src.id=d.source_warehouse_id WHERE 1=1', []
    if ids is not None:
        if not ids:
            return []
        placeholders = ','.join('?' for _ in ids)
        sql += f' AND (d.warehouse_id IN ({placeholders}) OR d.source_warehouse_id IN ({placeholders}))'
        params += ids + ids
    for key in ('kind', 'status', 'warehouse_id'):
        if request.args.get(key):
            sql += f' AND d.{key}=?'
            params.append(request.args[key])
    return [dict(r) for r in conn.execute(sql + ' ORDER BY d.id DESC LIMIT 300', params).fetchall()]


@workflow_bp.route('/api/inv/operations/<int:did>')
@login_required
@inv_view_required
@transact
def detail(conn, did):
    doc = document(conn, did)
    doc['lines'] = lines(conn, did)
    doc['events'] = [dict(r) for r in conn.execute('SELECT * FROM inv_document_events WHERE document_id=? ORDER BY id', (did,)).fetchall()]
    if doc['kind'] in ('replenishment', 'transfer'):
        doc['remaining'] = remaining(conn, doc)
        doc['received'] = received(conn, did)
        doc['receipts'] = [dict(r) for r in conn.execute('SELECT id,status,created_name,created_at FROM inv_documents WHERE parent_id=? ORDER BY id', (did,)).fetchall()]
    review_wid = doc['source_warehouse_id'] if doc['kind'] == 'transfer' else doc['warehouse_id']
    doc['can_review'] = (doc['status'] == 'submitted' and permissions(conn, review_wid)['can_approve']
                         and doc['created_by'] != int(current_user.id))
    doc['can_receive'] = (doc['kind'] in ('replenishment', 'transfer') and doc['status'] in ('open', 'in_transit', 'partial')
                          and permissions(conn, doc['warehouse_id'])['can_receive'])
    return doc


@workflow_bp.route('/api/inv/operations', methods=['POST'])
@login_required
@inv_view_required
@transact
def create(conn):
    data = request.get_json(silent=True)
    return create_document(conn, data)


def create_document(conn, data, *, admin_fulfillment=None):
    # Only the dedicated, authenticated route supplies this server-side context.
    # JSON flags on the ordinary document API never enable immediate admin entry.
    if admin_fulfillment is not None:
        require(current_user.username == 'admin', '仅 admin 可快速调整库存', 403)
    require(isinstance(data, dict), '请求必须是 JSON 对象')
    key = str(data.get('request_key', ''))
    require(re.fullmatch(r'[A-Za-z0-9_-]{16,80}', key), '请提供有效的防重复提交标识')
    digest = hashlib.sha256(json.dumps(data, sort_keys=True, ensure_ascii=False).encode()).hexdigest()
    # Serialize the same key even before its document exists (PG; SQLite uses BEGIN IMMEDIATE).
    if hasattr(conn, '_raw'):
        conn.execute('SELECT pg_advisory_xact_lock(?)', (int.from_bytes(hashlib.sha256(key.encode()).digest()[:8], 'big', signed=True),))
    previous = conn.execute('SELECT * FROM inv_documents WHERE request_key=?', (key,)).fetchone()
    if previous:
        require(previous['created_by'] == int(current_user.id) and previous['request_hash'] == digest,
                '提交标识已被其他内容使用，请刷新后核查', 409)
        document(conn, previous['id'])
        return {'id': previous['id'], 'status': previous['status'], 'replayed': True}
    kind = data.get('kind')
    require(kind in ('replenishment', 'transfer', 'receipt', 'stocktake'), '未知单据类型')
    wid = integer(data.get('warehouse_id'), 1)
    warehouse(conn, wid)
    quick_skus = None
    if admin_fulfillment is not None:
        require(kind == 'stocktake' and not data.get('auto_approve'), '快捷调整仅允许管理员盘点')
        scope = fulfillment_stocktake_context(conn, admin_fulfillment)
        require(integer(data.get('fulfillment_revision'), 1) == scope['revision'],
                '履约计划已变化，请重新打开库存调整', 409)
        quick_skus = {item['sku_id'] for item in scope['items']}
    automatic = data.get('auto_approve', False)
    require(isinstance(automatic, bool), '自动审批参数必须为布尔值')
    require(not automatic or kind == 'stocktake', '临时授权仅允许库存盘点')
    grant = None
    if automatic:
        grants = active_grants(conn, current_user.id, wid, lock=True)
        grant = next((g for g in grants if g['id'] == integer(data.get('grant_id'), 1)), None)
        require(grant is not None, '临时库存授权已过期、撤销或不属于此仓库，请联系管理员', 403)
    capability = 'can_stocktake' if kind == 'stocktake' else 'can_receive'
    source, source_authority, parent = None, None, None
    if kind == 'transfer':
        source = integer(data.get('source_warehouse_id'), 1)
        require(source != wid and visible(wid), '调出仓和调入仓必须不同且可见')
        source_authority = warehouse(conn, source, destination=False)['authority']
        allowed(conn, source, 'can_transfer')
    else:
        allowed(conn, wid, capability)
    if kind == 'receipt':
        parent = document(conn, integer(data.get('parent_id'), 1), lock=True)
        require(parent['kind'] in ('replenishment', 'transfer') and parent['warehouse_id'] == wid,
                '收货单必须关联本仓的补货或调拨单')
        require(parent['status'] in ('open', 'in_transit', 'partial'), '上游单据尚未调出或已经结束', 409)
        pending = conn.execute("SELECT id FROM inv_documents WHERE parent_id=? AND status='submitted'", (parent['id'],)).fetchone()
        require(not pending, '该单已有待审核收货，请先审核，避免重复登记', 409)
    reference = str(data.get('reference', '')).strip()
    note = str(data.get('note', '')).strip()
    require(reference and len(reference) <= 200 and note and len(note) <= 2000,
            '请填写到货批次/调拨凭证/盘点编号及说明（分别不超过200/2000字）')
    if kind == 'receipt':
        require(not conn.execute("SELECT id FROM inv_documents WHERE parent_id=? AND reference=? AND status<>'rejected'", (parent['id'], reference)).fetchone(),
                '该到货批次已登记，请检查原收货单', 409)
    items = data.get('items')
    require(isinstance(items, list) and 0 < len(items) <= 500, '需要1到500条明细')
    clean, seen = [], set()
    balances = remaining(conn, parent) if parent else {}
    for item in items:
        require(isinstance(item, dict), '明细格式错误')
        if quick_skus is not None:
            require(all(not isinstance(item.get(field), bool)
                        and re.fullmatch(r'\d+', str(item.get(field))) for field in
                        ('sku_id', 'qty', 'baseline', 'baseline_reserved', 'baseline_movement')),
                    '请提交完整的实盘数量和库存快照')
        sid = integer(item.get('sku_id'), 1)
        if quick_skus is not None:
            require(sid in quick_skus, '只能调整该订单在本仓的受管商品', 403)
        require(sid not in seen, '同一单据不能重复填写同一SKU')
        seen.add(sid)
        require(eligible_sku(conn, wid, sid), 'SKU不属于该仓的受管库存；请先由库存负责人确认仓库SKU映射')
        if source and source_authority == 'local':
            require(eligible_sku(conn, source, sid), '调出仓未管理该SKU')
        qty = integer(item.get('qty'), 1 if kind in ('replenishment', 'transfer') else 0)
        damaged = integer(item.get('damaged_qty')) if kind == 'receipt' else 0
        short = integer(item.get('short_qty')) if kind == 'receipt' else 0
        if parent:
            require(sid in balances and 0 < qty + damaged + short <= balances[sid], '实收、破损、确认少收合计不能超过待收数量')
        baseline = baseline_reserved = movement = 0
        if kind == 'stocktake':
            baseline = integer(item.get('baseline'))
            baseline_reserved = integer(item.get('baseline_reserved'))
            movement = integer(item.get('baseline_movement'))
            require(stock(conn, wid, sid, lock=True) == (baseline, baseline_reserved, movement),
                    '盘点期间库存发生变化，请重新核对实物并加载账面数量', 409)
            require(qty >= baseline_reserved, '实盘低于订单预留：请先核查缺货订单并处理预留', 409)
        clean.append((sid, qty, damaged, short, baseline, baseline_reserved, movement))
    status = 'open' if kind == 'replenishment' else 'submitted'
    uid, name = actor()
    cur = conn.execute('''INSERT INTO inv_documents(kind,status,warehouse_id,source_warehouse_id,source_authority,parent_id,
        reference,note,request_key,request_hash,created_by,created_name,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (kind, status, wid, source, source_authority, parent['id'] if parent else None, reference, note, key, digest, uid, name, now()))
    did = cur.lastrowid
    for item in clean:
        conn.execute('''INSERT INTO inv_document_lines(document_id,sku_id,qty,damaged_qty,short_qty,
            baseline,baseline_reserved,baseline_movement) VALUES (?,?,?,?,?,?,?,?)''', (did, *item))
    event(conn, did, 'created', {'status': status, 'reference': reference})
    if admin_fulfillment is not None:
        doc = document(conn, did)
        doc['order_id'] = admin_fulfillment['order_id']
        apply_stocktake(conn, doc, lines(conn, did), reviewer='admin 快捷盘点直接入账')
        status = 'approved'
        review_note = '内置 admin 在履约详情执行快捷盘点，直接入账'
        conn.execute('''UPDATE inv_documents SET status=?,reviewed_by=?,reviewed_name=?,review_note=?,reviewed_at=?
            WHERE id=?''', (status, uid, name, review_note, now(), did))
        event(conn, did, 'admin_stocktake_approved', {
            'status': status, 'note': review_note, 'authorization': 'builtin_admin',
            'order_id': admin_fulfillment['order_id'], 'fulfillment_id': admin_fulfillment['id'],
            'warehouse_id': wid, 'operator_id': uid,
        })
    elif grant:
        # Document, quantity movement and automatic approval are a single transaction.
        require(grant['expires_at'] > utc_now(), '临时库存授权已到期，请重新申请', 403)
        doc = document(conn, did)
        apply_stocktake(conn, doc, lines(conn, did), reviewer='系统自动审批')
        require(grant['expires_at'] > utc_now(), '临时库存授权已到期，本次未入账', 403)
        status = 'approved'
        review_note = f"依据临时授权#{grant['id']}（授权人：{grant['created_name']}，有效至：{grant['expires_at']}）自动通过"
        conn.execute('''UPDATE inv_documents SET status=?,reviewed_name=?,review_note=?,reviewed_at=?
            WHERE id=?''', (status, '系统自动审批', review_note, now(), did))
        event(conn, did, 'auto_approved', {'status': status, 'note': review_note,
              'grant_id': grant['id'], 'granted_by': grant['created_by'],
              'expires_at': grant['expires_at'], 'operator_id': uid}, system=True)
    return {'id': did, 'status': status}


def movement(conn, doc, sid, qty, movement_type, wid, apply=True, reviewer=None):
    uid, name = actor()
    return record_movement(conn, warehouse_id=wid, sku_id=sid, qty_delta=qty,
        movement_type=movement_type, ref_type='inventory_document', ref_id=str(doc['id']),
        order_id=doc.get('order_id'), operator_id=uid, operator_name=name,
        note=f"单据#{doc['id']} {doc['reference']}；提交:{doc['created_name']}；审核:{reviewer or name}", apply_stock=apply)


def apply_stocktake(conn, doc, items, reviewer=None):
    for item in items:
        actual = stock(conn, doc['warehouse_id'], item['sku_id'], lock=True)
        require(actual == (item['baseline'], item['baseline_reserved'], item['baseline_movement']),
                '账面已发生出入库或预留变化，不能覆盖新库存；请重新盘点', 409)
        require(item['qty'] >= actual[1], '实盘低于预留，请先处理受影响订单', 409)
        delta = item['qty'] - actual[0]
        if delta:
            movement(conn, doc, item['sku_id'], delta, 'adjust', doc['warehouse_id'], reviewer=reviewer)


def fulfillment_stocktake_context(conn, fulfillment):
    """Read the real order scope, including mapped but unallocated shortage lines."""
    state = conn.execute('SELECT revision FROM oms_order_fulfillment_state WHERE order_id=?',
                         (fulfillment['order_id'],)).fetchone()
    require(state and state['revision'] == fulfillment['revision']
            and fulfillment['status'] not in ('superseded', 'cancelled'),
            '履约计划已变化或取消，请重新打开当前履约详情', 409)
    require(fulfillment['mode'] == 'internal', '外部仓不能在此修改库存', 409)
    wh = warehouse(conn, fulfillment['warehouse_id'])
    rows = conn.execute('''SELECT oi.id,oi.sku_id,oi.name,oi.shortage_qty,oi.ordered_qty,
        k.sku_code,k.name AS sku_name FROM oms_order_items oi
        LEFT JOIN inv_skus k ON k.id=oi.sku_id WHERE oi.order_id=? AND (
            oi.shortage_qty>0 OR EXISTS(SELECT 1 FROM oms_fulfillment_items fi
                WHERE fi.fulfillment_id=? AND fi.order_item_id=oi.id))
        ORDER BY oi.sku_id,oi.id''', (fulfillment['order_id'], fulfillment['id'])).fetchall()
    items, unavailable = {}, []
    for row in rows:
        sid = row['sku_id']
        if not sid or not eligible_sku(conn, wh['id'], sid):
            unavailable.append({'name': row['name'], 'shortage_qty': row['shortage_qty']})
            continue
        if sid not in items:
            on_hand, reserved, last_movement = stock(conn, wh['id'], sid)
            items[sid] = {'sku_id': sid, 'sku_code': row['sku_code'], 'name': row['sku_name'],
                          'on_hand': on_hand, 'reserved': reserved, 'available': on_hand-reserved,
                          'movement': last_movement, 'ordered_qty': 0, 'shortage_qty': 0}
        items[sid]['ordered_qty'] += int(row['ordered_qty'])
        items[sid]['shortage_qty'] += int(row['shortage_qty'])
    return {'fulfillment_id': fulfillment['id'], 'order_id': fulfillment['order_id'],
            'order_number': fulfillment['order_number'], 'revision': fulfillment['revision'],
            'warehouse_id': wh['id'], 'warehouse_name': wh['name'],
            'items': list(items.values()), 'unavailable': unavailable}


@workflow_bp.route('/api/inv/operations/fulfillment/<fulfillment_id>/stocktake', methods=['GET', 'POST'])
@login_required
@transact
def fulfillment_stocktake(conn, fulfillment_id):
    require(current_user.username == 'admin', '仅 admin 可快速调整库存', 403)
    fulfillment = conn.execute('''SELECT f.*,o.number AS order_number FROM oms_fulfillments f
        JOIN orders o ON o.id=f.order_id WHERE f.id=?''', (fulfillment_id,)).fetchone()
    require(fulfillment is not None, '履约单不存在', 404)
    if request.method == 'GET':
        return fulfillment_stocktake_context(conn, fulfillment)
    data = request.get_json(silent=True)
    require(isinstance(data, dict), '请求必须是 JSON 对象')
    require(not set(data) - {'request_key', 'note', 'items', 'revision'}, '包含不支持的库存调整参数')
    items = data.get('items')
    require(isinstance(items, list) and 0 < len(items) <= 500 and all(isinstance(i, dict) for i in items),
            '需要1到500条库存明细')
    require(isinstance(data.get('note'), str) and data['note'].strip(), '请填写库存调整原因')
    # Deterministic ordering also locks SKU rows in a consistent order on PostgreSQL.
    items = sorted(items, key=lambda item: integer(item.get('sku_id'), 1))
    key = data.get('request_key', '')
    payload = {'kind': 'stocktake', 'warehouse_id': fulfillment['warehouse_id'],
               'reference': f"履约盘点 {fulfillment['id']} {key}", 'note': data['note'].strip(),
               'request_key': key, 'items': items, 'fulfillment_id': fulfillment['id'],
               'fulfillment_revision': data.get('revision'), 'order_id': fulfillment['order_id']}
    # Existing receipt lookup precedes plan/snapshot checks, so a timed-out request
    # remains safely replayable even after the restock worker replans the order.
    return create_document(conn, payload, admin_fulfillment=fulfillment)


@workflow_bp.route('/api/inv/operations/<int:did>/review', methods=['POST'])
@login_required
@inv_view_required
@transact
def review(conn, did):
    # Always lock a parent before its receipt: same order as receipt creation.
    preview = document(conn, did)
    parent = document(conn, preview['parent_id'], lock=True) if preview['parent_id'] else None
    doc = document(conn, did, lock=True)
    review_wid = doc['source_warehouse_id'] if doc['kind'] == 'transfer' else doc['warehouse_id']
    allowed(conn, review_wid, 'can_approve')
    require(doc['created_by'] != int(current_user.id), '不能审核自己提交的单据，请由另一名库存负责人审核', 403)
    data = request.get_json(silent=True) or {}
    decision = data.get('decision')
    require(decision in ('approve', 'reject'), '审核动作无效')
    note = str(data.get('note', '')).strip()
    require(note and len(note) <= 2000, '请填写审核意见（不超过2000字）')
    if doc['reviewed_at']:
        require((decision == 'reject') == (doc['status'] == 'rejected'), '单据已按另一审核结果处理', 409)
        return {'id': did, 'status': doc['status'], 'replayed': True}
    require(doc['status'] == 'submitted', '单据不是待审核状态', 409)
    items = lines(conn, did)
    status = 'rejected'
    if decision == 'approve':
        warehouse(conn, doc['warehouse_id'])
        status = 'approved'
        if doc['kind'] == 'transfer':
            source = warehouse(conn, review_wid, destination=False)
            require(source['authority'] == doc['source_authority'], '调出仓库存管理方式已变化，请驳回后重新登记', 409)
            for item in items:
                if source['authority'] == 'local':
                    on_hand, reserved, _ = stock(conn, review_wid, item['sku_id'], lock=True)
                    require(on_hand - reserved >= item['qty'], '调出仓可用库存不足，不能挪用订单预留', 409)
                    movement(conn, doc, item['sku_id'], -item['qty'], 'transfer_out', review_wid)
                else:
                    event(conn, did, 'external_dispatch_confirmed', {'sku_id': item['sku_id'], 'qty': item['qty'],
                          'reference': doc['reference'], 'external_stock_modified': False})
            status = 'in_transit'
        elif doc['kind'] == 'receipt':
            require(parent and parent['status'] in ('open', 'in_transit', 'partial'), '上游单据已结束', 409)
            balances = remaining(conn, parent)
            for item in items:
                require(item['qty'] + item['damaged_qty'] + item['short_qty'] <= balances[item['sku_id']],
                        '收货数量超过剩余数量，请驳回并重新核对', 409)
                if item['qty']:
                    stock(conn, doc['warehouse_id'], item['sku_id'], lock=True)
                    movement(conn, doc, item['sku_id'], item['qty'],
                             'transfer_in' if parent['kind'] == 'transfer' else 'purchase_in', doc['warehouse_id'])
        elif doc['kind'] == 'stocktake':
            apply_stocktake(conn, doc, items)
        else:
            raise WorkflowError('此类型不需要审核')
    uid, name = actor()
    conn.execute('''UPDATE inv_documents SET status=?,reviewed_by=?,reviewed_name=?,review_note=?,reviewed_at=? WHERE id=?''',
                 (status, uid, name, note, now(), did))
    event(conn, did, decision, {'note': note, 'status': status})
    if parent and decision == 'approve':
        pending = remaining(conn, parent)
        totals = received(conn, parent['id'])
        has_difference = any(r['damaged'] or r['short'] for r in totals.values())
        state = 'partial' if any(pending.values()) else ('completed_with_difference' if has_difference else 'completed')
        conn.execute('UPDATE inv_documents SET status=? WHERE id=?', (state, parent['id']))
        event(conn, parent['id'], 'receipt_approved', {'receipt_id': did, 'status': state})
    return {'id': did, 'status': status}


@workflow_bp.route('/api/inv/operations/<int:did>/cancel', methods=['POST'])
@login_required
@inv_view_required
@transact
def cancel(conn, did):
    doc = document(conn, did, lock=True)
    require(doc['created_by'] == int(current_user.id) or can_manage_inventory(), '不能撤销他人的单据', 403)
    require(doc['status'] in ('submitted', 'open'), '已入账/已调出单据不能撤销，请走更正单据', 409)
    require(not conn.execute("SELECT id FROM inv_documents WHERE parent_id=? AND status NOT IN ('rejected','cancelled')", (did,)).fetchone(),
            '已有收货记录，不能撤销', 409)
    note = str((request.get_json(silent=True) or {}).get('note', '')).strip()
    require(note and len(note) <= 2000, '请填写撤销原因')
    conn.execute("UPDATE inv_documents SET status='cancelled' WHERE id=?", (did,))
    event(conn, did, 'cancelled', {'note': note})
    return {'id': did, 'status': 'cancelled'}


@workflow_bp.route('/api/inv/operations/permissions', methods=['GET', 'PUT'])
@login_required
@transact
def permission_settings(conn):
    # Granting warehouse approval is an administrator action, not a warehouse approver action.
    require(current_user.username == 'admin', '仅内置管理员可修改仓库操作授权', 403)
    if request.method == 'PUT':
        data = request.get_json(silent=True) or {}
        uid, wid = integer(data.get('user_id'), 1), integer(data.get('warehouse_id'), 1)
        require(conn.execute('SELECT id FROM users WHERE id=?', (uid,)).fetchone(), '账号不存在')
        warehouse(conn, wid, destination=False)
        values = [integer(data.get(key)) for key in CAPABILITIES]
        require(all(value in (0, 1) for value in values), '权限值必须为0或1')
        before = conn.execute('SELECT * FROM inv_operation_permissions WHERE user_id=? AND warehouse_id=?', (uid, wid)).fetchone()
        conn.execute('''INSERT INTO inv_operation_permissions
            (user_id,warehouse_id,can_receive,can_transfer,can_stocktake,can_approve) VALUES (?,?,?,?,?,?)
            ON CONFLICT(user_id,warehouse_id) DO UPDATE SET can_receive=excluded.can_receive,
            can_transfer=excluded.can_transfer,can_stocktake=excluded.can_stocktake,can_approve=excluded.can_approve''',
            (uid, wid, *values))
        event(conn, None, 'permission_changed', {'user_id': uid, 'warehouse_id': wid,
              'before': dict(before) if before else None, 'after': dict(zip(CAPABILITIES, values))})
    return {'users': [dict(r) for r in conn.execute('SELECT id,name,username,can_manage_inventory FROM users ORDER BY id').fetchall()],
            'permissions': [dict(r) for r in conn.execute('SELECT * FROM inv_operation_permissions ORDER BY user_id,warehouse_id').fetchall()]}


@workflow_bp.route('/settings/inventory-access')
@login_required
def temporary_access_page():
    if current_user.username != 'admin':
        return '仅内置管理员可配置临时库存权限', 403
    return render_template('inv_temporary_access.html')


@workflow_bp.route('/api/inv/operations/temporary-access', methods=['GET', 'POST'])
@login_required
@transact
def temporary_access(conn):
    require(current_user.username == 'admin', '仅内置管理员可配置临时库存权限', 403)
    require(_table_exists(conn, 'inv_temporary_stock_grants'), '临时库存授权尚未完成数据库升级', 503)
    if request.method == 'POST':
        data = request.get_json(silent=True) or {}
        uid, wid = integer(data.get('user_id'), 1), integer(data.get('warehouse_id'), 1)
        require(conn.execute('SELECT id FROM users WHERE id=?', (uid,)).fetchone(), '账号不存在')
        warehouse(conn, wid)
        reason = str(data.get('reason', '')).strip()
        require(0 < len(reason) <= 2000, '请填写临时授权原因（不超过2000字）')
        try:
            expires = datetime.datetime.fromisoformat(str(data.get('expires_at', '')).replace('Z', '+00:00'))
            require(expires.tzinfo is not None, '到期时间必须包含时区')
            expiry = expires.astimezone(datetime.timezone.utc).isoformat(timespec='microseconds')
        except ValueError:
            raise WorkflowError('到期时间格式无效')
        key = str(data.get('request_key', ''))
        require(re.fullmatch(r'[A-Za-z0-9_-]{16,80}', key), '请提供有效的防重复提交标识')
        if hasattr(conn, '_raw'):
            conn.execute('SELECT pg_advisory_xact_lock(?)',
                         (int.from_bytes(hashlib.sha256(('grant:' + key).encode()).digest()[:8], 'big', signed=True),))
        previous = conn.execute('SELECT * FROM inv_temporary_stock_grants WHERE request_key=?', (key,)).fetchone()
        if previous:
            require((previous['user_id'], previous['warehouse_id'], previous['expires_at'], previous['reason'])
                    == (uid, wid, expiry, reason), '该提交标识已用于其他授权内容', 409)
            return {'id': previous['id'], 'replayed': True}
        require(expiry > utc_now(), '到期时间必须晚于当前时间')
        aid, name = actor()
        cur = conn.execute('''INSERT INTO inv_temporary_stock_grants
            (user_id,warehouse_id,expires_at,reason,created_by,created_name,created_at,request_key)
            VALUES (?,?,?,?,?,?,?,?)''', (uid, wid, expiry, reason, aid, name, now(), key))
        event(conn, None, 'temporary_stock_granted', {'grant_id': cur.lastrowid, 'user_id': uid,
              'warehouse_id': wid, 'expires_at': expiry, 'reason': reason})
        return {'id': cur.lastrowid}
    grants = [dict(r) for r in conn.execute('''SELECT g.*,u.name AS user_name,u.username,w.name AS warehouse_name
        FROM inv_temporary_stock_grants g JOIN users u ON u.id=g.user_id
        JOIN warehouses w ON w.id=g.warehouse_id ORDER BY g.id DESC LIMIT 500''').fetchall()]
    for grant in grants:
        grant['active'] = not grant['revoked_at'] and grant['expires_at'] > utc_now()
    return {'grants': grants,
            'users': [dict(r) for r in conn.execute('SELECT id,name,username FROM users ORDER BY id').fetchall()],
            'warehouses': [dict(r) for r in conn.execute('''SELECT w.id,w.name FROM warehouses w
                LEFT JOIN oms_warehouse_integrations wi ON wi.warehouse_id=w.id
                WHERE w.is_active=1 AND COALESCE(wi.inventory_authority,'local')='local' ORDER BY w.name''').fetchall()]}


@workflow_bp.route('/api/inv/operations/temporary-access/<int:gid>/revoke', methods=['POST'])
@login_required
@transact
def revoke_temporary_access(conn, gid):
    require(current_user.username == 'admin', '仅内置管理员可撤销临时库存权限', 403)
    grant = conn.execute('SELECT * FROM inv_temporary_stock_grants WHERE id=?' + lock_suffix(conn), (gid,)).fetchone()
    require(grant is not None, '授权不存在', 404)
    if grant['revoked_at']:
        return {'id': gid, 'replayed': True}
    reason = str((request.get_json(silent=True) or {}).get('reason', '')).strip()
    require(0 < len(reason) <= 2000, '请填写撤销原因（不超过2000字）')
    conn.execute('UPDATE inv_temporary_stock_grants SET revoked_at=?,revoked_by=?,revoke_reason=? WHERE id=?',
                 (now(), current_user.id, reason, gid))
    event(conn, None, 'temporary_stock_revoked', {'grant_id': gid, 'user_id': grant['user_id'],
          'warehouse_id': grant['warehouse_id'], 'reason': reason})
    return {'id': gid}
