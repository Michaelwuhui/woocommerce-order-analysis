"""Authenticated, party-scoped API for the modular reconciliation workspace."""
import csv
import io
import secrets
from functools import wraps
from flask import Blueprint, jsonify, render_template, request, session, Response, abort
from flask_login import login_required, current_user
from inv_common import get_conn, can_manage_inventory
from reconciliation_core import *
import reconciliation_inventory as inventory
import reconciliation_ledger as ledger
import reconciliation_history as history

bp = Blueprint('reconciliation_v2', __name__)


@bp.get('/partner-reconciliation/history')
@login_required
def history_page():
    conn = get_conn()
    try:
        editor(conn)
    finally:
        conn.close()
    session.setdefault('rec_csrf', secrets.token_urlsafe(32))
    return render_template('reconciliation_history.html', rec_csrf=session['rec_csrf'])


def _history_listing(conn):
    return [{k: r[k] for k in ('id', 'name', 'created_at')} for r in rows(conn,
        "SELECT id,name,created_at FROM rec_objects WHERE kind='history_draft' ORDER BY created_at DESC,id DESC LIMIT 100")]


def scope(conn):
    if current_user.username == 'admin':
        return None
    if not current_user.can_view_reconciliation():
        abort(403)
    explicit = rows(conn, 'SELECT party_id FROM rec_user_parties WHERE user_id=?', (current_user.id,))
    if explicit:
        return {r['party_id'] for r in explicit}
    legacy = current_user.get_accessible_partner_ids()
    if legacy is None:
        return None
    return {p['id'] for p in objects(conn, 'party') if p['data'].get('legacy_partner_id') in legacy}


def permitted(conn, party):
    allowed = scope(conn)
    if allowed is not None and party not in allowed:
        abort(403)


def editor(conn):
    if scope(conn) is not None or not current_user.can_edit_reconciliation() or not current_user.can_view_costs():
        abort(403)


def api(write=False):
    def decorate(fn):
        @wraps(fn)
        @login_required
        def wrapped(*args, **kwargs):
            conn = get_conn()
            try:
                if not ready(conn):
                    return jsonify(error='模块未安装，请执行显式数据库迁移'), 503
                scope(conn)
                if write:
                    if not session.get('rec_csrf') or not secrets.compare_digest(request.headers.get('X-Rec-CSRF', ''), session['rec_csrf']):
                        abort(403)
                    editor(conn)
                result = fn(conn, *args, **kwargs)
                if write:
                    conn.commit()
                return result
            except (ReconciliationError, ValueError, KeyError, TypeError) as exc:
                conn.rollback()
                return jsonify(error=str(exc)), 400
            except Exception:
                conn.rollback()
                raise
            finally:
                conn.close()
        return wrapped
    return decorate


@bp.app_errorhandler(ReconciliationError)
def domain_error(exc):
    return jsonify(error=str(exc), code='reconciliation_validation'), 400


@bp.get('/partner-reconciliation/modular')
@login_required
def page():
    if not current_user.can_view_reconciliation():
        abort(403)
    session.setdefault('rec_csrf', secrets.token_urlsafe(32))
    return render_template('reconciliation_v2.html', rec_csrf=session['rec_csrf'])


def safe_statement(row, external):
    result = dict(row)
    snap = json.loads(result['snapshot'])
    if external:
        snap['entries'] = [{k: e.get(k) for k in ('id', 'order_id', 'site', 'category', 'amount', 'currency', 'occurred_at')} for e in snap['entries']]
        snap.pop('actor', None)
    result['snapshot'] = snap
    return result


@bp.get('/api/reconciliation-v2/state')
@api()
def state(conn):
    allowed = scope(conn)
    private = allowed is None and current_user.can_view_costs()
    parties = [p for p in objects(conn, 'party') if allowed is None or p['id'] in allowed]
    statements = [safe_statement(r, not private) for r in rows(conn, 'SELECT * FROM rec_statements ORDER BY created_at DESC,id') if allowed is None or r['party_id'] in allowed]
    cash = [dict(r) for r in rows(conn, 'SELECT * FROM rec_cash ORDER BY occurred_at DESC') if allowed is None or r['party_id'] in allowed]
    for r in cash:
        if not private:
            r.pop('data', None)
    result = {'parties': parties, 'statements': statements, 'cash': cash, 'balances': ledger.balances(conn, allowed),
              'can_edit': bool(private and current_user.can_edit_reconciliation()), 'can_inventory': bool(private and can_manage_inventory())}
    if private:
        result.update(profit_report=ledger.profit_report(conn), profiles=objects(conn,'profile'), pools=objects(conn, 'pool'), contracts=objects(conn, 'contract'), fees=objects(conn, 'fee'),
            batches=rows(conn, '''SELECT b.id,b.batch_no,b.warehouse_id,b.sku_id,b.qty_remaining,b.expiry_date,
                r.pool_id,r.unit_cost,r.currency,r.reserved FROM inv_batches b LEFT JOIN rec_batches r ON r.batch_id=b.id ORDER BY b.id DESC LIMIT 500'''),
            warehouses=rows(conn, 'SELECT id,name FROM warehouses ORDER BY id'),
            orders=[{**r,'config':json.loads(r['data'])} for r in rows(conn, '''SELECT r.*,o.number,o.source,o.status,COALESCE(f.aggregate_status,'unallocated') AS fulfillment_status FROM rec_orders r JOIN orders o ON o.id=r.order_id LEFT JOIN oms_order_fulfillment_state f ON f.order_id=r.order_id ORDER BY r.created_at DESC LIMIT 200''')],
            bindings=rows(conn, 'SELECT * FROM rec_user_parties'))
    return jsonify(result)


@bp.post('/api/reconciliation-v2/config')
@api(write=True)
def config(conn):
    data = request.get_json()
    ident = save_object(conn, data['kind'], data['name'], data['data'], current_user.id)
    return jsonify(id=ident)


@bp.post('/api/reconciliation-v2/inventory')
@api(write=True)
def inventory_action(conn):
    if not can_manage_inventory():
        abort(403)
    data = request.get_json()
    action = data['action']
    if action == 'bind':
        inventory.bind_batch(conn, integer(data['batch_id'], 1), data['pool_id'], data.get('unit_cost'), data['currency'], current_user.id)
    elif action == 'classify':
        inventory.register_unclassified_batch(conn, integer(data['warehouse_id'],1), integer(data['sku_id'],1), data['quantity'], data['batch_no'], data['pool_id'], data.get('unit_cost'), data['currency'], current_user.id, data.get('expiry'))
    elif action == 'move':
        inventory.move_batch(conn, integer(data['batch_id'],1), integer(data['warehouse_id'],1), data['quantity'], data['reference'], data['reason'], current_user.id)
    elif action == 'return':
        inventory.return_source(conn, data['source_id'], data['quantity'], data['salable'], data['reason'], data['reference'], current_user.id)
    elif action == 'transfer':
        inventory.transfer_ownership(conn, integer(data['batch_id'], 1), data['pool_id'], data['reason'], current_user.id)
    else:
        raise ReconciliationError('操作无效')
    return jsonify(ok=True)


@bp.post('/api/reconciliation-v2/order')
@api(write=True)
def order_action(conn):
    data = request.get_json(); order_id = data['order_id']; action = data['action']
    if action == 'configure':
        configure_order(conn, order_id, data['data'], current_user.id)
        result = {'ok': True}
    elif action in {'preview', 'plan'}:
        if not can_manage_inventory():
            abort(403)
        from fulfillment_service import plan_order
        result = plan_order(conn, order_id, actor={'id': current_user.id, 'name': current_user.username}, commit=False)
        result['sources'] = [{**r, 'snapshot': json.loads(r['snapshot'])} for r in rows(conn, 'SELECT * FROM rec_sources WHERE order_id=? AND (reserved>0 OR shipped>0)', (order_id,))]
        if action == 'preview':
            conn.rollback()
    elif action == 'correct':
        ledger.financial_correction(conn, order_id, data['values'], data['reference'], data['reason'], current_user.id)
        result = {'ok': True}
    elif action == 'recognize':
        result = ledger.recognize(conn, order_id, current_user.id)
    else:
        raise ReconciliationError('操作无效')
    return jsonify(ledger.serializable(result))


@bp.get('/api/reconciliation-v2/order/<path:order_id>')
@api()
def order_detail(conn, order_id):
    editor(conn)
    result = ledger.economics(conn, order_id)
    result['shipments'] = ledger.shipments(conn, order_id)
    result['source_stock'] = rows(conn, 'SELECT id,batch_id,quantity,reserved,shipped,returned FROM rec_sources WHERE order_id=?', (order_id,))
    return jsonify(ledger.serializable(result))


@bp.post('/api/reconciliation-v2/bill')
@api(write=True)
def bill(conn):
    data = request.get_json()
    ledger.record_bill(conn, data['shipment_id'], data['data'], current_user.id)
    return jsonify(ok=True)


@bp.post('/api/reconciliation-v2/statement')
@api(write=True)
def statement(conn):
    data = request.get_json()
    if data['action'] == 'create':
        ident = ledger.create_statement(conn, data['party_id'], data['currency'], data['start'], data['end'], data['bucket'], current_user.id, data.get('timezone','UTC'))
    else:
        ident = data['id']
        ledger.statement_transition(conn, ident, data['status'], current_user.id, data.get('reason', ''))
    return jsonify(id=ident)


@bp.post('/api/reconciliation-v2/cash')
@api(write=True)
def cash(conn):
    data = request.get_json()
    if data['action'] == 'record':
        ident = ledger.record_cash(conn, data['party_id'], data['currency'], data['amount'], data['occurred_at'], data['reference'], data['proof'], current_user.id)
    elif data['action'] == 'offset':
        ledger.authorized_offset(conn, data['positive_statement'], data['negative_statement'], data['amount'], data['reference'], data['reason'], current_user.id)
        ident = data['reference']
    elif data['action'] == 'allocate':
        if not data.get('request_key'):
            raise ReconciliationError('资金核销需要唯一请求编号')
        ledger.allocate_cash(conn, data['cash_id'], data['statement_id'], data['amount'], current_user.id, data.get('request_key'))
        ident = data['cash_id']
    else:
        raise ReconciliationError('操作无效')
    return jsonify(id=ident)


@bp.post('/api/reconciliation-v2/adjustment')
@api(write=True)
def adjust(conn):
    d = request.get_json()
    ledger.adjustment(conn, d['party_id'], d['amount'], d['currency'], d['category'], d['reference'], d['reason'], current_user.id)
    return jsonify(ok=True)


@bp.post('/api/reconciliation-v2/access')
@api(write=True)
def access(conn):
    if current_user.username != 'admin':
        abort(403)
    data = request.get_json()
    object_(conn, data['party_id'], 'party')
    if not conn.execute('SELECT id FROM users WHERE id=?', (data['user_id'],)).fetchone():
        raise ReconciliationError('用户不存在')
    conn.execute('INSERT INTO rec_user_parties(user_id,party_id) VALUES (?,?) ON CONFLICT DO NOTHING', (data['user_id'], data['party_id']))
    audit(conn, uid(), 'access_bound', None, data, current_user.id)
    return jsonify(ok=True)


@bp.get('/api/reconciliation-v2/export/<ident>')
@api()
def export(conn, ident):
    row = conn.execute('SELECT * FROM rec_statements WHERE id=?', (ident,)).fetchone()
    if not row:
        abort(404)
    permitted(conn, row['party_id'])
    # Deliberately small export: never expose embedded source/other-party snapshots.
    out = io.StringIO(); writer = csv.writer(out)
    writer.writerow(['订单', '站点', '类型', '金额', '币种', '发生时间'])
    for e in json.loads(row['snapshot'])['entries']:
        cells = [e.get(k) or '' for k in ('order_id','site','category','amount','currency','occurred_at')]
        writer.writerow(["'" + str(v) if index != 3 and str(v).startswith(('=', '+', '-', '@', '\t', '\r')) else v for index,v in enumerate(cells)])
    return Response('\ufeff' + out.getvalue(), content_type='text/csv; charset=utf-8', headers={'Content-Disposition': 'attachment; filename="reconciliation.csv"'})


@bp.get('/api/reconciliation-v2/history/state')
@api()
def history_state(conn):
    editor(conn)
    return jsonify(rules=objects(conn, 'history_rule'), drafts=_history_listing(conn),
                   parties=objects(conn,'party'), pools=objects(conn,'pool'), contracts=objects(conn,'contract'),
                   warehouses=rows(conn,'SELECT id,name FROM warehouses ORDER BY id'))


@bp.post('/api/reconciliation-v2/history/rule')
@api(write=True)
def history_rule(conn):
    data = request.get_json()
    return jsonify(id=history.save_rule(conn, data['name'], data['data'], current_user.id))


@bp.post('/api/reconciliation-v2/history/preview')
@api(write=True)
def history_preview(conn):
    data = request.get_json()
    return jsonify(history.preview(conn, data['rule_id'], data['month'], data['currency'], data.get('reship_decisions')))


@bp.post('/api/reconciliation-v2/history/draft')
@api(write=True)
def history_draft(conn):
    return jsonify(id=history.create_draft(conn, request.get_json(), current_user.id))


@bp.get('/api/reconciliation-v2/history/draft/<ident>')
@api()
def history_detail(conn, ident):
    editor(conn)
    return jsonify(object_(conn, ident, 'history_draft'))


@bp.get('/api/reconciliation-v2/history/export/<ident>')
@api()
def history_export(conn, ident):
    editor(conn)
    snap = object_(conn, ident, 'history_draft')['data']['snapshot']
    out=io.StringIO(); writer=csv.writer(out)
    writer.writerow(['订单ID','订单号','站点','出库日期','状态','原币','商品净收入','客户运费','收入合计','收入人民币','应付物流费','物流费人民币','管理费人民币','运费净收益','供货货值原币','供货货值人民币','已匹配货值人民币小计','商品毛利原币','团队贡献利润原币','成本状态','缺失成本商品','成本记录ID','汇率月份','对人民币汇率','其他待核实'])
    for r in snap['rows']:
        values=[r.get(k) for k in ('order_id','number','site','shipped_at','state','currency','goods_income','shipping_income','revenue','revenue_cny','freight','freight_cny','management_cny','shipping_net')]
        values += [r.get('supplier_goods_value'),r.get('supplier_goods_value_cny'),r.get('known_goods_value_cny'),
                   r.get('product_profit'),r.get('contribution_profit'),r.get('cost_status','旧快照未计算成本'),
                   '；'.join(r.get('cost_missing') or []),
                   ','.join(str(line['cost_id']) for line in r.get('cost_lines') or []),
                   (r['rate'] or {}).get('month'),(r['rate'] or {}).get('rate'),'；'.join(r['pending'])]
        # Prefix textual formula triggers; numeric negative amounts remain numeric.
        writer.writerow(["'"+str(v) if i in (0,1,2,3,4,5,19,20,21,22,24) and str(v or '').startswith(('=','+','-','@','\t','\r')) else ('待成本' if v is None and i in (14,15,17,18) else '待核实' if v is None else v) for i,v in enumerate(values)])
    return Response('\ufeff'+out.getvalue(),content_type='text/csv; charset=utf-8',headers={'Content-Disposition':'attachment; filename="historical-reconciliation.csv"'})
