"""
inv_notify.py — 模块 8:通知中心(补货 / 滞销 / 临期 提醒;站内 + 邮件)。

三类提醒:
  - restock(补货):某仓某 SKU 可用(on_hand-reserved)<= 补货点
    (inv_skus.reorder_point,未设则用全局默认设置 inv_default_reorder_point)。
  - near_expiry / expired(临期/过期):复用 inv_batches 的到期判定。
  - slow_moving(滞销):有库存(on_hand>阈值)但近 N 天无销售出库(sale_out)。

投递:
  - 站内:写 inv_notifications,导航栏小铃铛轮询未读数 + 下拉查看。
  - 邮件:可选,SMTP 配置存 settings 表(inv_smtp_*),未配置则静默跳过(不报错)。

去重:同一问题(dedup_key)若已有 unread 通知,则不重复创建。
"""

import json
import datetime
from flask import Blueprint, render_template, request, jsonify
from flask_login import login_required

from inv_common import get_conn, inv_view_required, inv_manage_required, warehouse_scope_clause
import inv_batches

inv_notify_bp = Blueprint('inv_notify', __name__)

DEFAULT_REORDER_POINT = 10
DEFAULT_NEAR_DAYS = 90
DEFAULT_SLOW_DAYS = 60
DEFAULT_SLOW_MIN_STOCK = 1


# ───────────────────────── 设置读取 ─────────────────────────

def _setting(conn, key, default=None):
    try:
        r = conn.execute('SELECT value FROM settings WHERE key=?', (key,)).fetchone()
        return r['value'] if r and r['value'] not in (None, '') else default
    except Exception:
        return default


def _int_setting(conn, key, default):
    v = _setting(conn, key, None)
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


# ───────────────────────── 创建(去重) ─────────────────────────

def _create(conn, ntype, severity, title, body, dedup_key,
            sku_id=None, warehouse_id=None, ref_type=None, ref_id=None):
    """创建通知;若已有同 dedup_key 的 unread 则跳过。返回 id 或 None(跳过)。"""
    if dedup_key:
        ex = conn.execute("SELECT id FROM inv_notifications WHERE dedup_key=? AND status='unread'",
                          (dedup_key,)).fetchone()
        if ex:
            return None
    cur = conn.execute('''INSERT INTO inv_notifications
        (ntype, severity, title, body, sku_id, warehouse_id, ref_type, ref_id, dedup_key)
        VALUES (?,?,?,?,?,?,?,?,?)''',
        (ntype, severity, title, body, sku_id, warehouse_id, ref_type, ref_id, dedup_key))
    return cur.lastrowid


# ───────────────────────── 扫描器 ─────────────────────────

def scan_restock(conn):
    """补货:可用 <= 补货点。"""
    default_rp = _int_setting(conn, 'inv_default_reorder_point', DEFAULT_REORDER_POINT)
    rows = conn.execute('''
        SELECT st.warehouse_id, st.sku_id, st.on_hand, st.reserved,
               (st.on_hand - st.reserved) AS available,
               k.sku_code, k.name AS sku_name, COALESCE(k.reorder_point,0) AS rp,
               w.name AS warehouse_name
        FROM inv_stock st JOIN inv_skus k ON k.id=st.sku_id JOIN warehouses w ON w.id=st.warehouse_id
        WHERE k.is_active=1
    ''').fetchall()
    n = 0
    for r in rows:
        threshold = r['rp'] if r['rp'] and r['rp'] > 0 else default_rp
        if r['available'] <= threshold:
            sev = 'danger' if r['available'] <= 0 else 'warning'
            sold = (r['available'] <= 0)
            title = (f"售罄补货:{r['sku_code']} @ {r['warehouse_name']}" if sold
                     else f"低库存补货:{r['sku_code']} @ {r['warehouse_name']}")
            body = f"{r['sku_name']} 可用 {r['available']}(现存 {r['on_hand']} - 预留 {r['reserved']}),补货点 {threshold}"
            if _create(conn, 'restock', sev, title, body,
                       f"restock:{r['warehouse_id']}:{r['sku_id']}",
                       sku_id=r['sku_id'], warehouse_id=r['warehouse_id']):
                n += 1
    return n


def scan_expiry(conn):
    """临期/过期:复用批次到期判定。"""
    near_days = _int_setting(conn, 'inv_near_days', DEFAULT_NEAR_DAYS)
    rows = conn.execute('''SELECT b.id, b.warehouse_id, b.sku_id, b.batch_no, b.expiry_date, b.qty_remaining,
                                  k.sku_code, w.name AS warehouse_name
                           FROM inv_batches b JOIN inv_skus k ON k.id=b.sku_id JOIN warehouses w ON w.id=b.warehouse_id
                           WHERE b.qty_remaining>0 AND b.expiry_date IS NOT NULL''').fetchall()
    n = 0
    for r in rows:
        st = inv_batches.batch_status(r['expiry_date'], near_days)
        if st == 'expired':
            if _create(conn, 'expired', 'danger',
                       f"已过期:{r['sku_code']} 批次 {r['batch_no'] or r['id']} @ {r['warehouse_name']}",
                       f"到期 {r['expiry_date']},剩余 {r['qty_remaining']} 件需处理",
                       f"expired:{r['id']}", sku_id=r['sku_id'], warehouse_id=r['warehouse_id'],
                       ref_type='batch', ref_id=str(r['id'])):
                n += 1
        elif st == 'near':
            if _create(conn, 'near_expiry', 'warning',
                       f"临期:{r['sku_code']} 批次 {r['batch_no'] or r['id']} @ {r['warehouse_name']}",
                       f"到期 {r['expiry_date']},剩余 {r['qty_remaining']} 件,请优先发货",
                       f"near:{r['id']}", sku_id=r['sku_id'], warehouse_id=r['warehouse_id'],
                       ref_type='batch', ref_id=str(r['id'])):
                n += 1
    return n


def scan_slow_moving(conn):
    """滞销:有库存但近 N 天无销售出库。"""
    days = _int_setting(conn, 'inv_slow_days', DEFAULT_SLOW_DAYS)
    min_stock = _int_setting(conn, 'inv_slow_min_stock', DEFAULT_SLOW_MIN_STOCK)
    cutoff = (datetime.date.today() - datetime.timedelta(days=days)).isoformat()
    rows = conn.execute('''
        SELECT st.warehouse_id, st.sku_id, st.on_hand, k.sku_code, k.name AS sku_name, w.name AS warehouse_name,
               (SELECT COUNT(*) FROM inv_movements m
                WHERE m.warehouse_id=st.warehouse_id AND m.sku_id=st.sku_id
                  AND m.movement_type='sale_out' AND m.ts >= ?) AS recent_sales
        FROM inv_stock st JOIN inv_skus k ON k.id=st.sku_id JOIN warehouses w ON w.id=st.warehouse_id
        WHERE k.is_active=1 AND st.on_hand > ?
    ''', (cutoff, min_stock)).fetchall()
    n = 0
    for r in rows:
        if r['recent_sales'] == 0:
            if _create(conn, 'slow_moving', 'info',
                       f"滞销提醒:{r['sku_code']} @ {r['warehouse_name']}",
                       f"{r['sku_name']} 现存 {r['on_hand']},近 {days} 天无销售出库",
                       f"slow:{r['warehouse_id']}:{r['sku_id']}",
                       sku_id=r['sku_id'], warehouse_id=r['warehouse_id']):
                n += 1
    return n


def generate_all(conn, commit=True):
    counts = {
        'restock': scan_restock(conn),
        'expiry': scan_expiry(conn),
        'slow_moving': scan_slow_moving(conn),
    }
    if commit:
        conn.commit()
    return counts


# ───────────────────────── 邮件(可选) ─────────────────────────

def send_pending_emails(conn, commit=True):
    """把未发邮件的未读通知发出去(若配置了 SMTP)。返回发送数。"""
    host = _setting(conn, 'inv_smtp_host')
    to_addr = _setting(conn, 'inv_smtp_to')
    if not host or not to_addr:
        return 0  # 未配置 → 静默跳过
    import smtplib
    from email.mime.text import MIMEText
    port = _int_setting(conn, 'inv_smtp_port', 465)
    user = _setting(conn, 'inv_smtp_user')
    pwd = _setting(conn, 'inv_smtp_pass')
    from_addr = _setting(conn, 'inv_smtp_from') or user or to_addr
    use_ssl = (_setting(conn, 'inv_smtp_ssl', '1') == '1')

    rows = conn.execute("SELECT * FROM inv_notifications WHERE status='unread' AND emailed=0 ORDER BY id").fetchall()
    if not rows:
        return 0
    body_lines = [f"[{r['severity'].upper()}] {r['title']}\n  {r['body'] or ''}" for r in rows]
    msg = MIMEText('库存提醒汇总:\n\n' + '\n\n'.join(body_lines), 'plain', 'utf-8')
    msg['Subject'] = f"库存提醒 {len(rows)} 条 - 苍赋管理系统"
    msg['From'] = from_addr
    msg['To'] = to_addr
    try:
        if use_ssl:
            s = smtplib.SMTP_SSL(host, port, timeout=30)
        else:
            s = smtplib.SMTP(host, port, timeout=30); s.starttls()
        if user and pwd:
            s.login(user, pwd)
        s.sendmail(from_addr, [a.strip() for a in to_addr.split(',')], msg.as_string())
        s.quit()
    except Exception as e:
        # 邮件失败不影响站内通知
        return -1
    ids = [r['id'] for r in rows]
    conn.execute(f"UPDATE inv_notifications SET emailed=1 WHERE id IN ({','.join('?'*len(ids))})", ids)
    if commit:
        conn.commit()
    return len(ids)


# ───────────────────────────── API ─────────────────────────────

@inv_notify_bp.route('/inventory/notifications')
@login_required
@inv_view_required
def notifications_page():
    return render_template('inv_notifications.html')


@inv_notify_bp.route('/api/inv/notifications', methods=['GET'])
@login_required
@inv_view_required
def list_notifications():
    status = request.args.get('status', 'unread')
    ntype = request.args.get('type')
    try:
        limit = min(500, int(request.args.get('limit') or 100))
    except ValueError:
        limit = 100
    conn = get_conn()
    sql = 'SELECT * FROM inv_notifications WHERE 1=1'
    params = []
    if status and status != 'all':
        sql += ' AND status=?'; params.append(status)
    if ntype:
        sql += ' AND ntype=?'; params.append(ntype)
    # 合伙人只看自己仓的通知(warehouse_id 为空的全局通知所有人可见)
    sc, sp = warehouse_scope_clause('warehouse_id')
    if sc:
        sql += ' AND (warehouse_id IS NULL' + sc.replace(' AND ', ' OR ', 1) + ')'
        params += sp
    sql += ' ORDER BY id DESC LIMIT ?'; params.append(limit)
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@inv_notify_bp.route('/api/inv/notifications/unread-count', methods=['GET'])
@login_required
@inv_view_required
def unread_count():
    conn = get_conn()
    try:
        sc, sp = warehouse_scope_clause('warehouse_id')
        scope = ''
        if sc:
            scope = ' AND (warehouse_id IS NULL' + sc.replace(' AND ', ' OR ', 1) + ')'
        n = conn.execute(f"SELECT COUNT(*) FROM inv_notifications WHERE status='unread'{scope}", sp).fetchone()[0]
        by = {r['ntype']: r['n'] for r in conn.execute(
            f"SELECT ntype, COUNT(*) AS n FROM inv_notifications WHERE status='unread'{scope} GROUP BY ntype", sp).fetchall()}
        return jsonify({'unread': n, 'by_type': by})
    finally:
        conn.close()


@inv_notify_bp.route('/api/inv/notifications/<int:nid>/<action>', methods=['POST'])
@login_required
@inv_view_required
def update_notification(nid, action):
    if action not in ('read', 'dismiss', 'unread'):
        return jsonify({'error': '未知操作'}), 400
    status = {'read': 'read', 'dismiss': 'dismissed', 'unread': 'unread'}[action]
    conn = get_conn()
    try:
        conn.execute('UPDATE inv_notifications SET status=?, updated_at=CURRENT_TIMESTAMP WHERE id=?', (status, nid))
        conn.commit()
        return jsonify({'success': True})
    finally:
        conn.close()


@inv_notify_bp.route('/api/inv/notifications/mark-all-read', methods=['POST'])
@login_required
@inv_view_required
def mark_all_read():
    conn = get_conn()
    try:
        conn.execute("UPDATE inv_notifications SET status='read', updated_at=CURRENT_TIMESTAMP WHERE status='unread'")
        conn.commit()
        return jsonify({'success': True})
    finally:
        conn.close()


@inv_notify_bp.route('/api/inv/notifications/scan', methods=['POST'])
@login_required
@inv_manage_required
def api_scan():
    """触发扫描生成通知 + 尝试发邮件。"""
    conn = get_conn()
    try:
        counts = generate_all(conn)
        emailed = send_pending_emails(conn)
        return jsonify({'generated': counts, 'emailed': emailed})
    finally:
        conn.close()
