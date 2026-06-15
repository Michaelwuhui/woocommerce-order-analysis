"""Customer blocklist enforcement — auto-cancel COD orders from blacklisted phones.

Layer 1 of the COD-refuser defense. A customer is blocked *by phone* (emails
rotate; phone + address are the stable identity — see app.py's identity merge).
When enabled, any of their COD orders that are still PRE-SHIPMENT get
auto-cancelled via the WooCommerce REST API and an order note is added.

Prepaid orders (payment_method != 'cod') are NEVER touched: a blocked refuser who
pays upfront is harmless and should still ship — that is the whole point of the
"force prepayment" strategy. Already-shipped orders are never touched either.

This module is import-safe for the hourly cron (auto_sync.py): it must NOT import
the Flask app. It only needs a DB connection (sqlite3.Row factory) handed in.
"""
import json
import requests
from datetime import datetime

from oid_utils import woo_post_id  # raw WC post id for REST write-back

GLOBAL_ENABLE_KEY = 'blocklist_auto_cancel_enabled'

# Circuit breaker: if a single run would cancel more than this many orders, abort
# the WHOLE run and log loudly instead of cancelling anything. Guards against a
# bug (e.g. a bad phone match) nuking the store. Real blocklist activity is a
# trickle; raise this only if you deliberately blocklist many active customers.
MAX_CANCELS_PER_RUN = 20

_API_HEADERS = {
    "User-Agent": "WooCommerce API Client-Python/3.0.0",
    "Content-Type": "application/json",
    "Accept": "application/json",
}


# --- phone normalization -----------------------------------------------------
# MUST stay in sync with app.py:_normalize_phone / _is_placeholder_phone.
# Duplicated (not imported) so the cron can enforce without importing the app.
def _is_placeholder_phone(digits):
    if not digits:
        return True
    if len(set(digits)) <= 1:
        return True
    if len(digits) >= 6:
        diffs = {int(digits[i + 1]) - int(digits[i]) for i in range(len(digits) - 1)}
        if diffs == {1} or diffs == {-1}:
            return True
    return False


def normalize_phone(p):
    """Digits only, folded to the last 9 to drop +48/+61/+971 country codes.
    Returns None for too-short or obvious-placeholder numbers."""
    digits = ''.join(c for c in (p or '') if c.isdigit())
    if len(digits) < 7:
        return None
    canonical = digits[-9:] if len(digits) >= 9 else digits
    if _is_placeholder_phone(canonical) or _is_placeholder_phone(digits):
        return None
    return canonical


# --- helpers -----------------------------------------------------------------
def is_globally_enabled(conn):
    """Global kill-switch (settings.blocklist_auto_cancel_enabled)."""
    row = conn.execute(
        "SELECT value FROM settings WHERE key = ?", (GLOBAL_ENABLE_KEY,)
    ).fetchone()
    if not row:
        return False
    return str(row[0]).strip().lower() in ('1', 'true', 'yes', 'on')


def get_blocked_phones(conn):
    """{normalized_phone: row} for entries with auto_cancel on."""
    rows = conn.execute(
        "SELECT * FROM blocked_customers WHERE auto_cancel = 1"
    ).fetchall()
    return {r['phone']: r for r in rows if r['phone']}


def find_cancellable_orders(conn):
    """COD + pre-shipment orders whose normalized phone is on the blocklist.

    Pre-shipment = pending/processing, plus on-hold ONLY on sites where on-hold
    is NOT the local "已发货" convention (cod_on_hold_is_shipped=0, i.e. AU/AE).
    A shipped order (PL on-hold, shipped/delivered/partial/completed) is never
    a candidate. Already-cancelled/failed/refunded orders are excluded by status.
    Returns a list of (order_row, block_row) tuples.
    """
    blocked = get_blocked_phones(conn)
    if not blocked:
        return []
    on_hold_shipped = {
        r['url'] for r in conn.execute(
            "SELECT url FROM sites WHERE cod_on_hold_is_shipped = 1"
        ).fetchall()
    }
    rows = conn.execute(
        """
        SELECT id, number, source, status, total, currency, billing
        FROM orders
        WHERE payment_method = 'cod'
          AND status IN ('pending', 'processing', 'on-hold')
        """
    ).fetchall()
    out = []
    for o in rows:
        if o['status'] == 'on-hold' and o['source'] in on_hold_shipped:
            continue  # PL convention: on-hold == already shipped — never cancel
        try:
            billing = json.loads(o['billing'] or '{}')
        except (TypeError, ValueError):
            billing = {}
        ph = normalize_phone(billing.get('phone'))
        if ph and ph in blocked:
            out.append((o, blocked[ph]))
    return out


def _load_sites(conn):
    return {r['url']: r for r in conn.execute("SELECT * FROM sites").fetchall()}


def _cancel_remote(site, oid, reason_note):
    """PUT status=cancelled (+ best-effort note) on WooCommerce.
    Returns (ok: bool, detail: str). Mirrors app.update_order_status."""
    wid = woo_post_id(oid)
    status_url = f"{site['url']}/wp-json/wc/v3/orders/{wid}"
    try:
        resp = requests.put(
            status_url, json={'status': 'cancelled'},
            auth=(site['consumer_key'], site['consumer_secret']),
            timeout=60, headers=_API_HEADERS,
        )
    except Exception as e:
        return False, f"请求异常: {e}"
    text = resp.text or ''
    if text.strip().startswith('<!') or text.strip().startswith('<html'):
        return False, "WP返回HTML(可能WAF/认证问题)"
    if resp.status_code not in (200, 201):
        return False, f"API {resp.status_code}: {text[:160]}"
    try:
        requests.post(
            f"{status_url}/notes",
            json={'note': reason_note, 'customer_note': False},
            auth=(site['consumer_key'], site['consumer_secret']),
            timeout=30, headers=_API_HEADERS,
        )
    except Exception:
        pass  # note is non-critical
    return True, "ok"


def _log_row(conn, order, phone, result, detail, actor, now):
    conn.execute(
        """INSERT INTO blocked_cancel_log
           (order_id, order_number, phone, source, total, currency,
            old_status, result, detail, actor, created_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
        (order['id'], order['number'], phone, order['source'], order['total'],
         order['currency'], order['status'], result, detail, actor, now),
    )


def enforce(conn, progress=None, dry_run=False, actor='auto'):
    """Cancel every cancellable blocklisted COD order. Returns a summary dict.

    Idempotent: already-cancelled orders drop out of the candidate set, so this
    is safe to run every sync. Network/site errors leave the order untouched and
    are retried next run.
    """
    def _log(m):
        if progress:
            progress(m)

    candidates = find_cancellable_orders(conn)
    summary = {'checked': len(candidates), 'cancelled': 0, 'errors': 0,
               'skipped': 0, 'dry_run': dry_run, 'aborted': False, 'details': []}
    if not candidates:
        return summary

    # Circuit breaker — fail closed (cancel nothing) and ask for human review.
    if len(candidates) > MAX_CANCELS_PER_RUN:
        _log(f"[blocklist] 安全阀触发：候选 {len(candidates)} 单 > 上限 "
             f"{MAX_CANCELS_PER_RUN}，本次不执行任何取消，请人工核查 blocked_customers。")
        summary['errors'] = len(candidates)
        summary['aborted'] = True
        return summary

    sites = _load_sites(conn)
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    for order, block in candidates:
        oid = order['id']
        site = sites.get(order['source'])
        phone = block['phone']
        reason = (block['reason'] or '').strip() or '惯性COD拒收'
        rec = {'order_id': oid, 'number': order['number'], 'source': order['source'],
               'total': order['total'], 'currency': order['currency'],
               'status': order['status']}

        if not site:
            summary['skipped'] += 1
            rec['result'], rec['detail'] = 'skipped', '站点配置不存在'
            summary['details'].append(rec)
            continue
        if ('api_write_status' in site.keys() and site['api_write_status'] == 'error'):
            summary['skipped'] += 1
            rec['result'], rec['detail'] = 'skipped', '站点无API写权限'
            summary['details'].append(rec)
            continue

        note = (f"系统自动取消：手机号 {phone} 在客户拦截名单中（{reason}）。"
                f"如需购买请改为预付/转账，确认到账后我们再安排发货。")

        if dry_run:
            summary['cancelled'] += 1
            rec['result'], rec['detail'] = 'would-cancel', note
            summary['details'].append(rec)
            _log(f"[blocklist][dry] 将取消 #{order['number']} "
                 f"({order['total']} {order['currency']}) @ {order['source']}")
            continue

        ok, detail = _cancel_remote(site, oid, note)
        if ok:
            conn.execute("UPDATE orders SET status='cancelled' WHERE id=?", (oid,))
            _log_row(conn, order, phone, 'cancelled', note, actor, now)
            conn.execute(
                """UPDATE blocked_customers
                   SET cancelled_count = COALESCE(cancelled_count,0)+1,
                       last_cancelled_at = ?
                   WHERE phone = ?""", (now, phone))
            conn.commit()
            summary['cancelled'] += 1
            rec['result'], rec['detail'] = 'cancelled', 'ok'
            _log(f"[blocklist] 已取消 #{order['number']} "
                 f"({order['total']} {order['currency']}) @ {order['source']}")
        else:
            _log_row(conn, order, phone, 'error', detail, actor, now)
            conn.commit()
            summary['errors'] += 1
            rec['result'], rec['detail'] = 'error', detail
            _log(f"[blocklist] 取消失败 #{order['number']}: {detail}")
        summary['details'].append(rec)

    return summary
