"""Auto-confirm carrier-delivered COD orders — automation on top of the
「待确认结局」queue.

When enabled (settings.auto_confirm_delivered_enabled), every COD order sitting
in the 待确认结局 queue whose carrier_status is 'delivered' (🟢物流已签收) is
automatically confirmed — the unattended equivalent of a human clicking 已签收
(or the "批量确认所有「物流已签收」" button). For each such order it:
  1. sets the local delivery_confirmed flag (so it leaves the queue), and
  2. pushes WooCommerce status -> 'completed' (this fires WC's 'completed'
     customer email, same as a manual 已签收), adding an order note.

ONLY carrier-confirmed deliveries are touched. returned / in_transit / unknown /
problem-return / undelivered / already-confirmed orders are NEVER auto-confirmed.
Default OFF — nothing happens until the operator turns the switch on.

Ordering vs the manual confirm: a manual confirm sets the local flag FIRST
(best-effort WC after), because a human is present to reconcile. Here, unattended,
we push WC FIRST and only set the local flag on success — so a WC failure leaves
the order UNCONFIRMED and it is retried next run, instead of silently drifting
out of the queue while the store still shows on-hold. WC PUT completed is
idempotent, so a retry after a lost-response is harmless.

Import-safe for the hourly cron (auto_sync.py): does NOT import the Flask app —
it only needs a DB connection (sqlite3.Row factory) handed in.
"""
import requests
from datetime import datetime

from oid_utils import woo_post_id  # raw WC post id for REST write-back

ENABLE_KEY = 'auto_confirm_delivered_enabled'
DEFAULT_PENDING_OUTCOME_DAYS = 7
COUNTRY_PENDING_OUTCOME_DAYS = {
    'PL': 1,
    'AU': 14,
}
# Forward-only "effective start" timestamp, stamped (via SQL datetime('now'), so
# it matches DB timestamps) every time the switch is turned on. It is kept for
# audit/visibility, but auto-confirm now follows the queue state itself: if a COD
# order is already old enough to be in 待确认结局 and the carrier says delivered,
# it can be confirmed. This matches the operator workflow after the PL gate moved
# from 14 days to next-day review.
SINCE_KEY = 'auto_confirm_delivered_since'

# Per-run cap. When first switched on there can be a large backlog of delivered
# orders; draining at most this many per hourly run spreads the WooCommerce
# writes / customer emails over a few runs instead of one giant burst. The
# remainder is picked up on the next run (the candidate set shrinks as orders
# get confirmed, so it always converges).
MAX_PER_RUN = 300

_API_HEADERS = {
    "User-Agent": "WooCommerce API Client-Python/3.0.0",
    "Content-Type": "application/json",
    "Accept": "application/json",
}

_NOTE = "系统自动确认「已签收」（物流已签收 / carrier delivered）。"


def is_enabled(conn):
    """Master switch (settings.auto_confirm_delivered_enabled). Default OFF."""
    row = conn.execute(
        "SELECT value FROM settings WHERE key = ?", (ENABLE_KEY,)
    ).fetchone()
    if not row:
        return False
    return str(row[0]).strip().lower() in ('1', 'true', 'yes', 'on')


def get_since(conn):
    """Forward-only effective-start (settings.auto_confirm_delivered_since).
    None if never set. Only deliveries with carrier_status_at >= this are
    auto-confirmed."""
    row = conn.execute("SELECT value FROM settings WHERE key = ?", (SINCE_KEY,)).fetchone()
    return row[0] if row and row[0] else None


def find_confirmable_orders(conn, since):
    """COD orders in the 待确认结局 queue whose carrier status is delivered.

    Same candidate definition as /api/shipping/pending-outcome, narrowed to
    carrier_status='delivered'. The `since` argument is retained for API
    compatibility and audit context, but does not exclude already-actionable
    queue rows.
    """
    ready_sql = _ready_sql()
    return conn.execute(
        """
        SELECT o.id, o.number, o.source, o.status
        FROM orders o
        LEFT JOIN sites s ON o.source = s.url
        LEFT JOIN shipping_logs sl ON sl.id = (
            SELECT id FROM shipping_logs WHERE order_id = o.id ORDER BY id DESC LIMIT 1
        )
        WHERE o.status IN ('on-hold', 'shipped', 'partial-shipped')
          AND o.payment_method = 'cod'
          AND COALESCE(o.is_undelivered, 0) = 0
          AND COALESCE(o.is_problem_return, 0) = 0
          AND COALESCE(o.delivery_confirmed, 0) = 0
          AND o.carrier_status = 'delivered'
          AND o.carrier_status_at IS NOT NULL
          AND """ + ready_sql + """
        ORDER BY o.date_created ASC
        """
    ).fetchall()


def count_confirmable(conn, since):
    """Cheap COUNT for the UI (how many would be auto-confirmed right now)."""
    if not since:
        return 0
    ready_sql = _ready_sql()
    return conn.execute(
        """
        SELECT COUNT(*)
        FROM orders o
        LEFT JOIN sites s ON o.source = s.url
        LEFT JOIN shipping_logs sl ON sl.id = (
            SELECT id FROM shipping_logs WHERE order_id = o.id ORDER BY id DESC LIMIT 1
        )
        WHERE o.status IN ('on-hold', 'shipped', 'partial-shipped')
          AND o.payment_method = 'cod'
          AND COALESCE(o.is_undelivered, 0) = 0
          AND COALESCE(o.is_problem_return, 0) = 0
          AND COALESCE(o.delivery_confirmed, 0) = 0
          AND o.carrier_status = 'delivered'
          AND o.carrier_status_at IS NOT NULL
          AND """ + ready_sql + """
        """
    ).fetchone()[0]


def _pending_age_case_sql():
    """Country-specific age gate matching the shipping pending-outcome queue."""
    return (
        "CASE COALESCE(s.country, '') "
        "WHEN 'PL' THEN 1 "
        "WHEN 'AU' THEN 14 "
        f"ELSE {DEFAULT_PENDING_OUTCOME_DAYS} END"
    )


def _ready_sql():
    """Country-specific age gate matching the shipping pending-outcome queue."""
    age_case = _pending_age_case_sql()
    shipped = "datetime(replace(substr(COALESCE(sl.shipped_at, o.date_modified, o.date_created), 1, 19), 'T', ' '))"
    return f"{shipped} <= datetime('now', '-' || {age_case} || ' days')"


def _load_sites(conn):
    return {r['url']: r for r in conn.execute("SELECT * FROM sites").fetchall()}


def _complete_remote(site, oid):
    """PUT status=completed on WooCommerce. Returns (ok: bool, detail: str).
    Mirrors app.confirm_order_delivery's WC push (fires the completed email)."""
    wid = woo_post_id(oid)
    url = f"{site['url']}/wp-json/wc/v3/orders/{wid}"
    try:
        resp = requests.put(
            url, json={'status': 'completed'},
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
    return True, "ok"


def _confirm_local(conn, oid, now):
    """Local confirm: set the queue-clearing flag + an attributable note.
    delivery_confirmed_by stays NULL (no human); the note carries attribution."""
    conn.execute(
        "UPDATE orders SET delivery_confirmed = 1, delivery_confirmed_at = ?, "
        "delivery_confirmed_by = NULL WHERE id = ?", (now, oid))
    conn.execute(
        "INSERT INTO order_notes (order_id, note, date_created, customer_note, author, added_by_user) "
        "VALUES (?, ?, ?, 0, ?, 1)", (oid, _NOTE, now, '系统自动确认'))


def enforce(conn, progress=None, dry_run=False, actor='auto'):
    """Confirm every carrier-delivered COD order in the queue. Returns a summary.

    Idempotent: a confirmed order (delivery_confirmed=1) drops out of the
    candidate set, so this is safe to run every hour. A WC push failure leaves
    the order untouched (no local flag set) and is retried next run, so local
    and remote never drift apart.
    """
    def _log(m):
        if progress:
            progress(m)

    since = get_since(conn)
    if not since:
        # Enabled but no effective-start (e.g. the setting was flipped directly
        # in the DB, bypassing the UI which always stamps it). Establish the
        # window NOW so the backlog is never retroactively confirmed; this run
        # confirms nothing.
        if not dry_run:
            conn.execute(
                "INSERT OR REPLACE INTO settings (key, value) VALUES (?, datetime('now'))",
                (SINCE_KEY,))
            conn.commit()
        _log("[auto-confirm] 无生效起点，已设为 now，本次不追溯历史单")
        return {'checked': 0, 'confirmed': 0, 'synced': 0, 'local_only': 0,
                'errors': 0, 'capped': 0, 'dry_run': dry_run, 'no_since': True}

    candidates = find_confirmable_orders(conn, since)
    summary = {'checked': len(candidates), 'confirmed': 0, 'synced': 0,
               'local_only': 0, 'errors': 0, 'capped': 0, 'dry_run': dry_run, 'since': since}
    if not candidates:
        return summary

    if len(candidates) > MAX_PER_RUN:
        summary['capped'] = len(candidates) - MAX_PER_RUN
        candidates = candidates[:MAX_PER_RUN]
        _log(f"[auto-confirm] 候选 {summary['checked']} 单 > 单次上限 {MAX_PER_RUN}，"
             f"本次处理 {MAX_PER_RUN} 单，剩余 {summary['capped']} 单下次继续。")

    sites = _load_sites(conn)
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    for o in candidates:
        oid, num = o['id'], o['number']

        if dry_run:
            summary['confirmed'] += 1
            continue

        site = sites.get(o['source'])
        no_write = (not site or not site['consumer_key'] or not site['consumer_secret']
                    or ('api_write_status' in site.keys() and site['api_write_status'] == 'error'))

        # Already completed at the store, or no write access: confirm locally
        # only — no WC call, no extra email. Mirrors the manual confirm's
        # "站点已是已完成 / 仅本地标记签收" branches.
        if o['status'] == 'completed' or no_write:
            _confirm_local(conn, oid, now)
            conn.commit()
            summary['confirmed'] += 1
            summary['local_only'] += 1
            if no_write and o['status'] != 'completed':
                _log(f"[auto-confirm] #{num} 无写权限，仅本地标记签收")
            continue

        ok, detail = _complete_remote(site, oid)
        if ok:
            conn.execute("UPDATE orders SET status='completed' WHERE id=?", (oid,))
            _confirm_local(conn, oid, now)
            conn.commit()
            summary['confirmed'] += 1
            summary['synced'] += 1
        else:
            summary['errors'] += 1
            _log(f"[auto-confirm] 失败 #{num} @ {o['source']}: {detail}（保持未确认，下次重试）")

    return summary
