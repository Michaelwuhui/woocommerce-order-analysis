"""Auto-confirm terminal carrier outcomes on top of the 「待确认结局」 queue.

When enabled (settings.auto_confirm_delivered_enabled), every COD order sitting
in the 待确认结局 queue whose carrier_status is 'delivered' (🟢物流已签收) is
automatically confirmed — the unattended equivalent of a human clicking 已签收
(or the "批量确认所有「物流已签收」" button). For each such order it:
  1. sets the local delivery_confirmed flag (so it leaves the queue), and
  2. pushes WooCommerce status -> 'completed' (this fires WC's 'completed'
     customer email, same as a manual 已签收), adding an order note.

Delivery and return automation have independent, default-OFF switches.  The
delivery path only touches carrier-confirmed deliveries.  The return path only
touches ``carrier_status='returned'`` orders and mirrors the human ``拒收``
action: it marks the order undelivered and uses the configured return freight
policy (falling back to the order shipping fee). attention / in_transit / unknown / problem-return outcomes are
never auto-resolved.

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
from return_shipping_loss import load_policy, quote_loss, required_loss

from oid_utils import woo_post_id  # raw WC post id for REST write-back

ENABLE_KEY = 'auto_confirm_delivered_enabled'
RETURNED_ENABLE_KEY = 'auto_confirm_returned_enabled'
DEFAULT_PENDING_OUTCOME_DAYS = 7
COUNTRY_PENDING_OUTCOME_DAYS = {
    'PL': 1,
    'AU': 14,
}
# Activation timestamp, stamped (via SQL datetime('now'), so it matches DB
# timestamps) every time a switch is turned on. It is kept for audit/visibility
# and guards against a raw setting change bypassing the admin endpoint. Once
# activated, automation follows the queue state itself, including a safe backlog.
SINCE_KEY = 'auto_confirm_delivered_since'
RETURNED_SINCE_KEY = 'auto_confirm_returned_since'

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
_RETURNED_NOTE_PREFIX = "系统自动确认「未送达/退回」（物流明确退回 / carrier returned）"


def is_enabled(conn):
    """Master switch (settings.auto_confirm_delivered_enabled). Default OFF."""
    row = conn.execute(
        "SELECT value FROM settings WHERE key = ?", (ENABLE_KEY,)
    ).fetchone()
    if not row:
        return False
    return str(row[0]).strip().lower() in ('1', 'true', 'yes', 'on')


def is_returned_enabled(conn):
    """Return auto-confirm switch. Independent from delivered automation."""
    row = conn.execute(
        "SELECT value FROM settings WHERE key = ?", (RETURNED_ENABLE_KEY,)
    ).fetchone()
    if not row:
        return False
    return str(row[0]).strip().lower() in ('1', 'true', 'yes', 'on')


def get_since(conn):
    """Activation/audit timestamp for delivered automation; None if never set."""
    row = conn.execute("SELECT value FROM settings WHERE key = ?", (SINCE_KEY,)).fetchone()
    return row[0] if row and row[0] else None


def get_returned_since(conn):
    """Audit timestamp written whenever automatic return confirmation is enabled."""
    row = conn.execute(
        "SELECT value FROM settings WHERE key = ?", (RETURNED_SINCE_KEY,)
    ).fetchone()
    return row[0] if row and row[0] else None


def find_confirmable_orders(conn, since, order_id=None, include_completed=False):
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
        WHERE o.status IN (""" + ("'completed'," if include_completed else "") + """'on-hold', 'shipped', 'partial-shipped')
          AND o.payment_method = 'cod'
          AND COALESCE(o.is_undelivered, 0) = 0
          AND COALESCE(o.is_problem_return, 0) = 0
          AND COALESCE(o.delivery_confirmed, 0) = 0
          AND o.carrier_status = 'delivered'
          AND o.carrier_status_at IS NOT NULL
          AND (NOT EXISTS (
                 SELECT 1 FROM oms_order_fulfillment_state ofs WHERE ofs.order_id=o.id
               ) OR EXISTS (
                 SELECT 1 FROM oms_order_fulfillment_state ofs
                 WHERE ofs.order_id=o.id AND ofs.aggregate_status='delivered'
               ))
          AND """ + ready_sql + (" AND o.id=?" if order_id is not None else "") + """
        ORDER BY o.date_created ASC
        """, (str(order_id),) if order_id is not None else ()
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
          AND (NOT EXISTS (
                 SELECT 1 FROM oms_order_fulfillment_state ofs WHERE ofs.order_id=o.id
               ) OR EXISTS (
                 SELECT 1 FROM oms_order_fulfillment_state ofs
                 WHERE ofs.order_id=o.id AND ofs.aggregate_status='delivered'
               ))
          AND """ + ready_sql + """
        """
    ).fetchone()[0]


def _source_scope(allowed_sources):
    """Return an optional source SQL suffix and its bound parameters."""
    if allowed_sources is None:
        return '', []
    sources = list(allowed_sources)
    if not sources:
        return ' AND 1=0', []
    placeholders = ','.join('?' for _ in sources)
    return f' AND o.source IN ({placeholders})', sources


def _returned_safety_sql():
    """Only auto-resolve an unambiguous whole-order return.

    Legacy orders must have at most one unique tracking number. Managed OMS
    orders are eligible only when every active fulfillment has shipments and
    every one of those shipments is returned/cancelled, with at least one real
    return. A split parcel whose first tracking number returned therefore stays
    in the queue for human review instead of incorrectly failing the whole order.
    """
    return """
      AND (
        (
          NOT EXISTS (
            SELECT 1 FROM oms_order_fulfillment_state ofs
            WHERE ofs.order_id = o.id
          )
          AND (
            SELECT COUNT(DISTINCT NULLIF(trim(slg.tracking_number), ''))
            FROM shipping_logs slg WHERE slg.order_id = o.id
          ) <= 1
        )
        OR
        (
          EXISTS (
            SELECT 1 FROM oms_order_fulfillment_state ofs
            WHERE ofs.order_id=o.id AND ofs.aggregate_status='returned'
              AND COALESCE(ofs.manual_review,0)=0 AND COALESCE(ofs.has_shortage,0)=0
          )
          AND NOT EXISTS (SELECT 1 FROM oms_fulfillments f WHERE f.order_id=o.id)
          AND EXISTS (
            SELECT 1 FROM oms_domain_events ev WHERE ev.aggregate_type='order'
              AND ev.aggregate_id=o.id AND ev.event_type='legacy_carrier_outcome_reconciled'
              AND ev.to_status='returned' AND ev.actor_type='system'
          )
        )
        OR
        (
          EXISTS (
            SELECT 1
            FROM oms_order_fulfillment_state ofs
            JOIN oms_fulfillments f
              ON f.order_id = ofs.order_id AND f.revision = ofs.revision
            JOIN oms_shipments osh ON osh.fulfillment_id = f.id
            WHERE ofs.order_id = o.id
              AND f.status != 'superseded'
              AND osh.status = 'returned'
          )
          AND NOT EXISTS (
            SELECT 1
            FROM oms_order_fulfillment_state ofs
            JOIN oms_fulfillments f
              ON f.order_id = ofs.order_id AND f.revision = ofs.revision
            WHERE ofs.order_id = o.id
              AND f.status != 'superseded'
              AND (
                NOT EXISTS (
                  SELECT 1 FROM oms_shipments osh
                  WHERE osh.fulfillment_id = f.id
                )
                OR EXISTS (
                  SELECT 1 FROM oms_shipments osh
                  WHERE osh.fulfillment_id = f.id
                    AND osh.status NOT IN ('returned', 'cancelled')
                )
              )
          )
        )
      )
    """


def _returned_candidate_sql(select_clause, allowed_sources=None):
    source_sql, params = _source_scope(allowed_sources)
    sql = f"""
        SELECT {select_clause}
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
          AND o.carrier_status = 'returned'
          AND o.carrier_status_at IS NOT NULL
          {_returned_safety_sql()}
          AND {_ready_sql()}
          {source_sql}
    """
    return sql, params


def find_returnable_orders(conn, since, allowed_sources=None):
    """Unambiguous carrier-returned orders currently in 待确认结局.

    ``since`` is an activation/audit guard, matching the existing delivered
    switch. Once the switch has an activation timestamp, both the current safe
    backlog and future returned outcomes are eligible.
    """
    if not since:
        return []
    return list_returnable_orders(conn, allowed_sources)


def list_returnable_orders(conn, allowed_sources=None):
    """Return the current safe whole-order return candidates.

    Unlike :func:`find_returnable_orders`, this operator-facing helper does not
    depend on the automation switch having been activated.  It powers the
    explicit ``批量确认所有物流退回`` action while sharing the exact same status,
    age and split-parcel safety rules as the background worker.
    """
    sql, params = _returned_candidate_sql(
        "o.id, o.number, o.source, o.status, o.shipping_total, o.currency",
        allowed_sources,
    )
    return conn.execute(sql + " ORDER BY o.date_created ASC", params).fetchall()


def returnable_stats(conn, allowed_sources=None):
    """Count safe return candidates using the same policy as confirmation."""
    candidates = list_returnable_orders(conn, allowed_sources)
    policy = load_policy(conn)
    quotes = [quote_loss(conn, order, policy) for order in candidates]
    return {
        'candidates': sum(q['amount'] is not None for q in quotes),
        'shipping_loss': round(sum(q['amount'] for q in quotes if q['amount'] is not None), 2),
    }


def count_returnable(conn, since, allowed_sources=None):
    """Cheap UI count; disabled until the switch has an activation timestamp."""
    if not since:
        return 0
    return returnable_stats(conn, allowed_sources)['candidates']


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


def _mark_returned_local(
    conn,
    order,
    now,
    *,
    actor_name='系统自动确认',
    actor_user_id=None,
    batch=False,
    policy=None,
):
    """Atomically mirror the manual 拒收 action for one carrier return."""
    loss = required_loss(conn, order, policy)
    currency = str(order['currency'] or '').strip()
    actor_name = str(actor_name or '系统自动确认').strip()
    if batch:
        short_note = (
            f"订单被 {actor_name} 批量确认为「未送达/物流退回」"
            "（物流明确退回 / carrier returned）"
            f"，运费损失 {loss:.2f}"
        )
    else:
        short_note = f"{_RETURNED_NOTE_PREFIX}，运费损失 {loss:.2f}"
    if currency:
        short_note += f" {currency}"
    cursor = conn.execute(
        """
        UPDATE orders
           SET is_undelivered = 1,
               shipping_loss_amount = ?,
               undelivered_at = ?,
               undelivered_by = ?,
               undelivered_note = ?
         WHERE id = ?
           AND status IN ('on-hold', 'shipped', 'partial-shipped')
           AND payment_method = 'cod'
           AND COALESCE(is_undelivered, 0) = 0
           AND COALESCE(is_problem_return, 0) = 0
           AND COALESCE(delivery_confirmed, 0) = 0
           AND carrier_status = 'returned'
        """,
        (loss, now, actor_user_id, short_note, order['id']),
    )
    changed = int(cursor.rowcount or 0) == 1
    if not changed:
        return False, 0.0
    conn.execute(
        """
        INSERT INTO order_notes
            (order_id, note, date_created, customer_note, author, added_by_user)
        VALUES (?, ?, ?, 0, ?, 1)
        """,
        (order['id'], short_note, now, actor_name),
    )
    return True, loss


def confirm_returned_batch(
    conn,
    *,
    actor_name,
    actor_user_id,
    allowed_sources=None,
    dry_run=False,
):
    """Confirm every currently safe carrier return for an operator.

    This is deliberately independent of the automatic-return switch.  The
    candidate query and conditional update make the operation idempotent and
    safe to run while the background worker is active.  No WooCommerce status
    or customer notification is changed.
    """
    candidates = list_returnable_orders(conn, allowed_sources)
    summary = {
        'checked': len(candidates),
        'marked': 0,
        'skipped': 0,
        'shipping_loss': 0.0,
        'shipping_loss_by_currency': {},
        'dry_run': dry_run,
    }

    policy = load_policy(conn)
    for order in candidates:
        quote = quote_loss(conn, order, policy)
        if quote['amount'] is None:
            summary['skipped'] += 1
            continue
        loss = quote['amount']
        if not dry_run:
            changed, loss = _mark_returned_local(
                conn,
                order,
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                actor_name=actor_name,
                actor_user_id=actor_user_id,
                batch=True,
                policy=policy,
            )
            if not changed:
                summary['skipped'] += 1
                continue

        currency = str(order['currency'] or 'N/A').strip() or 'N/A'
        summary['marked'] += 1
        summary['shipping_loss'] += loss
        summary['shipping_loss_by_currency'][currency] = (
            summary['shipping_loss_by_currency'].get(currency, 0.0) + loss
        )

    summary['shipping_loss'] = round(summary['shipping_loss'], 2)
    summary['shipping_loss_by_currency'] = {
        currency: round(amount, 2)
        for currency, amount in sorted(summary['shipping_loss_by_currency'].items())
    }
    if not dry_run:
        conn.commit()
    return summary


def enforce_returned(conn, progress=None, dry_run=False, actor='auto'):
    """Auto-mark unambiguous carrier returns as undelivered.

    This is local and idempotent: the conditional update owns the transition,
    while ``is_undelivered=1`` removes the order from future candidate scans.
    It intentionally does not alter WooCommerce status or generate a customer
    email. Generated reconciliation statements remain snapshots and are not
    silently regenerated by this automation.
    """
    def _log(message):
        if progress:
            progress(message)

    since = get_returned_since(conn)
    if not since:
        if not dry_run:
            conn.execute(
                "INSERT OR REPLACE INTO settings (key, value) VALUES (?, datetime('now'))",
                (RETURNED_SINCE_KEY,),
            )
            conn.commit()
        _log("[auto-return] 无生效起点，已设为 now，本次不处理")
        return {
            'checked': 0, 'marked': 0, 'shipping_loss': 0.0,
            'skipped': 0, 'capped': 0, 'dry_run': dry_run, 'no_since': True,
        }

    candidates = find_returnable_orders(conn, since)
    summary = {
        'checked': len(candidates), 'marked': 0, 'shipping_loss': 0.0,
        'skipped': 0, 'capped': 0, 'dry_run': dry_run, 'since': since,
    }
    if len(candidates) > MAX_PER_RUN:
        summary['capped'] = len(candidates) - MAX_PER_RUN
        candidates = candidates[:MAX_PER_RUN]
        _log(
            f"[auto-return] 候选 {summary['checked']} 单 > 单次上限 {MAX_PER_RUN}，"
            f"本次处理 {MAX_PER_RUN} 单，剩余 {summary['capped']} 单下次继续。"
        )

    policy = load_policy(conn)
    for order in candidates:
        quote = quote_loss(conn, order, policy)
        if quote['amount'] is None:
            summary['skipped'] += 1
            _log(f"[auto-return] #{order['number']} 需人工处理：{quote['message']}")
            continue
        if dry_run:
            summary['marked'] += 1
            summary['shipping_loss'] += quote['amount']
            continue
        changed, loss = _mark_returned_local(
            conn, order, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), policy=policy
        )
        if not changed:
            conn.rollback()
            summary['skipped'] += 1
            continue
        conn.commit()
        summary['marked'] += 1
        summary['shipping_loss'] += loss
        _log(
            f"[auto-return] #{order['number']} 已标记未送达，"
            f"运费损失 {loss:.2f} {order['currency'] or ''}".rstrip()
        )

    summary['shipping_loss'] = round(summary['shipping_loss'], 2)
    return summary
