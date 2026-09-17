"""Bounded, restart-safe automatic delivery confirmation for Celery workers."""
from datetime import datetime
import json
from urllib.parse import urljoin, urlsplit

import requests

import auto_confirm
from external_operations import begin_operation, transition_operation
from oid_utils import woo_post_id


PAYLOAD = {"target_status": "completed"}
TIMEOUT = (5, 12)
WRITABLE_STATUSES = {"on-hold", "shipped", "partial-shipped", "processing"}


def eligible(conn, order_id, include_completed=False):
    if not auto_confirm.is_enabled(conn) or not auto_confirm.get_since(conn):
        return None
    rows = auto_confirm.find_confirmable_orders(conn, auto_confirm.get_since(conn), order_id, include_completed)
    return dict(rows[0]) if rows else None


def dispatch_candidates(conn, limit=48):
    """Cooldown failed/ambiguous operations without starving later orders."""
    if not auto_confirm.is_enabled(conn) or not auto_confirm.get_since(conn):
        return []
    blocked = {
        str(row["order_id"])
        for row in conn.execute("""
            SELECT order_id FROM external_operations
            WHERE operation_type='confirm_delivery'
              AND (status='cancelled' OR
                   (status IN ('pending','failed','reconciliation_required','external_success')
                    AND updated_at > CURRENT_TIMESTAMP - INTERVAL '10 minutes'))
        """).fetchall()
    }
    return [str(row["id"]) for row in auto_confirm.find_confirmable_orders(
        conn, auto_confirm.get_since(conn)
    ) if str(row["id"]) not in blocked][:limit]


class ReadError(ValueError):
    """A credential-free reason safe to retain in the operation ledger."""


def _remote_order(session, url, auth, expected_id):
    origin = urlsplit(url)
    for _ in range(3):
        response = session.get(url, auth=auth, headers=auto_confirm._API_HEADERS,
                               timeout=TIMEOUT, allow_redirects=False)
        if response.status_code not in {301, 302, 303, 307, 308}:
            break
        target = urljoin(url, response.headers.get("Location") or "")
        parsed = urlsplit(target)
        # Only the store's HTTPS bare/www canonical alias may receive its key.
        hosts = {origin.hostname, "www." + origin.hostname.removeprefix("www.")}
        hosts.add(origin.hostname.removeprefix("www."))
        if (parsed.scheme != "https" or parsed.hostname not in hosts
                or parsed.port != origin.port or parsed.username or parsed.password
                or parsed.path != origin.path or parsed.query or parsed.fragment or target == url):
            raise ReadError("read_unsafe_redirect")
        url = target
    else:
        raise ReadError("read_redirect_loop")
    if response.status_code != 200:
        raise ReadError(f"read_http_{response.status_code}")
    try:
        data = response.json()
        if not isinstance(data, dict) or str(data.get("id")) != str(expected_id):
            raise ReadError("read_wrong_order")
        if not isinstance(data.get("status"), str):
            raise ReadError("read_missing_status")
        data["_verified_url"] = url
        return data
    except (TypeError, AttributeError, requests.exceptions.JSONDecodeError):
        raise ReadError("read_invalid_json") from None


def _transition(conn, op, status, **kwargs):
    transition_operation(conn, op["operation_id"], status, **kwargs)
    op["status"] = status
    conn.commit()


def _record_phase(conn, op, phase):
    conn.execute("""UPDATE external_operations
        SET external_evidence=?::jsonb,updated_at=CURRENT_TIMESTAMP
        WHERE operation_id=? AND status='pending'""",
        (json.dumps({"automatic": True, "phase": phase}), op["operation_id"]))
    conn.commit()


def _interrupted_before_write(conn, op):
    row = conn.execute("SELECT external_evidence FROM external_operations WHERE operation_id=?",
                       (op["operation_id"],)).fetchone()
    evidence = row[0] if row else None
    if isinstance(evidence, str):
        evidence = json.loads(evidence)
    return isinstance(evidence, dict) and evidence.get("phase") == "preflight"


def _finish_local(conn, order_id, op, remote, wrote):
    evidence = {"order_id": remote["id"], "status": remote["status"],
                "verified_by": "GET", "automatic": True, "write_performed": wrote}
    if op["status"] in {"pending", "reconciliation_required"}:
        _transition(conn, op, "external_success", external_reference=str(remote["id"]), evidence=evidence)
    # Re-check after I/O so a concurrent manual exception is never cleared.
    current = eligible(conn, order_id, include_completed=True)
    if not current:
        if op["status"] == "external_success":
            _transition(conn, op, "reconciliation_required", error="eligibility_changed_after_readback", evidence=evidence)
        return {"order_id": order_id, "result": "eligibility_changed"}
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    changed = conn.execute("""
        UPDATE orders SET status='completed',delivery_confirmed=1,
               delivery_confirmed_at=?,delivery_confirmed_by=NULL
        WHERE id=? AND COALESCE(delivery_confirmed,0)=0
          AND COALESCE(is_undelivered,0)=0 AND COALESCE(is_problem_return,0)=0
          AND carrier_status='delivered'
    """, (now, order_id)).rowcount
    if changed:
        conn.execute("""INSERT INTO order_notes
            (order_id,note,date_created,customer_note,author,added_by_user)
            VALUES (?,?,?,0,?,1)""",
            (order_id, auto_confirm._NOTE, now, "系统自动确认"))
        conn.execute("""UPDATE oms_order_fulfillment_state
            SET completion_sync_status='synced',updated_at=CURRENT_TIMESTAMP
            WHERE order_id=? AND aggregate_status='delivered'""", (order_id,))
    if op["status"] == "external_success":
        transition_operation(conn, op["operation_id"], "local_committed", evidence=evidence)
    conn.commit()
    return {"order_id": order_id, "result": "confirmed" if changed else "already_confirmed", "wrote": wrote}


def process_delivered_order(conn, order_id, session=None):
    """One order, at most one PUT; uncertain outcomes only receive GET retries.

    The shared confirm_delivery ledger also excludes concurrent manual writes.
    A committed pending operation survives worker death before/after the PUT.
    """
    own_session = session is None
    session = session or requests.Session()
    lock_key = "auto-confirm-delivered:" + str(order_id)
    locked = False
    try:
        locked = bool(conn.execute("SELECT pg_try_advisory_lock(hashtextextended(?,0))", (lock_key,)).fetchone()[0])
        conn.commit()
        if not locked:
            return {"order_id": order_id, "result": "busy"}
        order = eligible(conn, order_id)
        if not order:
            return {"order_id": order_id, "result": "ineligible"}
        # Managed fulfillments have their own completion outbox. Do not race it.
        ofs = conn.execute("SELECT completion_sync_status FROM oms_order_fulfillment_state WHERE order_id=?", (order_id,)).fetchone()
        if ofs and ofs[0] in {"pending", "running"}:
            return {"order_id": order_id, "result": "fulfillment_completion_pending"}
        site = conn.execute("SELECT id,url,consumer_key,consumer_secret FROM sites WHERE url=?", (order["source"],)).fetchone()
        if not site or not site["consumer_key"] or not site["consumer_secret"]:
            return {"order_id": order_id, "result": "missing_credentials"}
        op = begin_operation(conn, operation_type="confirm_delivery", order_id=order_id,
                             site_id=site["id"], request_payload=PAYLOAD,
                             created_by="celery-beat:auto-delivered")
        if op['should_execute']:
            _record_phase(conn, op, 'preflight')
        else:
            conn.commit()
        if op["status"] == "cancelled":
            return {"order_id": order_id, "result": "cancelled_operation"}
        # Another executor may still be running. Only reconcile a stale claim.
        if not op["should_execute"] and op["status"] == "pending":
            age = conn.execute("SELECT EXTRACT(EPOCH FROM (CURRENT_TIMESTAMP-updated_at)) FROM external_operations WHERE operation_id=?", (op["operation_id"],)).fetchone()[0]
            if float(age) < 120:
                return {"order_id": order_id, "result": "operation_in_progress"}
            if _interrupted_before_write(conn, op):
                _transition(conn, op, 'failed', error='worker_interrupted_before_write',
                            evidence={'automatic': True, 'phase': 'preflight', 'retry_safe': True})
                return {"order_id": order_id, "result": "retry_ready"}
        conn.commit()
        wid = woo_post_id(order_id)
        url = site["url"].rstrip("/") + f"/wp-json/wc/v3/orders/{wid}"
        auth = (site["consumer_key"], site["consumer_secret"])
        try:
            remote = _remote_order(session, url, auth, wid)
        except (requests.RequestException, ValueError) as exc:
            target = "failed" if op["should_execute"] else "reconciliation_required"
            if op["status"] in {"pending", "external_success", "reconciliation_required"}:
                detail = str(exc) if isinstance(exc, ReadError) else type(exc).__name__
                _transition(conn, op, target, error="preflight:" + detail)
            return {"order_id": order_id, "result": "read_failed"}
        url = remote["_verified_url"]
        if remote["status"] == "completed":
            return _finish_local(conn, order_id, op, remote, False)
        if not op["should_execute"]:
            if op["status"] in {"pending", "external_success", "reconciliation_required"}:
                _transition(conn, op, "reconciliation_required", error="remote_not_completed_after_uncertain_operation")
            return {"order_id": order_id, "result": "reconciliation_required"}
        if remote["status"] not in WRITABLE_STATUSES or not eligible(conn, order_id):
            _transition(conn, op, "cancelled", error="remote_status_or_local_eligibility_changed",
                        evidence={"status": remote["status"], "verified_by": "GET"})
            return {"order_id": order_id, "result": "status_conflict", "remote_status": remote["status"]}
        conn.commit()
        # Persist the boundary before calling WooCommerce. An interrupted
        # read-only preflight can be retried; a started write must reconcile.
        _record_phase(conn, op, 'write_started')
        try:
            response = session.put(url, auth=auth, headers=auto_confirm._API_HEADERS,
                                   json={"status": "completed"}, timeout=TIMEOUT, allow_redirects=False)
            # Authentication/validation/rate-limit rejections have no mutation.
            if response.status_code in {400, 401, 403, 404, 405, 422, 429}:
                _transition(conn, op, "failed", error=f"write_http_{response.status_code}")
                return {"order_id": order_id, "result": "write_rejected", "http": response.status_code}
        except requests.RequestException:
            pass  # Read back even on timeout: the store may have committed.
        try:
            remote = _remote_order(session, url, auth, wid)
        except (requests.RequestException, ValueError):
            _transition(conn, op, "reconciliation_required", error="post_write_readback_failed")
            return {"order_id": order_id, "result": "reconciliation_required"}
        if remote["status"] != "completed":
            _transition(conn, op, "reconciliation_required", error="post_write_status_not_completed",
                        evidence={"status": remote["status"], "verified_by": "GET"})
            return {"order_id": order_id, "result": "reconciliation_required"}
        return _finish_local(conn, order_id, op, remote, True)
    finally:
        conn.rollback()
        if locked:
            conn.execute("SELECT pg_advisory_unlock(hashtextextended(?,0))", (lock_key,))
            conn.commit()
        if own_session:
            session.close()
