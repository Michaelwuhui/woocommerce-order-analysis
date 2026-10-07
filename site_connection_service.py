"""Retire a Woo connection without removing its order/synchronization identity.

Only the live API configuration is removed.  Historical rows keep referencing
the original site ID and URL; no site or synchronization row is deleted.
"""
from __future__ import annotations

import json
from urllib.parse import urlsplit


ARCHIVED_STATUS = "archived"


class SiteConnectionError(ValueError):
    def __init__(self, message, *, code="SITE_BUSY", status=409, details=None):
        super().__init__(message)
        self.code = code
        self.status = status
        self.details = details or {}


def _postgres(conn):
    return hasattr(conn, "_raw")


def _value(row, key, default=None):
    try:
        return row[key]
    except (KeyError, IndexError, TypeError):
        return default


def _exists(conn, table):
    if _postgres(conn):
        return bool(conn.execute("SELECT to_regclass(?)", (table,)).fetchone()[0])
    return bool(conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)
    ).fetchone())


def _columns(conn, table):
    if _postgres(conn):
        return {row[0] for row in conn.execute(
            "SELECT column_name FROM information_schema.columns "
            "WHERE table_schema=current_schema() AND table_name=?", (table,)
        ).fetchall()}
    return {row[1] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()}


def archived_site_ids(conn):
    # Old SQLite fixtures/installations predate api_status. PostgreSQL uses the
    # migrated schema and must fail visibly if its required column is absent.
    if not _postgres(conn) and "api_status" not in _columns(conn, "sites"):
        return set()
    return {int(row[0]) for row in conn.execute(
        "SELECT id FROM sites WHERE api_status=?", (ARCHIVED_STATUS,)
    ).fetchall()}


def is_site_archived(site_or_conn, site_id=None):
    if site_id is None:
        return _value(site_or_conn, "api_status") == ARCHIVED_STATUS
    return int(site_id) in archived_site_ids(site_or_conn)


def filter_active_sites(conn, rows):
    archived = archived_site_ids(conn)
    return [row for row in rows if not is_site_archived(row)
            and _value(row, "id", _value(row, "site_id")) not in archived]


def lock_active_sites(conn, site_ids):
    """Keep the site config stable until the caller's task INSERT commits.

    The archive transaction takes FOR UPDATE on the same row. IDs are locked in
    order so a two-site clone cannot deadlock another clone. Missing legacy
    SQLite sites tables are supported only for queue-only test installations.
    """
    ids = sorted({int(value) for value in site_ids})
    if not ids:
        return []
    if not _postgres(conn) and not _exists(conn, "sites"):
        return []
    has_marker = _postgres(conn) or "api_status" in _columns(conn, "sites")
    columns = "id,api_status" if has_marker else "id"
    placeholders = ",".join("?" for _ in ids)
    lock = " FOR KEY SHARE" if _postgres(conn) else ""
    rows = conn.execute(
        f"SELECT {columns} FROM sites WHERE id IN ({placeholders}) ORDER BY id" + lock,
        tuple(ids),
    ).fetchall()
    if {int(row["id"]) for row in rows} != set(ids):
        raise SiteConnectionError("站点不存在", code="SITE_NOT_FOUND", status=404)
    if any(is_site_archived(row) for row in rows):
        raise SiteConnectionError("站点连接已移除，请先重新配置连接", code="SITE_ARCHIVED")
    return rows


def _json(value):
    if isinstance(value, dict):
        return value
    try:
        parsed = json.loads(value or "{}")
        return parsed if isinstance(parsed, dict) else {}
    except (TypeError, ValueError):
        return {}


def _endpoint(url):
    parts = urlsplit(str(url or "").strip())
    return parts.netloc.lower() + parts.path.rstrip("/")


def _scope_has_site(scope, site_id):
    ids = [value for key in ("site_ids", "target_site_ids")
           for value in (scope.get(key) or [])]
    return str(scope.get("source_site_id")) == str(site_id) or any(str(value) == str(site_id) for value in ids)


def _one(conn, sql, params):
    return bool(conn.execute(sql, params).fetchone())


def _busy_reasons(conn, site):
    site_id = int(site["id"])
    reasons = []
    if _exists(conn, "sync_site_progress"):
        if _exists(conn, "sync_runs") and _one(conn, """
            SELECT 1 FROM sync_site_progress p JOIN sync_runs r ON r.run_id=p.run_id
            WHERE p.site_id=? AND (p.status NOT IN ('success','error','auth_error','cancelled')
              OR r.status IN ('queued','running','recovering','cancelling')) LIMIT 1
        """, (site_id,)):
            reasons.append("order_sync")
        elif _one(conn, "SELECT 1 FROM sync_site_progress WHERE site_id=? "
                  "AND status NOT IN ('success','error','auth_error','cancelled') LIMIT 1", (site_id,)):
            reasons.append("order_sync")
    if _exists(conn, "sync_page_dispatches") and _one(conn,
        "SELECT 1 FROM sync_page_dispatches WHERE site_id=? "
        "AND status NOT IN ('completed','cancelled','error','auth_error') LIMIT 1", (site_id,)):
        reasons.append("sync_pages")
    if _exists(conn, "sync_page_receipts") and "post_commit_status" in _columns(conn, "sync_page_receipts"):
        if _one(conn, "SELECT 1 FROM sync_page_receipts WHERE site_id=? "
                "AND COALESCE(post_commit_status,'pending') NOT IN ('completed','skipped') LIMIT 1", (site_id,)):
            reasons.append("sync_post_commit")
    if _exists(conn, "sync_task_outbox"):
        payloads = conn.execute("SELECT payload FROM sync_task_outbox "
            "WHERE status IN ('pending','publishing','error')").fetchall()
        if any(str(_json(row[0]).get("site_id")) == str(site_id) for row in payloads):
            reasons.append("sync_outbox")
    if _exists(conn, "external_operations") and _one(conn,
        "SELECT 1 FROM external_operations WHERE site_id=? "
        "AND status NOT IN ('local_committed','notified','failed','cancelled') LIMIT 1", (site_id,)):
        reasons.append("external_operation")
    if _exists(conn, "orders") and _exists(conn, "oms_integration_jobs"):
        if _one(conn, """SELECT 1 FROM oms_integration_jobs j JOIN orders o
            ON j.aggregate_type='order' AND j.aggregate_id=o.id
            WHERE o.source=? AND j.job_type='COMPLETE_WOOCOMMERCE_ORDER'
              AND j.status IN ('pending','retry','running') LIMIT 1""", (site["url"],)):
            reasons.append("order_fulfillment_job")
        if _exists(conn, "oms_shipments") and _exists(conn, "oms_fulfillments"):
            if _one(conn, """SELECT 1 FROM oms_integration_jobs j
                JOIN oms_shipments s ON j.aggregate_type='shipment' AND j.aggregate_id=s.id
                JOIN oms_fulfillments f ON f.id=s.fulfillment_id JOIN orders o ON o.id=f.order_id
                WHERE o.source=? AND j.job_type='SYNC_SHIPMENT_TO_WOOCOMMERCE'
                  AND j.status IN ('pending','retry','running') LIMIT 1""", (site["url"],)):
                reasons.append("shipment_website_job")
            if _exists(conn, "oms_shipment_notifications") and _one(conn, """
                SELECT 1 FROM oms_shipment_notifications n JOIN oms_shipments s ON s.id=n.shipment_id
                JOIN oms_fulfillments f ON f.id=s.fulfillment_id JOIN orders o ON o.id=f.order_id
                WHERE o.source=? AND n.status IN ('pending','retry','running','sending')
                  AND s.status NOT IN ('cancelled','label_pending','label_ready') LIMIT 1""", (site["url"],)):
                reasons.append("shipment_notification")
        if _exists(conn, "order_notification_jobs") and _one(conn, """
            SELECT 1 FROM oms_integration_jobs j JOIN order_notification_jobs n
              ON j.aggregate_type='order_notification' AND j.aggregate_id=n.id
            JOIN orders o ON o.id=n.order_id WHERE o.source=? AND j.job_type='ORDER_NOTIFICATION'
              AND j.status IN ('pending','retry','running') LIMIT 1""", (site["url"],)):
            reasons.append("order_notification")
    if _exists(conn, "orders") and _exists(conn, "oms_fulfillments") and _one(conn,
        "SELECT 1 FROM oms_fulfillments f JOIN orders o ON o.id=f.order_id WHERE o.source=? "
        "AND f.status IN ('submitting','submission_unknown','cancel_pending','cancel_rejected') LIMIT 1", (site["url"],)):
        reasons.append("warehouse_result_unknown")
    if _exists(conn, "product_clone_jobs") and _one(conn,
        "SELECT 1 FROM product_clone_jobs WHERE (source_site_id=? OR target_site_id=?) "
        "AND status IN ('queued','running') LIMIT 1", (site_id, site_id)):
        reasons.append("product_clone")
    if _exists(conn, "inv_push_locks") and _one(conn,
        "SELECT 1 FROM inv_push_locks WHERE site_id=? LIMIT 1", (site_id,)):
        reasons.append("inventory_push_lock")
    if _exists(conn, "inv_push_runs") and _one(conn,
        "SELECT 1 FROM inv_push_runs WHERE site_id=? "
        "AND status NOT IN ('success','partial','error','skipped','cancelled','failed','completed','succeeded') LIMIT 1", (site_id,)):
        reasons.append("inventory_push")
    if _exists(conn, "stock_sync_jobs") and _exists(conn, "stock_sync_plan_items"):
        if _one(conn, """SELECT 1 FROM stock_sync_jobs j
            JOIN stock_sync_plan_items i ON i.plan_id=j.plan_id WHERE i.site_id=?
              AND j.status IN ('queued','running','cancel_requested','requires_review') LIMIT 1""", (site_id,)):
            reasons.append("stock_sync")
        if _exists(conn, "stock_sync_plans"):
            active = conn.execute("""SELECT p.request_json FROM stock_sync_jobs j
                JOIN stock_sync_plans p ON p.id=j.plan_id
                WHERE j.status IN ('queued','running','cancel_requested','requires_review')""").fetchall()
            if any(str(_json(row[0]).get("source_site_id")) == str(site_id) for row in active):
                reasons.append("stock_reference")
    if _exists(conn, "stock_sync_work"):
        if _exists(conn, "stock_sync_catalog_snapshots"):
            snapshots = conn.execute("""SELECT s.site_id,s.scope_json FROM stock_sync_work w
                JOIN stock_sync_catalog_snapshots s ON s.id=w.object_id
                WHERE w.status IN ('queued','running')""").fetchall()
            if any(str(row["site_id"]) == str(site_id) or _scope_has_site(_json(row["scope_json"]), site_id)
                   for row in snapshots):
                reasons.append("stock_snapshot")
        if _exists(conn, "stock_sync_plans"):
            plans = conn.execute("""SELECT p.request_json FROM stock_sync_work w
                JOIN stock_sync_plans p ON p.id=w.object_id WHERE w.status IN ('queued','running')""").fetchall()
            if any(_scope_has_site(_json(row[0]), site_id) for row in plans):
                reasons.append("stock_preview")
    if _exists(conn, "stock_sync_resource_leases"):
        endpoint = _endpoint(site["url"])
        leases = conn.execute("SELECT resource_key FROM stock_sync_resource_leases").fetchall()
        if any(str(row[0]) == "endpoint:" + endpoint
               or str(row[0]).startswith(endpoint + "/products/") for row in leases):
            reasons.append("product_write_lease")
        if _exists(conn, "stock_sync_plan_items") and _one(conn,
            "SELECT 1 FROM stock_sync_resource_leases l JOIN stock_sync_plan_items i "
            "ON i.resource_key=l.resource_key WHERE i.site_id=? LIMIT 1", (site_id,)):
            reasons.append("stock_write_lease")
    return sorted(set(reasons))


def _protect_master(conn, site):
    if not _exists(conn, "product_masters") or "product_master_id" not in _columns(conn, "sites"):
        return
    masters = conn.execute("SELECT id,url FROM product_masters").fetchall()
    ids = [row["id"] for row in masters if _endpoint(row["url"]) == _endpoint(site["url"])]
    for master_id in ids:
        if _one(conn, "SELECT 1 FROM sites WHERE product_master_id=? AND id<>? "
                "AND COALESCE(api_status,'')<>? LIMIT 1", (master_id, site["id"], ARCHIVED_STATUS)):
            raise SiteConnectionError("该连接仍是其他站点的商品 Master，请先解除相关配置",
                                      code="MASTER_IN_USE")


def _public_site(site):
    data = {key: _value(site, key) for key in
            ("id", "url", "manager", "country", "product_master_id", "api_status")}
    data["configured"] = bool(_value(site, "consumer_key") and _value(site, "consumer_secret"))
    return data


def remove_connection(conn, site_id, actor_id=None, actor_name=""):
    if isinstance(site_id, bool) or not str(site_id).isdigit() or int(site_id) <= 0:
        raise SiteConnectionError("站点 ID 无效", code="INVALID_SITE_ID", status=400)
    site_id = int(site_id)
    try:
        if not _postgres(conn) and not getattr(conn, "in_transaction", False):
            conn.execute("BEGIN IMMEDIATE")
        lock = " FOR UPDATE NOWAIT" if _postgres(conn) else ""
        site = conn.execute("SELECT * FROM sites WHERE id=?" + lock, (site_id,)).fetchone()
        if not site:
            raise SiteConnectionError("站点不存在", code="SITE_NOT_FOUND", status=404)
        if _postgres(conn):
            acquired = conn.execute("SELECT pg_try_advisory_xact_lock(hashtextextended(?,0))",
                                    (f"woo-sync-site:{site_id}",)).fetchone()[0]
            if not acquired:
                raise SiteConnectionError("该站点正在同步，请等待任务结束",
                                          details={"reasons": ["order_sync_lock"]})
        already_archived = is_site_archived(site)
        if not already_archived:
            _protect_master(conn, site)
            reasons = _busy_reasons(conn, site)
            if reasons:
                raise SiteConnectionError("该站点仍有进行中或待核对的任务，请先处理后再移除连接",
                                          details={"reasons": reasons})
        before = _public_site(site)
        conn.execute("UPDATE sites SET consumer_key='',consumer_secret='',"
                     "product_master_id=NULL,api_status=? WHERE id=?", (ARCHIVED_STATUS, site_id))
        if not already_archived and _exists(conn, "inv_site_sync_audit"):
            after = dict(before, configured=False, product_master_id=None, api_status=ARCHIVED_STATUS)
            conn.execute("INSERT INTO inv_site_sync_audit "
                         "(site_id,action,before_json,after_json,operator_id,operator_name) VALUES (?,?,?,?,?,?)",
                         (site_id, "connection_removed", json.dumps(before, ensure_ascii=False),
                          json.dumps(after, ensure_ascii=False), actor_id, str(actor_name or "")[:255]))
        conn.commit()
        return {"success": True, "id": site_id, "url": site["url"], "archived": True,
                "already_archived": already_archived}
    except Exception as exc:
        conn.rollback()
        if getattr(exc, "sqlstate", None) == "55P03" or getattr(exc, "pgcode", None) == "55P03":
            raise SiteConnectionError("该站点配置正在使用中，请稍后再试",
                                      details={"reasons": ["site_configuration_lock"]}) from None
        raise
