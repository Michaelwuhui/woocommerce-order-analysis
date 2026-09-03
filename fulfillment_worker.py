#!/usr/bin/env python3
"""Durable worker for multi-warehouse fulfillment integrations."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import random
import socket
import db_backend as sqlite3
import sys
import time
import traceback
import uuid
from datetime import datetime, timedelta, timezone
from typing import Any
from urllib.parse import unquote

import carrier_tracking
from fulfillment_common import get_conn, json_dump, json_load, redact, utcnow
from fulfillment_service import (
    DomainError,
    SHIPMENT_PROGRESS,
    add_tracking_event,
    build_poland_wms_payload,
    build_wms_payload,
    create_shipment,
    enqueue_job,
    mark_manual_review,
    mark_shipment_shipped,
    plan_order,
    recompute_order_status,
    record_event,
    transition_fulfillment,
)
from fulfillment_woocommerce import WooError, complete_order, sync_shipment
from hungary_wms import (
    WmsClient,
    WmsError,
    WmsResult,
    normalize_wms_fulfillment_status,
    normalize_wms_tracking_status,
)
from poland_wms import (
    PolandWmsClient,
    PolandWmsError,
    PolandWmsResult,
    normalize_poland_tracking_status,
)
from order_notification_service import (
    NotificationPermanent,
    NotificationRetry,
    enqueue_notification_failure_alert,
    process_notification_alert,
    process_notification_job,
)


WORKER_ID = f"{socket.gethostname()}:{os.getpid()}:{uuid.uuid4().hex[:8]}"
LEASE_SECONDS = 180


class RetryTask(RuntimeError):
    def __init__(self, message: str, *, code: str = "retry", delay_seconds: int | None = None):
        super().__init__(message)
        self.code = code
        self.delay_seconds = delay_seconds


def _future(seconds: int) -> str:
    return (datetime.now(timezone.utc) + timedelta(seconds=seconds)).replace(microsecond=0).isoformat()


def _bucket(minutes: int = 60) -> str:
    now = datetime.now(timezone.utc)
    minute = (now.minute // minutes) * minutes if minutes < 60 else 0
    return now.replace(minute=minute, second=0, microsecond=0).strftime("%Y%m%d%H%M")


def claim_job(conn) -> dict | None:
    conn.execute("BEGIN" if sqlite3.is_postgres_backend() else "BEGIN IMMEDIATE")
    conn.execute(
        '''UPDATE oms_integration_jobs
           SET status='retry', locked_at=NULL, locked_by=NULL, lease_expires_at=NULL,
               updated_at=CURRENT_TIMESTAMP
           WHERE status='running' AND lease_expires_at IS NOT NULL
             AND lease_expires_at < ?''',
        (utcnow(),),
    )
    lock_clause = " FOR UPDATE SKIP LOCKED" if sqlite3.is_postgres_backend() else ""
    row = conn.execute(
        '''SELECT * FROM oms_integration_jobs
           WHERE status IN ('pending','retry')
             AND available_at <= ?
             AND attempts < max_attempts
           ORDER BY available_at, id LIMIT 1''' + lock_clause,
        (utcnow(),),
    ).fetchone()
    if not row:
        conn.commit()
        return None
    expires = _future(LEASE_SECONDS)
    conn.execute(
        '''UPDATE oms_integration_jobs
           SET status='running', attempts=attempts+1, locked_at=?, locked_by=?,
               lease_expires_at=?, updated_at=CURRENT_TIMESTAMP
           WHERE id=? AND status IN ('pending','retry')''',
        (utcnow(), WORKER_ID, expires, row["id"]),
    )
    if conn.execute("SELECT changes()").fetchone()[0] != 1:
        conn.rollback()
        return None
    conn.commit()
    return dict(conn.execute("SELECT * FROM oms_integration_jobs WHERE id=?", (row["id"],)).fetchone())


def finish_job(conn, job_id: int, result: Any = None):
    conn.execute(
        '''UPDATE oms_integration_jobs
           SET status='succeeded', completed_at=?, locked_at=NULL, locked_by=NULL,
               lease_expires_at=NULL, last_error=NULL, last_error_code=NULL,
               updated_at=CURRENT_TIMESTAMP
           WHERE id=?''',
        (utcnow(), job_id),
    )
    conn.commit()


def retry_job(conn, job: dict, error: Exception, *, delay_seconds: int | None = None, code: str | None = None):
    attempts = int(job["attempts"] or 1)
    if delay_seconds is None:
        delay_seconds = min(3600, (2 ** min(attempts, 7)) * 15) + random.randint(0, 15)
    if attempts >= int(job["max_attempts"] or 10):
        dead_job(conn, job, error, code=code or "max_attempts")
        return
    conn.execute(
        '''UPDATE oms_integration_jobs
           SET status='retry', available_at=?, locked_at=NULL, locked_by=NULL,
               lease_expires_at=NULL, last_error_code=?, last_error=?,
               updated_at=CURRENT_TIMESTAMP WHERE id=?''',
        (_future(delay_seconds), code or getattr(error, "code", "retry"), str(error)[:1000], job["id"]),
    )
    if job.get("aggregate_type") == "order_notification":
        try:
            enqueue_notification_failure_alert(
                conn,
                job,
                phase="delayed",
                error_code=code or getattr(error, "code", "retry"),
            )
        except Exception:
            print("failed to queue delayed notification alert", file=sys.stderr, flush=True)
            print(traceback.format_exc(), file=sys.stderr, flush=True)
    conn.commit()


def dead_job(conn, job: dict, error: Exception, *, code: str | None = None):
    error_code = code or getattr(error, "code", "failed")
    error_summary = str(error)[:1000]
    conn.execute(
        '''UPDATE oms_integration_jobs
           SET status='dead_letter', locked_at=NULL, locked_by=NULL,
               lease_expires_at=NULL, last_error_code=?, last_error=?,
               updated_at=CURRENT_TIMESTAMP WHERE id=?''',
        (error_code, error_summary, job["id"]),
    )
    if job.get("aggregate_type") == "order" and job.get("aggregate_id"):
        mark_manual_review(conn, job["aggregate_id"], f"异步任务失败: {job['job_type']} - {str(error)[:300]}", commit=False)
    elif job.get("aggregate_type") == "order_notification" and job.get("aggregate_id"):
        notification_job = conn.execute(
            "SELECT status FROM order_notification_jobs WHERE id=?",
            (job["aggregate_id"],),
        ).fetchone()
        if notification_job:
            previous_status = notification_job["status"]
            preserve_statuses = {
                "SENT",
                "MANUAL_REVIEW",
                "READY_PREVIEW",
                "READY_MANUAL",
                "SKIPPED",
            }
            domain_status_changed = previous_status not in preserve_statuses and previous_status != "DEAD_LETTER"
            if domain_status_changed:
                conn.execute(
                    '''UPDATE order_notification_jobs
                       SET status='DEAD_LETTER',last_error_code=?,last_error_summary=?,
                           updated_at=CURRENT_TIMESTAMP
                       WHERE id=?''',
                    (error_code, error_summary, job["aggregate_id"]),
                )
            conn.execute(
                '''INSERT INTO notification_audit_logs
                   (action,object_type,object_id,actor_type,before_summary,after_summary)
                   VALUES ('notification_queue_dead_letter','order_notification_job',?,'system',?,?)''',
                (
                    job["aggregate_id"],
                    json_dump({"status": previous_status}),
                    json_dump(
                        {
                            "status": "DEAD_LETTER" if domain_status_changed else previous_status,
                            "domain_status_changed": domain_status_changed,
                            "queue_job_id": job["id"],
                            "attempts": int(job.get("attempts") or 0),
                            "max_attempts": int(job.get("max_attempts") or 0),
                            "error_code": error_code,
                            "error_summary": error_summary,
                        }
                    ),
                ),
            )
            try:
                enqueue_notification_failure_alert(
                    conn,
                    job,
                    phase="final",
                    error_code=error_code,
                )
            except Exception:
                print("failed to queue final notification alert", file=sys.stderr, flush=True)
                print(traceback.format_exc(), file=sys.stderr, flush=True)
    conn.commit()


def _audit_wms_success(conn, job: dict, operation: str, result: WmsResult, correlation_id: str):
    conn.execute(
        '''INSERT INTO oms_external_api_calls
           (correlation_id, job_id, provider, operation, method, endpoint,
            request_hash, request_redacted, response_http_code, response_code,
            response_redacted, duration_ms, attempt, outcome)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,'success')''',
        (
            correlation_id,
            job["id"],
            "hungary_wms",
            operation,
            result.method,
            result.endpoint,
            hashlib.sha256(json_dump(result.request_redacted).encode()).hexdigest(),
            json_dump(result.request_redacted),
            result.http_status,
            str(result.business_code),
            json_dump(redact(result.raw)),
            result.duration_ms,
            job["attempts"],
        ),
    )
    conn.commit()


def _audit_wms_error(conn, job: dict, operation: str, error: WmsError, correlation_id: str):
    conn.execute(
        '''INSERT INTO oms_external_api_calls
           (correlation_id, job_id, provider, operation, method, endpoint,
            response_http_code, response_code, response_redacted, duration_ms,
            attempt, outcome, error)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (
            correlation_id,
            job["id"],
            "hungary_wms",
            operation,
            "POST",
            operation,
            error.http_status,
            str(error.business_code or error.code),
            json_dump(redact(error.response)) if error.response is not None else None,
            error.duration_ms,
            job["attempts"],
            "unknown" if error.unknown_outcome else "error",
            str(error)[:1000],
        ),
    )
    conn.commit()


def _audit_poland_success(
    conn,
    job: dict,
    operation: str,
    result: PolandWmsResult,
    correlation_id: str,
):
    request_redacted = redact(result.request_redacted)
    conn.execute(
        '''INSERT INTO oms_external_api_calls
           (correlation_id, job_id, provider, operation, method, endpoint,
            request_hash, request_redacted, response_http_code, response_code,
            response_redacted, duration_ms, attempt, outcome)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,'success')''',
        (
            correlation_id,
            job["id"],
            "poland_wms",
            operation,
            result.method,
            result.endpoint,
            hashlib.sha256(json_dump(request_redacted).encode()).hexdigest(),
            json_dump(request_redacted),
            result.http_status,
            str(result.business_code),
            json_dump(redact(result.raw)),
            result.duration_ms,
            job["attempts"],
        ),
    )
    conn.commit()


def _audit_poland_error(
    conn,
    job: dict,
    operation: str,
    error: PolandWmsError,
    correlation_id: str,
):
    conn.execute(
        '''INSERT INTO oms_external_api_calls
           (correlation_id, job_id, provider, operation, method, endpoint,
            response_http_code, response_code, response_redacted, duration_ms,
            attempt, outcome, error)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (
            correlation_id,
            job["id"],
            "poland_wms",
            operation,
            "POST" if operation in {"authenticate", "create_order", "cancel_order"} else "GET",
            operation,
            error.http_status,
            str(error.business_code or error.code),
            json_dump(redact(error.response)) if error.response is not None else None,
            error.duration_ms,
            job["attempts"],
            "unknown" if error.unknown_outcome else "error",
            str(error)[:1000],
        ),
    )
    conn.commit()


def _recursive_field(value: Any, names: tuple[str, ...]):
    if isinstance(value, dict):
        for name in names:
            if value.get(name) not in (None, ""):
                return value.get(name)
        for nested in value.values():
            hit = _recursive_field(nested, names)
            if hit not in (None, ""):
                return hit
    elif isinstance(value, list):
        for nested in value:
            hit = _recursive_field(nested, names)
            if hit not in (None, ""):
                return hit
    return None


def _wms_error_message(data: Any) -> str | None:
    if isinstance(data, dict):
        value = data.get("errorMsg") or data.get("error")
        if value:
            return str(value)
        for nested in data.values():
            found = _wms_error_message(nested)
            if found:
                return found
    elif isinstance(data, list):
        for nested in data:
            found = _wms_error_message(nested)
            if found:
                return found
    return None


def handle_refresh_stock(conn, job: dict, payload: dict):
    row = conn.execute(
        "SELECT * FROM oms_warehouse_integrations WHERE provider='hungary_wms' AND is_enabled=1 LIMIT 1"
    ).fetchone()
    if not row:
        return {"skipped": "integration_disabled"}
    client = WmsClient(base_url=row["base_url"])
    correlation = uuid.uuid4().hex
    try:
        result = client.inventory(row["external_code"] or "HU01")
    except WmsError as exc:
        _audit_wms_error(conn, job, "inventory", exc, correlation)
        raise
    _audit_wms_success(conn, job, "inventory", result, correlation)
    items = result.data if isinstance(result.data, list) else []
    synced = 0
    for item in items:
        if not isinstance(item, dict):
            continue
        barcode = str(item.get("product_sku_barcode") or item.get("productSkuBarcode") or "").strip()
        if not barcode:
            continue
        sku = conn.execute(
            "SELECT id FROM inv_skus WHERE UPPER(barcode)=UPPER(?) OR UPPER(sku_code)=UPPER(?) LIMIT 1",
            (barcode, barcode),
        ).fetchone()
        sku_id = sku["id"] if sku else None
        conn.execute(
            '''INSERT INTO oms_external_stock
               (warehouse_id, sku_barcode, sku_id, quantity, lock_quantity,
                available_quantity, source_updated_at, synced_at, raw_json)
               VALUES (?,?,?,?,?,?,?,?,?)
               ON CONFLICT(warehouse_id, sku_barcode) DO UPDATE SET
                 sku_id=excluded.sku_id, quantity=excluded.quantity,
                 lock_quantity=excluded.lock_quantity,
                 available_quantity=excluded.available_quantity,
                 source_updated_at=excluded.source_updated_at,
                 synced_at=excluded.synced_at, raw_json=excluded.raw_json''',
            (
                row["warehouse_id"], barcode, sku_id,
                int(item.get("quantity") or 0), int(item.get("lock_quantity") or 0),
                int(item.get("available_quantity") or 0),
                item.get("updated_at") or item.get("updateTime"), utcnow(), json_dump(item),
            ),
        )
        if sku_id:
            conn.execute(
                '''INSERT INTO oms_sku_warehouses (sku_id, warehouse_id, is_enabled, notes)
                   VALUES (?, ?, 1, '由匈牙利 WMS 库存自动识别')
                   ON CONFLICT(sku_id, warehouse_id) DO UPDATE SET
                     is_enabled=1, updated_at=CURRENT_TIMESTAMP''',
                (sku_id, row["warehouse_id"]),
            )
        synced += 1
    conn.commit()
    return {"synced": synced}


def handle_submit_wms(conn, job: dict, payload: dict):
    fid = payload.get("fulfillment_id") or job["aggregate_id"]
    fulfillment = conn.execute("SELECT * FROM oms_fulfillments WHERE id=?", (fid,)).fetchone()
    if not fulfillment:
        raise DomainError("履约单不存在", "fulfillment_not_found")
    if fulfillment["status"] in {"accepted", "picking", "packed", "shipped", "delivered"}:
        return {"noop": fulfillment["status"]}
    if fulfillment["status"] == "submission_unknown":
        enqueue_job(
            conn, "VERIFY_WMS_SUBMISSION", "fulfillment", fid,
            f"verify:{fulfillment['idempotency_key']}:{job['attempts']}", {"fulfillment_id": fid},
            available_at=_future(30),
        )
        conn.commit()
        return {"verify": True}
    if fulfillment["status"] not in {"ready_to_submit", "failed_retryable"}:
        raise DomainError(f"履约状态 {fulfillment['status']} 不能提交 WMS", "not_submittable")

    request_payload = build_wms_payload(conn, fid)
    payload_hash = hashlib.sha256(json_dump(request_payload).encode()).hexdigest()
    transition_fulfillment(
        conn, fid, "submitting", reason="开始提交匈牙利 WMS",
        extra_updates={"submitted_at": utcnow(), "payload_hash": payload_hash},
    )
    conn.commit()
    integration = conn.execute(
        "SELECT * FROM oms_warehouse_integrations WHERE warehouse_id=?", (fulfillment["warehouse_id"],)
    ).fetchone()
    client = WmsClient(base_url=integration["base_url"] if integration else None)
    correlation = uuid.uuid4().hex
    try:
        result = client.create_invoices([request_payload])
    except WmsError as exc:
        _audit_wms_error(conn, job, "create_invoice", exc, correlation)
        current = conn.execute("SELECT status FROM oms_fulfillments WHERE id=?", (fid,)).fetchone()["status"]
        if exc.unknown_outcome and current == "submitting":
            transition_fulfillment(
                conn, fid, "submission_unknown", reason=str(exc), correlation_id=correlation,
                extra_updates={"last_error_code": exc.code, "last_error_message": str(exc)},
            )
            enqueue_job(
                conn, "VERIFY_WMS_SUBMISSION", "fulfillment", fid,
                f"verify:{fulfillment['idempotency_key']}:{job['attempts']}", {"fulfillment_id": fid},
                available_at=_future(30),
            )
            conn.commit()
            return {"unknown": True}
        if exc.retryable and current == "submitting":
            transition_fulfillment(
                conn, fid, "failed_retryable", reason=str(exc), correlation_id=correlation,
                extra_updates={"last_error_code": exc.code, "last_error_message": str(exc)},
            )
            conn.commit()
        raise
    _audit_wms_success(conn, job, "create_invoice", result, correlation)
    error_message = _wms_error_message(result.data)
    if error_message:
        transition_fulfillment(
            conn, fid, "rejected", reason=error_message, correlation_id=correlation,
            extra_updates={"last_error_code": "wms_rejected", "last_error_message": error_message},
        )
        mark_manual_review(conn, fulfillment["order_id"], f"匈牙利 WMS 拒绝履约: {error_message}", commit=False)
        conn.commit()
        raise DomainError(error_message, "wms_rejected")

    pick_code = _recursive_field(result.data, ("pickCode",))
    tracking_number = _recursive_field(
        result.data, ("expressCode", "trackingNumber", "tracking_number")
    )
    label_url = _recursive_field(result.data, ("labelUrl", "labelURL"))
    transition_fulfillment(
        conn, fid, "accepted", reason="WMS 已接单", correlation_id=correlation,
        extra_updates={
            "external_pick_code": str(pick_code) if pick_code else None,
            "external_label_url": str(label_url) if label_url else None,
            "last_error_code": None, "last_error_message": None,
        },
    )
    if tracking_number:
        create_shipment(
            conn, fid, str(tracking_number), carrier_slug="wms-auto", carrier_name="WMS动态物流",
            label_url=str(label_url) if label_url else None,
            external_shipment_id=str(tracking_number), tracking_source="hungary_wms",
            initial_status="label_ready" if label_url else "label_pending", commit=False,
        )
    enqueue_job(
        conn, "POLL_WMS_STATUS", "fulfillment", fid,
        f"wms-status:{fid}:{_bucket(10)}", {"fulfillment_id": fid}, available_at=_future(60),
    )
    if not label_url:
        enqueue_job(
            conn, "FETCH_WMS_LABEL", "fulfillment", fid,
            f"wms-label:{fid}:r{fulfillment['revision']}", {"fulfillment_id": fid}, available_at=_future(30),
        )
    conn.commit()
    return {
        "accepted": True, "pick_code": bool(pick_code),
        "tracking_number": bool(tracking_number), "label": bool(label_url),
    }


def _poland_integration(conn, warehouse_id: int):
    row = conn.execute(
        """SELECT * FROM oms_warehouse_integrations
           WHERE warehouse_id=? AND provider='poland_wms'""",
        (warehouse_id,),
    ).fetchone()
    if not row or not row["is_enabled"]:
        raise DomainError("新波兰仓接口尚未启用", "poland_wms_disabled")
    return row


def _poland_client(integration) -> tuple[PolandWmsClient, dict]:
    if not integration:
        raise PolandWmsError(
            "新波兰仓接口配置不存在",
            code="integration_missing",
        )
    config = json_load(integration["config_json"], {}) or {}
    if not isinstance(config, dict):
        config = {}
    client = PolandWmsClient(
        order_base_url=integration["base_url"],
        print_base_url=config.get("print_base_url"),
    )
    return client, config


def _poland_result_row(data: Any) -> dict:
    if isinstance(data, dict):
        nested = data.get("data")
        if isinstance(nested, list):
            return next((row for row in nested if isinstance(row, dict)), {})
        if isinstance(nested, dict):
            return nested
        return data
    if isinstance(data, list):
        return next((row for row in data if isinstance(row, dict)), {})
    return {}


def _poland_external_fields(data: Any) -> dict:
    row = _poland_result_row(data)
    return {
        "order_id": row.get("order_id") or row.get("orderId"),
        "document_code": (
            row.get("order_customerinvoicecode")
            or row.get("documentCode")
            or row.get("reference_number")
        ),
        "tracking_number": (
            row.get("tracking_number")
            or row.get("trackingNumber")
            or row.get("order_serveinvoicecode")
            or row.get("order_transfercode")
        ),
        "carrier_name": (
            row.get("post_customername")
            or row.get("carrier_name")
            or row.get("carrierName")
            or "新波兰仓物流"
        ),
        "message": unquote(str(row.get("message") or row.get("msg") or "")),
        "row": row,
    }


def _create_poland_shipment(
    conn,
    fulfillment,
    *,
    tracking_number: str,
    carrier_name: str,
    label_url: str | None,
):
    shipment = create_shipment(
        conn,
        fulfillment["id"],
        str(tracking_number),
        carrier_slug="poland-wms",
        carrier_name=carrier_name or "新波兰仓物流",
        label_url=label_url,
        external_shipment_id=str(tracking_number),
        tracking_source="poland_wms",
        initial_status="label_ready" if label_url else "label_pending",
        commit=False,
    )
    enqueue_job(
        conn,
        "POLL_SHIPMENT_TRACKING",
        "shipment",
        shipment["id"],
        f"track:{shipment['id']}:provider-created",
        {"shipment_id": shipment["id"]},
        available_at=_future(120),
    )
    return shipment


def handle_submit_poland_wms(conn, job: dict, payload: dict):
    fid = payload.get("fulfillment_id") or job["aggregate_id"]
    fulfillment = conn.execute(
        "SELECT * FROM oms_fulfillments WHERE id=?", (fid,)
    ).fetchone()
    if not fulfillment:
        raise DomainError("履约单不存在", "fulfillment_not_found")
    if fulfillment["provider"] != "poland_wms":
        raise DomainError("该履约单不属于新波兰仓", "provider_mismatch")
    if fulfillment["status"] in {"accepted", "picking", "packed", "shipped", "delivered"}:
        return {"noop": fulfillment["status"]}
    if fulfillment["status"] == "submission_unknown":
        enqueue_job(
            conn,
            "VERIFY_PL_SUBMISSION",
            "fulfillment",
            fid,
            f"verify-pl:{fulfillment['idempotency_key']}:{job['attempts']}",
            {"fulfillment_id": fid},
            available_at=_future(30),
        )
        conn.commit()
        return {"verify": True}
    if fulfillment["status"] not in {"ready_to_submit", "failed_retryable"}:
        raise DomainError(
            f"履约状态 {fulfillment['status']} 不能提交新波兰仓",
            "not_submittable",
        )

    integration = _poland_integration(conn, fulfillment["warehouse_id"])
    client, config = _poland_client(integration)
    auth_correlation = uuid.uuid4().hex
    try:
        auth = client.authenticate()
    except PolandWmsError as exc:
        _audit_poland_error(conn, job, "authenticate", exc, auth_correlation)
        raise
    _audit_poland_success(conn, job, "authenticate", auth, auth_correlation)
    auth_data = auth.data if isinstance(auth.data, dict) else {}
    request_payload = build_poland_wms_payload(
        conn,
        fid,
        customer_id=str(auth_data["customer_id"]),
        customer_userid=str(auth_data["customer_userid"]),
    )
    payload_hash = hashlib.sha256(json_dump(request_payload).encode()).hexdigest()
    transition_fulfillment(
        conn,
        fid,
        "submitting",
        reason="开始提交新波兰仓华磊接口",
        extra_updates={"submitted_at": utcnow(), "payload_hash": payload_hash},
    )
    conn.commit()

    correlation = uuid.uuid4().hex
    try:
        result = client.create_order(request_payload)
    except PolandWmsError as exc:
        _audit_poland_error(conn, job, "create_order", exc, correlation)
        current = conn.execute(
            "SELECT status FROM oms_fulfillments WHERE id=?", (fid,)
        ).fetchone()["status"]
        if exc.unknown_outcome and current == "submitting":
            transition_fulfillment(
                conn,
                fid,
                "submission_unknown",
                reason=str(exc),
                correlation_id=correlation,
                extra_updates={
                    "last_error_code": exc.code,
                    "last_error_message": str(exc),
                },
            )
            enqueue_job(
                conn,
                "VERIFY_PL_SUBMISSION",
                "fulfillment",
                fid,
                f"verify-pl:{fulfillment['idempotency_key']}:{job['attempts']}",
                {"fulfillment_id": fid},
                available_at=_future(30),
            )
            conn.commit()
            return {"unknown": True}
        if exc.retryable and current == "submitting":
            transition_fulfillment(
                conn,
                fid,
                "failed_retryable",
                reason=str(exc),
                correlation_id=correlation,
                extra_updates={
                    "last_error_code": exc.code,
                    "last_error_message": str(exc),
                },
            )
            conn.commit()
        elif current == "submitting":
            transition_fulfillment(
                conn,
                fid,
                "rejected",
                reason=str(exc),
                correlation_id=correlation,
                extra_updates={
                    "last_error_code": exc.code,
                    "last_error_message": str(exc),
                },
            )
            mark_manual_review(
                conn,
                fulfillment["order_id"],
                f"新波兰仓拒绝履约: {str(exc)[:300]}",
                commit=False,
            )
            conn.commit()
        raise

    _audit_poland_success(conn, job, "create_order", result, correlation)
    fields = _poland_external_fields(result.data)
    external_order_id = fields["order_id"]
    tracking_number = fields["tracking_number"]
    if not external_order_id:
        transition_fulfillment(
            conn,
            fid,
            "rejected",
            reason="新波兰仓返回成功但缺少 order_id，无法生成标签",
            correlation_id=correlation,
        )
        mark_manual_review(
            conn,
            fulfillment["order_id"],
            "新波兰仓返回成功但缺少 order_id，需人工核对外部系统",
            commit=False,
        )
        conn.commit()
        return {"accepted": False, "manual_review": True}

    label_url = client.label_url(
        str(external_order_id),
        print_type=str(config.get("print_type") or "lab10_10"),
        label_format=config.get("label_format"),
        print_goods=bool(config.get("print_goods")),
    )
    transition_fulfillment(
        conn,
        fid,
        "accepted",
        reason="新波兰仓已接单",
        correlation_id=correlation,
        extra_updates={
            # The legacy schema's generic external reference field stores the
            # HuaLei order_id; it is required for label printing.
            "external_pick_code": str(external_order_id),
            "external_label_url": label_url,
            "last_error_code": None,
            "last_error_message": None,
        },
    )
    fulfillment = conn.execute(
        "SELECT * FROM oms_fulfillments WHERE id=?", (fid,)
    ).fetchone()
    if tracking_number:
        _create_poland_shipment(
            conn,
            fulfillment,
            tracking_number=str(tracking_number),
            carrier_name=str(fields["carrier_name"]),
            label_url=label_url,
        )
    else:
        enqueue_job(
            conn,
            "POLL_PL_TRACKING_NUMBER",
            "fulfillment",
            fid,
            f"pl-tracking-number:{fid}",
            {"fulfillment_id": fid},
            available_at=_future(120),
        )
    conn.commit()
    return {
        "accepted": True,
        "external_order_id": bool(external_order_id),
        "tracking_number": bool(tracking_number),
        "label": True,
    }


def _query_poland_order(conn, job: dict, fulfillment, operation: str):
    integration = _poland_integration(conn, fulfillment["warehouse_id"])
    client, config = _poland_client(integration)
    correlation = uuid.uuid4().hex
    try:
        result = client.tracking_number(
            document_code=fulfillment["external_invoice_code"]
        )
    except PolandWmsError as exc:
        _audit_poland_error(conn, job, operation, exc, correlation)
        raise
    _audit_poland_success(conn, job, operation, result, correlation)
    return client, config, _poland_external_fields(result.data), correlation


def handle_verify_poland_submission(conn, job: dict, payload: dict):
    fid = payload.get("fulfillment_id") or job["aggregate_id"]
    fulfillment = conn.execute(
        "SELECT * FROM oms_fulfillments WHERE id=?", (fid,)
    ).fetchone()
    if not fulfillment:
        raise DomainError("履约单不存在", "fulfillment_not_found")
    if fulfillment["status"] != "submission_unknown":
        return {"noop": fulfillment["status"]}
    client, config, fields, correlation = _query_poland_order(
        conn, job, fulfillment, "verify_submission"
    )
    found = bool(
        fields["order_id"]
        or fields["tracking_number"]
        or fields["document_code"] == fulfillment["external_invoice_code"]
    )
    if not found:
        if int(job["attempts"]) < 3:
            raise RetryTask(
                "新波兰仓暂未查到原单号，继续核验",
                code="not_found_yet",
                delay_seconds=60,
            )
        reason = (
            "新波兰仓创建请求结果连续无法确认；为避免重复出库，"
            "系统已停止自动重提并转人工核对"
        )
        transition_fulfillment(
            conn,
            fid,
            "manual_review",
            reason=reason,
            correlation_id=correlation,
        )
        mark_manual_review(conn, fulfillment["order_id"], reason, commit=False)
        conn.commit()
        return {"not_found": True, "manual_review": True, "resubmit": False}

    external_order_id = fields["order_id"]
    label_url = (
        client.label_url(
            str(external_order_id),
            print_type=str(config.get("print_type") or "lab10_10"),
            label_format=config.get("label_format"),
            print_goods=bool(config.get("print_goods")),
        )
        if external_order_id
        else fulfillment["external_label_url"]
    )
    transition_fulfillment(
        conn,
        fid,
        "accepted",
        reason="超时后按原单号确认新波兰仓已接单",
        correlation_id=correlation,
        extra_updates={
            "external_pick_code": (
                str(external_order_id)
                if external_order_id
                else fulfillment["external_pick_code"]
            ),
            "external_label_url": label_url,
            "last_error_code": None,
            "last_error_message": None,
        },
    )
    fulfillment = conn.execute(
        "SELECT * FROM oms_fulfillments WHERE id=?", (fid,)
    ).fetchone()
    if fields["tracking_number"]:
        _create_poland_shipment(
            conn,
            fulfillment,
            tracking_number=str(fields["tracking_number"]),
            carrier_name=str(fields["carrier_name"]),
            label_url=label_url,
        )
    else:
        enqueue_job(
            conn,
            "POLL_PL_TRACKING_NUMBER",
            "fulfillment",
            fid,
            f"pl-tracking-number:{fid}",
            {"fulfillment_id": fid},
            available_at=_future(120),
        )
    conn.commit()
    return {"confirmed": True}


def handle_poll_poland_tracking_number(conn, job: dict, payload: dict):
    fid = payload.get("fulfillment_id") or job["aggregate_id"]
    fulfillment = conn.execute(
        "SELECT * FROM oms_fulfillments WHERE id=?", (fid,)
    ).fetchone()
    if not fulfillment or fulfillment["status"] in {
        "shipped", "delivered", "cancelled", "returned", "superseded"
    }:
        return {"noop": True}
    existing = conn.execute(
        "SELECT * FROM oms_shipments WHERE fulfillment_id=? ORDER BY created_at LIMIT 1",
        (fid,),
    ).fetchone()
    if existing and existing["tracking_number"]:
        return {"noop": "tracking_exists"}

    client, config, fields, correlation = _query_poland_order(
        conn, job, fulfillment, "tracking_number"
    )
    if not fields["tracking_number"]:
        if int(job["attempts"]) >= 8:
            reason = "新波兰仓长时间未返回运单号，已转人工核对"
            transition_fulfillment(
                conn,
                fid,
                "manual_hold",
                reason=reason,
                correlation_id=correlation,
            )
            mark_manual_review(conn, fulfillment["order_id"], reason, commit=False)
            conn.commit()
            return {"manual_review": True}
        raise RetryTask(
            "新波兰仓运单号尚未生成",
            code="tracking_number_not_ready",
            delay_seconds=300,
        )

    external_order_id = fields["order_id"] or fulfillment["external_pick_code"]
    label_url = fulfillment["external_label_url"]
    if not label_url and external_order_id:
        label_url = client.label_url(
            str(external_order_id),
            print_type=str(config.get("print_type") or "lab10_10"),
            label_format=config.get("label_format"),
            print_goods=bool(config.get("print_goods")),
        )
        conn.execute(
            """UPDATE oms_fulfillments
               SET external_pick_code=COALESCE(external_pick_code, ?),
                   external_label_url=?, updated_at=CURRENT_TIMESTAMP
               WHERE id=?""",
            (str(external_order_id), label_url, fid),
        )
    fulfillment = conn.execute(
        "SELECT * FROM oms_fulfillments WHERE id=?", (fid,)
    ).fetchone()
    shipment = _create_poland_shipment(
        conn,
        fulfillment,
        tracking_number=str(fields["tracking_number"]),
        carrier_name=str(fields["carrier_name"]),
        label_url=label_url,
    )
    conn.commit()
    return {"tracking_number": True, "shipment_id": shipment["id"]}


def handle_cancel_poland_wms(conn, job: dict, payload: dict):
    fid = payload.get("fulfillment_id") or job["aggregate_id"]
    fulfillment = conn.execute(
        "SELECT * FROM oms_fulfillments WHERE id=?", (fid,)
    ).fetchone()
    if not fulfillment:
        raise DomainError("履约单不存在", "fulfillment_not_found")
    if fulfillment["status"] == "cancelled":
        return {"noop": "cancelled"}
    if fulfillment["provider"] != "poland_wms":
        raise DomainError("该履约单不属于新波兰仓", "provider_mismatch")
    if fulfillment["status"] != "cancel_pending":
        raise DomainError(
            f"履约状态 {fulfillment['status']} 不能执行新波兰仓取消",
            "not_cancellable",
        )
    external_order_id = fulfillment["external_pick_code"]
    if not external_order_id:
        reason = "新波兰仓缺少外部 order_id，无法调用取消接口"
        transition_fulfillment(conn, fid, "manual_review", reason=reason)
        mark_manual_review(conn, fulfillment["order_id"], reason, commit=False)
        conn.commit()
        return {"manual_required": True}

    integration = _poland_integration(conn, fulfillment["warehouse_id"])
    client, _ = _poland_client(integration)
    auth_correlation = uuid.uuid4().hex
    try:
        auth = client.authenticate()
    except PolandWmsError as exc:
        _audit_poland_error(conn, job, "authenticate", exc, auth_correlation)
        if exc.retryable:
            raise
        reason = f"新波兰仓取消前认证失败: {str(exc)[:300]}"
        transition_fulfillment(conn, fid, "manual_review", reason=reason)
        mark_manual_review(conn, fulfillment["order_id"], reason, commit=False)
        conn.commit()
        return {"manual_required": True, "reason": exc.code}
    _audit_poland_success(conn, job, "authenticate", auth, auth_correlation)
    auth_data = auth.data if isinstance(auth.data, dict) else {}

    correlation = uuid.uuid4().hex
    try:
        result = client.cancel_order(
            order_id=str(external_order_id),
            customer_id=str(auth_data["customer_id"]),
            reason=str(payload.get("reason") or "客户取消/人工调整"),
        )
    except PolandWmsError as exc:
        _audit_poland_error(conn, job, "cancel_order", exc, correlation)
        if exc.unknown_outcome:
            reason = (
                "新波兰仓取消请求结果未知；已停止自动重试，"
                "请在对方系统按外部 order_id 人工核验"
            )
            transition_fulfillment(
                conn,
                fid,
                "manual_review",
                reason=reason,
                correlation_id=correlation,
            )
            mark_manual_review(conn, fulfillment["order_id"], reason, commit=False)
            conn.commit()
            return {"manual_required": True, "unknown": True}
        reason = f"新波兰仓取消被拒绝: {str(exc)[:300]}"
        transition_fulfillment(
            conn,
            fid,
            "cancel_rejected",
            reason=reason,
            correlation_id=correlation,
        )
        mark_manual_review(conn, fulfillment["order_id"], reason, commit=False)
        conn.commit()
        return {"cancelled": False, "rejected": True, "reason": exc.code}

    _audit_poland_success(conn, job, "cancel_order", result, correlation)
    transition_fulfillment(
        conn,
        fid,
        "cancelled",
        reason="新波兰仓取消接口已确认成功",
        correlation_id=correlation,
    )
    recompute_order_status(conn, fulfillment["order_id"], commit=False)
    conn.commit()
    return {"cancelled": True}


def _single_status_row(result: WmsResult) -> dict | None:
    data = result.data
    if isinstance(data, list):
        return next((row for row in data if isinstance(row, dict)), None)
    return data if isinstance(data, dict) else None


def _query_wms_fulfillment(conn, job: dict, fulfillment, operation: str) -> tuple[dict | None, str]:
    integration = conn.execute(
        "SELECT * FROM oms_warehouse_integrations WHERE warehouse_id=?", (fulfillment["warehouse_id"],)
    ).fetchone()
    client = WmsClient(base_url=integration["base_url"] if integration else None)
    correlation = uuid.uuid4().hex
    try:
        result = client.invoice_status([fulfillment["external_invoice_code"]])
    except WmsError as exc:
        _audit_wms_error(conn, job, operation, exc, correlation)
        raise
    _audit_wms_success(conn, job, operation, result, correlation)
    return _single_status_row(result), correlation


def handle_verify_submission(conn, job: dict, payload: dict):
    fid = payload.get("fulfillment_id") or job["aggregate_id"]
    fulfillment = conn.execute("SELECT * FROM oms_fulfillments WHERE id=?", (fid,)).fetchone()
    if not fulfillment:
        raise DomainError("履约单不存在", "fulfillment_not_found")
    if fulfillment["status"] != "submission_unknown":
        return {"noop": fulfillment["status"]}
    row, correlation = _query_wms_fulfillment(conn, job, fulfillment, "verify_submission")
    if not row:
        if job["attempts"] < 3:
            raise RetryTask("WMS 暂未返回该 invoiceCode，继续查询", code="not_found_yet", delay_seconds=60)
        # The supplier has not confirmed invoiceCode idempotency or the final
        # consistency window.  Re-submitting after an unknown timeout could
        # create a duplicate outbound order, so stop automation and require a
        # human to reconcile with the supplier.
        reason = "WMS 提交结果连续查询仍未知；为避免重复出库，已停止自动重提并转人工核对"
        transition_fulfillment(
            conn, fid, "manual_review", reason=reason, correlation_id=correlation
        )
        mark_manual_review(conn, fulfillment["order_id"], reason, commit=False)
        conn.commit()
        return {"not_found": True, "manual_review": True, "resubmit": False}
    pick_code = row.get("pickCode") or fulfillment["external_pick_code"]
    tracking_number = row.get("expressCode") or row.get("trackingNumber")
    transition_fulfillment(
        conn, fid, "accepted", reason="超时后查询确认 WMS 已接单", correlation_id=correlation,
        extra_updates={"external_pick_code": pick_code, "last_error_code": None, "last_error_message": None},
    )
    if tracking_number:
        create_shipment(
            conn, fid, str(tracking_number), carrier_slug="wms-auto", carrier_name="WMS动态物流",
            external_shipment_id=str(tracking_number), tracking_source="hungary_wms",
            initial_status="label_pending", commit=False,
        )
    enqueue_job(
        conn, "POLL_WMS_STATUS", "fulfillment", fid,
        f"wms-status:{fid}:{_bucket(10)}", {"fulfillment_id": fid}, available_at=_future(60),
    )
    conn.commit()
    return {"confirmed": True}


def handle_fetch_label(conn, job: dict, payload: dict):
    fid = payload.get("fulfillment_id") or job["aggregate_id"]
    fulfillment = conn.execute("SELECT * FROM oms_fulfillments WHERE id=?", (fid,)).fetchone()
    if not fulfillment:
        raise DomainError("履约单不存在", "fulfillment_not_found")
    integration = conn.execute(
        "SELECT * FROM oms_warehouse_integrations WHERE warehouse_id=?", (fulfillment["warehouse_id"],)
    ).fetchone()
    client = WmsClient(base_url=integration["base_url"] if integration else None)
    correlation = uuid.uuid4().hex
    try:
        result = client.labels([fulfillment["external_invoice_code"]])
    except WmsError as exc:
        _audit_wms_error(conn, job, "get_label", exc, correlation)
        raise
    _audit_wms_success(conn, job, "get_label", result, correlation)
    label = _recursive_field(result.data, ("labelUrl", "labelURL"))
    if not label:
        raise RetryTask("WMS 暂未生成面单", code="label_not_ready", delay_seconds=120)
    conn.execute(
        "UPDATE oms_fulfillments SET external_label_url=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
        (str(label), fid),
    )
    conn.execute(
        '''UPDATE oms_shipments SET label_url=?, status=CASE WHEN status='label_pending' THEN 'label_ready' ELSE status END,
                  updated_at=CURRENT_TIMESTAMP WHERE fulfillment_id=?''',
        (str(label), fid),
    )
    conn.commit()
    return {"label": True}


def handle_cancel_wms(conn, job: dict, payload: dict):
    fid = payload.get("fulfillment_id") or job["aggregate_id"]
    fulfillment = conn.execute("SELECT * FROM oms_fulfillments WHERE id=?", (fid,)).fetchone()
    if not fulfillment:
        raise DomainError("履约单不存在", "fulfillment_not_found")
    if fulfillment["status"] == "cancelled":
        return {"noop": "cancelled"}
    if fulfillment["status"] != "cancel_pending":
        raise DomainError(f"履约状态 {fulfillment['status']} 不能执行 WMS 取消", "not_cancellable")
    # Backward-compatible guard for any cancellation job queued by an older
    # web process.  The supplier does not expose cancellation through the API.
    reason = "匈牙利 WMS 不支持 API 取消；请在物流群联系对方运营人工拦截"
    transition_fulfillment(
        conn, fid, "manual_review", reason=reason
    )
    mark_manual_review(conn, fulfillment["order_id"], reason, commit=False)
    conn.commit()
    return {"manual_required": True, "cancelled": False}


def handle_poll_wms_status(conn, job: dict, payload: dict):
    fid = payload.get("fulfillment_id") or job["aggregate_id"]
    fulfillment = conn.execute("SELECT * FROM oms_fulfillments WHERE id=?", (fid,)).fetchone()
    if not fulfillment or fulfillment["status"] in {"delivered", "cancelled", "returned", "superseded"}:
        return {"noop": True}
    row, correlation = _query_wms_fulfillment(conn, job, fulfillment, "invoice_status")
    if not row:
        raise RetryTask("WMS 状态暂不可用", code="status_not_ready", delay_seconds=120)
    raw_status = str(row.get("statusName") or row.get("status") or "")
    normalized = normalize_wms_fulfillment_status(raw_status)
    pick_code = row.get("pickCode") or fulfillment["external_pick_code"]
    tracking_number = row.get("expressCode") or row.get("trackingNumber")
    if normalized == "stock_shortage":
        if fulfillment["status"] != "stock_shortage":
            transition_fulfillment(conn, fid, "stock_shortage", reason=raw_status, correlation_id=correlation)
        conn.execute(
            "UPDATE oms_order_fulfillment_state SET has_shortage=1, manual_review=1, aggregate_status='manual_review', manual_reason=?, updated_at=CURRENT_TIMESTAMP WHERE order_id=?",
            (f"匈牙利 WMS 缺货: {raw_status}", fulfillment["order_id"]),
        )
    elif normalized in {"picking", "packed"} and fulfillment["status"] != normalized:
        transition_fulfillment(conn, fid, normalized, reason=raw_status, correlation_id=correlation)
    elif normalized == "shipped":
        shipment = conn.execute("SELECT * FROM oms_shipments WHERE fulfillment_id=? ORDER BY created_at LIMIT 1", (fid,)).fetchone()
        if not shipment and tracking_number:
            shipment = create_shipment(
                conn, fid, str(tracking_number), carrier_slug="wms-auto", carrier_name="WMS动态物流",
                external_shipment_id=str(tracking_number), tracking_source="hungary_wms",
                initial_status="label_pending", commit=False,
            )
        if shipment:
            mark_shipment_shipped(conn, shipment["id"], reason=f"WMS 状态: {raw_status}", commit=False)
    elif normalized == "cancelled" and fulfillment["status"] != "cancelled":
        transition_fulfillment(conn, fid, "cancelled", reason=raw_status, correlation_id=correlation)
    if pick_code and not fulfillment["external_pick_code"]:
        conn.execute("UPDATE oms_fulfillments SET external_pick_code=? WHERE id=?", (str(pick_code), fid))
    if normalized not in {"cancelled", "shipped"}:
        enqueue_job(
            conn, "POLL_WMS_STATUS", "fulfillment", fid,
            f"wms-status:{fid}:{_future(600)[:16]}", {"fulfillment_id": fid}, available_at=_future(600),
        )
    conn.commit()
    return {"raw_status": raw_status, "normalized": normalized}


def _extract_wms_tracking(data: Any) -> list[dict]:
    events = []
    candidates = data if isinstance(data, list) else [data]
    for item in candidates:
        if not isinstance(item, dict):
            continue
        nested = item.get("events") or item.get("details") or item.get("trackList")
        if isinstance(nested, list):
            events.extend(_extract_wms_tracking(nested))
            continue
        status = item.get("status") or item.get("trackStatus") or item.get("statusCode")
        if status:
            events.append({
                "raw_status": str(status),
                "event_at": item.get("time") or item.get("eventTime") or item.get("trackTime"),
                "location": item.get("location") or item.get("address"),
                "description": item.get("description") or item.get("content") or item.get("trackContent"),
                "raw": item,
            })
    return events


def _extract_poland_tracking(data: Any) -> list[dict]:
    events = []

    def visit(value: Any, inherited_status: str | None = None):
        if isinstance(value, list):
            for item in value:
                visit(item, inherited_status)
            return
        if not isinstance(value, dict):
            return
        raw_status = str(
            value.get("trackStatus")
            or value.get("businessStatus")
            or value.get("status")
            or inherited_status
            or ""
        ).strip()
        nested_keys = (
            "trackDetails",
            "trackList",
            "track_list",
            "details",
            "events",
            "tracks",
            "data",
        )
        nested_found = False
        for key in nested_keys:
            nested = value.get(key)
            if isinstance(nested, (list, dict)):
                nested_found = True
                visit(nested, raw_status)
        description = (
            value.get("track_content")
            or value.get("trackContent")
            or value.get("description")
            or value.get("content")
        )
        event_at = (
            value.get("track_date")
            or value.get("trackDate")
            or value.get("eventTime")
            or value.get("time")
        )
        if description or event_at or (raw_status and not nested_found):
            events.append(
                {
                    "raw_status": raw_status or str(description or ""),
                    "event_at": event_at,
                    "location": (
                        value.get("location")
                        or value.get("track_location")
                        or value.get("trackLocation")
                    ),
                    "description": description,
                    "raw": value,
                }
            )

    visit(data)
    events.sort(key=lambda item: str(item.get("event_at") or ""))
    return events


def handle_poll_tracking(conn, job: dict, payload: dict):
    sid = payload.get("shipment_id") or job["aggregate_id"]
    shipment = conn.execute(
        '''SELECT s.*, f.provider, f.warehouse_id FROM oms_shipments s
           JOIN oms_fulfillments f ON f.id=s.fulfillment_id WHERE s.id=?''',
        (sid,),
    ).fetchone()
    if not shipment or shipment["status"] in {"delivered", "returned", "cancelled"}:
        return {"noop": True}
    tracking = shipment["tracking_number"]
    official_ok = third_party_ok = False
    if shipment["provider"] == "hungary_wms":
        integration = conn.execute(
            "SELECT * FROM oms_warehouse_integrations WHERE warehouse_id=?", (shipment["warehouse_id"],)
        ).fetchone()
        client = WmsClient(base_url=integration["base_url"] if integration else None)
        correlation = uuid.uuid4().hex
        try:
            result = client.track(tracking)
            _audit_wms_success(conn, job, "tracking", result, correlation)
            events = _extract_wms_tracking(result.data)
            if not events:
                status = _recursive_field(result.data, ("status", "trackStatus", "statusCode"))
                if status:
                    events = [{"raw_status": str(status), "raw": result.data}]
            for event in events:
                add_tracking_event(
                    conn, sid, "hungary_wms", normalize_wms_tracking_status(event["raw_status"]),
                    raw_status=event["raw_status"], event_at=event.get("event_at"),
                    location=event.get("location"), description=event.get("description"),
                    raw_payload=event.get("raw"), correlation_id=correlation, commit=False,
                )
            official_ok = bool(events)
        except WmsError as exc:
            _audit_wms_error(conn, job, "tracking", exc, correlation)
    elif shipment["provider"] == "poland_wms":
        integration = conn.execute(
            "SELECT * FROM oms_warehouse_integrations WHERE warehouse_id=?",
            (shipment["warehouse_id"],),
        ).fetchone()
        correlation = uuid.uuid4().hex
        try:
            client, _ = _poland_client(integration)
            result = client.track(tracking)
            _audit_poland_success(conn, job, "tracking", result, correlation)
            events = _extract_poland_tracking(result.data)
            for event in events:
                normalized = normalize_poland_tracking_status(event["raw_status"])
                current = conn.execute(
                    "SELECT status FROM oms_shipments WHERE id=?", (sid,)
                ).fetchone()["status"]
                if SHIPMENT_PROGRESS.get(current, 0) < 2 and SHIPMENT_PROGRESS.get(
                    normalized, 0
                ) >= 2:
                    mark_shipment_shipped(
                        conn,
                        sid,
                        reason=f"新波兰仓官方轨迹: {event['raw_status']}",
                        commit=False,
                    )
                add_tracking_event(
                    conn,
                    sid,
                    "poland_wms",
                    normalized,
                    raw_status=event["raw_status"],
                    event_at=event.get("event_at"),
                    location=event.get("location"),
                    description=event.get("description"),
                    raw_payload=event.get("raw"),
                    correlation_id=correlation,
                    commit=False,
                )
            official_ok = bool(events)
        except PolandWmsError as exc:
            _audit_poland_error(conn, job, "tracking", exc, correlation)

    key_row = conn.execute("SELECT value FROM settings WHERE key='track718_api_key'").fetchone()
    key = key_row["value"] if key_row else None
    carrier = carrier_tracking.classify_carrier(shipment["carrier_slug"], tracking)
    try:
        if carrier == "expressone_hu":
            result = carrier_tracking.expressone_hu_detail(tracking)
            if result.get("ok"):
                outcome = result.get("outcome")
                normalized = {
                    "in_transit": "in_transit", "attention": "exception",
                    "returned": "returned", "delivered": "delivered",
                }.get(outcome, "exception")
                for event in (result.get("events") or [])[:20]:
                    add_tracking_event(
                        conn, sid, "expressone_hu", normalized,
                        raw_status=str(event.get("code") or outcome or ""),
                        event_at=event.get("time"), description=event.get("status"),
                        raw_payload=result, commit=False,
                    )
                official_ok = True
            # The official Express One page is authoritative; Track718 remains
            # an independent supplemental source when configured.
            if key:
                third = carrier_tracking.track718_detail(
                    tracking, key, code=None, poll=2, poll_wait=1
                )
                if third.get("ok"):
                    outcome = third.get("outcome")
                    normalized = {
                        "in_transit": "in_transit", "attention": "exception",
                        "returned": "returned", "delivered": "delivered",
                    }.get(outcome, "exception")
                    for event in (third.get("events") or [{}])[:20]:
                        add_tracking_event(
                            conn, sid, "track718", normalized,
                            raw_status=str(third.get("result") or outcome or ""),
                            event_at=event.get("time"), description=event.get("status"),
                            raw_payload=third, commit=False,
                        )
                    third_party_ok = True
        elif carrier == "inpost":
            result = carrier_tracking.inpost_status(tracking)
            if result.get("ok"):
                outcome = result.get("outcome")
                normalized = {"in_transit": "in_transit", "attention": "exception", "returned": "returned", "delivered": "delivered"}.get(outcome, "exception")
                add_tracking_event(
                    conn, sid, "inpost", normalized, raw_status=result.get("raw"),
                    event_at=result.get("last_event_at"), description=result.get("last_event"),
                    raw_payload=result, commit=False,
                )
                third_party_ok = True
        elif key:
            code = (
                "dpd-pl" if carrier == "dpd"
                else carrier_tracking.TRACK718_PACKETA if carrier == "packeta"
                else None
            )
            result = carrier_tracking.track718_detail(tracking, key, code=code, poll=2, poll_wait=1)
            if result.get("ok"):
                outcome = result.get("outcome")
                normalized = {"in_transit": "in_transit", "attention": "exception", "returned": "returned", "delivered": "delivered"}.get(outcome, "exception")
                events = result.get("events") or [{}]
                for event in events[:20]:
                    add_tracking_event(
                        conn, sid, "track718", normalized,
                        raw_status=str(result.get("result") or outcome or ""),
                        event_at=event.get("time"), description=event.get("status"),
                        raw_payload=result, commit=False,
                    )
                third_party_ok = True
    except Exception:
        pass
    current = conn.execute("SELECT status FROM oms_shipments WHERE id=?", (sid,)).fetchone()["status"]
    if current not in {"delivered", "returned", "cancelled"}:
        enqueue_job(
            conn, "POLL_SHIPMENT_TRACKING", "shipment", sid,
            f"track:{sid}:{_future(21600)[:13]}", {"shipment_id": sid}, available_at=_future(21600),
        )
    conn.commit()
    if not official_ok and not third_party_ok:
        raise RetryTask("官方和第三方物流暂未返回可用轨迹", code="tracking_not_ready", delay_seconds=900)
    return {"official": official_ok, "third_party": third_party_ok, "status": current}


def handle_reconcile(conn, job: dict, payload: dict):
    issues = 0
    checks = conn.execute(
        '''SELECT ofs.order_id, ofs.aggregate_status, o.status AS woo_status
           FROM oms_order_fulfillment_state ofs JOIN orders o ON o.id=ofs.order_id'''
    ).fetchall()
    for row in checks:
        issue = None
        if row["woo_status"] == "completed" and row["aggregate_status"] != "delivered":
            issue = ("completed_before_all_delivered", "danger")
        elif row["aggregate_status"] == "delivered" and row["woo_status"] != "completed":
            issue = ("delivered_not_completed", "warning")
        if issue:
            key = f"{issue[0]}:{row['order_id']}"
            conn.execute(
                '''INSERT INTO oms_reconciliation_issues
                   (issue_type, severity, aggregate_type, aggregate_id, dedup_key, detail_json)
                   VALUES (?,?,'order',?,?,?)
                   ON CONFLICT(dedup_key, status) DO UPDATE SET
                     last_seen_at=CURRENT_TIMESTAMP, detail_json=excluded.detail_json''',
                (issue[0], issue[1], row["order_id"], key, json_dump(dict(row))),
            )
            issues += 1
    conn.commit()
    return {"issues": issues}


HANDLERS = {
    "PLAN_ORDER": lambda c, j, p: plan_order(c, p.get("order_id") or j["aggregate_id"]),
    "REFRESH_HU_STOCK": handle_refresh_stock,
    "SUBMIT_HU_FULFILLMENT": handle_submit_wms,
    "SUBMIT_PL_FULFILLMENT": handle_submit_poland_wms,
    "VERIFY_WMS_SUBMISSION": handle_verify_submission,
    "VERIFY_PL_SUBMISSION": handle_verify_poland_submission,
    "POLL_PL_TRACKING_NUMBER": handle_poll_poland_tracking_number,
    "CANCEL_PL_FULFILLMENT": handle_cancel_poland_wms,
    "FETCH_WMS_LABEL": handle_fetch_label,
    "CANCEL_HU_FULFILLMENT": handle_cancel_wms,
    "POLL_WMS_STATUS": handle_poll_wms_status,
    "SYNC_SHIPMENT_TO_WOOCOMMERCE": lambda c, j, p: sync_shipment(c, p.get("shipment_id") or j["aggregate_id"]),
    "POLL_SHIPMENT_TRACKING": handle_poll_tracking,
    "RECOMPUTE_ORDER_STATUS": lambda c, j, p: recompute_order_status(c, p.get("order_id") or j["aggregate_id"]),
    "COMPLETE_WOOCOMMERCE_ORDER": lambda c, j, p: complete_order(c, p.get("order_id") or j["aggregate_id"]),
    "RECONCILE": handle_reconcile,
    "ORDER_NOTIFICATION": process_notification_job,
    "ORDER_NOTIFICATION_ALERT": process_notification_alert,
}


def dispatch(conn, job: dict):
    handler = HANDLERS.get(job["job_type"])
    if not handler:
        raise DomainError(f"未知任务类型: {job['job_type']}", "unknown_job_type")
    return handler(conn, job, json_load(job.get("payload_json"), {}) or {})


def seed_periodic_jobs(conn):
    if not conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='oms_integration_jobs'").fetchone():
        return
    enabled = conn.execute(
        "SELECT value FROM settings WHERE key='oms_fulfillment_enabled'"
    ).fetchone()
    if not enabled or str(enabled["value"]).strip().lower() not in {"1", "true", "yes", "on"}:
        return
    hour = datetime.now(timezone.utc).strftime("%Y%m%d%H")
    day = datetime.now(timezone.utc).strftime("%Y%m%d")
    if conn.execute("SELECT 1 FROM oms_warehouse_integrations WHERE provider='hungary_wms' AND is_enabled=1").fetchone():
        enqueue_job(conn, "REFRESH_HU_STOCK", "warehouse", "HU", f"hu-stock:{hour}", {})
    enqueue_job(conn, "RECONCILE", "system", "fulfillment", f"reconcile:{day}:{hour}", {})
    for row in conn.execute(
        '''SELECT id FROM oms_fulfillments
           WHERE provider='hungary_wms' AND status IN ('accepted','picking','packed')'''
    ).fetchall():
        enqueue_job(
            conn, "POLL_WMS_STATUS", "fulfillment", row["id"],
            f"wms-status:{row['id']}:{_bucket(10)}", {"fulfillment_id": row["id"]},
        )
    for row in conn.execute(
        '''SELECT id FROM oms_fulfillments
           WHERE provider='poland_wms' AND status='accepted'
             AND NOT EXISTS (
               SELECT 1 FROM oms_shipments s
               WHERE s.fulfillment_id=oms_fulfillments.id
                 AND s.tracking_number IS NOT NULL
             )'''
    ).fetchall():
        enqueue_job(
            conn,
            "POLL_PL_TRACKING_NUMBER",
            "fulfillment",
            row["id"],
            f"pl-tracking-number:{row['id']}",
            {"fulfillment_id": row["id"]},
        )
    for row in conn.execute(
        "SELECT id FROM oms_shipments WHERE status NOT IN ('delivered','returned','cancelled') AND tracking_number IS NOT NULL"
    ).fetchall():
        enqueue_job(
            conn, "POLL_SHIPMENT_TRACKING", "shipment", row["id"],
            f"track:{row['id']}:{hour}", {"shipment_id": row["id"]},
        )
    conn.commit()


def run_one(conn) -> bool:
    job = claim_job(conn)
    if not job:
        return False
    try:
        result = dispatch(conn, job)
        finish_job(conn, job["id"], result)
    except RetryTask as exc:
        conn.rollback()
        retry_job(conn, job, exc, delay_seconds=exc.delay_seconds, code=exc.code)
    except NotificationRetry as exc:
        conn.rollback()
        retry_job(conn, job, exc, delay_seconds=exc.delay_seconds, code=exc.code)
    except NotificationPermanent as exc:
        conn.rollback()
        dead_job(conn, job, exc, code=exc.code)
    except (WmsError, PolandWmsError) as exc:
        conn.rollback()
        if exc.retryable:
            retry_job(conn, job, exc, code=exc.code)
        else:
            dead_job(conn, job, exc, code=exc.code)
    except WooError as exc:
        conn.rollback()
        if exc.retryable or exc.unknown_outcome:
            retry_job(conn, job, exc, code=exc.code)
        else:
            dead_job(conn, job, exc, code=exc.code)
    except DomainError as exc:
        conn.rollback()
        dead_job(conn, job, exc, code=exc.code)
    except Exception as exc:
        print(traceback.format_exc(), file=sys.stderr, flush=True)
        # PostgreSQL rejects every statement after one statement in the
        # transaction fails.  Clear that failed transaction before recording
        # the durable retry; SQLite accepts the same rollback safely.
        conn.rollback()
        retry_job(conn, job, exc, code="unhandled")
    return True


def run_iteration(conn, *, seed: bool = False) -> bool:
    """Run one worker iteration without exiting on a transient SQLite lock."""

    try:
        if seed:
            seed_periodic_jobs(conn)
        return run_one(conn)
    except sqlite3.OperationalError as exc:
        if "locked" not in str(exc).lower() and "busy" not in str(exc).lower():
            raise
        try:
            conn.rollback()
        except sqlite3.Error:
            pass
        print(
            f"[worker] SQLite 暂时繁忙，保留任务并等待下一轮: {exc}",
            file=sys.stderr,
            flush=True,
        )
        return False


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--once", action="store_true")
    parser.add_argument("--idle-sleep", type=float, default=5.0)
    args = parser.parse_args()
    conn = get_conn()
    last_seed = 0.0
    try:
        while True:
            should_seed = time.monotonic() - last_seed > 60
            worked = run_iteration(conn, seed=should_seed)
            if should_seed:
                last_seed = time.monotonic()
            if args.once:
                break
            if not worked:
                time.sleep(max(0.5, args.idle_sleep))
    finally:
        conn.close()


if __name__ == "__main__":
    main()
