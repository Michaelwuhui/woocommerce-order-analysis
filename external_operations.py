"""Idempotency and reconciliation state for external WooCommerce mutations."""

from __future__ import annotations

import hashlib
import json
import uuid
from typing import Any


ALLOWED_TRANSITIONS = {
    "pending": {"external_success", "reconciliation_required", "failed", "cancelled"},
    "external_success": {"local_committed", "reconciliation_required"},
    "reconciliation_required": {"external_success", "local_committed", "failed"},
    "local_committed": {"notified"},
    "notified": set(),
    "failed": set(),
    "cancelled": set(),
}


def canonical_hash(payload: dict[str, Any]) -> str:
    encoded = json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def make_idempotency_key(
    operation_type: str,
    order_id: str,
    payload: dict[str, Any],
) -> str:
    return (
        str(operation_type)
        + ":"
        + str(order_id)
        + ":"
        + canonical_hash(payload)
    )


def begin_operation(
    connection,
    *,
    operation_type: str,
    order_id: str,
    site_id: int | None,
    request_payload: dict[str, Any],
    created_by: str | None,
) -> dict[str, Any]:
    """Create a pending operation or return the matching prior operation."""

    request_hash = canonical_hash(request_payload)
    key = make_idempotency_key(operation_type, order_id, request_payload)
    operation_id = str(uuid.uuid4())
    cursor = connection.execute(
        """
        INSERT INTO external_operations
            (operation_id,operation_type,idempotency_key,order_id,site_id,
             request_hash,request_payload,status,attempts,created_by)
        VALUES (?,?,?,?,?,?,?::jsonb,'pending',1,?)
        ON CONFLICT(idempotency_key) DO NOTHING
        """,
        (
            operation_id,
            operation_type,
            key,
            str(order_id),
            site_id,
            request_hash,
            json.dumps(
                request_payload,
                ensure_ascii=False,
                sort_keys=True,
                separators=(",", ":"),
            ),
            str(created_by or "")[:255],
        ),
    )
    created = int(cursor.rowcount or 0) == 1
    row = connection.execute(
        """
        SELECT operation_id,operation_type,idempotency_key,order_id,site_id,
               request_hash,status,external_reference,attempts,created_at,
               external_succeeded_at,local_committed_at,notified_at,last_error
        FROM external_operations WHERE idempotency_key=? FOR UPDATE
        """,
        (key,),
    ).fetchone()
    resumed = False
    if not created and row and str(row["status"]) == "failed":
        # A prior write was definitively rejected before any external side
        # effect. A later, explicit HTTP submission may safely retry it, but
        # the row lock ensures only one concurrent caller becomes executor.
        # Ambiguous/pending operations are never resumed here.
        connection.execute(
            """
            UPDATE external_operations
            SET status='pending',attempts=attempts+1,last_error=NULL,
                updated_at=CURRENT_TIMESTAMP
            WHERE operation_id=? AND status='failed'
            """,
            (str(row["operation_id"]),),
        )
        resumed = True
        row = connection.execute(
            """
            SELECT operation_id,operation_type,idempotency_key,order_id,site_id,
                   request_hash,status,external_reference,attempts,created_at,
                   external_succeeded_at,local_committed_at,notified_at,last_error
            FROM external_operations WHERE operation_id=?
            """,
            (str(row["operation_id"]),),
        ).fetchone()
    result = dict(row)
    result["operation_id"] = str(result["operation_id"])
    result["created"] = created
    result["resumed"] = resumed
    result["should_execute"] = created or resumed
    return result


def transition_operation(
    connection,
    operation_id: str,
    target_status: str,
    *,
    external_reference: str | None = None,
    evidence: dict[str, Any] | None = None,
    error: str | None = None,
) -> dict[str, Any]:
    row = connection.execute(
        """
        SELECT operation_id,status FROM external_operations
        WHERE operation_id=? FOR UPDATE
        """,
        (operation_id,),
    ).fetchone()
    if not row:
        raise KeyError(operation_id)
    current = str(row["status"])
    if current == target_status:
        return {"operation_id": str(row["operation_id"]), "status": current}
    if target_status not in ALLOWED_TRANSITIONS.get(current, set()):
        raise ValueError(f"invalid external operation transition {current}->{target_status}")
    timestamps = {
        "external_success": "external_succeeded_at=CURRENT_TIMESTAMP,",
        "local_committed": "local_committed_at=CURRENT_TIMESTAMP,",
        "notified": "notified_at=CURRENT_TIMESTAMP,",
    }
    timestamp_sql = timestamps.get(target_status, "")
    connection.execute(
        f"""
        UPDATE external_operations
        SET status=?,{timestamp_sql}
            external_reference=COALESCE(?,external_reference),
            external_evidence=CASE
                WHEN ?::jsonb='{{}}'::jsonb THEN external_evidence
                ELSE ?::jsonb
            END,
            last_error=?,updated_at=CURRENT_TIMESTAMP
        WHERE operation_id=?
        """,
        (
            target_status,
            external_reference,
            json.dumps(evidence or {}, ensure_ascii=False, separators=(",", ":")),
            json.dumps(evidence or {}, ensure_ascii=False, separators=(",", ":")),
            str(error)[:2000] if error else None,
            operation_id,
        ),
    )
    return {"operation_id": operation_id, "status": target_status}


def operation_by_key(
    connection,
    operation_type: str,
    order_id: str,
    request_payload: dict[str, Any],
) -> dict[str, Any] | None:
    key = make_idempotency_key(operation_type, order_id, request_payload)
    row = connection.execute(
        """
        SELECT operation_id,status,external_reference,last_error,updated_at
        FROM external_operations WHERE idempotency_key=?
        """,
        (key,),
    ).fetchone()
    if not row:
        return None
    result = dict(row)
    result["operation_id"] = str(result["operation_id"])
    return result
