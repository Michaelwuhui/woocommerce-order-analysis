"""Replan current managed-product shortages into the temporary transit stock.

The command is dry-run by default. ``--apply`` commits the database changes.
It never calls an external service and fails if planning creates a WMS submit
job, which protects the paused Hungary/new-Poland integrations.
"""

from __future__ import annotations

import argparse
import json
import os
import sys


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from fulfillment_service import DomainError, order_contains_managed_product, plan_order
from inv_common import get_conn
from managed_transfer_catalog import JYJG_TRANSIT_WAREHOUSE_CODE


def _external_submit_jobs(conn) -> int:
    return int(conn.execute(
        """SELECT COUNT(*) FROM oms_integration_jobs
           WHERE job_type IN ('SUBMIT_HU_FULFILLMENT','SUBMIT_PL_FULFILLMENT')"""
    ).fetchone()[0])


def run(*, apply: bool) -> dict:
    conn = get_conn()
    report = {
        "mode": "apply" if apply else "dry_run",
        "planned": [],
        "noop": [],
        "skipped": [],
        "errors": [],
    }
    try:
        warehouse = conn.execute(
            "SELECT id,name,is_active FROM warehouses WHERE code=?",
            (JYJG_TRANSIT_WAREHOUSE_CODE,),
        ).fetchone()
        if not warehouse or not warehouse["is_active"]:
            raise RuntimeError("金毅金谷临时中转仓未启用，请先执行迁移017")
        # Establish an outer transaction before per-order SAVEPOINTs. Without
        # this, releasing the first savepoint would commit a nominal dry-run.
        conn.execute("BEGIN")
        before_jobs = _external_submit_jobs(conn)
        orders = conn.execute(
            """SELECT o.id,o.number,o.status,o.source,
                      COALESCE(s.country,'') AS market,
                      COALESCE(fs.has_shortage,0) AS has_shortage,
                      fs.aggregate_status
               FROM orders o
               JOIN sites s ON s.url=o.source
               JOIN oms_order_fulfillment_state fs ON fs.order_id=o.id
               WHERE o.status IN ('processing','offline','on-hold')
                 AND fs.has_shortage=1
               ORDER BY o.date_created,o.id"""
        ).fetchall()
        for order in orders:
            if not order_contains_managed_product(conn, order["id"]):
                continue
            conn.execute("SAVEPOINT replan_one")
            try:
                result = plan_order(
                    conn,
                    order["id"],
                    actor={"type": "system", "name": "JYJG临时中转仓上线"},
                    commit=False,
                )
                allocations = []
                for assignment in result.get("assignments") or []:
                    wh = conn.execute(
                        "SELECT code,name FROM warehouses WHERE id=?",
                        (assignment["warehouse_id"],),
                    ).fetchone()
                    allocations.append({
                        "warehouse_code": wh["code"] if wh else None,
                        "warehouse_name": wh["name"] if wh else None,
                        "quantity": sum(int(line["qty"]) for line in assignment["lines"]),
                    })
                entry = {
                    "order_id": order["id"],
                    "order_number": order["number"],
                    "market": order["market"],
                    "action": result["action"],
                    "aggregate_status": result.get("aggregate_status")
                    or ("stock_shortage" if result.get("shortages") else "allocated"),
                    "allocations": allocations,
                    "shortages": result.get("shortages") or [],
                }
                report["noop" if result["action"] == "noop" else "planned"].append(entry)
                conn.execute("RELEASE SAVEPOINT replan_one")
            except DomainError as exc:
                conn.execute("ROLLBACK TO SAVEPOINT replan_one")
                conn.execute("RELEASE SAVEPOINT replan_one")
                target = "skipped" if exc.code == "fulfillment_already_started" else "errors"
                report[target].append({
                    "order_id": order["id"],
                    "order_number": order["number"],
                    "market": order["market"],
                    "code": exc.code,
                    "error": str(exc),
                })

        after_jobs = _external_submit_jobs(conn)
        if after_jobs != before_jobs:
            raise RuntimeError("受控重分仓意外生成了外部WMS提交任务，已阻止提交")
        stock = conn.execute(
            """SELECT COUNT(*) AS sku_count,
                      COALESCE(SUM(st.on_hand),0) AS on_hand,
                      COALESCE(SUM(st.reserved),0) AS reserved,
                      COALESCE(SUM(st.on_hand-st.reserved),0) AS available
               FROM inv_stock st WHERE st.warehouse_id=?""",
            (warehouse["id"],),
        ).fetchone()
        report["inventory"] = dict(stock)
        report["external_submit_jobs_created"] = after_jobs - before_jobs
        if report["errors"]:
            conn.rollback()
            report["committed"] = False
        elif apply:
            conn.commit()
            report["committed"] = True
        else:
            conn.rollback()
            report["committed"] = False
        return report
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--apply", action="store_true", help="commit replanning changes")
    args = parser.parse_args()
    print(json.dumps(run(apply=args.apply), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
