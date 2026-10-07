#!/usr/bin/env python3
"""Standalone worker for durable WooCommerce product clone jobs."""

from __future__ import annotations

import argparse
import os
import socket
import db_backend as sqlite3
import time
import traceback
import uuid
from site_connection_service import is_site_archived

from product_clone_jobs import (
    claim_clone_job,
    fail_clone_job,
    finish_clone_job,
    init_product_clone_jobs,
    recover_interrupted_jobs,
    save_clone_progress,
    set_current_product,
)


WORKER_ID = f"{socket.gethostname()}:{os.getpid()}:{uuid.uuid4().hex[:8]}"


def get_connection(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path, timeout=30)
    conn.row_factory = sqlite3.Row
    return conn


def resolve_site_for_worker(conn, site_id: int, *, get_api_endpoint):
    """Resolve credentials after the web route has already authorized the job.

    The interactive resolver reads Flask-Login's ``current_user`` and therefore
    cannot be called from a standalone process.  The queued job is immutable and
    was permission-checked at submission, so the worker only resolves the saved
    site ID and its configured product master here.
    """
    site = conn.execute(
        """SELECT id, url, consumer_key, consumer_secret, product_master_id, manager
           FROM sites WHERE id=?""",
        (site_id,),
    ).fetchone()
    if not site:
        raise ValueError(f"站点 {site_id} 不存在")
    if is_site_archived(conn, site_id):
        raise ValueError("站点连接已删除，未执行克隆")
    api_url, ck, cs = get_api_endpoint(conn, site)
    if not (api_url and ck and cs):
        raise ValueError("站点未配置完整的 WC REST API 凭据")
    return site, api_url, ck, cs


def resolve_catalog_clone_site(conn, site_id, actor_id):
    """New catalog jobs use real site IDs and recheck live manager authority."""
    from product_manager_catalog import _configured
    actor = conn.execute(
        "SELECT id,username,name,role,can_manage_products,can_manage_own_products FROM users WHERE id=?",
        (actor_id,),
    ).fetchone()
    if not actor or (actor["username"] != "admin" and actor["can_manage_products"] != 1):
        raise ValueError("克隆发起人的产品管理权限已撤销")
    site = conn.execute(
        "SELECT id,url,consumer_key,consumer_secret,product_master_id,manager FROM sites WHERE id=?",
        (site_id,),
    ).fetchone()
    if not site:
        raise ValueError("克隆站点不存在")
    if is_site_archived(conn, site_id):
        raise ValueError("站点连接已删除，未执行克隆")
    scoped = actor["username"] != "admin" and (actor["role"] != "admin" or actor["can_manage_own_products"] == 1)
    name = str(actor["name"] or "").strip()
    if scoped and (not name or name != str(site["manager"] or "").strip()):
        raise ValueError("克隆站点负责人已变化，当前发起人没有操作权限")
    if not _configured(site):
        raise ValueError("实际站点直接 API 凭据不完整")
    return site, site["url"].rstrip("/"), site["consumer_key"], site["consumer_secret"]


def process_clone_job(conn, job: dict, *, clone_one, resolve_site) -> dict:
    """Process one claimed job and persist progress after every product."""
    direct = (job.get("options") or {}).get("_catalog_direct_sites") is True
    try:
        if any(is_site_archived(conn, int(job[key])) for key in ("source_site_id", "target_site_id")):
            raise ValueError("站点连接已删除，未执行克隆")
        if direct:
            source = resolve_catalog_clone_site(conn, int(job["source_site_id"]), job["created_by_id"])
            target = resolve_catalog_clone_site(conn, int(job["target_site_id"]), job["created_by_id"])
        else:
            source = resolve_site(conn, int(job["source_site_id"]))
            target = resolve_site(conn, int(job["target_site_id"]))
        _, src_url, src_ck, src_cs = source
        _, tgt_url, tgt_ck, tgt_cs = target
        if direct and (src_url != job["options"].get("_catalog_source_url")
                       or tgt_url != job["options"].get("_catalog_target_url")):
            raise ValueError("克隆站点 URL 已变化，未执行本任务")
    except Exception as exc:
        fail_clone_job(conn, job["id"], f"站点配置读取失败: {exc}")
        return {"success": [], "failed": []}

    results = job.get("results") or {"success": [], "failed": []}
    completed_ids = {
        int(item.get("source_id") or item.get("product_id"))
        for bucket in ("success", "failed")
        for item in (results.get(bucket) or [])
        if item.get("source_id") or item.get("product_id")
    }
    for product_id in job.get("product_ids") or []:
        product_id = int(product_id)
        if product_id in completed_ids:
            continue
        set_current_product(conn, job["id"], product_id)
        try:
            clone_options = dict(job.get("options") or {})
            # Recheck both legacy and direct jobs immediately before every
            # remote write, including images and variants of an existing parent.
            def check_clone_write():
                if any(is_site_archived(conn, int(job[key])) for key in ("source_site_id", "target_site_id")):
                    raise ValueError("克隆执行期间站点连接已删除，未继续写入")
                if direct:
                    fresh_source = resolve_catalog_clone_site(conn, int(job["source_site_id"]), job["created_by_id"])
                    fresh_target = resolve_catalog_clone_site(conn, int(job["target_site_id"]), job["created_by_id"])
                else:
                    fresh_source = resolve_site(conn, int(job["source_site_id"]))
                    fresh_target = resolve_site(conn, int(job["target_site_id"]))
                if fresh_source[1:] != source[1:] or fresh_target[1:] != target[1:]:
                    raise ValueError("克隆执行期间站点 URL 或 API 凭据已变化，未继续写入")
            check_clone_write()
            # A process-local callable never enters the persisted job JSON.
            clone_options["_catalog_write_check"] = check_clone_write
            result = clone_one(
                src_url, src_ck, src_cs, tgt_url, tgt_ck, tgt_cs,
                product_id, clone_options,
            )
        except Exception as exc:
            result = {"error": f"未知错误: {exc}"}
        if direct and result.get("partial_clone"):
            results["failed"].append({"product_id": product_id, "target_id": result.get("new_id"),
                "partial_clone": True, "warnings": result.get("warnings") or [],
                "error": "目标商品已存在，但变体或图片迁移尚未完整；请核对现有目标商品，不要直接重复克隆。"})
        elif result.get("error"):
            results["failed"].append({"product_id": product_id, "error": result["error"]})
        else:
            results["success"].append(
                {
                    "source_id": product_id,
                    "target_id": result["new_id"],
                    "name": result.get("name"),
                    "sku": result.get("sku"),
                    "permalink": result.get("permalink"),
                    "warnings": result.get("warnings") or [],
                    "skipped_existing": bool(result.get("skipped_existing")),
                }
            )
        save_clone_progress(conn, job["id"], results)

    finish_clone_job(conn, job["id"], results)
    return results


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--database", default="woocommerce_orders.db")
    parser.add_argument("--idle-sleep", type=float, default=2.0)
    parser.add_argument("--once", action="store_true")
    args = parser.parse_args()

    conn = get_connection(args.database)
    init_product_clone_jobs(conn)
    interrupted = recover_interrupted_jobs(conn)
    if interrupted:
        print(f"marked {interrupted} interrupted product clone job(s)", flush=True)

    # Imported only after queue recovery so Flask's startup initialization does
    # not hide an interrupted clone from the worker lifecycle.
    import app as app_module

    while True:
        job = claim_clone_job(conn, WORKER_ID)
        if not job:
            if args.once:
                break
            time.sleep(max(0.2, args.idle_sleep))
            continue
        print(f"processing product clone job {job['id']} ({job['total_count']} products)", flush=True)
        try:
            process_clone_job(
                conn,
                job,
                clone_one=app_module._clone_one_product,
                resolve_site=lambda db, site_id: resolve_site_for_worker(
                    db, site_id, get_api_endpoint=app_module.get_product_api_endpoint
                ),
            )
        except Exception as exc:
            fail_clone_job(conn, job["id"], f"worker 异常: {exc}")
            traceback.print_exc()
        if args.once:
            break
    conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
