"""WooCommerce inventory publishing with per-site, fail-closed automation.

The inventory ledger remains authoritative. Automatic publishing is disabled
globally and per site until an administrator explicitly enables it. Observe
mode performs the same calculation and audit path without writing WooCommerce.
"""

import json
import db_backend as sqlite3
import uuid
from functools import wraps

from flask import Blueprint, jsonify, render_template, request
from flask_login import current_user, login_required

import inv_allocator
from inv_common import (
    _deny,
    can_manage_inventory,
    can_view_inventory,
    current_operator,
    get_conn,
    inv_admin_required,
)


inv_push_bp = Blueprint("inv_push", __name__)

SYNC_MODES = {"off", "observe", "live", "paused"}
ALLOCATION_STRATEGIES = {"quota", "mirror"}
DEFAULT_SYNC_CONFIG = {
    "mode": "off",
    "interval_minutes": 15,
    "allocation_strategy": "quota",
    "allocation_weight": 1,
    "safety_stock": 0,
    "failure_threshold": 3,
    "consecutive_failures": 0,
    "last_attempt_at": None,
    "last_success_at": None,
    "next_run_at": None,
    "last_error": None,
    "paused_reason": None,
}


def _table_exists(conn, name):
    return bool(
        conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)
        ).fetchone()
    )


def _table_columns(conn, name):
    if not _table_exists(conn, name):
        return set()
    return {row["name"] for row in conn.execute(f"PRAGMA table_info({name})").fetchall()}


def _setting(conn, key, default=None):
    if not _table_exists(conn, "settings"):
        return default
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row["value"] if row and row["value"] is not None else default


def _setting_int(conn, key, default):
    try:
        return int(_setting(conn, key, default))
    except (TypeError, ValueError):
        return int(default)


def global_sync_enabled(conn):
    return str(_setting(conn, "inv_auto_push_global_enabled", "0")).lower() in {
        "1", "true", "yes", "on",
    }


def get_site_sync_config(conn, site_id):
    config = dict(DEFAULT_SYNC_CONFIG)
    config["site_id"] = int(site_id)
    if not _table_exists(conn, "inv_site_sync_config"):
        return config
    row = conn.execute(
        "SELECT * FROM inv_site_sync_config WHERE site_id=?", (site_id,)
    ).fetchone()
    if row:
        config.update(dict(row))
    return config


# ---------------------------------------------------------------------------
# Stock calculation
# ---------------------------------------------------------------------------


def _serving_warehouses(conn, market):
    """Return quantity-authoritative serving warehouses for a market."""
    cands = inv_allocator.candidate_warehouses(conn, market)
    ids = [int(c["warehouse_id"]) for c in cands]
    if not ids:
        ids = [
            int(row["id"])
            for row in conn.execute(
                "SELECT id FROM warehouses WHERE country=? AND is_active=1", (market,)
            ).fetchall()
        ]
    if not ids:
        return []
    marks = ",".join("?" * len(ids))
    rows = conn.execute(
        f"""SELECT w.id,COALESCE(wi.inventory_authority,'local') AS inventory_authority,
                   COALESCE(wi.config_json,'{{}}') AS config_json
            FROM warehouses w LEFT JOIN oms_warehouse_integrations wi ON wi.warehouse_id=w.id
            WHERE w.id IN ({marks}) AND w.is_active=1""",
        ids,
    ).fetchall()
    result = []
    for row in rows:
        try:
            config = json.loads(row["config_json"] or "{}")
        except (TypeError, ValueError):
            config = {}
        authority = row["inventory_authority"] or "local"
        if authority == "manual_partner" or config.get("requires_quantity_inventory") is False:
            continue
        if authority in ("local", "external_wms"):
            result.append({"warehouse_id": int(row["id"]), "authority": authority})
    return sorted(result, key=lambda row: row["warehouse_id"])


def _available_sku(conn, sku_id, warehouses):
    available = 0
    local_ids = [row["warehouse_id"] for row in warehouses if row["authority"] == "local"]
    external_ids = [
        row["warehouse_id"] for row in warehouses if row["authority"] == "external_wms"
    ]
    if local_ids:
        marks = ",".join("?" * len(local_ids))
        row = conn.execute(
            f"""SELECT COALESCE(SUM(
                         CASE WHEN st.on_hand - st.reserved > 0
                              THEN st.on_hand - st.reserved ELSE 0 END
                       ),0) AS available
                FROM inv_stock st JOIN oms_sku_warehouses sw
                  ON sw.warehouse_id=st.warehouse_id
                 AND sw.sku_id=st.sku_id AND sw.is_enabled=1
                WHERE st.sku_id=? AND st.warehouse_id IN ({marks})""",
            [sku_id] + local_ids,
        ).fetchone()
        available += max(0, int(row["available"] or 0))
    if external_ids:
        marks = ",".join("?" * len(external_ids))
        row = conn.execute(
            f"""SELECT COALESCE(SUM(
                         CASE WHEN available_quantity > 0
                              THEN available_quantity ELSE 0 END
                       ),0) AS available
                FROM oms_external_stock WHERE sku_id=? AND warehouse_id IN ({marks})""",
            [sku_id] + external_ids,
        ).fetchone()
        available += max(0, int(row["available"] or 0))
    return available


def _weighted_share(total, participants, site_id):
    """Distribute integer SKU units deterministically using largest remainder."""
    participants = sorted(participants, key=lambda row: int(row["site_id"]))
    weights = {int(row["site_id"]): max(1, int(row["weight"] or 1)) for row in participants}
    weight_total = sum(weights.values())
    if total <= 0 or not weight_total or int(site_id) not in weights:
        return 0
    shares = {sid: (total * weight) // weight_total for sid, weight in weights.items()}
    remainder = total - sum(shares.values())
    ranked = sorted(
        weights,
        key=lambda sid: (-((total * weights[sid]) % weight_total), sid),
    )
    for sid in ranked[:remainder]:
        shares[sid] += 1
    return shares[int(site_id)]


def _quota_participants(conn, sku_id, warehouse_ids, country_cache):
    rows = conn.execute(
        """SELECT DISTINCT s.id AS site_id,s.country,
                  COALESCE(c.allocation_weight,1) AS weight,
                  COALESCE(c.safety_stock,0) AS safety_stock,
                  COALESCE(c.mode,'off') AS mode,
                  COALESCE(c.allocation_strategy,'quota') AS allocation_strategy
           FROM inv_site_sku_map m JOIN sites s ON s.id=m.site_id
           LEFT JOIN inv_site_sync_config c ON c.site_id=s.id
           WHERE m.sku_id=? AND m.is_active=1""",
        (sku_id,),
    ).fetchall()
    result = []
    target = tuple(sorted(warehouse_ids))
    for row in rows:
        market = (row["country"] or "").upper()
        if market not in country_cache:
            country_cache[market] = tuple(
                item["warehouse_id"] for item in _serving_warehouses(conn, market)
            )
        if country_cache[market] == target:
            result.append(dict(row))
    return result


def compute_site_stock(conn, site_id, use_sync_strategy=False):
    """Calculate publishable stock for every active mapping of one site.

    Legacy/manual calculations retain mirrored stock. Automatic/observe runs
    pass use_sync_strategy=True and use the configured shared-stock quota.
    """
    site = conn.execute("SELECT id,url,country FROM sites WHERE id=?", (site_id,)).fetchone()
    if not site:
        return []
    market = (site["country"] or "").upper()
    warehouses = _serving_warehouses(conn, market)
    warehouse_ids = [row["warehouse_id"] for row in warehouses]
    maps = conn.execute(
        """SELECT m.*,k.sku_code,k.name AS sku_name
           FROM inv_site_sku_map m JOIN inv_skus k ON k.id=m.sku_id
           WHERE m.site_id=? AND m.is_active=1""",
        (site_id,),
    ).fetchall()
    config = get_site_sync_config(conn, site_id)
    use_quota = use_sync_strategy and config["allocation_strategy"] == "quota"
    country_cache = {market: tuple(warehouse_ids)}
    out = []
    for mapping in maps:
        pool_available = _available_sku(conn, mapping["sku_id"], warehouses)
        allocated = pool_available
        participant_count = 1
        if use_quota and _table_exists(conn, "inv_site_sync_config"):
            participants = _quota_participants(
                conn, mapping["sku_id"], warehouse_ids, country_cache
            )
            participant_count = len(participants) or 1
            safety_stock = max(
                [int(row.get("safety_stock") or 0) for row in participants]
                or [int(config["safety_stock"] or 0)]
            )
            distributable = max(0, pool_available - safety_stock)
            allocated = _weighted_share(distributable, participants, site_id)
        qty_per_item = int(mapping["qty_per_item"] or 1)
        out.append(
            {
                "site_id": int(site_id),
                "source": site["url"],
                "market": market,
                "wc_product_id": mapping["wc_product_id"],
                "wc_variation_id": mapping["wc_variation_id"] or 0,
                "sku_id": mapping["sku_id"],
                "sku_code": mapping["sku_code"],
                "sku_name": mapping["sku_name"],
                "qty_per_item": qty_per_item,
                "available_sku": pool_available,
                "allocated_sku": allocated,
                "publishable": allocated // qty_per_item,
                "allocation_strategy": "quota" if use_quota else "mirror",
                "allocation_participants": participant_count,
                "serving_warehouses": warehouse_ids,
            }
        )
    return out


def _sku_has_quantity_source(conn, sku_id, warehouses):
    for warehouse in warehouses:
        warehouse_id = warehouse["warehouse_id"]
        if warehouse["authority"] == "local":
            row = conn.execute(
                """SELECT 1 FROM oms_sku_warehouses sw JOIN inv_stock st
                     ON st.warehouse_id=sw.warehouse_id AND st.sku_id=sw.sku_id
                   WHERE sw.warehouse_id=? AND sw.sku_id=? AND sw.is_enabled=1""",
                (warehouse_id, sku_id),
            ).fetchone()
        else:
            row = conn.execute(
                "SELECT 1 FROM oms_external_stock WHERE warehouse_id=? AND sku_id=?",
                (warehouse_id, sku_id),
            ).fetchone()
        if row:
            return True
    return False


def site_sync_readiness(conn, site_id, for_live=False):
    site = conn.execute("SELECT id,url,country FROM sites WHERE id=?", (site_id,)).fetchone()
    errors, warnings = [], []
    if not site:
        return {"ready": False, "errors": ["站点不存在"], "warnings": []}
    market = (site["country"] or "").upper()
    warehouses = _serving_warehouses(conn, market)
    maps = conn.execute(
        "SELECT sku_id FROM inv_site_sku_map WHERE site_id=? AND is_active=1",
        (site_id,),
    ).fetchall()
    if not maps:
        errors.append("没有已启用的 SKU 映射")
    if not warehouses:
        errors.append("该市场没有维护真实数量的服务仓")
    missing_source = [
        int(row["sku_id"])
        for row in maps
        if warehouses and not _sku_has_quantity_source(conn, row["sku_id"], warehouses)
    ]
    if missing_source:
        errors.append(f"{len(missing_source)} 个映射 SKU 没有库存数据源")

    external_ids = [
        row["warehouse_id"] for row in warehouses if row["authority"] == "external_wms"
    ]
    external_columns = _table_columns(conn, "oms_external_stock")
    if external_ids and {"source_updated_at", "synced_at"}.issubset(external_columns):
        freshness = max(
            5, _setting_int(conn, "inv_auto_push_external_freshness_minutes", 180)
        )
        marks = ",".join("?" * len(external_ids))
        stale = conn.execute(
            f"""SELECT COUNT(*) AS n FROM oms_external_stock
                WHERE warehouse_id IN ({marks}) AND sku_id IN
                    (SELECT sku_id FROM inv_site_sku_map WHERE site_id=? AND is_active=1)
                  AND (COALESCE(source_updated_at,synced_at) IS NULL
                       OR datetime(COALESCE(source_updated_at,synced_at)) < datetime('now',?))""",
            external_ids + [site_id, f"-{freshness} minutes"],
        ).fetchone()["n"]
        if stale:
            errors.append(f"{stale} 条外部 WMS 库存超过 {freshness} 分钟未更新")

    unmapped = 0
    if _table_exists(conn, "inv_site_product_catalog"):
        unmapped = conn.execute(
            """SELECT COUNT(*) AS n FROM inv_site_product_catalog c
               WHERE c.site_id=? AND c.candidate_sku_id IS NOT NULL
                 AND NOT EXISTS (
                    SELECT 1 FROM inv_site_sku_map m
                    WHERE m.site_id=c.site_id AND m.wc_product_id=c.wc_product_id
                      AND COALESCE(m.wc_variation_id,0)=COALESCE(c.wc_variation_id,0)
                      AND m.is_active=1
                 )""",
            (site_id,),
        ).fetchone()["n"]
        if unmapped:
            warnings.append(f"{unmapped} 个目录候选商品尚未确认映射")

    config = get_site_sync_config(conn, site_id)
    shared_not_live = []
    shared_non_quota = []
    shared_site_ids = []
    if (
        maps
        and warehouses
        and config["allocation_strategy"] == "quota"
        and _table_exists(conn, "inv_site_sync_config")
    ):
        target_warehouses = [row["warehouse_id"] for row in warehouses]
        cache = {market: tuple(target_warehouses)}
        for mapping in maps:
            for participant in _quota_participants(
                conn, mapping["sku_id"], target_warehouses, cache
            ):
                participant_id = int(participant["site_id"])
                if participant_id == int(site_id):
                    continue
                shared_site_ids.append(participant_id)
                if participant["mode"] != "live":
                    shared_not_live.append(participant_id)
                if participant["allocation_strategy"] != "quota":
                    shared_non_quota.append(participant_id)
        shared_not_live = sorted(set(shared_not_live))
        shared_non_quota = sorted(set(shared_non_quota))
        shared_site_ids = sorted(set(shared_site_ids))
        if shared_not_live:
            message = f"共享同一库存池的 {len(shared_not_live)} 个站点尚未进入正式同步"
            if for_live:
                errors.append(message)
            else:
                warnings.append(message)
        if shared_non_quota:
            message = f"共享库存池中有 {len(shared_non_quota)} 个站点未使用权重配额"
            if for_live:
                errors.append(message)
            else:
                warnings.append(message)
    elif maps and warehouses and config["allocation_strategy"] == "mirror":
        target_warehouses = [row["warehouse_id"] for row in warehouses]
        cache = {market: tuple(target_warehouses)}
        for mapping in maps:
            shared_site_ids.extend(
                int(participant["site_id"])
                for participant in _quota_participants(
                    conn, mapping["sku_id"], target_warehouses, cache
                )
                if int(participant["site_id"]) != int(site_id)
            )
        shared_site_ids = sorted(set(shared_site_ids))
        if shared_site_ids:
            message = "镜像库存不能用于多站点共享库存池，请改用按权重配额"
            if for_live:
                errors.append(message)
            else:
                warnings.append(message)

    return {
        "ready": not errors,
        "errors": errors,
        "warnings": warnings,
        "mapping_count": len(maps),
        "warehouse_ids": [row["warehouse_id"] for row in warehouses],
        "missing_source_count": len(missing_source),
        "unmapped_candidate_count": int(unmapped or 0),
        "shared_not_live_site_ids": shared_not_live,
        "shared_non_quota_site_ids": shared_non_quota,
        "shared_site_ids": shared_site_ids,
    }


# ---------------------------------------------------------------------------
# WooCommerce write and verification
# ---------------------------------------------------------------------------


def _resource_url(api_url, product_id, variation_id):
    resource = f"{api_url.rstrip('/')}/wp-json/wc/v3/products/{product_id}"
    if variation_id:
        resource += f"/variations/{variation_id}"
    return resource


def _get_stock_state(api_url, ck, cs, product_id, variation_id):
    import requests as req
    from app import _parse_wc_response

    try:
        response = req.get(
            _resource_url(api_url, product_id, variation_id),
            auth=(ck, cs),
            timeout=45,
            headers={
                "User-Agent": "WooCommerce API Client-Python/3.0.0",
                "Accept": "application/json",
            },
        )
    except req.RequestException as exc:
        return None, f"读取 Woo 库存失败: {exc}"
    payload, error = _parse_wc_response(response)
    if error:
        return None, error
    try:
        quantity = int(payload.get("stock_quantity") or 0)
    except (TypeError, ValueError):
        quantity = 0
    return {"manage_stock": bool(payload.get("manage_stock")), "stock_quantity": quantity}, None


def _put_stock(api_url, ck, cs, product_id, variation_id, qty):
    import requests as req
    from app import _build_product_update_payload, _parse_wc_response

    payload, error = _build_product_update_payload(
        {"manage_stock": True, "stock_quantity": qty}
    )
    if error:
        return False, error
    try:
        response = req.put(
            _resource_url(api_url, product_id, variation_id),
            auth=(ck, cs),
            json=payload,
            timeout=90,
            headers={
                "User-Agent": "WooCommerce API Client-Python/3.0.0",
                "Content-Type": "application/json",
                "Accept": "application/json",
            },
        )
    except req.RequestException as exc:
        return False, f"写入 Woo 库存失败: {exc}"
    _payload, error = _parse_wc_response(response)
    return error is None, error


def _sync_one_stock(api_url, ck, cs, item, only_changed=True):
    before, error = _get_stock_state(
        api_url, ck, cs, item["wc_product_id"], item["wc_variation_id"]
    )
    if error:
        return "error", None, None, error
    desired = int(item["publishable"])
    previous = int(before["stock_quantity"])
    if only_changed and before["manage_stock"] and previous == desired:
        return "unchanged", previous, previous, None

    write_ok, write_error = _put_stock(
        api_url, ck, cs, item["wc_product_id"], item["wc_variation_id"], desired
    )
    after, readback_error = _get_stock_state(
        api_url, ck, cs, item["wc_product_id"], item["wc_variation_id"]
    )
    if after and after["manage_stock"] and int(after["stock_quantity"]) == desired:
        return "ok", previous, desired, None
    if readback_error:
        detail = f"{write_error or '写入结果未知'}；回读失败: {readback_error}"
    elif write_ok:
        detail = f"写入后回读不一致，期望 {desired}，实际 {after['stock_quantity']}"
    else:
        detail = write_error or "Woo 库存写入失败"
    return "error", previous, after["stock_quantity"] if after else None, detail


def _latest_pushed_quantities(conn, site_id):
    rows = conn.execute(
        """SELECT l.wc_product_id,COALESCE(l.wc_variation_id,0) AS wc_variation_id,
                  l.pushed_qty
           FROM inv_push_logs l JOIN (
               SELECT wc_product_id,COALESCE(wc_variation_id,0) AS variation_id,MAX(id) AS max_id
               FROM inv_push_logs
               WHERE site_id=? AND status IN ('ok','unchanged')
               GROUP BY wc_product_id,COALESCE(wc_variation_id,0)
           ) latest ON latest.max_id=l.id""",
        (site_id,),
    ).fetchall()
    return {
        (int(row["wc_product_id"]), int(row["wc_variation_id"])): int(
            row["pushed_qty"] or 0
        )
        for row in rows
    }


def detect_mass_drop(conn, site_id, items):
    previous = _latest_pushed_quantities(conn, site_id)
    comparable = []
    zeroed = 0
    large_drops = 0
    max_drop = max(1, min(100, _setting_int(conn, "inv_auto_push_max_drop_percent", 80)))
    for item in items:
        key = (int(item["wc_product_id"]), int(item["wc_variation_id"] or 0))
        old = previous.get(key)
        if old is None or old <= 0:
            continue
        new = int(item["publishable"])
        comparable.append((old, new))
        if new == 0:
            zeroed += 1
        if new <= old * (100 - max_drop) / 100:
            large_drops += 1
    if not comparable:
        return None
    zero_guard = max(
        1, min(100, _setting_int(conn, "inv_auto_push_zero_guard_percent", 50))
    )
    if zeroed and zeroed * 100 >= len(comparable) * zero_guard:
        return f"安全拦截: {zeroed}/{len(comparable)} 个已同步商品将变为 0 库存"
    if len(comparable) >= 4 and large_drops * 100 >= len(comparable) * 50:
        return (
            f"安全拦截: {large_drops}/{len(comparable)} 个商品库存将下降 "
            f"{max_drop}% 以上"
        )
    return None


def _site_api_credentials(conn, site):
    from app import get_product_api_endpoint

    return get_product_api_endpoint(conn, site)


def push_site(
    conn,
    site_id,
    dry_run=True,
    only_changed=True,
    *,
    use_sync_strategy=False,
    operator=None,
    enforce_safety=True,
):
    """Calculate and optionally publish one site, logging every mapped item."""
    operator_id, operator_name = operator or current_operator()
    site = conn.execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone()
    items = compute_site_stock(conn, site_id, use_sync_strategy=use_sync_strategy)
    result = {
        "site_id": int(site_id),
        "source": site["url"] if site else None,
        "dry_run": bool(dry_run),
        "total": len(items),
        "ok": 0,
        "unchanged": 0,
        "dry": 0,
        "error": 0,
        "items": [],
    }
    if not site:
        result["fatal"] = "站点不存在"
        return result

    readiness = site_sync_readiness(conn, site_id, for_live=not dry_run)
    result["readiness"] = readiness
    if not dry_run and not readiness["ready"]:
        result["fatal"] = "；".join(readiness["errors"])
        result["error"] = len(items)
        return result
    if not dry_run and enforce_safety:
        safety_error = detect_mass_drop(conn, site_id, items)
        if safety_error:
            result["fatal"] = safety_error
            result["error"] = len(items)
            return result

    api_url = consumer_key = consumer_secret = None
    if not dry_run:
        api_url, consumer_key, consumer_secret = _site_api_credentials(conn, site)
        if not (api_url and consumer_key and consumer_secret):
            result["fatal"] = "站点缺少可用的商品 API 凭据"
            result["error"] = len(items)
            return result

    for item in items:
        previous = remote = error = None
        if dry_run:
            status = "dry"
            result["dry"] += 1
        else:
            status, previous, remote, error = _sync_one_stock(
                api_url, consumer_key, consumer_secret, item, only_changed=only_changed
            )
            if status == "ok":
                result["ok"] += 1
            elif status == "unchanged":
                result["unchanged"] += 1
            else:
                result["error"] += 1
        conn.execute(
            """INSERT INTO inv_push_logs
               (site_id,source,wc_product_id,wc_variation_id,sku_id,prev_qty,
                pushed_qty,status,error,operator_id,operator_name)
               VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
            (
                site_id, item["source"], item["wc_product_id"], item["wc_variation_id"],
                item["sku_id"], previous, item["publishable"], status, error,
                operator_id, operator_name,
            ),
        )
        result["items"].append(
            {
                "wc_product_id": item["wc_product_id"],
                "sku_code": item["sku_code"],
                "publishable": item["publishable"],
                "status": status,
                "previous": previous,
                "remote": remote,
                "error": error,
            }
        )
    conn.commit()
    return result


# ---------------------------------------------------------------------------
# Scheduler, locks, circuit breaker and configuration audit
# ---------------------------------------------------------------------------


def _acquire_site_lock(conn, site_id):
    token = uuid.uuid4().hex
    conn.execute(
        "DELETE FROM inv_push_locks WHERE acquired_at < datetime('now','-30 minutes')"
    )
    try:
        conn.execute(
            "INSERT INTO inv_push_locks (site_id,lock_token) VALUES (?,?)",
            (site_id, token),
        )
        conn.commit()
        return token
    except sqlite3.IntegrityError:
        conn.rollback()
        return None


def _release_site_lock(conn, site_id, token):
    conn.execute(
        "DELETE FROM inv_push_locks WHERE site_id=? AND lock_token=?", (site_id, token)
    )
    conn.commit()


def _next_run_sql(conn, site_id):
    conn.execute(
        """UPDATE inv_site_sync_config
           SET next_run_at=datetime('now','+' || interval_minutes || ' minutes'),
               updated_at=CURRENT_TIMESTAMP
           WHERE site_id=?""",
        (site_id,),
    )


def execute_site_sync(
    conn,
    site_id,
    *,
    trigger_type="scheduler",
    force_dry_run=None,
    operator=None,
):
    """Run one configured site with locking, run audit and failure suspension."""
    config = get_site_sync_config(conn, site_id)
    if trigger_type == "scheduler":
        if not global_sync_enabled(conn):
            return {"site_id": site_id, "status": "skipped", "reason": "全局自动同步已关闭"}
        if config["mode"] not in ("observe", "live"):
            return {"site_id": site_id, "status": "skipped", "reason": "站点未启用"}
        due = conn.execute(
            """SELECT 1 FROM inv_site_sync_config
               WHERE site_id=? AND (next_run_at IS NULL OR datetime(next_run_at)<=datetime('now'))""",
            (site_id,),
        ).fetchone()
        if not due:
            return {"site_id": site_id, "status": "skipped", "reason": "尚未到执行时间"}

    dry_run = config["mode"] != "live"
    if force_dry_run is not None:
        dry_run = bool(force_dry_run)
    operator_id, operator_name = operator or current_operator()
    if not operator_name and trigger_type == "scheduler":
        operator_name = "system:auto_inventory_push"

    token = _acquire_site_lock(conn, site_id)
    if not token:
        return {"site_id": site_id, "status": "skipped", "reason": "该站点已有同步任务运行中"}
    run_id = None
    try:
        cursor = conn.execute(
            """INSERT INTO inv_push_runs
               (site_id,trigger_type,configured_mode,dry_run,operator_id,operator_name)
               VALUES (?,?,?,?,?,?)""",
            (
                site_id, trigger_type, config["mode"], 1 if dry_run else 0,
                operator_id, operator_name,
            ),
        )
        run_id = cursor.lastrowid
        conn.commit()
        result = push_site(
            conn,
            site_id,
            dry_run=dry_run,
            only_changed=True,
            use_sync_strategy=True,
            operator=(operator_id, operator_name),
        )
        if result.get("fatal") or result["error"]:
            run_status = "partial" if result["ok"] or result["unchanged"] else "error"
        else:
            run_status = "success"
        conn.execute(
            """UPDATE inv_push_runs SET status=?,total_count=?,ok_count=?,
                      unchanged_count=?,error_count=?,fatal_error=?,finished_at=CURRENT_TIMESTAMP
               WHERE id=?""",
            (
                run_status, result["total"], result["ok"], result["unchanged"],
                result["error"], result.get("fatal"), run_id,
            ),
        )
        if _table_exists(conn, "inv_site_sync_config"):
            if run_status == "success":
                conn.execute(
                    """UPDATE inv_site_sync_config
                       SET consecutive_failures=0,last_attempt_at=CURRENT_TIMESTAMP,
                           last_success_at=CURRENT_TIMESTAMP,last_error=NULL,paused_reason=NULL
                       WHERE site_id=?""",
                    (site_id,),
                )
            else:
                error_text = result.get("fatal") or f"{result['error']} 个商品同步失败"
                conn.execute(
                    """UPDATE inv_site_sync_config
                       SET consecutive_failures=consecutive_failures+1,
                           last_attempt_at=CURRENT_TIMESTAMP,last_error=?
                       WHERE site_id=?""",
                    (error_text, site_id),
                )
                failed = conn.execute(
                    "SELECT consecutive_failures,failure_threshold FROM inv_site_sync_config WHERE site_id=?",
                    (site_id,),
                ).fetchone()
                if failed and failed["consecutive_failures"] >= failed["failure_threshold"]:
                    conn.execute(
                        """UPDATE inv_site_sync_config
                           SET mode='paused',paused_reason=?,updated_at=CURRENT_TIMESTAMP
                           WHERE site_id=?""",
                        (f"连续 {failed['consecutive_failures']} 次失败，系统自动暂停", site_id),
                    )
            _next_run_sql(conn, site_id)
        conn.commit()
        result["status"] = run_status
        result["run_id"] = run_id
        return result
    except Exception as exc:
        conn.rollback()
        if run_id:
            conn.execute(
                """UPDATE inv_push_runs SET status='error',fatal_error=?,finished_at=CURRENT_TIMESTAMP
                   WHERE id=?""",
                (str(exc), run_id),
            )
        if _table_exists(conn, "inv_site_sync_config"):
            conn.execute(
                """UPDATE inv_site_sync_config
                   SET consecutive_failures=consecutive_failures+1,
                       last_attempt_at=CURRENT_TIMESTAMP,last_error=?
                   WHERE site_id=?""",
                (str(exc), site_id),
            )
            failed = conn.execute(
                "SELECT consecutive_failures,failure_threshold FROM inv_site_sync_config WHERE site_id=?",
                (site_id,),
            ).fetchone()
            if failed and failed["consecutive_failures"] >= failed["failure_threshold"]:
                conn.execute(
                    """UPDATE inv_site_sync_config
                       SET mode='paused',paused_reason=?,updated_at=CURRENT_TIMESTAMP
                       WHERE site_id=?""",
                    (f"连续 {failed['consecutive_failures']} 次异常，系统自动暂停", site_id),
                )
            _next_run_sql(conn, site_id)
        conn.commit()
        raise
    finally:
        _release_site_lock(conn, site_id, token)


def scheduler_site_ids(conn):
    if not _table_exists(conn, "inv_site_sync_config") or not global_sync_enabled(conn):
        return []
    return [
        int(row["site_id"])
        for row in conn.execute(
            """SELECT site_id FROM inv_site_sync_config
               WHERE mode IN ('observe','live')
                 AND (next_run_at IS NULL OR datetime(next_run_at)<=datetime('now'))
               ORDER BY site_id"""
        ).fetchall()
    ]


def _audit_config(conn, site_id, action, before, after, operator):
    operator_id, operator_name = operator
    conn.execute(
        """INSERT INTO inv_site_sync_audit
           (site_id,action,before_json,after_json,operator_id,operator_name)
           VALUES (?,?,?,?,?,?)""",
        (
            site_id, action,
            json.dumps(before, ensure_ascii=False, sort_keys=True) if before is not None else None,
            json.dumps(after, ensure_ascii=False, sort_keys=True) if after is not None else None,
            operator_id, operator_name,
        ),
    )


def update_site_sync_config(conn, site_id, changes, operator=None):
    site = conn.execute("SELECT 1 FROM sites WHERE id=?", (site_id,)).fetchone()
    if not site:
        raise ValueError("站点不存在")
    before = get_site_sync_config(conn, site_id)
    mode = str(changes.get("mode", before["mode"])).strip().lower()
    strategy = str(
        changes.get("allocation_strategy", before["allocation_strategy"])
    ).strip().lower()
    if mode not in SYNC_MODES:
        raise ValueError("无效的同步模式")
    if strategy not in ALLOCATION_STRATEGIES:
        raise ValueError("无效的库存分配策略")
    try:
        interval = max(
            5, min(1440, int(changes.get("interval_minutes", before["interval_minutes"])))
        )
        weight = max(
            1, min(1000, int(changes.get("allocation_weight", before["allocation_weight"])))
        )
        safety = max(0, int(changes.get("safety_stock", before["safety_stock"])))
        threshold = max(
            1, min(20, int(changes.get("failure_threshold", before["failure_threshold"])))
        )
    except (TypeError, ValueError):
        raise ValueError("间隔、权重、安全库存和失败阈值必须是整数")
    operator = operator or current_operator()
    reset_failures = mode in ("off", "observe", "live") and before["mode"] == "paused"
    conn.execute(
        """INSERT INTO inv_site_sync_config
           (site_id,mode,interval_minutes,allocation_strategy,allocation_weight,
            safety_stock,failure_threshold,consecutive_failures,next_run_at,
            last_error,paused_reason,updated_by,updated_by_name)
           VALUES (?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP,?,?,?,?)
           ON CONFLICT(site_id) DO UPDATE SET
             mode=excluded.mode,interval_minutes=excluded.interval_minutes,
             allocation_strategy=excluded.allocation_strategy,
             allocation_weight=excluded.allocation_weight,
             safety_stock=excluded.safety_stock,failure_threshold=excluded.failure_threshold,
             consecutive_failures=excluded.consecutive_failures,
             next_run_at=CURRENT_TIMESTAMP,last_error=excluded.last_error,
             paused_reason=excluded.paused_reason,updated_by=excluded.updated_by,
             updated_by_name=excluded.updated_by_name,updated_at=CURRENT_TIMESTAMP""",
        (
            site_id, mode, interval, strategy, weight, safety, threshold,
            0 if reset_failures else int(before["consecutive_failures"] or 0),
            None if reset_failures else before["last_error"],
            None if reset_failures else before["paused_reason"],
            operator[0], operator[1],
        ),
    )
    after = get_site_sync_config(conn, site_id)
    _audit_config(conn, site_id, "update_site_config", before, after, operator)
    conn.commit()
    return after


# ---------------------------------------------------------------------------
# Site-scoped permissions and API
# ---------------------------------------------------------------------------


def _is_superadmin():
    return getattr(current_user, "username", None) == "admin"


def _has_global_site_scope():
    return _is_superadmin() or getattr(current_user, "role", None) in ("admin", "viewer")


def _scoped_site_ids(conn):
    if _has_global_site_scope():
        return None
    user_id = current_user.id
    explicit = {
        int(row["site_id"])
        for row in conn.execute(
            "SELECT site_id FROM user_site_permissions WHERE user_id=?", (user_id,)
        ).fetchall()
    }
    countries = {
        row["country"]
        for row in conn.execute(
            "SELECT country FROM user_country_permissions WHERE user_id=?", (user_id,)
        ).fetchall()
    }
    excluded = {
        int(row["site_id"])
        for row in conn.execute(
            "SELECT site_id FROM user_site_exclusions WHERE user_id=?", (user_id,)
        ).fetchall()
    }
    name = (getattr(current_user, "name", None) or "").strip()
    rows = conn.execute("SELECT id,country,manager FROM sites").fetchall()
    allowed = set(explicit)
    for row in rows:
        if row["country"] in countries or (name and (row["manager"] or "").strip() == name):
            allowed.add(int(row["id"]))
    return sorted(allowed - excluded)


def _site_allowed(conn, site_id):
    scope = _scoped_site_ids(conn)
    return scope is None or int(site_id) in scope


def _site_rows_for_current_user(conn):
    scope = _scoped_site_ids(conn)
    sql = "SELECT id,url,country,manager FROM sites"
    params = []
    if scope is not None:
        if not scope:
            return []
        sql += f" WHERE id IN ({','.join('?' * len(scope))})"
        params.extend(scope)
    sql += " ORDER BY country,url"
    return conn.execute(sql, params).fetchall()


def _json_error(message, status=400):
    return jsonify({"error": message}), status


def sync_inventory_view_required(func):
    """Exclude warehouse-only partners from Woo stock publishing controls."""
    @wraps(func)
    def wrapper(*args, **kwargs):
        if not (_is_superadmin() or can_view_inventory() or can_manage_inventory()):
            if request.path.startswith("/api/"):
                return _json_error("您没有 Woo 库存同步权限", 403)
            return _deny("Woo 库存同步需要站点库存权限，请联系管理员。")
        return func(*args, **kwargs)
    return wrapper


@inv_push_bp.route("/inventory/push")
@login_required
@sync_inventory_view_required
def push_page():
    conn = get_conn()
    try:
        sites = [dict(row) for row in _site_rows_for_current_user(conn)]
    finally:
        conn.close()
    return render_template("inv_push.html", sites=sites)


@inv_push_bp.route("/api/inv/site-stock/<int:site_id>", methods=["GET"])
@login_required
@sync_inventory_view_required
def api_site_stock(site_id):
    conn = get_conn()
    try:
        if not _site_allowed(conn, site_id):
            return _json_error("您没有该站点的库存查看权限", 403)
        use_strategy = request.args.get("strategy") == "auto"
        return jsonify(compute_site_stock(conn, site_id, use_sync_strategy=use_strategy))
    finally:
        conn.close()


@inv_push_bp.route("/api/inv/push/<int:site_id>", methods=["POST"])
@login_required
@sync_inventory_view_required
def api_push_site(site_id):
    data = request.get_json(silent=True) or {}
    dry_run = bool(data.get("dry_run", True))
    conn = get_conn()
    try:
        if not _site_allowed(conn, site_id):
            return _json_error("您没有该站点的库存同步权限", 403)
        if not dry_run and not (_is_superadmin() or can_manage_inventory()):
            return _json_error("您只能执行不写入 WooCommerce 的同步演练", 403)
        result = execute_site_sync(
            conn,
            site_id,
            trigger_type="manual",
            force_dry_run=dry_run,
            operator=current_operator(),
        )
        if result.get("fatal"):
            return jsonify(result), 409
        return jsonify(result)
    finally:
        conn.close()


@inv_push_bp.route("/api/inv/sync-configs", methods=["GET"])
@login_required
@sync_inventory_view_required
def api_sync_configs():
    conn = get_conn()
    try:
        rows = []
        for site in _site_rows_for_current_user(conn):
            config = get_site_sync_config(conn, site["id"])
            readiness = site_sync_readiness(
                conn, site["id"], for_live=config["mode"] == "live"
            )
            latest_run = None
            if _table_exists(conn, "inv_push_runs"):
                run = conn.execute(
                    """SELECT status,total_count,ok_count,unchanged_count,error_count,
                              fatal_error,started_at,finished_at
                       FROM inv_push_runs WHERE site_id=? ORDER BY id DESC LIMIT 1""",
                    (site["id"],),
                ).fetchone()
                latest_run = dict(run) if run else None
            rows.append(
                {
                    **dict(site),
                    **config,
                    "readiness": readiness,
                    "latest_run": latest_run,
                    "can_edit": bool(_is_superadmin() or can_manage_inventory()),
                }
            )
        return jsonify(
            {
                "global_enabled": global_sync_enabled(conn),
                "can_edit_global": _is_superadmin(),
                "sites": rows,
            }
        )
    finally:
        conn.close()


@inv_push_bp.route("/api/inv/sync-config/<int:site_id>", methods=["PUT"])
@login_required
@sync_inventory_view_required
def api_update_sync_config(site_id):
    if not (_is_superadmin() or can_manage_inventory()):
        return _json_error("您没有库存同步配置权限", 403)
    conn = get_conn()
    try:
        if not _site_allowed(conn, site_id):
            return _json_error("您没有该站点的配置权限", 403)
        try:
            config = update_site_sync_config(
                conn, site_id, request.get_json(silent=True) or {}, current_operator()
            )
        except ValueError as exc:
            return _json_error(str(exc), 400)
        readiness = site_sync_readiness(
            conn, site_id, for_live=config["mode"] == "live"
        )
        if config["mode"] == "live" and not readiness["ready"]:
            update_site_sync_config(
                conn, site_id, {**config, "mode": "off"}, current_operator()
            )
            return _json_error(
                "未通过正式同步就绪检查，已保持关闭：" + "；".join(readiness["errors"]),
                409,
            )
        return jsonify({"success": True, "config": config, "readiness": readiness})
    finally:
        conn.close()


@inv_push_bp.route("/api/inv/sync-config/bulk", methods=["POST"])
@login_required
@inv_admin_required
def api_bulk_sync_config():
    if not _is_superadmin():
        return _json_error("批量开关仅限超级管理员", 403)
    data = request.get_json(silent=True) or {}
    try:
        site_ids = sorted({int(value) for value in data.get("site_ids", [])})
    except (TypeError, ValueError):
        return _json_error("站点 ID 格式错误")
    if not site_ids:
        return _json_error("请至少选择一个站点")
    conn = get_conn()
    updated, errors = [], []
    try:
        original = {}
        for site_id in site_ids:
            try:
                original[site_id] = get_site_sync_config(conn, site_id)
                update_site_sync_config(conn, site_id, data, current_operator())
                updated.append(site_id)
            except ValueError as exc:
                errors.append({"site_id": site_id, "error": str(exc)})
        if data.get("mode") == "live" and not errors:
            for site_id in updated:
                readiness = site_sync_readiness(conn, site_id, for_live=True)
                if not readiness["ready"]:
                    errors.append(
                        {"site_id": site_id, "error": "；".join(readiness["errors"])}
                    )
        if errors:
            for site_id in updated:
                update_site_sync_config(
                    conn, site_id, original[site_id], current_operator()
                )
            updated = []
        return jsonify({"success": not errors, "updated": updated, "errors": errors})
    finally:
        conn.close()


@inv_push_bp.route("/api/inv/sync-global", methods=["PUT"])
@login_required
@inv_admin_required
def api_update_sync_global():
    if not _is_superadmin():
        return _json_error("全局开关仅限超级管理员", 403)
    enabled = bool((request.get_json(silent=True) or {}).get("enabled", False))
    conn = get_conn()
    try:
        before = {"enabled": global_sync_enabled(conn)}
        conn.execute(
            "INSERT OR REPLACE INTO settings (key,value) VALUES (?,?)",
            ("inv_auto_push_global_enabled", "1" if enabled else "0"),
        )
        _audit_config(
            conn, None, "update_global_switch", before, {"enabled": enabled},
            current_operator(),
        )
        conn.commit()
        return jsonify({"success": True, "enabled": enabled})
    finally:
        conn.close()


@inv_push_bp.route("/api/inv/push-logs", methods=["GET"])
@login_required
@sync_inventory_view_required
def api_push_logs():
    site_id = request.args.get("site_id")
    try:
        limit = min(2000, int(request.args.get("limit") or 200))
    except ValueError:
        limit = 200
    conn = get_conn()
    try:
        if site_id and not _site_allowed(conn, int(site_id)):
            return _json_error("您没有该站点的库存日志权限", 403)
        sql = """SELECT pl.*,k.sku_code FROM inv_push_logs pl
                 LEFT JOIN inv_skus k ON k.id=pl.sku_id WHERE 1=1"""
        params = []
        if site_id:
            sql += " AND pl.site_id=?"
            params.append(int(site_id))
        else:
            scope = _scoped_site_ids(conn)
            if scope is not None:
                if not scope:
                    return jsonify([])
                sql += f" AND pl.site_id IN ({','.join('?' * len(scope))})"
                params.extend(scope)
        sql += " ORDER BY pl.id DESC LIMIT ?"
        params.append(limit)
        rows = conn.execute(sql, params).fetchall()
        return jsonify([dict(row) for row in rows])
    finally:
        conn.close()
