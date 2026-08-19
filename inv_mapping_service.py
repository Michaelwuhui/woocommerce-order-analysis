"""Warehouse-first WooCommerce product mapping helpers.

The warehouse SKU is the canonical fulfilment identity.  WooCommerce products
and variations are only site-specific aliases of that SKU.  This module keeps
the read-only Woo catalogue scan and the deterministic candidate rules away
from the Flask views so they can be tested without a request context.
"""

from __future__ import annotations

import html
import json
import re
import unicodedata
from collections import Counter, defaultdict


def normalize_identifier(value):
    value = html.unescape(str(value or ""))
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def _identifier_tokens(value):
    value = html.unescape(str(value or ""))
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch)).lower()
    return re.findall(r"[a-z0-9]+", value)


def _family_matches(sku, product_name):
    name_key = normalize_identifier(product_name)
    if sku["family_key"] and sku["family_key"] in name_key:
        return True
    # Warehouses and shops sometimes omit a marketing token (for example
    # "RandM") while keeping the brand, model and puff number.  Treat this as
    # a review-level family match, never as an automatic exact match.
    ignored = {"puff", "puffs", "randm", "disposable", "vape"}
    family_tokens = [t for t in sku.get("family_tokens", []) if t not in ignored]
    product_tokens = set(_identifier_tokens(product_name))
    numbers = [t for t in family_tokens if any(ch.isdigit() for ch in t)]
    words = [t for t in family_tokens if t not in numbers]
    return bool(
        numbers
        and all(token in product_tokens for token in numbers)
        and len([token for token in words if token in product_tokens]) >= min(2, len(words))
    )


def _managed_family(notes, name):
    for line in str(notes or "").splitlines():
        try:
            payload = json.loads(line)
        except (TypeError, ValueError):
            continue
        family = payload.get("managed_family") if isinstance(payload, dict) else None
        if family:
            return str(family).replace("-", " ")
    return str(name or "").split(" - ", 1)[0].strip()


def warehouse_skus(conn, warehouse_id):
    rows = conn.execute(
        """SELECT k.id,k.sku_code,k.name,k.barcode,k.flavor,k.notes,k.is_active,
                  COALESCE(st.on_hand,0) AS on_hand,COALESCE(st.reserved,0) AS reserved
           FROM oms_sku_warehouses sw
           JOIN inv_skus k ON k.id=sw.sku_id
           LEFT JOIN inv_stock st ON st.warehouse_id=sw.warehouse_id AND st.sku_id=sw.sku_id
           WHERE sw.warehouse_id=? AND sw.is_enabled=1 AND k.is_active=1
           ORDER BY k.sku_code""",
        (warehouse_id,),
    ).fetchall()
    result = []
    for row in rows:
        item = dict(row)
        item["family"] = _managed_family(item.get("notes"), item.get("name"))
        item["family_key"] = normalize_identifier(item["family"])
        item["family_tokens"] = _identifier_tokens(item["family"])
        item["name_key"] = normalize_identifier(item.get("name"))
        item["flavor_key"] = normalize_identifier(item.get("flavor"))
        item["sku_key"] = normalize_identifier(item.get("sku_code"))
        item["barcode_key"] = normalize_identifier(item.get("barcode"))
        item["available"] = int(item.get("on_hand") or 0) - int(item.get("reserved") or 0)
        result.append(item)
    return result


def warehouse_rows(conn, allowed_ids=None):
    sql = """SELECT w.id,w.name,w.code,w.country,
                    COALESCE(wi.inventory_authority,'local') AS inventory_authority,
                    COUNT(DISTINCT sw.sku_id) AS sku_count
             FROM warehouses w
             JOIN oms_sku_warehouses sw ON sw.warehouse_id=w.id AND sw.is_enabled=1
             LEFT JOIN oms_warehouse_integrations wi ON wi.warehouse_id=w.id
             WHERE w.is_active=1"""
    params = []
    if allowed_ids is not None:
        if not allowed_ids:
            return []
        sql += f" AND w.id IN ({','.join('?' * len(allowed_ids))})"
        params.extend(allowed_ids)
    sql += " GROUP BY w.id ORDER BY w.country,w.name"
    return [dict(row) for row in conn.execute(sql, params).fetchall()]


def serving_sites(conn, warehouse_id):
    markets = [
        str(row["market_code"] or "").upper()
        for row in conn.execute(
            """SELECT market_code FROM inv_market_warehouses
               WHERE warehouse_id=? AND is_active=1 ORDER BY market_code""",
            (warehouse_id,),
        ).fetchall()
        if row["market_code"]
    ]
    if not markets:
        wh = conn.execute("SELECT country FROM warehouses WHERE id=?", (warehouse_id,)).fetchone()
        if wh and wh["country"]:
            markets = [str(wh["country"]).upper()]
    if not markets:
        return []
    marks = ",".join("?" * len(markets))
    return [
        dict(row)
        for row in conn.execute(
            f"""SELECT id,url,country,manager,
                       CASE WHEN COALESCE(consumer_key,'')<>'' AND COALESCE(consumer_secret,'')<>''
                            THEN 1 ELSE 0 END AS has_api
                FROM sites WHERE UPPER(COALESCE(country,'')) IN ({marks})
                ORDER BY country,url""",
            markets,
        ).fetchall()
    ]


def _candidate_for_product(product, skus):
    sku_key = normalize_identifier(product.get("wc_sku"))
    name_key = normalize_identifier(product.get("name"))

    if sku_key:
        exact = [s for s in skus if sku_key in (s["sku_key"], s["barcode_key"]) and sku_key]
        if len(exact) == 1:
            return exact[0]["id"], "exact_sku", 100

    exact = [s for s in skus if name_key and s["name_key"] == name_key]
    if len(exact) == 1:
        return exact[0]["id"], "exact_name", 98

    review = []
    for sku in skus:
        family_ok = _family_matches(sku, product.get("name"))
        flavor_ok = bool(sku["flavor_key"] and sku["flavor_key"] in name_key)
        if family_ok and flavor_ok:
            review.append(sku)
    if len(review) == 1:
        return review[0]["id"], "review_family_flavor", 85
    return None, None, 0


def _parent_is_relevant(product, skus):
    sku_key = normalize_identifier(product.get("sku"))
    if sku_key and any(sku_key in (s["sku_key"], s["barcode_key"]) for s in skus):
        return True
    return any(_family_matches(s, product.get("name")) for s in skus)


def _response_json(response, label):
    if response.status_code >= 400:
        try:
            message = (response.json() or {}).get("message")
        except Exception:
            message = None
        raise RuntimeError(f"{label}失败：HTTP {response.status_code} {message or response.text[:160]}")
    try:
        payload = response.json()
    except Exception as exc:
        raise RuntimeError(f"{label}返回的不是 JSON") from exc
    if not isinstance(payload, list):
        raise RuntimeError(f"{label}返回格式异常")
    return payload


def _woo_pages(session, url, auth, params, label, max_pages=100):
    rows = []
    for page in range(1, max_pages + 1):
        page_params = dict(params, per_page=100, page=page)
        response = session.get(
            url,
            auth=auth,
            params=page_params,
            timeout=90,
            headers={"User-Agent": "WooCommerce API Client-Python/3.0.0", "Accept": "application/json"},
        )
        batch = _response_json(response, label)
        rows.extend(batch)
        total_pages = int(response.headers.get("X-WP-TotalPages") or 0)
        if not batch or len(batch) < 100 or (total_pages and page >= total_pages):
            break
    return rows


def _variation_name(parent_name, variation):
    options = []
    for attr in variation.get("attributes") or []:
        option = str(attr.get("option") or "").strip()
        if option and option not in options:
            options.append(option)
    return " - ".join([str(parent_name or "").strip()] + options) if options else str(parent_name or "")


def scan_site_catalog(conn, site_id, warehouse_id, session=None):
    """Read relevant published products from one Woo site and cache candidates.

    This function performs GET requests only.  It never changes WooCommerce.
    Variable-product variations are fetched only for warehouse-related parent
    products, keeping the scan bounded even on large stores.
    """
    import requests

    site = conn.execute(
        """SELECT id,url,consumer_key,consumer_secret FROM sites WHERE id=?""",
        (site_id,),
    ).fetchone()
    if not site:
        raise ValueError("站点不存在")
    skus = warehouse_skus(conn, warehouse_id)
    if not skus:
        raise ValueError("所选仓库没有已启用 SKU")
    if not site["consumer_key"] or not site["consumer_secret"]:
        raise ValueError("站点缺少 WooCommerce API 凭据")

    conn.execute(
        """INSERT INTO inv_site_catalog_scans
             (site_id,warehouse_id,status,started_at,finished_at,total_products,error)
           VALUES (?,?,'running',CURRENT_TIMESTAMP,NULL,0,NULL)
           ON CONFLICT(site_id,warehouse_id) DO UPDATE SET
             status='running',started_at=CURRENT_TIMESTAMP,finished_at=NULL,error=NULL""",
        (site_id, warehouse_id),
    )
    conn.commit()
    client = session or requests.Session()
    base = str(site["url"]).rstrip("/") + "/wp-json/wc/v3/products"
    auth = (site["consumer_key"], site["consumer_secret"])
    cached = []
    try:
        products = _woo_pages(client, base, auth, {"status": "publish"}, "读取商品")
        for product in products:
            if not _parent_is_relevant(product, skus):
                continue
            ptype = str(product.get("type") or "simple")
            if ptype == "variable":
                variations = _woo_pages(
                    client,
                    f"{base}/{int(product['id'])}/variations",
                    auth,
                    {"status": "publish"},
                    f"读取商品 {product.get('id')} 变体",
                )
                for variation in variations:
                    cached.append({
                        "wc_product_id": int(product["id"]),
                        "wc_variation_id": int(variation["id"]),
                        "wc_sku": variation.get("sku") or "",
                        "name": _variation_name(product.get("name"), variation),
                        "product_type": "variation",
                        "manage_stock": 1 if variation.get("manage_stock") else 0,
                        "stock_quantity": variation.get("stock_quantity"),
                        "permalink": product.get("permalink") or "",
                    })
            else:
                cached.append({
                    "wc_product_id": int(product["id"]),
                    "wc_variation_id": 0,
                    "wc_sku": product.get("sku") or "",
                    "name": product.get("name") or "",
                    "product_type": ptype,
                    "manage_stock": 1 if product.get("manage_stock") else 0,
                    "stock_quantity": product.get("stock_quantity"),
                    "permalink": product.get("permalink") or "",
                })

        conn.execute(
            "DELETE FROM inv_site_product_catalog WHERE site_id=? AND warehouse_id=?",
            (site_id, warehouse_id),
        )
        for item in cached:
            candidate_id, method, confidence = _candidate_for_product(item, skus)
            conn.execute(
                """INSERT INTO inv_site_product_catalog
                     (site_id,warehouse_id,wc_product_id,wc_variation_id,wc_sku,name,
                      product_type,manage_stock,stock_quantity,permalink,
                      candidate_sku_id,match_method,match_confidence,scanned_at)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)""",
                (
                    site_id, warehouse_id, item["wc_product_id"], item["wc_variation_id"],
                    item["wc_sku"], item["name"], item["product_type"], item["manage_stock"],
                    item["stock_quantity"], item["permalink"], candidate_id, method, confidence,
                ),
            )
        conn.execute(
            """UPDATE inv_site_catalog_scans SET status='success',finished_at=CURRENT_TIMESTAMP,
                      total_products=?,error=NULL WHERE site_id=? AND warehouse_id=?""",
            (len(cached), site_id, warehouse_id),
        )
        conn.commit()
        return {"site_id": site_id, "warehouse_id": warehouse_id, "total_products": len(cached)}
    except Exception as exc:
        conn.execute(
            """UPDATE inv_site_catalog_scans SET status='failed',finished_at=CURRENT_TIMESTAMP,error=?
               WHERE site_id=? AND warehouse_id=?""",
            (str(exc)[:500], site_id, warehouse_id),
        )
        conn.commit()
        raise


def _catalog_rows_with_map(conn, site_id, warehouse_id):
    rows = conn.execute(
        """SELECT c.*,k.sku_code AS candidate_sku_code,k.name AS candidate_sku_name,
                  m.id AS map_id,m.sku_id AS mapped_sku_id,m.qty_per_item,
                  mk.sku_code AS mapped_sku_code
           FROM inv_site_product_catalog c
           LEFT JOIN inv_skus k ON k.id=c.candidate_sku_id
           LEFT JOIN inv_site_sku_map m
             ON m.site_id=c.site_id AND m.wc_product_id=c.wc_product_id
            AND COALESCE(m.wc_variation_id,0)=COALESCE(c.wc_variation_id,0) AND m.is_active=1
           LEFT JOIN inv_skus mk ON mk.id=m.sku_id
           WHERE c.site_id=? AND c.warehouse_id=?
           ORDER BY c.name,c.wc_variation_id""",
        (site_id, warehouse_id),
    ).fetchall()
    return [dict(row) for row in rows]


def mapping_overview(conn, warehouse_id):
    sku_count = len(warehouse_skus(conn, warehouse_id))
    result = []
    for site in serving_sites(conn, warehouse_id):
        scan = conn.execute(
            """SELECT status,started_at,finished_at,total_products,error
               FROM inv_site_catalog_scans WHERE site_id=? AND warehouse_id=?""",
            (site["id"], warehouse_id),
        ).fetchone()
        catalog = _catalog_rows_with_map(conn, site["id"], warehouse_id) if scan else []
        mapped_catalog = sum(1 for row in catalog if row.get("map_id"))
        exact = sum(
            1 for row in catalog
            if not row.get("map_id") and row.get("match_method") in ("exact_sku", "exact_name")
        )
        review = sum(
            1 for row in catalog
            if not row.get("map_id") and row.get("match_method") == "review_family_flavor"
        )
        unresolved = sum(
            1 for row in catalog if not row.get("map_id") and not row.get("candidate_sku_id")
        )
        mapped_skus = conn.execute(
            """SELECT COUNT(DISTINCT m.sku_id)
               FROM inv_site_sku_map m
               JOIN oms_sku_warehouses sw ON sw.sku_id=m.sku_id
               WHERE m.site_id=? AND sw.warehouse_id=? AND sw.is_enabled=1 AND m.is_active=1""",
            (site["id"], warehouse_id),
        ).fetchone()[0]
        if not scan:
            readiness = "not_scanned"
        elif scan["status"] == "failed":
            readiness = "scan_failed"
        elif not catalog:
            readiness = "not_listed"
        elif mapped_catalog == len(catalog):
            readiness = "ready"
        else:
            readiness = "action_needed"
        result.append({
            **site,
            "warehouse_sku_count": sku_count,
            "mapped_sku_count": int(mapped_skus or 0),
            "catalog_count": len(catalog),
            "mapped_catalog_count": mapped_catalog,
            "exact_candidate_count": exact,
            "review_candidate_count": review,
            "unresolved_count": unresolved,
            "readiness": readiness,
            "scan": dict(scan) if scan else None,
        })
    return result


def mapping_detail(conn, warehouse_id, site_id):
    skus = warehouse_skus(conn, warehouse_id)
    maps = conn.execute(
        """SELECT m.* FROM inv_site_sku_map m
           JOIN oms_sku_warehouses sw ON sw.sku_id=m.sku_id
           WHERE m.site_id=? AND sw.warehouse_id=? AND sw.is_enabled=1 AND m.is_active=1
           ORDER BY m.id""",
        (site_id, warehouse_id),
    ).fetchall()
    maps_by_sku = defaultdict(list)
    for row in maps:
        maps_by_sku[int(row["sku_id"])].append(dict(row))
    catalog = _catalog_rows_with_map(conn, site_id, warehouse_id)
    candidates_by_sku = defaultdict(list)
    for row in catalog:
        if row.get("candidate_sku_id") and not row.get("map_id"):
            candidates_by_sku[int(row["candidate_sku_id"])].append(row)
    sku_rows = []
    for sku in skus:
        item = dict(sku)
        item["mappings"] = maps_by_sku.get(int(sku["id"]), [])
        item["candidates"] = candidates_by_sku.get(int(sku["id"]), [])
        if item["mappings"]:
            item["status"] = "mapped"
        elif item["candidates"]:
            item["status"] = "candidate"
        else:
            item["status"] = "not_listed"
        sku_rows.append(item)
    scan = conn.execute(
        "SELECT * FROM inv_site_catalog_scans WHERE site_id=? AND warehouse_id=?",
        (site_id, warehouse_id),
    ).fetchone()
    return {
        "skus": sku_rows,
        "site_products": catalog,
        "scan": dict(scan) if scan else None,
    }


def apply_safe_mappings(conn, site_id, warehouse_id, operator_id=None, operator_name=None):
    """Create only unique exact SKU/name mappings from the latest scan."""
    rows = _catalog_rows_with_map(conn, site_id, warehouse_id)
    allowed = {int(s["id"]) for s in warehouse_skus(conn, warehouse_id)}
    wc_sku_counts = Counter(
        str(row.get("wc_sku") or "").strip().lower()
        for row in rows if str(row.get("wc_sku") or "").strip()
    )
    created = []
    skipped = []
    for row in rows:
        if row.get("map_id"):
            continue
        sku_id = row.get("candidate_sku_id")
        method = row.get("match_method")
        if not sku_id or int(sku_id) not in allowed or method not in ("exact_sku", "exact_name"):
            continue
        conflict = conn.execute(
            """SELECT id,sku_id FROM inv_site_sku_map
               WHERE site_id=? AND is_active=1 AND wc_product_id=?
                 AND COALESCE(wc_variation_id,0)=? LIMIT 1""",
            (site_id, row["wc_product_id"], row["wc_variation_id"] or 0),
        ).fetchone()
        if conflict:
            skipped.append({"catalog_id": row["id"], "reason": "conflict"})
            continue
        wc_sku = str(row.get("wc_sku") or "").strip() or None
        # A number of Woo variable products reuse one parent-level SKU on all
        # flavour variations.  Product+variation remains exact, but that shared
        # SKU is ambiguous and must not become a fallback resolver key.
        if wc_sku and wc_sku_counts[wc_sku.lower()] > 1:
            wc_sku = None
        if method == "exact_sku" and wc_sku:
            sku_conflict = conn.execute(
                """SELECT id,sku_id FROM inv_site_sku_map
                   WHERE site_id=? AND is_active=1 AND COALESCE(wc_sku,'')<>''
                     AND LOWER(wc_sku)=LOWER(?) AND sku_id<>? LIMIT 1""",
                (site_id, wc_sku, int(sku_id)),
            ).fetchone()
            if sku_conflict:
                skipped.append({"catalog_id": row["id"], "reason": "wc_sku_conflict"})
                continue
        cursor = conn.execute(
            """INSERT INTO inv_site_sku_map
                 (site_id,wc_product_id,wc_variation_id,wc_sku,raw_name,sku_id,qty_per_item,is_active)
               VALUES (?,?,?,?,?,?,1,1)""",
            (
                site_id, row["wc_product_id"], row["wc_variation_id"] or 0,
                wc_sku, row["name"], int(sku_id),
            ),
        )
        map_id = int(cursor.lastrowid)
        conn.execute(
            """INSERT INTO inv_mapping_audit
                 (action,site_id,warehouse_id,sku_id,map_id,match_method,operator_id,operator_name,payload_json)
               VALUES ('auto_create',?,?,?,?,?,?,?,?)""",
            (
                site_id, warehouse_id, int(sku_id), map_id, method,
                operator_id, operator_name,
                json.dumps({
                    "wc_product_id": row["wc_product_id"],
                    "wc_variation_id": row["wc_variation_id"],
                    "wc_sku": row["wc_sku"],
                    "name": row["name"],
                }, ensure_ascii=False, separators=(",", ":")),
            ),
        )
        created.append({"map_id": map_id, "sku_id": int(sku_id), "catalog_id": row["id"]})
    conn.commit()
    return {"created": len(created), "skipped": len(skipped), "items": created, "skipped_items": skipped}


def confirm_mapping_candidates(
    conn, site_id, warehouse_id, catalog_ids, operator_id=None, operator_name=None
):
    """Confirm a user-selected batch of one-to-one catalogue candidates.

    Unlike ``apply_safe_mappings``, this accepts review-level candidates because
    an operator has explicitly checked them.  The durable identity remains the
    Woo product+variation pair.  A shared/ambiguous Woo SKU is omitted so it
    cannot route another flavour through the fallback resolver.
    """
    selected = set()
    for value in catalog_ids or []:
        try:
            selected.add(int(value))
        except (TypeError, ValueError):
            continue
    if not selected:
        return {"created": 0, "skipped": 0, "items": [], "skipped_items": []}
    if len(selected) > 500:
        raise ValueError("单次最多确认 500 条映射")

    rows = _catalog_rows_with_map(conn, site_id, warehouse_id)
    allowed = {int(s["id"]) for s in warehouse_skus(conn, warehouse_id)}
    wc_sku_counts = Counter(
        str(row.get("wc_sku") or "").strip().lower()
        for row in rows if str(row.get("wc_sku") or "").strip()
    )
    created = []
    skipped = []
    found = set()
    for row in rows:
        catalog_id = int(row["id"])
        if catalog_id not in selected:
            continue
        found.add(catalog_id)
        sku_id = row.get("candidate_sku_id")
        if row.get("map_id"):
            skipped.append({"catalog_id": catalog_id, "reason": "already_mapped"})
            continue
        if not sku_id or int(sku_id) not in allowed or not row.get("match_method"):
            skipped.append({"catalog_id": catalog_id, "reason": "candidate_missing"})
            continue
        conflict = conn.execute(
            """SELECT id FROM inv_site_sku_map
               WHERE site_id=? AND wc_product_id=?
                 AND COALESCE(wc_variation_id,0)=? AND is_active=1 LIMIT 1""",
            (site_id, row["wc_product_id"], row["wc_variation_id"] or 0),
        ).fetchone()
        if conflict:
            skipped.append({"catalog_id": catalog_id, "reason": "conflict"})
            continue

        wc_sku = str(row.get("wc_sku") or "").strip() or None
        if wc_sku and wc_sku_counts[wc_sku.lower()] > 1:
            wc_sku = None
        if wc_sku:
            fallback_conflict = conn.execute(
                """SELECT id FROM inv_site_sku_map
                   WHERE site_id=? AND COALESCE(wc_sku,'')<>'' AND LOWER(wc_sku)=LOWER(?)
                     AND sku_id<>? AND is_active=1 LIMIT 1""",
                (site_id, wc_sku, int(sku_id)),
            ).fetchone()
            if fallback_conflict:
                wc_sku = None

        cursor = conn.execute(
            """INSERT INTO inv_site_sku_map
                 (site_id,wc_product_id,wc_variation_id,wc_sku,raw_name,sku_id,qty_per_item,is_active)
               VALUES (?,?,?,?,?,?,1,1)""",
            (
                site_id, row["wc_product_id"], row["wc_variation_id"] or 0,
                wc_sku, row["name"], int(sku_id),
            ),
        )
        map_id = int(cursor.lastrowid)
        conn.execute(
            """INSERT INTO inv_mapping_audit
                 (action,site_id,warehouse_id,sku_id,map_id,match_method,
                  operator_id,operator_name,payload_json)
               VALUES ('batch_confirm',?,?,?,?,?,?,?,?)""",
            (
                site_id, warehouse_id, int(sku_id), map_id,
                "manual_confirm:" + str(row.get("match_method") or "candidate"),
                operator_id, operator_name,
                json.dumps({
                    "catalog_id": catalog_id,
                    "wc_product_id": row["wc_product_id"],
                    "wc_variation_id": row["wc_variation_id"],
                    "wc_sku": row["wc_sku"],
                    "stored_wc_sku": wc_sku,
                    "name": row["name"],
                }, ensure_ascii=False, separators=(",", ":")),
            ),
        )
        created.append({"map_id": map_id, "sku_id": int(sku_id), "catalog_id": catalog_id})

    for catalog_id in sorted(selected - found):
        skipped.append({"catalog_id": catalog_id, "reason": "not_found"})
    conn.commit()
    return {
        "created": len(created), "skipped": len(skipped),
        "items": created, "skipped_items": skipped,
    }
