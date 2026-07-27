"""Tracking-plugin detection shared by the shipping workflow and tests."""


def detect_site_tracking_format(conn, site_url):
    """Return ast, villatheme, custom_lineitem, or unknown for a site.

    Existing shipment metadata is authoritative.  A brand-new site may not
    have local shipment history yet, so its public REST namespaces are used as
    a non-mutating fallback.
    """

    rows = conn.execute(
        """SELECT meta_data, line_items FROM orders
           WHERE source = ? AND status IN ('on-hold','shipped','completed')
           ORDER BY date_modified DESC LIMIT 10""",
        (site_url,),
    ).fetchall()

    ast = villa = custom = 0
    for row in rows:
        metadata = row["meta_data"] or ""
        line_items = row["line_items"] or ""
        if "_wc_shipment_tracking_items" in metadata:
            ast += 1
        if "_vi_wot_order_item_tracking_data" in line_items:
            villa += 1
        elif '"key":"tracking_number"' in line_items or '"key": "tracking_number"' in line_items:
            custom += 1

    if ast and ast >= max(villa, custom):
        return "ast"
    if villa and villa >= custom:
        return "villatheme"
    if custom:
        return "custom_lineitem"

    try:
        import requests

        root = requests.get(
            f"{site_url.rstrip('/')}/wp-json/",
            timeout=6,
            headers={"User-Agent": "WooCommerce Order Analysis/1.0"},
        )
        if root.status_code == 200:
            namespaces = root.json().get("namespaces") or []
            if any(str(ns).lower().startswith("wc-shipment-tracking/") for ns in namespaces):
                return "ast"
    except Exception:
        # The caller retains its established safe fallback: AST-shaped
        # metadata and a customer-note notification.
        pass
    return "unknown"
