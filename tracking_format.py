"""Tracking-plugin detection shared by the shipping workflow and tests."""

from order_shipments import detect_tracking_format_rows


AST_NAMESPACE_PREFIXES = (
    "wc-ast-pro/",
    "wc-ast/",
    "wc-shipment-tracking/",
)


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

    detected = detect_tracking_format_rows(rows)
    if detected != "unknown":
        return detected

    try:
        import requests

        root = requests.get(
            f"{site_url.rstrip('/')}/wp-json/",
            timeout=6,
            headers={"User-Agent": "WooCommerce Order Analysis/1.0"},
        )
        if root.status_code == 200:
            namespaces = root.json().get("namespaces") or []
            if any(
                str(namespace).strip().lower().startswith(AST_NAMESPACE_PREFIXES)
                for namespace in namespaces
            ):
                return "ast"
    except Exception:
        # The caller retains its established safe fallback: AST-shaped
        # metadata and a customer-note notification.
        pass
    return "unknown"
