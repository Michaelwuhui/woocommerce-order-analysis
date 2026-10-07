"""Request-scoped product recognition and one-pass historical trend building.

There is no process-global cache: permissions, rules and mappings are loaded
fresh by each endpoint and no order data is retained between requests.
"""
from datetime import datetime, timedelta


class RequestProductParser:
    def __init__(self, parser, brands, series):
        self.parser = parser
        self.brands = brands
        self.series = series
        self.results = {}

    def __call__(self, name):
        if name not in self.results:
            # Explicit series is essential: the legacy wrapper fetches it from
            # the database when omitted, even if a caller ignores that field.
            self.results[name] = self.parser(name, self.brands, self.series)
        return self.results[name]


def build_weekly_trends(orders, top_products, top_flavors, top_flavor_qtys,
                       manual_mappings, *, parse_items, full_product_name,
                       normalize_raw_name, normalize_flavor, parse_product):
    """Keep the two existing chart classification rules while reading once.

    The flavor chart intentionally uses full-name manual overrides and variation
    metadata only. The product chart also uses base-name overrides and parsed
    name fallbacks. Series remains absent from its historical key, as before.
    """
    flavor_keys = {flavor.upper().strip(): flavor for flavor in top_flavors}
    product_keys = {}
    for product in top_products[:10]:
        brand = product.get("brand") or "Unknown"
        puffs = product.get("puffs") or "N/A"
        flavor = normalize_flavor(product.get("flavor") or "")
        key = f"{brand}|{puffs}|{flavor}"
        product_keys[key] = {
            "label": f"{brand} {puffs} {product.get('flavor', '')[:15]}",
            "pageTotal": product["quantity"],
        }

    weekly_flavors, weekly_products = {}, {}
    for order in orders:
        items = parse_items(order["line_items"])
        if not isinstance(items, list):
            continue
        created = order["date_created"]
        if not created:
            continue
        try:
            date = datetime.strptime(created[:10], "%Y-%m-%d")
            year, week, _ = date.isocalendar()
            week_key = f"{year}-W{week:02d}"
            label = (date - timedelta(days=date.weekday())).strftime("%m/%d")
        except Exception:
            continue
        flavors = weekly_flavors.setdefault(week_key, {"label": label, "flavors": {}})["flavors"]
        products = weekly_products.setdefault(week_key, {"label": label, "products": {}})["products"]
        flavor_valid = product_valid = True
        for item in items:
            try:
                quantity = item.get("quantity", 0)
                name = item.get("name", "")
                full_name, meta_flavor, meta_puffs = full_product_name(item)
                full_key = normalize_raw_name(full_name)
            except Exception:
                # Each former pass stopped this order at the malformed line.
                break
            if flavor_valid:
                try:
                    if full_key in manual_mappings and manual_mappings[full_key].get("flavor"):
                        flavor = manual_mappings[full_key]["flavor"]
                    else:
                        flavor = meta_flavor or "未知口味"
                    matched = flavor_keys.get(flavor.upper().strip())
                    if matched is not None:
                        flavors[matched] = flavors.get(matched, 0) + quantity
                except Exception:
                    flavor_valid = False
            if product_valid:
                try:
                    name_key = normalize_raw_name(name)
                    mapping = manual_mappings.get(full_key) or manual_mappings.get(name_key)
                    if mapping:
                        brand = mapping.get("brand") or "Unknown"
                        puffs = mapping.get("puffs") or meta_puffs
                        flavor = mapping.get("flavor") or meta_flavor or ""
                    else:
                        parsed = parse_product(name)
                        brand = parsed.get("brand") or "Unknown"
                        puffs = meta_puffs or parsed.get("puffs")
                        flavor = meta_flavor or parsed.get("flavor") or ""
                    key = f"{brand}|{puffs or 'N/A'}|{normalize_flavor(flavor)}"
                    if key in product_keys:
                        products[key] = products.get(key, 0) + quantity
                except Exception:
                    product_valid = False
            if not flavor_valid and not product_valid:
                break

    weeks = sorted(weekly_flavors)[-8:]
    weekly_trend = {
        "weeks": [weekly_flavors[key]["label"] for key in weeks],
        "flavors": top_flavors,
        "datasets": [{"flavor": flavor,
                      "data": [weekly_flavors[key]["flavors"].get(flavor, 0) for key in weeks],
                      "pageTotal": top_flavor_qtys.get(flavor, 0)} for flavor in top_flavors],
    }
    product_trend = {
        "weeks": [weekly_products[key]["label"] for key in weeks],
        "products": [info["label"] for info in product_keys.values()],
        "datasets": [{"label": info["label"],
                      "data": [weekly_products[week_key]["products"].get(key, 0) for week_key in weeks],
                      "pageTotal": info["pageTotal"]} for key, info in product_keys.items()],
    }
    return weekly_trend, product_trend
