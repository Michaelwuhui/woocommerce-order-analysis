"""Import-time production smoke test; does not send external requests."""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app


rules = {rule.rule for rule in app.url_map.iter_rules()}
required = {
    "/fulfillment",
    "/api/fulfillment/orders",
    "/api/fulfillment/webhook/wms",
    "/api/fulfillment/<fulfillment_id>/shipment",
    "/api/fulfillment/<fulfillment_id>/submit-wms",
    "/api/fulfillment/<fulfillment_id>/cancel",
}
missing = required - rules
assert not missing, f"missing fulfillment routes: {sorted(missing)}"

for template in ("base.html", "orders.html", "shipping.html", "fulfillment.html"):
    app.jinja_env.get_template(template)

print(f"app_smoke=ok routes={len(rules)} fulfillment_routes={len([r for r in rules if 'fulfillment' in r])}")
