"""Composite order-id helpers (cross-site collision fix).

Order numbers are only unique *within* one WooCommerce store. We run many
stores that each number orders 1..N independently, so the same numeric id
(e.g. 9009) exists on several sites. The local DB used the bare WC id as the
`orders` primary key, so a same-numbered order from a second site collided
with the first under `ON CONFLICT(id)` and the two were silently merged.

Fix: the local primary key is now a *surrogate* id that is globally unique by
construction:

    oid = "<sites.id>-<woo_id>"      e.g.  "14-9009"  (vapesklep)  vs  "10-9009" (vapepolska)

`woo_id` keeps the raw per-site WooCommerce post id, which is what the WC REST
API needs for write-back (/orders/<woo_id>...). `number` stays the customer
facing display number. The surrogate is opaque everywhere else — the app
passes `id` around as a string token and never does arithmetic on it.

This module is the single source of truth for the surrogate format. It is
imported by the sync (sync_utils.py, 1.wooorders_sqlite.py), the web app
(app.py), the migration script, and any tool that writes orders, so they all
agree on exactly the same id.
"""


def make_oid(site_id, woo_id):
    """Build the surrogate order id for a (site, woo_id) pair.

    site_id : int  — sites.id (stable per-store row id)
    woo_id  : int/str — the raw WooCommerce order/post id (orders.number's id sibling)
    """
    return f"{int(site_id)}-{woo_id}"


def woo_post_id(oid):
    """Extract the raw WooCommerce post id from a surrogate (or a bare id).

    Use this anywhere the value is handed to the WC REST API. Accepts:
      "14-9009" -> "9009"   (surrogate)
      "9009"    -> "9009"   (legacy / already-bare, e.g. test orders)
       9009     -> "9009"
    """
    if oid is None:
        return None
    s = str(oid)
    # surrogate is "<siteid>-<wooid>"; woo_id never contains '-', so the last
    # segment after the final '-' is always the raw WC id.
    return s.rsplit("-", 1)[-1] if "-" in s else s


def is_surrogate(oid):
    """True if oid already looks like a "<siteid>-<wooid>" surrogate."""
    if oid is None:
        return False
    s = str(oid)
    if "-" not in s:
        return False
    a, b = s.split("-", 1)
    return a.isdigit() and b.isdigit()


# ---- source(url) -> sites.id resolution, cached per connection -------------

_site_id_cache = {}


def site_id_for_source(conn, source):
    """Map an order's source URL to its sites.id. Cached in-process.

    Returns int site id, or None if the source isn't a known site (caller must
    decide what to do — sync should skip / log rather than mis-key).
    """
    if source is None:
        return None
    if source in _site_id_cache:
        return _site_id_cache[source]
    row = conn.execute("SELECT id FROM sites WHERE url = ?", (source,)).fetchone()
    sid = row[0] if row else None
    _site_id_cache[source] = sid
    return sid


def clear_site_id_cache():
    _site_id_cache.clear()
