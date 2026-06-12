#!/usr/bin/env python3
"""One-time migration: bare WC order id  ->  cross-site-safe surrogate id.

    orders.id        9009            -> "14-9009"   (<sites.id>-<woo_id>)
    orders.woo_id    (new column)    <- 9009        (raw WC post id for REST)

Re-keys every order-id reference (order_notes, shipping_logs,
reconciliation_statement_orders, orders_archive) and resets fulfillment-status
columns that were inherited from a *different* site's order that previously
shared the same bare id (detected as: status timestamp predates the order's
own creation).

Safe to re-run: every step is guarded on `instr(id,'-')=0`, so already-migrated
rows are skipped. Dry-run by default; pass --live to write. All writes happen
in a single transaction.

Usage:
    python migrate_composite_id.py            # dry-run report, no writes
    python migrate_composite_id.py --live     # perform the migration
"""
import argparse
import sqlite3
import sys

DB = "woocommerce_orders.db"

# fulfillment-status column groups to clean when they predate the order.
# (group label, gate column, predicate timestamp col, columns reset to value)
STALE_GROUPS = [
    ("carrier_status", "carrier_status", "carrier_status_at",
     {"carrier_status": None, "carrier_status_at": None}),
    ("delivery_confirmed", "delivery_confirmed", "delivery_confirmed_at",
     {"delivery_confirmed": 0, "delivery_confirmed_at": None, "delivery_confirmed_by": None}),
    ("is_undelivered", "is_undelivered", "undelivered_at",
     {"is_undelivered": 0, "undelivered_at": None, "undelivered_by": None,
      "undelivered_note": None, "shipping_loss_amount": 0}),
    ("is_problem_return", "is_problem_return", "problem_return_at",
     {"is_problem_return": 0, "problem_return_type": None, "product_loss_amount": 0,
      "problem_return_at": None, "problem_return_by": None, "problem_return_note": None,
      "problem_return_evidence": None}),
]

# predicate: a status timestamp that is non-empty AND earlier than the order's
# own creation => the status belongs to a previous, different-site occupant.
def _predate(ts_col):
    return (f"{ts_col} IS NOT NULL AND {ts_col} <> '' "
            f"AND datetime({ts_col}) < datetime(replace(date_created,'T',' '))")


def report(conn):
    c = conn.cursor()
    total = c.execute("SELECT COUNT(*) FROM orders").fetchone()[0]
    bare = c.execute("SELECT COUNT(*) FROM orders WHERE instr(id,'-')=0").fetchone()[0]
    migrated = total - bare
    unmappable = c.execute(
        "SELECT COUNT(*) FROM orders o LEFT JOIN sites s ON s.url=o.source "
        "WHERE s.id IS NULL AND instr(o.id,'-')=0").fetchone()[0]
    print(f"  orders total={total}  to-migrate(bare)={bare}  already-migrated={migrated}  "
          f"unmappable-source={unmappable}")

    for tbl, col, src_join in [
        ("order_notes", "order_id", None),
        ("shipping_logs", "order_id", "source"),
        ("reconciliation_statement_orders", "order_id", None),
        ("orders_archive", "id", None),
    ]:
        try:
            n = c.execute(f"SELECT COUNT(*) FROM {tbl} WHERE instr({col},'-')=0").fetchone()[0]
            print(f"  {tbl}.{col}: {n} bare rows to re-key")
        except sqlite3.OperationalError as e:
            print(f"  {tbl}: skip ({e})")

    print("  --- stale fulfillment-status to reset (status predates creation) ---")
    for label, gate, ts, _ in STALE_GROUPS:
        n = c.execute(
            f"SELECT COUNT(*) FROM orders WHERE COALESCE({gate},0) NOT IN (0,'') "
            f"AND {_predate(ts)}").fetchone()[0]
        print(f"    {label}: {n}")
    # sample
    print("  --- sample (first 8 stale carrier/delivery) ---")
    for row in c.execute(
        "SELECT id, source, status, substr(date_created,1,16) dc, carrier_status, "
        "substr(carrier_status_at,1,16) cs_at, delivery_confirmed dconf "
        "FROM orders WHERE (" + _predate("carrier_status_at") + ") OR "
        "(delivery_confirmed=1 AND " + _predate("delivery_confirmed_at") + ") "
        "ORDER BY date_created DESC LIMIT 8"):
        print("    ", tuple(row))


def migrate(conn):
    c = conn.cursor()
    # 1) add woo_id column if missing
    cols = {r[1] for r in c.execute("PRAGMA table_info(orders)")}
    if "woo_id" not in cols:
        print("  + ALTER orders ADD COLUMN woo_id INTEGER")
        c.execute("ALTER TABLE orders ADD COLUMN woo_id INTEGER")
    acols = {r[1] for r in c.execute("PRAGMA table_info(orders_archive)")} \
        if c.execute("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='orders_archive'").fetchone()[0] else set()
    if acols and "woo_id" not in acols:
        c.execute("ALTER TABLE orders_archive ADD COLUMN woo_id INTEGER")

    # 2) orders: woo_id <- id ; id <- "<site_id>-<id>"  (only bare, mappable rows)
    n = c.execute("""
        UPDATE orders
           SET woo_id = CAST(id AS INTEGER),
               id = (SELECT s.id FROM sites s WHERE s.url = orders.source) || '-' || id
         WHERE instr(id,'-')=0
           AND source IN (SELECT url FROM sites)
    """).rowcount
    print(f"  orders re-keyed: {n}")

    # 2b) orders_archive (same scheme; usually empty)
    if acols:
        na = c.execute("""
            UPDATE orders_archive
               SET woo_id = CAST(id AS INTEGER),
                   id = (SELECT s.id FROM sites s WHERE s.url = orders_archive.source) || '-' || id
             WHERE instr(id,'-')=0 AND source IN (SELECT url FROM sites)
        """).rowcount
        print(f"  orders_archive re-keyed: {na}")

    # 3) order_notes: map numeric order_id -> surrogate via orders.woo_id (1:1 at migration time)
    nn = c.execute("""
        UPDATE order_notes
           SET order_id = (SELECT o.id FROM orders o WHERE o.woo_id = order_notes.order_id LIMIT 1)
         WHERE instr(order_id,'-')=0
           AND order_id IN (SELECT woo_id FROM orders)
    """).rowcount
    print(f"  order_notes re-keyed: {nn}  (orphans left as-is)")

    # 4) shipping_logs: re-key using its own source + the EXISTING order_id.
    #    shipping_logs.order_id historically held the WC post id (= orders.id),
    #    NOT woo_order_id (which stores order['number'] and differs on
    #    Sequential-Order-Number sites). Suffix must be order_id so the
    #    shipping_logs.order_id = orders.id join stays intact.
    ns = c.execute("""
        UPDATE shipping_logs
           SET order_id = (SELECT s.id FROM sites s WHERE s.url = shipping_logs.source) || '-' || order_id
         WHERE instr(order_id,'-')=0 AND source IN (SELECT url FROM sites)
    """).rowcount
    print(f"  shipping_logs re-keyed: {ns}")

    # 5) reconciliation_statement_orders: best-effort via orders.woo_id
    nr = c.execute("""
        UPDATE reconciliation_statement_orders
           SET order_id = (SELECT o.id FROM orders o WHERE o.woo_id = reconciliation_statement_orders.order_id LIMIT 1)
         WHERE instr(order_id,'-')=0
           AND order_id IN (SELECT woo_id FROM orders)
    """).rowcount
    print(f"  reconciliation_statement_orders re-keyed: {nr}  (orphans/unmatched left as-is)")

    # 6) clean stale fulfillment status inherited from a previous same-id occupant
    for label, gate, ts, resets in STALE_GROUPS:
        set_clause = ", ".join(
            f"{k} = {('NULL' if v is None else repr(v) if isinstance(v,str) else v)}"
            for k, v in resets.items())
        nc = c.execute(
            f"UPDATE orders SET {set_clause} "
            f"WHERE COALESCE({gate},0) NOT IN (0,'') AND {_predate(ts)}").rowcount
        print(f"  cleaned stale {label}: {nc}")


def verify(conn):
    c = conn.cursor()
    bad = c.execute("SELECT COUNT(*) FROM orders WHERE instr(id,'-')=0").fetchone()[0]
    nullwoo = c.execute("SELECT COUNT(*) FROM orders WHERE woo_id IS NULL").fetchone()[0]
    dup = c.execute("SELECT COUNT(*) FROM (SELECT id FROM orders GROUP BY id HAVING COUNT(*)>1)").fetchone()[0]
    print(f"  VERIFY: orders still-bare={bad}  null-woo_id={nullwoo}  duplicate-id={dup}")
    return bad == 0 and nullwoo == 0 and dup == 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--live", action="store_true", help="perform writes (default: dry-run)")
    ap.add_argument("--db", default=DB)
    args = ap.parse_args()

    conn = sqlite3.connect(args.db)
    conn.execute("PRAGMA foreign_keys=OFF")
    print(f"=== migrate_composite_id  db={args.db}  mode={'LIVE' if args.live else 'DRY-RUN'} ===")
    print("[before]")
    report(conn)

    if not args.live:
        print("\n(dry-run: no changes written. Re-run with --live to apply.)")
        return

    print("\n[migrating in one transaction]")
    try:
        conn.execute("BEGIN")
        migrate(conn)
        ok = verify(conn)
        if not ok:
            conn.rollback()
            print("!! verification FAILED -> rolled back, no changes written")
            sys.exit(1)
        conn.commit()
        print("  committed.")
    except Exception as e:
        conn.rollback()
        print(f"!! error -> rolled back: {e}")
        raise
    print("\n[after]")
    report(conn)


if __name__ == "__main__":
    main()
