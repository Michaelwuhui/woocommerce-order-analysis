"""Print compact invariants for a database copy after migration up/down."""

import sqlite3
import sys


db_file = sys.argv[1]
phase = sys.argv[2]
conn = sqlite3.connect(db_file)
try:
    print(f"integrity_{phase}={conn.execute('PRAGMA integrity_check').fetchone()[0]}")
    print(
        f"oms_tables_{phase}="
        + str(conn.execute(
            "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name LIKE 'oms_%'"
        ).fetchone()[0])
    )
    print(
        f"migrations_{phase}="
        + str(conn.execute(
            "SELECT group_concat(version, ',') FROM inv_schema_migrations"
        ).fetchone()[0])
    )
    if phase == "up":
        print(
            "flags_up="
            + str(conn.execute(
                "SELECT group_concat(key || '=' || value, ',') FROM settings WHERE key LIKE 'oms_%'"
            ).fetchone()[0])
        )
        print(
            "routes_up="
            + str(conn.execute(
                "SELECT group_concat(market_code || ':' || warehouse_id, ',') "
                "FROM inv_market_warehouses WHERE market_code IN ('PL','HU','CZ')"
            ).fetchone()[0])
        )
finally:
    conn.close()
