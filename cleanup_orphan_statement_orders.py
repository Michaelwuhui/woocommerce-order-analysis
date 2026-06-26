#!/usr/bin/env python3
"""One-time cleanup: remove orphan rows in reconciliation_statement_orders.

Background: api_delete_statement used to delete only the reconciliation_statements
row, leaving the frozen order snapshot rows behind (the schema declares
ON DELETE CASCADE, but PRAGMA foreign_keys is never enabled on the connection,
so the cascade never fired). This script deletes snapshot rows whose statement_id
no longer exists in reconciliation_statements.

Safe by default: prints what it WOULD delete. Pass --apply to actually delete.
reconciliation_audit_log is intentionally NOT touched.

Usage:
    python3 cleanup_orphan_statement_orders.py          # dry run
    python3 cleanup_orphan_statement_orders.py --apply   # really delete
"""
import sqlite3
import sys

DB_FILE = '/www/wwwroot/woo-analysis/woocommerce_orders.db'

ORPHAN_FILTER = (
    'statement_id NOT IN (SELECT id FROM reconciliation_statements)'
)


def main():
    apply = '--apply' in sys.argv[1:]
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            f'''SELECT statement_id, COUNT(*) AS n
                FROM reconciliation_statement_orders
                WHERE {ORPHAN_FILTER}
                GROUP BY statement_id
                ORDER BY statement_id'''
        ).fetchall()

        if not rows:
            print('No orphan snapshot rows. Nothing to do.')
            return

        total = sum(r['n'] for r in rows)
        print(f'Orphan statement_ids: {len(rows)}  |  orphan rows: {total}')
        for r in rows:
            print(f'  statement_id={r["statement_id"]:>4}  rows={r["n"]}')

        if not apply:
            print('\nDRY RUN — nothing deleted. Re-run with --apply to delete.')
            return

        cur = conn.execute(
            f'DELETE FROM reconciliation_statement_orders WHERE {ORPHAN_FILTER}'
        )
        conn.commit()
        print(f'\nDeleted {cur.rowcount} orphan rows.')

        remaining = conn.execute(
            f'''SELECT COUNT(*) FROM reconciliation_statement_orders
                WHERE {ORPHAN_FILTER}'''
        ).fetchone()[0]
        print(f'Remaining orphan rows: {remaining}')
    finally:
        conn.close()


if __name__ == '__main__':
    main()
