"""Explicit additive migration; does not authorize users or change inventory."""
from inv_common import get_conn


def migrate(conn):
    pk = 'BIGSERIAL PRIMARY KEY' if hasattr(conn, '_raw') else 'INTEGER PRIMARY KEY AUTOINCREMENT'
    conn.execute(f'''CREATE TABLE IF NOT EXISTS inv_temporary_stock_grants (
        id {pk}, user_id INTEGER NOT NULL REFERENCES users(id),
        warehouse_id INTEGER NOT NULL REFERENCES warehouses(id),
        expires_at TEXT NOT NULL, reason TEXT NOT NULL,
        created_by INTEGER NOT NULL REFERENCES users(id), created_name TEXT NOT NULL,
        created_at TEXT NOT NULL, request_key TEXT NOT NULL UNIQUE,
        revoked_at TEXT, revoked_by INTEGER REFERENCES users(id), revoke_reason TEXT)''')
    conn.execute('''CREATE INDEX IF NOT EXISTS idx_inv_temporary_stock_grants_scope
        ON inv_temporary_stock_grants(user_id,warehouse_id,expires_at) WHERE revoked_at IS NULL''')
    conn.commit()


if __name__ == '__main__':
    conn = get_conn()
    try:
        migrate(conn)
        print('Temporary inventory grant schema ready; no user grants or stock changed.')
    finally:
        conn.close()
