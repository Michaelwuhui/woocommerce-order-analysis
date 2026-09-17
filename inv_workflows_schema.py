"""Additive inventory workflow schema. Run explicitly; never during app import."""
from inv_common import get_conn


def migrate(conn):
    pk = 'BIGSERIAL PRIMARY KEY' if hasattr(conn, '_raw') else 'INTEGER PRIMARY KEY AUTOINCREMENT'
    statements = [
        '''CREATE TABLE IF NOT EXISTS inv_operation_permissions (
            user_id INTEGER NOT NULL REFERENCES users(id),
            warehouse_id INTEGER NOT NULL REFERENCES warehouses(id),
            can_receive INTEGER NOT NULL DEFAULT 0 CHECK(can_receive IN (0,1)),
            can_transfer INTEGER NOT NULL DEFAULT 0 CHECK(can_transfer IN (0,1)),
            can_stocktake INTEGER NOT NULL DEFAULT 0 CHECK(can_stocktake IN (0,1)),
            can_approve INTEGER NOT NULL DEFAULT 0 CHECK(can_approve IN (0,1)),
            PRIMARY KEY(user_id,warehouse_id))''',
        f'''CREATE TABLE IF NOT EXISTS inv_documents (
            id {pk}, kind TEXT NOT NULL CHECK(kind IN ('replenishment','transfer','receipt','stocktake')),
            status TEXT NOT NULL, warehouse_id INTEGER NOT NULL REFERENCES warehouses(id),
            source_warehouse_id INTEGER REFERENCES warehouses(id),
            source_authority TEXT,
            parent_id INTEGER REFERENCES inv_documents(id),
            reference TEXT NOT NULL, note TEXT NOT NULL,
            request_key TEXT NOT NULL UNIQUE, request_hash TEXT NOT NULL,
            created_by INTEGER NOT NULL REFERENCES users(id), created_name TEXT NOT NULL,
            reviewed_by INTEGER REFERENCES users(id), reviewed_name TEXT, review_note TEXT,
            created_at TEXT NOT NULL, reviewed_at TEXT)''',
        f'''CREATE TABLE IF NOT EXISTS inv_document_lines (
            id {pk}, document_id INTEGER NOT NULL REFERENCES inv_documents(id),
            sku_id INTEGER NOT NULL REFERENCES inv_skus(id),
            qty INTEGER NOT NULL DEFAULT 0 CHECK(qty>=0),
            damaged_qty INTEGER NOT NULL DEFAULT 0 CHECK(damaged_qty>=0),
            short_qty INTEGER NOT NULL DEFAULT 0 CHECK(short_qty>=0),
            baseline INTEGER NOT NULL DEFAULT 0, baseline_reserved INTEGER NOT NULL DEFAULT 0,
            baseline_movement INTEGER NOT NULL DEFAULT 0,
            UNIQUE(document_id,sku_id))''',
        f'''CREATE TABLE IF NOT EXISTS inv_document_events (
            id {pk}, document_id INTEGER REFERENCES inv_documents(id),
            action TEXT NOT NULL, actor_id INTEGER NOT NULL, actor_name TEXT NOT NULL,
            detail TEXT NOT NULL, created_at TEXT NOT NULL)''',
        'CREATE INDEX IF NOT EXISTS idx_inv_documents_warehouse ON inv_documents(warehouse_id,id)',
        'CREATE INDEX IF NOT EXISTS idx_inv_documents_parent ON inv_documents(parent_id,status)',
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_inv_document_reference ON inv_documents(kind,warehouse_id,reference) WHERE kind IN ('replenishment','transfer','stocktake') AND status NOT IN ('rejected','cancelled')",
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_inv_receipt_reference ON inv_documents(parent_id,reference) WHERE kind='receipt' AND status NOT IN ('rejected','cancelled')",
    ]
    for sql in statements:
        conn.execute(sql)
    # Existing shippers can report receipts/counts only in their existing local-stock warehouses.
    # Never grant approval, transfer, global inventory management, or change physical stock.
    conn.execute('''INSERT INTO inv_operation_permissions
        (user_id,warehouse_id,can_receive,can_stocktake)
        SELECT p.user_id,p.warehouse_id,1,1 FROM oms_warehouse_user_permissions p
        JOIN oms_warehouse_integrations wi ON wi.warehouse_id=p.warehouse_id
        WHERE p.can_ship=1 AND wi.inventory_authority='local'
        ON CONFLICT(user_id,warehouse_id) DO NOTHING''')
    conn.commit()


if __name__ == '__main__':
    connection = get_conn()
    try:
        migrate(connection)
        print('Inventory workflow schema ready; no stock or order changes.')
    finally:
        connection.close()
