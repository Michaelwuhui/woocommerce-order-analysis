"""Additive reconciliation migration. Never run implicitly on application import."""


def migrate(conn):
    statements = [
        """CREATE TABLE IF NOT EXISTS rec_objects (
            id TEXT PRIMARY KEY, kind TEXT NOT NULL, name TEXT NOT NULL,
            data TEXT NOT NULL, created_at TEXT NOT NULL, actor TEXT NOT NULL)""",
        "CREATE INDEX IF NOT EXISTS rec_objects_kind ON rec_objects(kind)",
        """CREATE TABLE IF NOT EXISTS rec_batches (
            batch_id INTEGER PRIMARY KEY, pool_id TEXT NOT NULL REFERENCES rec_objects(id),
            unit_cost TEXT, currency TEXT NOT NULL, reserved INTEGER NOT NULL DEFAULT 0,
            CHECK(reserved>=0))""",
        """CREATE TABLE IF NOT EXISTS rec_orders (
            order_id TEXT PRIMARY KEY, data TEXT NOT NULL, created_at TEXT NOT NULL)""",
        """CREATE TABLE IF NOT EXISTS rec_sources (
            id TEXT PRIMARY KEY, fulfillment_item_id INTEGER NOT NULL,
            order_id TEXT NOT NULL, batch_id INTEGER NOT NULL REFERENCES rec_batches(batch_id),
            quantity INTEGER NOT NULL, reserved INTEGER NOT NULL,
            shipped INTEGER NOT NULL DEFAULT 0, returned INTEGER NOT NULL DEFAULT 0,
            snapshot TEXT NOT NULL, CHECK(reserved>=0), CHECK(shipped>=returned))""",
        "CREATE INDEX IF NOT EXISTS rec_sources_item ON rec_sources(fulfillment_item_id)",
        "CREATE INDEX IF NOT EXISTS rec_sources_order ON rec_sources(order_id)",
        """CREATE TABLE IF NOT EXISTS rec_actions (
            id TEXT PRIMARY KEY, kind TEXT NOT NULL, order_id TEXT,
            data TEXT NOT NULL, created_at TEXT NOT NULL, actor TEXT NOT NULL)""",
        """CREATE TABLE IF NOT EXISTS rec_entries (
            id TEXT PRIMARY KEY, event_key TEXT NOT NULL, party_id TEXT NOT NULL,
            order_id TEXT, site TEXT, category TEXT NOT NULL,
            amount TEXT NOT NULL, currency TEXT NOT NULL, occurred_at TEXT NOT NULL,
            data TEXT NOT NULL, statement_id TEXT,
            UNIQUE(event_key,party_id,category,currency))""",
        "CREATE INDEX IF NOT EXISTS rec_entries_party ON rec_entries(party_id,currency,occurred_at)",
        """CREATE TABLE IF NOT EXISTS rec_statements (
            id TEXT PRIMARY KEY, party_id TEXT NOT NULL, currency TEXT NOT NULL,
            period_start TEXT NOT NULL, period_end TEXT NOT NULL,
            status TEXT NOT NULL, snapshot TEXT NOT NULL, created_at TEXT NOT NULL)""",
        """CREATE TABLE IF NOT EXISTS rec_cash (
            id TEXT PRIMARY KEY, party_id TEXT NOT NULL, currency TEXT NOT NULL,
            amount TEXT NOT NULL, occurred_at TEXT NOT NULL, reference TEXT NOT NULL,
            data TEXT NOT NULL, UNIQUE(party_id,currency,reference))""",
        """CREATE TABLE IF NOT EXISTS rec_cash_allocations (
            cash_id TEXT NOT NULL REFERENCES rec_cash(id),
            statement_id TEXT NOT NULL REFERENCES rec_statements(id),
            amount TEXT NOT NULL, PRIMARY KEY(cash_id,statement_id))""",
        """CREATE TABLE IF NOT EXISTS rec_user_parties (
            user_id INTEGER NOT NULL, party_id TEXT NOT NULL REFERENCES rec_objects(id),
            PRIMARY KEY(user_id,party_id))""",
    ]
    for sql in statements:
        conn.execute(sql)
    conn.commit()
