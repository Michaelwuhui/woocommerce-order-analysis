"""Versioned additive migration, invoked by a schema owner, never at app startup.

python stock_sync_schema.py
PostgreSQL runtime role needs SELECT/INSERT/UPDATE on stock_sync_* tables and
DELETE on stock_sync_resource_leases. Existing inventory tables remain read-only.
"""
import re
from stock_sync_common import connect, stamp

VERSION = '001_cross_site_stock_sync'


def migrate(c):
    ts = 'TIMESTAMPTZ' if hasattr(c, '_raw') else 'TEXT'
    statements = [
        f'CREATE TABLE IF NOT EXISTS stock_sync_schema_migrations(version TEXT PRIMARY KEY,applied_at {ts} NOT NULL)',
        f'''CREATE TABLE IF NOT EXISTS stock_sync_reference_sites(
            site_id INTEGER PRIMARY KEY,enabled INTEGER NOT NULL CHECK(enabled IN (0,1)),
            version INTEGER NOT NULL,actor_id INTEGER NOT NULL,updated_at {ts} NOT NULL)''',
        f'''CREATE TABLE IF NOT EXISTS stock_sync_site_capabilities(
            site_id INTEGER PRIMARY KEY,independent_write INTEGER NOT NULL DEFAULT 0,
            evidence TEXT NOT NULL,version INTEGER NOT NULL,actor_id INTEGER NOT NULL,updated_at {ts} NOT NULL)''',
        f'''CREATE TABLE IF NOT EXISTS stock_sync_bindings(
            map_id INTEGER PRIMARY KEY,site_id INTEGER NOT NULL,sku_id INTEGER NOT NULL,
            mapping_hash TEXT NOT NULL,baseline_json TEXT NOT NULL,policy TEXT NOT NULL,
            version INTEGER NOT NULL,updated_at {ts} NOT NULL)''',
        f'''CREATE TABLE IF NOT EXISTS stock_sync_controls(
            id TEXT PRIMARY KEY,map_id INTEGER NOT NULL,site_id INTEGER NOT NULL,sku_id INTEGER NOT NULL,
            mapping_hash TEXT NOT NULL,kind TEXT NOT NULL CHECK(kind IN ('manual_hold','manual_availability','reference_snapshot')),
            stock_status TEXT NOT NULL,protection TEXT NOT NULL,source_json TEXT NOT NULL DEFAULT '{{}}',
            reason TEXT NOT NULL,active INTEGER NOT NULL CHECK(active IN (0,1)),version INTEGER NOT NULL,
            actor_id INTEGER NOT NULL,created_at {ts} NOT NULL,released_by INTEGER,released_at {ts})''',
        "CREATE UNIQUE INDEX IF NOT EXISTS stock_sync_reference_active ON stock_sync_controls(map_id,kind) WHERE active=1 AND kind IN ('reference_snapshot','manual_availability')",
        'CREATE INDEX IF NOT EXISTS stock_sync_controls_map ON stock_sync_controls(map_id,active)',
        f'''CREATE TABLE IF NOT EXISTS stock_sync_catalog_snapshots(
            id TEXT PRIMARY KEY,actor_id INTEGER NOT NULL,site_id INTEGER,
            scope_json TEXT NOT NULL,status TEXT NOT NULL,complete INTEGER NOT NULL DEFAULT 0,
            items_json TEXT NOT NULL DEFAULT '[]',progress INTEGER NOT NULL DEFAULT 0,error TEXT,
            created_at {ts} NOT NULL,observed_at {ts},expires_at {ts} NOT NULL)''',
        f'''CREATE TABLE IF NOT EXISTS stock_sync_plans(
            id TEXT PRIMARY KEY,actor_id INTEGER NOT NULL,request_json TEXT NOT NULL,status TEXT NOT NULL,
            version INTEGER NOT NULL DEFAULT 1,summary_json TEXT NOT NULL DEFAULT '{{}}',error TEXT,
            created_at {ts} NOT NULL,expires_at {ts} NOT NULL)''',
        '''CREATE TABLE IF NOT EXISTS stock_sync_plan_items(
            id TEXT PRIMARY KEY,plan_id TEXT NOT NULL,resource_key TEXT NOT NULL,map_id INTEGER,
            site_id INTEGER NOT NULL,decision TEXT NOT NULL,detail_json TEXT NOT NULL,
            UNIQUE(plan_id,resource_key))''',
        f'''CREATE TABLE IF NOT EXISTS stock_sync_jobs(
            id TEXT PRIMARY KEY,plan_id TEXT NOT NULL UNIQUE,actor_id INTEGER NOT NULL,
            idempotency_key TEXT NOT NULL,request_hash TEXT NOT NULL,status TEXT NOT NULL,
            cancel_requested INTEGER NOT NULL DEFAULT 0,created_at {ts} NOT NULL,completed_at {ts},
            UNIQUE(actor_id,idempotency_key))''',
        f'''CREATE TABLE IF NOT EXISTS stock_sync_job_items(
            id TEXT PRIMARY KEY,job_id TEXT NOT NULL,plan_item_id TEXT NOT NULL,resource_key TEXT NOT NULL,
            status TEXT NOT NULL,attempts INTEGER NOT NULL DEFAULT 0,intent_json TEXT,
            before_json TEXT,after_json TEXT,error TEXT,updated_at {ts} NOT NULL,
            UNIQUE(job_id,resource_key))''',
        f'''CREATE TABLE IF NOT EXISTS stock_sync_events(
            id TEXT PRIMARY KEY,kind TEXT NOT NULL,actor_id INTEGER,object_id TEXT NOT NULL,
            detail_json TEXT NOT NULL,created_at {ts} NOT NULL)''',
        f'''CREATE TABLE IF NOT EXISTS stock_sync_work(
            id TEXT PRIMARY KEY,kind TEXT NOT NULL,object_id TEXT NOT NULL UNIQUE,status TEXT NOT NULL,
            worker_id TEXT,heartbeat_at {ts},created_at {ts} NOT NULL)''',
        'CREATE INDEX IF NOT EXISTS stock_sync_work_queue ON stock_sync_work(status,created_at)',
        'CREATE INDEX IF NOT EXISTS stock_sync_items_job ON stock_sync_job_items(job_id,status)',
        f'''CREATE TABLE IF NOT EXISTS stock_sync_resource_leases(
            resource_key TEXT PRIMARY KEY,owner TEXT NOT NULL,version INTEGER NOT NULL,
            heartbeat_at {ts} NOT NULL,quarantined INTEGER NOT NULL DEFAULT 0)''',
        f'''CREATE TABLE IF NOT EXISTS stock_sync_rate_limits(
            endpoint_key TEXT PRIMARY KEY,retry_at {ts} NOT NULL,reason TEXT NOT NULL)''',
    ]
    for statement in statements:
        if hasattr(c, '_raw'):
            # DDL must not pass through the legacy adapter's column-name-based
            # boolean rewriting. Use native types on PostgreSQL.
            for flag in ('enabled','active','complete','cancel_requested','quarantined','independent_write'):
                statement = re.sub(r'\b'+flag+r' INTEGER NOT NULL DEFAULT 0', flag+' BOOLEAN NOT NULL DEFAULT FALSE', statement)
                statement = statement.replace(flag+' INTEGER NOT NULL CHECK('+flag+' IN (0,1))',flag+' BOOLEAN NOT NULL')
            statement = statement.replace('WHERE active=1','WHERE active=TRUE')
            c._raw.execute(statement)
        else:
            c.execute(statement)
    if hasattr(c, '_metadata_cache'):
        # The adapter can have inspected the old schema before this migration.
        c._metadata_cache.clear()
    c.execute('INSERT INTO stock_sync_schema_migrations VALUES(?,?) ON CONFLICT(version) DO NOTHING', (VERSION, stamp()))
    c.commit()


if __name__ == '__main__':
    c = connect()
    try:
        migrate(c)
        print(VERSION + ': migrated; feature remains disabled unless explicitly enabled')
    finally:
        c.close()
