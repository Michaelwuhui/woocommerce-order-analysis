-- Temporary network/domain outages never disable a site or remove its orders.
SET ROLE woo_analysis_owner;
SET search_path TO public;

CREATE TABLE IF NOT EXISTS sync_site_health (
    site_id bigint PRIMARY KEY REFERENCES sites(id) ON DELETE CASCADE,
    failure_kind text NOT NULL CHECK (failure_kind IN (
        'dns','tls','timeout','connection','http','redirect','html','json'
    )),
    failure_count integer NOT NULL CHECK (failure_count > 0),
    next_check_at timestamptz NOT NULL,
    last_error text NOT NULL,
    last_failure_at timestamptz NOT NULL DEFAULT CURRENT_TIMESTAMP,
    last_failure_run_id uuid NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_sync_site_health_recheck
    ON sync_site_health(next_check_at,site_id);
GRANT SELECT,INSERT,UPDATE,DELETE ON sync_site_health TO woo_analysis_app;

RESET ROLE;
