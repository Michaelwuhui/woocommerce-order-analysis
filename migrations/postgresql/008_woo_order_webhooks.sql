-- WooCommerce 实时订单 Webhook 注册管理与投递收件箱表（配合每小时增量同步兜底）
SET ROLE woo_analysis_owner;
SET search_path TO public;

CREATE TABLE IF NOT EXISTS woo_site_webhooks (
    site_id bigint NOT NULL REFERENCES sites(id) ON DELETE CASCADE,
    topic text NOT NULL CHECK (topic IN ('order.created', 'order.updated')),
    secret text NOT NULL,
    previous_secret text,
    remote_webhook_id bigint,
    delivery_url text NOT NULL,
    remote_status text NOT NULL DEFAULT 'unknown',
    last_synced_at timestamptz,
    last_ping_at timestamptz,
    last_delivery_at timestamptz,
    last_error text,
    created_at timestamptz NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at timestamptz NOT NULL DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (site_id, topic)
);

CREATE INDEX IF NOT EXISTS idx_woo_site_webhooks_status
    ON woo_site_webhooks(remote_status, last_synced_at);

CREATE TABLE IF NOT EXISTS woo_webhook_deliveries (
    id bigserial PRIMARY KEY,
    delivery_key text NOT NULL UNIQUE,
    site_id bigint NOT NULL REFERENCES sites(id) ON DELETE CASCADE,
    woo_order_id bigint,
    topic text NOT NULL,
    event text,
    delivery_id text,
    payload_hash text NOT NULL,
    status text NOT NULL DEFAULT 'queued' CHECK (status IN (
        'ping', 'queued', 'coalesced', 'fetching', 'synced', 'ignored', 'error'
    )),
    order_oid text,
    order_status text,
    latency_ms integer,
    error_message text,
    received_at timestamptz NOT NULL DEFAULT CURRENT_TIMESTAMP,
    processed_at timestamptz
);

CREATE INDEX IF NOT EXISTS idx_woo_webhook_deliveries_received
    ON woo_webhook_deliveries(received_at DESC);

CREATE INDEX IF NOT EXISTS idx_woo_webhook_deliveries_site_received
    ON woo_webhook_deliveries(site_id, received_at DESC);

CREATE INDEX IF NOT EXISTS idx_woo_webhook_deliveries_pending
    ON woo_webhook_deliveries(status, received_at)
    WHERE status IN ('queued', 'fetching');

INSERT INTO settings (key, value)
VALUES ('woo_webhook_sync_enabled', 'true')
ON CONFLICT (key) DO NOTHING;

GRANT SELECT, INSERT, UPDATE, DELETE ON woo_site_webhooks, woo_webhook_deliveries TO woo_analysis_app;
GRANT USAGE, SELECT ON SEQUENCE woo_webhook_deliveries_id_seq TO woo_analysis_app;

RESET ROLE;
