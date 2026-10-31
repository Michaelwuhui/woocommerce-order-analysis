-- Frozen product-source bills: immutable revisions, never rewrite locked history.
SET ROLE woo_analysis_owner;
SET search_path TO public;
CREATE TABLE IF NOT EXISTS reconciliation_statement_versions (
    statement_id bigint NOT NULL,
    revision integer NOT NULL CHECK (revision > 0),
    summary_json text NOT NULL,
    lines_json text NOT NULL,
    fees_json text NOT NULL,
    previous_json text NOT NULL,
    created_at text NOT NULL,
    actor_id bigint,
    settlement_ready integer NOT NULL DEFAULT 1,
    packing_fee_cny text NOT NULL DEFAULT '0.00',
    pending_orders integer NOT NULL DEFAULT 0,
    PRIMARY KEY (statement_id, revision)
);
GRANT SELECT, INSERT ON reconciliation_statement_versions TO woo_analysis_app;
RESET ROLE;
