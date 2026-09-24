-- Clean synchronization uses the same exclusive, persistent run ledger as
-- quick/deep synchronization. No clean schedule is enabled by this migration.
SET ROLE woo_analysis_owner;
SET search_path TO public;

ALTER TABLE sync_runs DROP CONSTRAINT IF EXISTS sync_runs_mode_check;
ALTER TABLE sync_runs ADD CONSTRAINT sync_runs_mode_check
    CHECK (mode IN ('quick', 'auto', 'deep', 'clean'));

-- Woo order notes can contain operator/customer history. Preserve them with
-- the order; order_note_sync_state is only a resumable fetch cursor.
CREATE TABLE IF NOT EXISTS order_notes_archive (
    id bigint PRIMARY KEY,
    wc_note_id bigint,
    order_id text NOT NULL,
    note text,
    date_created text,
    customer_note boolean,
    author text,
    added_by_user boolean,
    archived_at timestamptz NOT NULL DEFAULT CURRENT_TIMESTAMP,
    archive_reason text NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_order_notes_archive_order
    ON order_notes_archive (order_id);
GRANT SELECT, INSERT, UPDATE, DELETE ON order_notes_archive TO woo_analysis_app;

RESET ROLE;
