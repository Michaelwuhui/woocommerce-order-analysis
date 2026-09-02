-- Persist the local follow-up hooks which run after an authoritative page
-- write.  A worker crash between the page COMMIT and those hooks must not
-- silently lose fulfillment planning or notification enqueueing.

SET ROLE woo_analysis_owner;
SET search_path TO public;

ALTER TABLE sync_page_receipts
    ADD COLUMN IF NOT EXISTS planning_candidates jsonb NOT NULL DEFAULT '[]'::jsonb,
    ADD COLUMN IF NOT EXISTS post_commit_status text NOT NULL DEFAULT 'pending',
    ADD COLUMN IF NOT EXISTS post_commit_attempts integer NOT NULL DEFAULT 0,
    ADD COLUMN IF NOT EXISTS post_commit_heartbeat_at timestamptz,
    ADD COLUMN IF NOT EXISTS post_commit_finished_at timestamptz,
    ADD COLUMN IF NOT EXISTS post_commit_error text;

DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1
        FROM pg_constraint
        WHERE conrelid = 'sync_page_receipts'::regclass
          AND conname = 'ck_sync_page_receipts_post_commit_status'
    ) THEN
        ALTER TABLE sync_page_receipts
            ADD CONSTRAINT ck_sync_page_receipts_post_commit_status
            CHECK (post_commit_status IN (
                'pending','processing','completed','skipped','error'
            ));
    END IF;
END
$$;

DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1
        FROM pg_constraint
        WHERE conrelid = 'sync_page_receipts'::regclass
          AND conname = 'ck_sync_page_receipts_post_commit_attempts'
    ) THEN
        ALTER TABLE sync_page_receipts
            ADD CONSTRAINT ck_sync_page_receipts_post_commit_attempts
            CHECK (post_commit_attempts >= 0);
    END IF;
END
$$;

CREATE INDEX IF NOT EXISTS idx_sync_page_receipts_post_commit_recovery
ON sync_page_receipts (post_commit_status, post_commit_heartbeat_at, committed_at)
WHERE post_commit_status IN ('pending','processing','error');

GRANT SELECT, INSERT, UPDATE, DELETE ON sync_page_receipts TO woo_analysis_app;

RESET ROLE;
