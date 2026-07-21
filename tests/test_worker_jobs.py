import sqlite3
import unittest
from unittest.mock import patch

from fulfillment_service import enqueue_job
from fulfillment_worker import claim_job, finish_job, handle_verify_submission, retry_job


class DurableJobTests(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(":memory:")
        self.db.row_factory = sqlite3.Row
        self.db.executescript(
            """
            CREATE TABLE oms_integration_jobs (
              id INTEGER PRIMARY KEY AUTOINCREMENT, job_type TEXT NOT NULL,
              aggregate_type TEXT, aggregate_id TEXT, idempotency_key TEXT NOT NULL UNIQUE,
              payload_json TEXT, payload_hash TEXT, status TEXT NOT NULL DEFAULT 'pending',
              attempts INTEGER NOT NULL DEFAULT 0, max_attempts INTEGER NOT NULL DEFAULT 10,
              available_at TEXT DEFAULT CURRENT_TIMESTAMP, locked_at TEXT, locked_by TEXT,
              lease_expires_at TEXT, last_error_code TEXT, last_error TEXT,
              created_at TEXT DEFAULT CURRENT_TIMESTAMP, updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
              completed_at TEXT
            );
            """
        )

    def tearDown(self):
        self.db.close()

    def test_idempotent_enqueue_claim_retry_and_finish(self):
        first = enqueue_job(self.db, "TEST", "shipment", "S1", "same-key", {"n": 1})
        second = enqueue_job(self.db, "TEST", "shipment", "S1", "same-key", {"n": 1})
        self.db.commit()
        self.assertEqual(first, second)
        self.assertEqual(1, self.db.execute("SELECT COUNT(*) FROM oms_integration_jobs").fetchone()[0])

        job = claim_job(self.db)
        self.assertEqual("running", job["status"])
        self.assertEqual(1, job["attempts"])
        retry_job(self.db, job, RuntimeError("temporary"), delay_seconds=-1, code="temporary")
        retried = claim_job(self.db)
        self.assertEqual(2, retried["attempts"])
        finish_job(self.db, retried["id"], {"ok": True})
        row = self.db.execute("SELECT * FROM oms_integration_jobs WHERE id=?", (retried["id"],)).fetchone()
        self.assertEqual("succeeded", row["status"])
        self.assertIsNotNone(row["completed_at"])

    def test_expired_lease_is_reclaimed(self):
        jid = enqueue_job(self.db, "TEST", "shipment", "S2", "lease-key", {})
        self.db.execute(
            """UPDATE oms_integration_jobs SET status='running', attempts=1,
                      lease_expires_at='2000-01-01T00:00:00+00:00' WHERE id=?""",
            (jid,),
        )
        self.db.commit()
        job = claim_job(self.db)
        self.assertEqual(jid, job["id"])
        self.assertEqual(2, job["attempts"])

    def test_unknown_wms_submission_never_auto_resubmits(self):
        fulfillment = {
            "id": "F1",
            "order_id": "O1",
            "status": "submission_unknown",
            "external_invoice_code": "INV-1",
        }

        class Result:
            def fetchone(self):
                return fulfillment

        class Conn:
            committed = False

            def execute(self, _sql, _params=()):
                return Result()

            def commit(self):
                self.committed = True

        conn = Conn()
        job = {"id": 9, "aggregate_id": "F1", "attempts": 3}
        with patch("fulfillment_worker._query_wms_fulfillment", return_value=(None, "corr")), \
             patch("fulfillment_worker.transition_fulfillment") as transition, \
             patch("fulfillment_worker.mark_manual_review") as manual, \
             patch("fulfillment_worker.enqueue_job") as enqueue:
            result = handle_verify_submission(conn, job, {"fulfillment_id": "F1"})

        self.assertEqual({"not_found": True, "manual_review": True, "resubmit": False}, result)
        transition.assert_called_once()
        self.assertEqual("manual_review", transition.call_args.args[2])
        manual.assert_called_once()
        enqueue.assert_not_called()
        self.assertTrue(conn.committed)


if __name__ == "__main__":
    unittest.main()
