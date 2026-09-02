import sqlite3

import product_clone_jobs

from product_clone_jobs import (
    claim_clone_job,
    enqueue_clone_job,
    get_clone_job,
    init_product_clone_jobs,
    recover_interrupted_jobs,
)
from product_clone_worker import process_clone_job, resolve_site_for_worker


def make_conn(tmp_path):
    conn = sqlite3.connect(tmp_path / "clone-jobs.db")
    conn.row_factory = sqlite3.Row
    init_product_clone_jobs(conn)
    return conn


def test_postgres_startup_verifies_migrated_table_without_ddl(monkeypatch):
    class Connection:
        def __init__(self):
            self.statements = []
            self.commits = 0

        def execute(self, statement):
            self.statements.append(statement)

        def commit(self):
            self.commits += 1

    monkeypatch.setattr(product_clone_jobs.db, "is_postgres_backend", lambda: True)
    conn = Connection()

    init_product_clone_jobs(conn)

    assert conn.statements == ["SELECT 1 FROM product_clone_jobs WHERE 1=0"]
    assert conn.commits == 1


def enqueue(conn, product_ids=(11, 12)):
    return enqueue_clone_job(
        conn,
        source_site_id=35,
        target_site_id=2,
        product_ids=list(product_ids),
        options={
            "include_variations": True,
            "include_images": True,
            "status_on_target": "draft",
        },
        target_url="https://target.example",
        created_by_id="1",
        created_by_name="管理员",
    )


def test_enqueue_claim_and_serialize(tmp_path):
    conn = make_conn(tmp_path)
    queued = enqueue(conn)
    assert queued["status"] == "queued"
    assert queued["product_ids"] == [11, 12]
    assert queued["options"]["status_on_target"] == "draft"
    assert queued["terminal"] is False

    claimed = claim_clone_job(conn, "test-worker")
    assert claimed["id"] == queued["id"]
    assert claimed["status"] == "running"
    assert claimed["worker_id"] == "test-worker"
    assert claim_clone_job(conn, "other-worker") is None
    conn.close()


def test_worker_persists_success_and_failure_progress(tmp_path):
    conn = make_conn(tmp_path)
    queued = enqueue(conn)
    job = claim_clone_job(conn, "test-worker")

    def resolve_site(_conn, site_id):
        return ({"id": site_id}, f"https://site-{site_id}.example", "ck", "cs")

    def clone_one(_su, _sck, _scs, _tu, _tck, _tcs, product_id, options):
        assert options["status_on_target"] == "draft"
        if product_id == 12:
            return {"error": "simulated failure"}
        return {
            "new_id": 1011,
            "name": "Product 11",
            "sku": "SKU-11",
            "permalink": "https://target.example/product-11",
            "warnings": [],
        }

    results = process_clone_job(
        conn, job, clone_one=clone_one, resolve_site=resolve_site
    )
    assert len(results["success"]) == 1
    assert len(results["failed"]) == 1

    saved = get_clone_job(conn, queued["id"])
    assert saved["status"] == "partial_failed"
    assert saved["completed_count"] == 2
    assert saved["success_count"] == 1
    assert saved["failed_count"] == 1
    assert saved["terminal"] is True
    assert saved["results"]["failed"][0]["product_id"] == 12
    conn.close()


def test_running_jobs_are_not_automatically_retried(tmp_path):
    conn = make_conn(tmp_path)
    queued = enqueue(conn, product_ids=(11,))
    claim_clone_job(conn, "dead-worker")

    assert recover_interrupted_jobs(conn) == 1
    saved = get_clone_job(conn, queued["id"])
    assert saved["status"] == "interrupted"
    assert saved["terminal"] is True
    assert "不会自动重试" in saved["last_error"]
    assert claim_clone_job(conn, "new-worker") is None
    conn.close()


def test_worker_site_resolver_does_not_require_flask_current_user(tmp_path):
    conn = make_conn(tmp_path)
    conn.execute(
        """CREATE TABLE sites (
               id INTEGER PRIMARY KEY, url TEXT, consumer_key TEXT,
               consumer_secret TEXT, product_master_id INTEGER, manager TEXT
           )"""
    )
    conn.execute(
        "INSERT INTO sites VALUES (35, 'https://source.example', 'ck', 'cs', NULL, 'Owner')"
    )
    conn.commit()

    site, api_url, ck, cs = resolve_site_for_worker(
        conn,
        35,
        get_api_endpoint=lambda _conn, row: (
            row["url"], row["consumer_key"], row["consumer_secret"]
        ),
    )
    assert site["id"] == 35
    assert (api_url, ck, cs) == ("https://source.example", "ck", "cs")
    conn.close()
