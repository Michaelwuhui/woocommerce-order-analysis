#!/usr/bin/env python3
"""Process-level recovery acceptance test using only local/test resources.

The caller must load the protected application environment and set
WOO_DB_NAME_OVERRIDE to a database ending in ``_test``.  The harness starts a
temporary Redis, a loopback-only WooCommerce mock, Celery workers, and a test
Gunicorn listener.  It never connects to a configured customer URL.
"""

from __future__ import annotations

import json
import os
import shutil
import signal
import socket
import subprocess
import sys
import tempfile
import threading
import time
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import urlopen


ROOT = Path(__file__).resolve().parents[1]
PYTHON = "/www/wwwroot/woo-analysis/venv/bin/python"
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def free_port() -> int:
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])
    finally:
        sock.close()


def wait_until(predicate, timeout=45.0, interval=0.1, label="condition"):
    deadline = time.monotonic() + timeout
    last_error = None
    while time.monotonic() < deadline:
        try:
            value = predicate()
            if value:
                return value
        except Exception as exc:  # services may still be starting
            last_error = exc
        time.sleep(interval)
    suffix = f": {type(last_error).__name__}" if last_error else ""
    raise AssertionError(f"timed out waiting for {label}{suffix}")


class MockState:
    def __init__(self):
        self.lock = threading.Lock()
        self.order_requests = 0
        self.first_seen = threading.Event()
        self.release_first = threading.Event()
        self.site_url = ""

    def next_order(self):
        with self.lock:
            self.order_requests += 1
            request_number = self.order_requests
        if request_number == 1:
            self.first_seen.set()
            self.release_first.wait(timeout=40)
        stamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
        return {
            "id": 992340001,
            "number": "process-recovery",
            "order_key": "pytest-sync-process-recovery",
            "status": "pending",
            "currency": "EUR",
            "date_created": stamp,
            "date_modified": stamp,
            "discount_total": "0.00",
            "shipping_total": "1.25",
            "total": f"{11 + request_number}.25",
            "total_tax": "0.00",
            "prices_include_tax": False,
            "billing": {},
            "shipping": {},
            "meta_data": [],
            "line_items": [],
            "tax_lines": [],
            "shipping_lines": [],
            "fee_lines": [],
            "coupon_lines": [],
            "refunds": [],
        }


def handler_for(state: MockState):
    class Handler(BaseHTTPRequestHandler):
        protocol_version = "HTTP/1.1"

        def do_GET(self):
            if self.path.split("?", 1)[0].endswith("/orders"):
                payload = [state.next_order()]
                headers = {"X-WP-TotalPages": "1"}
                self._send(payload, headers)
                return
            if "/notes" in self.path:
                self._send([])
                return
            self._send({"error": "not found"}, status=404)

        def _send(self, payload, headers=None, status=200):
            body = json.dumps(payload, separators=(",", ":")).encode("utf-8")
            try:
                self.send_response(status)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(body)))
                for key, value in (headers or {}).items():
                    self.send_header(key, value)
                self.end_headers()
                self.wfile.write(body)
            except (BrokenPipeError, ConnectionResetError):
                pass

        def log_message(self, _format, *_args):
            return

    return Handler


class ManagedProcess:
    def __init__(self, command, environment, log_path):
        self.log_handle = open(log_path, "ab", buffering=0)
        self.process = subprocess.Popen(
            command,
            cwd=ROOT,
            env=environment,
            stdin=subprocess.DEVNULL,
            stdout=self.log_handle,
            stderr=subprocess.STDOUT,
            start_new_session=True,
        )

    def assert_alive(self):
        if self.process.poll() is not None:
            raise AssertionError(f"process exited early with {self.process.returncode}")

    def stop(self, hard=False, timeout=15):
        if self.process.poll() is None:
            sig = signal.SIGKILL if hard else signal.SIGTERM
            try:
                os.killpg(self.process.pid, sig)
            except ProcessLookupError:
                pass
            try:
                self.process.wait(timeout=timeout)
            except subprocess.TimeoutExpired:
                os.killpg(self.process.pid, signal.SIGKILL)
                self.process.wait(timeout=10)
        self.log_handle.close()


def start_redis(environment, work_dir, port):
    command = [
        "/usr/bin/redis-server",
        "--bind", "127.0.0.1",
        "--port", str(port),
        "--protected-mode", "yes",
        "--dir", str(work_dir),
        "--dbfilename", "dump.rdb",
        "--appendonly", "yes",
        "--appendfsync", "everysec",
        "--save", "60", "1",
        "--maxmemory-policy", "noeviction",
        "--daemonize", "no",
    ]
    return ManagedProcess(command, environment, work_dir / "redis.log")


def start_worker(environment, work_dir, queue, name, concurrency=1):
    command = [
        PYTHON,
        "-m",
        "celery",
        "-A",
        "celery_app:celery_app",
        "worker",
        f"--queues={queue}",
        f"--concurrency={concurrency}",
        "--pool=prefork",
        "--prefetch-multiplier=1",
        f"--hostname={name}@%h",
        "--loglevel=INFO",
        "--without-gossip",
        "--without-mingle",
        "--without-heartbeat",
    ]
    return ManagedProcess(command, environment, work_dir / f"{name}.log")


def start_gunicorn(environment, work_dir, port):
    command = [
        str(ROOT / "venv/bin/gunicorn") if (ROOT / "venv/bin/gunicorn").exists() else "/www/wwwroot/woo-analysis/venv/bin/gunicorn",
        "--workers", "2",
        "--timeout", "30",
        "--bind", f"127.0.0.1:{port}",
        "app:app",
    ]
    return ManagedProcess(command, environment, work_dir / "gunicorn.log")


def main() -> int:
    database = os.getenv("WOO_DB_NAME_OVERRIDE", "")
    if not database.endswith("_test"):
        raise RuntimeError("WOO_DB_NAME_OVERRIDE must name an isolated *_test database")
    if os.getenv("WOO_DB_BACKEND", "").lower() not in {"postgres", "postgresql", "pg"}:
        raise RuntimeError("process recovery harness requires PostgreSQL")

    temporary = Path(tempfile.mkdtemp(prefix="woo-process-recovery-"))
    redis_port = free_port()
    http_port = free_port()
    gunicorn_port = free_port()
    environment = os.environ.copy()
    environment.update(
        {
            "WOO_CELERY_BROKER_URL": f"redis://127.0.0.1:{redis_port}/0",
            "CELERY_BROKER_URL": f"redis://127.0.0.1:{redis_port}/0",
            "WOO_CELERY_VISIBILITY_TIMEOUT": "5",
            "WOO_SYNC_RECOVERY_STALE_SECONDS": "3",
            "WOO_SYNC_HEARTBEAT_STALE_SECONDS": "2",
            "WOO_SYNC_POST_COMMIT_ACTIONS_ENABLED": "0",
            "WOO_SYNC_CONNECT_TIMEOUT": "2",
            "WOO_SYNC_READ_TIMEOUT": "45",
            "WOO_SYNC_IPV4_ONLY": "1",
            "WOO_DB_POOL_MAX": "2",
            "WOO_DB_APPLICATION_NAME": "woo-process-recovery",
            "PYTHONUNBUFFERED": "1",
        }
    )
    os.environ.update(environment)

    import redis

    import db_backend as db
    import sync_service
    from celery_app import celery_app
    from oid_utils import make_oid

    parsed_broker = urlparse(str(celery_app.conf.broker_url))
    if parsed_broker.hostname != "127.0.0.1" or parsed_broker.port != redis_port:
        raise AssertionError(
            "Celery broker mismatch "
            f"host={parsed_broker.hostname!r} port={parsed_broker.port!r} "
            f"expected_port={redis_port}"
        )

    client = redis.Redis.from_url(environment["WOO_CELERY_BROKER_URL"])
    processes = []
    redis_process = None
    http_server = None
    http_thread = None
    site_id = None
    order_id = None
    run_ids = []
    success = False

    def connection():
        value = db.connect("woocommerce_orders.db", timeout=10)
        value.row_factory = db.Row
        return value

    def run_status(run_id):
        conn = connection()
        try:
            row = conn.execute(
                "SELECT status FROM sync_runs WHERE run_id=?", (run_id,)
            ).fetchone()
            return str(row["status"]) if row else None
        finally:
            conn.close()

    def dispatch_status(run_id):
        conn = connection()
        try:
            row = conn.execute(
                "SELECT status FROM sync_page_dispatches WHERE run_id=? AND page=1",
                (run_id,),
            ).fetchone()
            return str(row["status"]) if row else None
        finally:
            conn.close()

    def outbox_state(run_id):
        conn = connection()
        try:
            row = conn.execute(
                """
                SELECT status,last_error FROM sync_task_outbox
                WHERE payload->>'run_id'=? ORDER BY outbox_id DESC LIMIT 1
                """,
                (run_id,),
            ).fetchone()
            if not row:
                return None
            error_type = None
            if row["last_error"]:
                error_type = str(row["last_error"]).split(":", 1)[0][:80]
            return str(row["status"]), error_type
        finally:
            conn.close()

    try:
        redis_process = start_redis(environment, temporary, redis_port)
        processes.append(redis_process)
        wait_until(lambda: client.ping(), label="temporary Redis")

        state = MockState()
        state.site_url = f"http://127.0.0.1:{http_port}"
        http_server = ThreadingHTTPServer(
            ("127.0.0.1", http_port), handler_for(state)
        )
        http_thread = threading.Thread(target=http_server.serve_forever, daemon=True)
        http_thread.start()

        conn = connection()
        try:
            row = conn.execute(
                """
                INSERT INTO sites(url,consumer_key,consumer_secret,api_status)
                VALUES (?,?,?,'unknown') RETURNING id
                """,
                (state.site_url, "mock-key", "mock-secret"),
            ).fetchone()
            site_id = int(row[0])
            conn.commit()
        finally:
            conn.close()
        order_id = make_oid(site_id, 992340001)

        first, created = sync_service.start_sync(
            mode="quick",
            created_by="pytest:process-recovery:first",
            site_ids=[site_id],
            params={"per_page": 50, "notes_per_page": 0},
            publish=True,
        )
        assert created is True
        run_ids.append(first["run_id"])
        state_row = outbox_state(first["run_id"])
        if not state_row or state_row[0] != "published":
            error_type = state_row[1] if state_row else "missing"
            raise AssertionError(f"outbox publish failed ({error_type})")
        wait_until(lambda: client.dbsize() > 0, label="persisted Redis queue")
        time.sleep(1.2)

        # A graceful Redis restart must retain the durable queued task through AOF.
        redis_process.stop()
        processes.remove(redis_process)
        redis_process = start_redis(environment, temporary, redis_port)
        processes.append(redis_process)
        wait_until(lambda: client.ping(), label="restarted Redis")
        assert client.dbsize() > 0

        gunicorn = start_gunicorn(environment, temporary, gunicorn_port)
        processes.append(gunicorn)
        wait_until(
            lambda: urlopen(f"http://127.0.0.1:{gunicorn_port}/login", timeout=2).status == 200,
            label="test Gunicorn",
        )
        fetch = start_worker(environment, temporary, "sync_fetch", "fetch-recovery")
        processes.append(fetch)
        wait_until(lambda: state.first_seen.is_set(), timeout=35, label="blocked fetch")
        wait_until(
            lambda: dispatch_status(first["run_id"]) == "fetching",
            label="fetch claim",
        )

        # Restarting Gunicorn cannot affect the independent Celery fetch process.
        gunicorn.stop()
        processes.remove(gunicorn)
        assert dispatch_status(first["run_id"]) == "fetching"
        gunicorn = start_gunicorn(environment, temporary, gunicorn_port)
        processes.append(gunicorn)
        wait_until(
            lambda: urlopen(f"http://127.0.0.1:{gunicorn_port}/login", timeout=2).status == 200,
            label="restarted test Gunicorn",
        )

        # Kill the fetch process group mid-request. Late ACK plus Redis visibility
        # timeout must cause the restarted worker to redeliver the same task.
        fetch.stop(hard=True)
        processes.remove(fetch)
        state.release_first.set()
        fetch = start_worker(environment, temporary, "sync_fetch", "fetch-recovered")
        processes.append(fetch)
        writer = start_worker(environment, temporary, "sync_write", "write-recovery")
        processes.append(writer)
        time.sleep(4)
        fetch_recovery = sync_service.recover_stale_work()
        assert fetch_recovery["dispatches"] >= 1
        wait_until(
            lambda: run_status(first["run_id"]) == "success",
            timeout=60,
            interval=0.25,
            label="fetch redelivery completion",
        )

        # Block a second run's update inside PostgreSQL, kill the writer process
        # group, then release the lock. The transaction must roll back and the
        # redelivery must produce one receipt and one order row.
        lock_connection = connection()
        lock_connection.execute(
            "SELECT id FROM orders WHERE id=? FOR UPDATE", (order_id,)
        ).fetchone()
        second, created = sync_service.start_sync(
            mode="quick",
            created_by="pytest:process-recovery:second",
            site_ids=[site_id],
            params={"per_page": 50, "notes_per_page": 0},
            publish=True,
        )
        assert created is True
        run_ids.append(second["run_id"])
        try:
            wait_until(
                lambda: dispatch_status(second["run_id"]) == "writing",
                timeout=35,
                label="blocked writer transaction",
            )
            writer.stop(hard=True)
            processes.remove(writer)
        finally:
            lock_connection.rollback()
            lock_connection.close()

        time.sleep(4)
        writer_recovery = sync_service.recover_stale_work()
        assert writer_recovery["dispatches"] >= 1
        writer = start_worker(environment, temporary, "sync_write", "write-recovered")
        processes.append(writer)
        wait_until(
            lambda: run_status(second["run_id"]) == "success",
            timeout=60,
            interval=0.25,
            label="writer redelivery completion",
        )

        conn = connection()
        try:
            receipt_count = int(
                conn.execute(
                    "SELECT count(*) FROM sync_page_receipts WHERE run_id=?",
                    (second["run_id"],),
                ).fetchone()[0]
            )
            order_count = int(
                conn.execute(
                    "SELECT count(*) FROM orders WHERE id=?", (order_id,)
                ).fetchone()[0]
            )
            assert receipt_count == 1
            assert order_count == 1
        finally:
            conn.close()

        result = {
            "ok": True,
            "redis_aof_restart": True,
            "gunicorn_restart_celery_independent": True,
            "fetch_worker_redelivery": True,
            "writer_worker_redelivery": True,
            "postgres_outbox_recovery": True,
            "writer_receipts": receipt_count,
            "order_rows": order_count,
            "mock_order_requests": state.order_requests,
        }
        print(json.dumps(result, sort_keys=True))
        success = True
        return 0
    finally:
        if http_server is not None:
            http_server.shutdown()
            http_server.server_close()
        for managed in reversed(processes):
            try:
                managed.stop()
            except Exception:
                pass
        try:
            conn = connection()
            try:
                if run_ids:
                    placeholders = ",".join("?" for _ in run_ids)
                    conn.execute(
                        f"DELETE FROM sync_task_outbox WHERE payload->>'run_id' IN ({placeholders})",
                        tuple(run_ids),
                    )
                    conn.execute(
                        f"DELETE FROM sync_runs WHERE run_id::text IN ({placeholders})",
                        tuple(run_ids),
                    )
                if order_id is not None:
                    conn.execute("DELETE FROM order_notes WHERE order_id=?", (order_id,))
                    conn.execute("DELETE FROM shipping_logs WHERE order_id=?", (order_id,))
                    conn.execute("DELETE FROM orders WHERE id=?", (order_id,))
                if site_id is not None:
                    conn.execute("DELETE FROM sites WHERE id=?", (site_id,))
                conn.commit()
            finally:
                conn.close()
        except Exception:
            pass
        db.close_pools()
        if success:
            shutil.rmtree(temporary, ignore_errors=True)
        else:
            print(f"process recovery logs retained at {temporary}", file=sys.stderr)


if __name__ == "__main__":
    raise SystemExit(main())
