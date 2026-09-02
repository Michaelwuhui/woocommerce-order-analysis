#!/usr/bin/env python3
"""Create a verified online backup for the configured production database.

SQLite mode uses the SQLite backup API, so a live WAL database is copied as a
single consistent snapshot.  PostgreSQL mode uses ``pg_dump`` custom format,
which takes a transactionally consistent online snapshot and can be inspected
or restored with ``pg_restore``.  Every archive has a SHA256 sidecar and a
small, non-secret manifest.

PostgreSQL restore (into a freshly created empty database):
  sha256sum -c <backup>.dump.sha256
  pg_restore --list <backup>.dump
  pg_restore --exit-on-error --clean --if-exists --no-owner --no-acl \
    --dbname=<restore_database> <backup>.dump

SQLite restore:
  sha256sum -c <backup>.db.gz.sha256
  gunzip -c <backup>.db.gz > restored.db
  sqlite3 restored.db "PRAGMA integrity_check;"
"""

from __future__ import annotations

import gzip
import hashlib
import json
import os
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path


APP_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DB = os.getenv("WOO_SQLITE_PATH", os.path.join(APP_DIR, "woocommerce_orders.db"))
BACKUP_DIR = os.getenv("WOO_BACKUP_DIR", "/www/backups/woo-orders")
KEEP_LOCAL = int(os.getenv("WOO_BACKUP_KEEP_LOCAL", "48"))
OFFSITE_CONFIG = os.getenv(
    "WOO_BACKUP_OFFSITE_CONFIG", os.path.join(APP_DIR, "backup_offsite.json")
)


def log(message):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}", flush=True)


def backend_name():
    value = os.getenv("WOO_DB_BACKEND", "sqlite").strip().lower()
    return "postgres" if value in {"postgres", "postgresql", "pg"} else "sqlite"


def make_consistent_snapshot(src_path, dst_path):
    """Use SQLite's online backup API; this is safe while WAL is active."""
    source = sqlite3.connect(src_path, timeout=30)
    destination = sqlite3.connect(dst_path)
    try:
        source.backup(destination)
    finally:
        destination.close()
        source.close()


def sqlite_manifest(db_path):
    connection = sqlite3.connect(db_path, timeout=30)
    try:
        result = connection.execute("PRAGMA integrity_check").fetchone()
        if not result or result[0] != "ok":
            raise RuntimeError("SQLite backup integrity_check failed")
        names = [
            row[0]
            for row in connection.execute(
                "SELECT name FROM sqlite_master "
                "WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
            )
        ]
        counts = {}
        for name in names:
            quoted = '"' + name.replace('"', '""') + '"'
            counts[name] = int(connection.execute(f"SELECT COUNT(*) FROM {quoted}").fetchone()[0])
        return {"integrity_check": "ok", "table_counts": counts}
    finally:
        connection.close()


def gzip_file(src_path, dst_path):
    with open(src_path, "rb") as source, gzip.open(dst_path, "wb", compresslevel=6) as target:
        shutil.copyfileobj(source, target, length=1024 * 1024)


def sha256_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _atomic_text(path, text):
    temporary = str(path) + ".tmp"
    with open(temporary, "w", encoding="utf-8") as handle:
        handle.write(text)
        handle.flush()
        os.fsync(handle.fileno())
    os.chmod(temporary, 0o600)
    os.replace(temporary, path)


def write_sidecars(archive_path, manifest):
    digest = sha256_file(archive_path)
    checksum_path = archive_path + ".sha256"
    manifest_path = archive_path + ".manifest.json"
    _atomic_text(checksum_path, f"{digest}  {os.path.basename(archive_path)}\n")
    _atomic_text(
        manifest_path,
        json.dumps(manifest, ensure_ascii=False, sort_keys=True, indent=2) + "\n",
    )
    return [archive_path, checksum_path, manifest_path]


def _backup_main_files(backup_dir, backend):
    if not os.path.isdir(backup_dir):
        return []
    if backend == "postgres":
        pattern = "woo_analysis_*.dump"
    else:
        pattern = "woocommerce_orders_*.db.gz"
    return sorted(str(path) for path in Path(backup_dir).glob(pattern))


def rotate_local(backup_dir, keep, backend):
    """Rotate only archives for the active backend; never delete rollback-era files."""
    files = _backup_main_files(backup_dir, backend)
    stale = files[:-keep] if keep > 0 else []
    removed = 0
    for archive in stale:
        for path in (archive, archive + ".sha256", archive + ".manifest.json"):
            try:
                os.remove(path)
            except FileNotFoundError:
                pass
            except OSError as exc:
                log(f"  滚动删除失败 {path}: {exc}")
        removed += 1
    if removed:
        log(f"  本地滚动：删除 {removed} 组旧的 {backend} 备份，保留最近 {keep} 组")


def upload_offsite(local_path):
    """Upload one archive or sidecar; an absent config is a deliberate no-op."""
    if not os.path.exists(OFFSITE_CONFIG):
        log("异地未配置（无 backup_offsite.json），跳过异地上传。")
        return
    try:
        with open(OFFSITE_CONFIG, encoding="utf-8") as handle:
            config = json.load(handle)
    except Exception as exc:
        log(f"异地配置读取失败：{exc}，跳过。")
        return
    mode = config.get("mode", "none")
    if mode == "none":
        log("异地配置 mode=none，跳过异地上传。")
    elif mode == "s3":
        _upload_s3(local_path, config.get("s3", {}))
    elif mode == "rsync":
        _upload_rsync(local_path, config.get("rsync", {}))
    else:
        log(f"未知异地模式 mode={mode}，跳过。")


def _upload_s3(local_path, config):
    try:
        import boto3
        from botocore.config import Config
    except ImportError:
        log("需要 boto3 才能上传 S3/R2；本次已跳过。")
        return
    try:
        client = boto3.client(
            "s3",
            endpoint_url=config.get("endpoint_url"),
            aws_access_key_id=config["access_key_id"],
            aws_secret_access_key=config["secret_access_key"],
            region_name=config.get("region", "auto"),
            config=Config(signature_version="s3v4"),
        )
        key = config.get("prefix", "") + os.path.basename(local_path)
        client.upload_file(local_path, config["bucket"], key)
        log(f"  已上传异地 S3/R2: s3://{config['bucket']}/{key}")
    except Exception as exc:
        log(f"  异地 S3/R2 上传失败：{exc}")


def _upload_rsync(local_path, config):
    target = config.get("target")
    if not target:
        log("  rsync 缺少 target，跳过。")
        return
    command = ["rsync", "-az"] + config.get("extra_args", []) + [local_path, target]
    try:
        subprocess.run(command, check=True, timeout=600)
        log(f"  已 rsync 到 {target}")
    except Exception as exc:
        log(f"  rsync 失败：{exc}")


def create_sqlite_backup(timestamp):
    if not os.path.exists(SRC_DB):
        raise RuntimeError(f"SQLite source does not exist: {SRC_DB}")
    final_path = os.path.join(BACKUP_DIR, f"woocommerce_orders_{timestamp}.db.gz")
    source_size = os.path.getsize(SRC_DB)
    log(f"开始 SQLite 一致性备份（源文件 {source_size / 1024 / 1024:.1f} MB）")
    # Keep the temporary artifact on the destination filesystem so the final
    # os.replace() remains atomic under systemd PrivateTmp and separate /www
    # mounts instead of failing with EXDEV.
    temporary_dir = tempfile.mkdtemp(prefix=".woo-sqlite-backup-", dir=BACKUP_DIR)
    try:
        snapshot_path = os.path.join(temporary_dir, "snapshot.db")
        compressed_path = os.path.join(temporary_dir, "snapshot.db.gz")
        started = time.monotonic()
        make_consistent_snapshot(SRC_DB, snapshot_path)
        manifest = sqlite_manifest(snapshot_path)
        gzip_file(snapshot_path, compressed_path)
        os.chmod(compressed_path, 0o600)
        os.replace(compressed_path, final_path)
        manifest.update(
            {
                "backend": "sqlite",
                "created_at": datetime.now(timezone.utc).isoformat(),
                "elapsed_seconds": round(time.monotonic() - started, 3),
            }
        )
        log(f"  integrity_check=ok，表数={len(manifest['table_counts'])}")
        return write_sidecars(final_path, manifest)
    finally:
        shutil.rmtree(temporary_dir, ignore_errors=True)


def _postgres_settings():
    settings = {
        "host": os.getenv("WOO_DB_HOST", "127.0.0.1"),
        "port": os.getenv("WOO_DB_PORT", "5432"),
        "dbname": os.getenv("WOO_DB_NAME_OVERRIDE") or os.getenv("WOO_DB_NAME"),
        "user": os.getenv("WOO_DB_USER"),
        "password": os.getenv("WOO_DB_PASSWORD"),
    }
    missing = [name for name in ("dbname", "user", "password") if not settings.get(name)]
    if missing:
        raise RuntimeError("PostgreSQL backup configuration is incomplete: " + ", ".join(missing))
    return settings


def _postgres_table_counts(settings):
    try:
        import psycopg
        from psycopg import sql
    except ImportError as exc:
        raise RuntimeError("psycopg is required for PostgreSQL backup verification") from exc
    connection = psycopg.connect(**settings, connect_timeout=10)
    try:
        connection.execute("SET TRANSACTION READ ONLY")
        rows = connection.execute(
            "SELECT tablename FROM pg_catalog.pg_tables "
            "WHERE schemaname='public' ORDER BY tablename"
        ).fetchall()
        counts = {}
        for (table_name,) in rows:
            query = sql.SQL("SELECT COUNT(*) FROM {}").format(sql.Identifier(table_name))
            counts[table_name] = int(connection.execute(query).fetchone()[0])
        return counts
    finally:
        connection.rollback()
        connection.close()


def create_postgres_backup(timestamp):
    settings = _postgres_settings()
    pg_dump = shutil.which("pg_dump")
    pg_restore = shutil.which("pg_restore")
    if not pg_dump or not pg_restore:
        raise RuntimeError("pg_dump and pg_restore must both be installed")
    final_path = os.path.join(BACKUP_DIR, f"woo_analysis_{timestamp}.dump")
    temporary_dir = tempfile.mkdtemp(prefix=".woo-postgres-backup-", dir=BACKUP_DIR)
    temporary_dump = os.path.join(temporary_dir, "database.dump")
    environment = os.environ.copy()
    environment["PGPASSWORD"] = settings["password"]
    command = [
        pg_dump,
        "--format=custom",
        "--compress=6",
        "--no-owner",
        "--no-acl",
        "--serializable-deferrable",
        "--host",
        settings["host"],
        "--port",
        str(settings["port"]),
        "--username",
        settings["user"],
        "--file",
        temporary_dump,
        settings["dbname"],
    ]
    log("开始 PostgreSQL 一致性自定义格式备份")
    started = time.monotonic()
    try:
        result = subprocess.run(
            command,
            env=environment,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            text=True,
            timeout=1800,
        )
        if result.returncode != 0:
            raise RuntimeError(f"pg_dump failed with exit code {result.returncode}")
        listing = subprocess.run(
            [pg_restore, "--list", temporary_dump],
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            text=True,
            timeout=120,
        )
        if listing.returncode != 0 or "TABLE DATA public orders" not in listing.stdout:
            raise RuntimeError("pg_restore archive verification failed")
        archive_entries = sum(
            1 for line in listing.stdout.splitlines() if line and not line.startswith(";")
        )
        table_counts = _postgres_table_counts(settings)
        os.chmod(temporary_dump, 0o600)
        os.replace(temporary_dump, final_path)
        manifest = {
            "backend": "postgres",
            "created_at": datetime.now(timezone.utc).isoformat(),
            "database": settings["dbname"],
            "pg_restore_list_ok": True,
            "archive_entries": archive_entries,
            "source_table_counts_observed_after_dump": table_counts,
            "elapsed_seconds": round(time.monotonic() - started, 3),
        }
        log(f"  pg_restore --list 校验通过，表数={len(table_counts)}")
        return write_sidecars(final_path, manifest)
    finally:
        environment["PGPASSWORD"] = ""
        shutil.rmtree(temporary_dir, ignore_errors=True)


def main():
    os.umask(0o077)
    os.makedirs(BACKUP_DIR, mode=0o700, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backend = backend_name()
    try:
        artifacts = (
            create_postgres_backup(timestamp)
            if backend == "postgres"
            else create_sqlite_backup(timestamp)
        )
    except Exception as exc:
        log(f"备份失败：{exc}")
        return 2
    rotate_local(BACKUP_DIR, KEEP_LOCAL, backend)
    for artifact in artifacts:
        upload_offsite(artifact)
    log(f"备份完成：{os.path.basename(artifacts[0])}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
