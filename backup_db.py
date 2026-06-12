#!/usr/bin/env python3
"""
backup_db.py — woo-analysis 订单库一致性热备 + 异地上传 (P0-b)

做什么：
  1. 用 SQLite 在线备份 API 对 woocommerce_orders.db 做一份**一致性快照**
     （WAL 安全，不停服务、不锁库；优于直接 cp 文件）。
  2. 对快照跑 PRAGMA integrity_check，确认备份本身没坏才保留。
  3. gzip 压缩，带时间戳存到本地备份目录（独立于应用目录与网站根目录）。
  4. 本地按数量滚动保留（默认 48 份）。
  5. 若配置了异地目标（backup_offsite.json），上传到异地（R2/S3 或 rsync）；
     未配置则跳过——异地目标定下来后填配置即可，无需改代码。

手动恢复：
  gunzip -c <backup>.db.gz > restored.db
  sqlite3 restored.db "PRAGMA integrity_check;"          # 校验
  # 确认无误后：停 gunicorn → 先把当前损坏库改名留存 → 用 restored.db 顶替 → 启动
"""
import gzip
import json
import os
import shutil
import sqlite3
import sys
import tempfile
import time
from datetime import datetime

# ---- 配置 ----
APP_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DB = os.path.join(APP_DIR, "woocommerce_orders.db")
BACKUP_DIR = "/www/backups/woo-orders"           # 本地暂存（应用目录 / 网站根之外）
KEEP_LOCAL = 48                                  # 本地保留份数（每小时一次 ≈ 最近 2 天）
OFFSITE_CONFIG = os.path.join(APP_DIR, "backup_offsite.json")


def log(msg):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}", flush=True)


def make_consistent_snapshot(src_path, dst_path):
    """用在线备份 API 生成一致性快照（WAL 安全，自动处理并发写入）。"""
    src = sqlite3.connect(src_path, timeout=30)
    dst = sqlite3.connect(dst_path)
    try:
        src.backup(dst)
    finally:
        dst.close()
        src.close()


def integrity_ok(db_path):
    conn = sqlite3.connect(db_path, timeout=30)
    try:
        row = conn.execute("PRAGMA integrity_check").fetchone()
        return bool(row) and row[0] == "ok"
    finally:
        conn.close()


def gzip_file(src_path, dst_path):
    with open(src_path, "rb") as fi, gzip.open(dst_path, "wb", compresslevel=6) as fo:
        shutil.copyfileobj(fi, fo, length=1024 * 1024)


def rotate_local(backup_dir, keep):
    files = sorted(
        os.path.join(backup_dir, f) for f in os.listdir(backup_dir)
        if f.startswith("woocommerce_orders_") and f.endswith(".db.gz")
    )
    stale = files[:-keep] if keep > 0 else []
    removed = 0
    for path in stale:
        try:
            os.remove(path)
            removed += 1
        except OSError as e:
            log(f"  滚动删除失败 {path}: {e}")
    if removed:
        log(f"  本地滚动：删除 {removed} 份旧备份，保留最近 {keep} 份")


def upload_offsite(local_gz):
    """按 backup_offsite.json 上传异地；未配置则跳过。"""
    if not os.path.exists(OFFSITE_CONFIG):
        log("异地未配置（无 backup_offsite.json），跳过异地上传。")
        return
    try:
        with open(OFFSITE_CONFIG, encoding="utf-8") as f:
            cfg = json.load(f)
    except Exception as e:
        log(f"异地配置读取失败：{e}，跳过。")
        return
    mode = cfg.get("mode", "none")
    if mode == "none":
        log("异地配置 mode=none，跳过异地上传。")
    elif mode == "s3":
        _upload_s3(local_gz, cfg.get("s3", {}))
    elif mode == "rsync":
        _upload_rsync(local_gz, cfg.get("rsync", {}))
    else:
        log(f"未知异地模式 mode={mode}，跳过。")


def _upload_s3(local_gz, s3cfg):
    try:
        import boto3
        from botocore.config import Config
    except ImportError:
        log("需要 boto3 才能上传 S3/R2：venv/bin/pip install boto3 后重试。本次已跳过。")
        return
    try:
        client = boto3.client(
            "s3",
            endpoint_url=s3cfg.get("endpoint_url"),
            aws_access_key_id=s3cfg["access_key_id"],
            aws_secret_access_key=s3cfg["secret_access_key"],
            region_name=s3cfg.get("region", "auto"),
            config=Config(signature_version="s3v4"),
        )
        key = s3cfg.get("prefix", "") + os.path.basename(local_gz)
        client.upload_file(local_gz, s3cfg["bucket"], key)
        log(f"  已上传异地 S3/R2: s3://{s3cfg['bucket']}/{key}")
    except Exception as e:
        log(f"  异地 S3/R2 上传失败：{e}")


def _upload_rsync(local_gz, rcfg):
    import subprocess
    target = rcfg.get("target")
    if not target:
        log("  rsync 缺少 target，跳过。")
        return
    cmd = ["rsync", "-az"] + rcfg.get("extra_args", []) + [local_gz, target]
    try:
        subprocess.run(cmd, check=True, timeout=600)
        log(f"  已 rsync 到 {target}")
    except Exception as e:
        log(f"  rsync 失败：{e}")


def main():
    if not os.path.exists(SRC_DB):
        log(f"源库不存在：{SRC_DB}，中止。")
        return 1
    os.makedirs(BACKUP_DIR, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    final_gz = os.path.join(BACKUP_DIR, f"woocommerce_orders_{ts}.db.gz")
    src_size = os.path.getsize(SRC_DB)
    log(f"开始备份：{SRC_DB} ({src_size / 1024 / 1024:.1f} MB) -> {final_gz}")

    tmpdir = tempfile.mkdtemp(prefix="woobak_")
    snap = os.path.join(tmpdir, "snap.db")
    try:
        t0 = time.time()
        make_consistent_snapshot(SRC_DB, snap)
        log(f"  快照完成（{time.time() - t0:.1f}s）")

        if not integrity_ok(snap):
            log("  ❌ 快照完整性校验失败（integrity_check != ok），丢弃，本次不产出备份。")
            return 2
        log("  完整性校验：ok")

        gzip_file(snap, final_gz)
        gz_size = os.path.getsize(final_gz)
        log(f"  压缩完成：{gz_size / 1024 / 1024:.1f} MB（压缩率 {gz_size / src_size:.0%}）")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    rotate_local(BACKUP_DIR, KEEP_LOCAL)
    upload_offsite(final_gz)
    log("备份完成 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
