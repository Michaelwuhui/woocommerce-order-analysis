"""Daily recoverable PostgreSQL/code/config bundle, verified on Google Drive.

Uses an existing private OAuth credential file. Never imports the web app or
the panel plugin, and never treats an upload request alone as backup success.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import random
import re
import shutil
import subprocess
import tarfile
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import urlsplit
from zoneinfo import ZoneInfo

import requests

from backup_alerts import STATE_DIR, atomic_json, now_iso, read_json, record_health, send_email


CONFIG_PATH = Path(os.getenv("WOO_DRIVE_BACKUP_CONFIG", "/etc/woo-analysis/drive-backup.json"))
DRIVE_API = "https://www.googleapis.com/drive/v3/files"
UPLOAD_API = "https://www.googleapis.com/upload/drive/v3/files"
MANAGED_BY = "woo-analysis-daily-v1"
FIELDS = "id,name,size,md5Checksum,sha256Checksum,parents,trashed,appProperties,webViewLink,createdTime"
TRANSIENT_REASONS = {"rateLimitExceeded", "userRateLimitExceeded", "quotaExceeded", "backendError"}
CHUNK_SIZE = 16 * 1024 * 1024


class BackupError(RuntimeError):
    def __init__(self, message, *, retryable=False):
        super().__init__(message)
        self.retryable = retryable


def log(message):
    print(f"[{now_iso()}] {message}", flush=True)


def file_hash(path, algorithm="sha256"):
    digest = hashlib.new(algorithm)
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(4 * 1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


class DriveClient:
    def __init__(self, credentials, *, session=None, sleep=time.sleep, attempts=8):
        self.credentials = credentials
        self.session = session or requests.Session()
        self.sleep = sleep
        self.attempts = attempts
        self.token = None

    def _pause(self, attempt):
        self.sleep(min(60, 2 ** (attempt + 1)) + random.random())

    def refresh(self):
        for attempt in range(self.attempts):
            try:
                response = self.session.post("https://oauth2.googleapis.com/token", data={
                    "grant_type": "refresh_token", "refresh_token": self.credentials["refresh_token"],
                    "client_id": self.credentials["client_id"], "client_secret": self.credentials["client_secret"],
                }, timeout=(10, 30))
                if response.status_code == 200:
                    self.token = response.json()["access_token"]
                    return
                if response.status_code < 500 and response.status_code != 429:
                    raise BackupError(f"Google OAuth HTTP {response.status_code}; 请检查授权")
            except requests.RequestException:
                pass
            if attempt + 1 < self.attempts:
                self._pause(attempt)
        raise BackupError("Google OAuth 暂时不可用，重试已用尽")

    def request(self, method, url, *, acceptable=(200,), attempts=None, **kwargs):
        parsed = urlsplit(url)
        if parsed.scheme != "https" or parsed.hostname != "www.googleapis.com":
            raise BackupError("拒绝非 Google 上传地址")
        if not self.token:
            self.refresh()
        tries = attempts or self.attempts
        refreshed = False
        for attempt in range(tries + 1):
            try:
                headers = dict(kwargs.pop("headers", {})) if attempt == 0 else headers
                headers["Authorization"] = "Bearer " + self.token
                response = self.session.request(method, url, headers=headers, timeout=(10, 120), **kwargs)
                if response.status_code in acceptable:
                    return response
                if response.status_code == 401 and not refreshed:
                    self.refresh()
                    refreshed = True
                    continue
                try:
                    errors = response.json().get("error", {}).get("errors", [])
                    reasons = {str(row.get("reason")) for row in errors}
                except (ValueError, AttributeError, TypeError):
                    reasons = set()
                transient = response.status_code in {408, 429, 500, 502, 503, 504} or (response.status_code == 403 and bool(reasons & TRANSIENT_REASONS))
                label = ",".join(sorted(reasons & TRANSIENT_REASONS))
                error = BackupError(f"Google Drive HTTP {response.status_code} {label}".strip(), retryable=transient)
            except requests.RequestException:
                error = BackupError("Google Drive 网络连接失败", retryable=True)
            if not error.retryable or attempt + 1 >= tries:
                raise error
            log(f"{error}; 退避重试 {attempt + 1}/{tries}")
            self._pause(attempt)
        raise BackupError("Google Drive 请求重试已用尽")

    def metadata(self, file_id, *, missing_ok=False):
        response = self.request("GET", DRIVE_API + "/" + file_id, params={"fields": FIELDS}, acceptable=(200, 404) if missing_ok else (200,))
        return None if response.status_code == 404 else response.json()

    def verify(self, receipt, folder_id):
        actual = self.metadata(receipt["file_id"])
        if actual.get("trashed") or folder_id not in actual.get("parents", []):
            raise BackupError("云端备份位置不符或已删除")
        if int(actual.get("size", -1)) != receipt["size"] or actual.get("md5Checksum") != receipt["md5"]:
            raise BackupError("Google Drive 文件大小或内容校验不一致")
        if actual.get("sha256Checksum") and actual["sha256Checksum"] != receipt["sha256"]:
            raise BackupError("Google Drive SHA256 校验不一致")
        return actual

    def upload(self, path, folder_id, receipt, save):
        """Persist a generated file ID and resumable URL before sending bytes."""
        if not receipt.get("file_id"):
            generated = self.request("GET", DRIVE_API + "/generateIds", params={"count": 1, "space": "drive", "type": "files"}).json()
            receipt["file_id"] = generated["ids"][0]
            save(receipt)
        existing = self.metadata(receipt["file_id"], missing_ok=True)
        if existing and existing.get("md5Checksum"):
            return self.verify(receipt, folder_id)
        size = receipt["size"]
        failures = 0
        restarts = 0
        offset = 0
        while True:
            if not receipt.get("upload_url"):
                metadata = {"id": receipt["file_id"], "name": Path(path).name, "parents": [folder_id], "appProperties": {"managed_by": MANAGED_BY, "backup_day": receipt["day"], "sha256": receipt["sha256"]}}
                start = self.request("POST", UPLOAD_API, params={"uploadType": "resumable", "fields": FIELDS}, json=metadata, headers={"X-Upload-Content-Type": "application/x-tar", "X-Upload-Content-Length": str(size)})
                receipt["upload_url"] = start.headers.get("Location")
                if not receipt["upload_url"]:
                    raise BackupError("Google Drive 未返回续传地址")
                save(receipt)
            # Always recover the server's offset, including after a process crash.
            status = self.request("PUT", receipt["upload_url"], headers={"Content-Length": "0", "Content-Range": f"bytes */{size}"}, data=b"", acceptable=(200, 201, 308, 404, 410))
            if status.status_code in {200, 201}:
                return self.verify(receipt, folder_id)
            if status.status_code in {404, 410}:
                receipt.pop("upload_url", None)
                save(receipt)
                restarts += 1
                if restarts > 2:
                    raise BackupError("Google Drive 续传会话反复失效")
                completed = self.metadata(receipt["file_id"], missing_ok=True)
                if completed and completed.get("md5Checksum"):
                    return self.verify(receipt, folder_id)
                continue
            matched = re.fullmatch(r"bytes=0-(\d+)", status.headers.get("Range", ""))
            offset = int(matched.group(1)) + 1 if matched else 0
            with Path(path).open("rb") as handle:
                handle.seek(offset)
                while offset < size:
                    chunk = handle.read(min(CHUNK_SIZE, size - offset))
                    if not chunk:
                        raise BackupError("本地备份文件在上传期间被截断")
                    try:
                        result = self.request("PUT", receipt["upload_url"], data=chunk, headers={"Content-Type": "application/x-tar", "Content-Range": f"bytes {offset}-{offset + len(chunk) - 1}/{size}"}, acceptable=(200, 201, 308), attempts=1)
                    except BackupError as exc:
                        if not exc.retryable or failures >= 7:
                            raise
                        self._pause(failures)
                        failures += 1
                        break  # query the committed server range, never guess
                    failures = 0
                    if result.status_code in {200, 201}:
                        return self.verify(receipt, folder_id)
                    matched = re.fullmatch(r"bytes=0-(\d+)", result.headers.get("Range", ""))
                    confirmed = int(matched.group(1)) + 1 if matched else 0
                    if confirmed <= offset or confirmed > offset + len(chunk):
                        raise BackupError("Google Drive 续传偏移异常")
                    offset = confirmed
                    handle.seek(offset)
                    log(f"已上传 {offset}/{size} 字节")

    def retain(self, folder_id, days, *, today):
        if days < 1:
            raise BackupError("云端保留天数必须至少为 1")
        cutoff = (today - timedelta(days=days - 1)).isoformat()
        query = f"'{folder_id}' in parents and trashed=false and appProperties has {{ key='managed_by' and value='{MANAGED_BY}' }}"
        token = None
        while True:
            params = {"q": query, "fields": "nextPageToken,files(id,appProperties)", "pageSize": 100}
            if token:
                params["pageToken"] = token
            page = self.request("GET", DRIVE_API, params=params).json()
            for item in page.get("files", []):
                day = item.get("appProperties", {}).get("backup_day", "")
                if re.fullmatch(r"\d{4}-\d{2}-\d{2}", day) and day < cutoff:
                    self.request("PATCH", DRIVE_API + "/" + item["id"], json={"trashed": True})
            token = page.get("nextPageToken")
            if not token:
                return


def newest_snapshot(directory, *, max_age_seconds=7200, now=None):
    now = time.time() if now is None else now
    candidates = sorted(Path(directory).glob("woo_analysis_*.dump"), reverse=True)
    for archive in candidates:
        if archive.is_symlink() or not re.fullmatch(r"woo_analysis_\d{8}_\d{6}\.dump", archive.name):
            continue
        checksum = Path(str(archive) + ".sha256")
        manifest_path = Path(str(archive) + ".manifest.json")
        if not checksum.is_file() or not manifest_path.is_file():
            continue  # hourly writer may still be publishing sidecars
        manifest = read_json(manifest_path)
        created = datetime.fromisoformat(manifest["created_at"]).timestamp()
        if now - created > max_age_seconds or created - now > 300:
            raise BackupError("最近完整 PostgreSQL 备份超过 2 小时或时间异常")
        if manifest.get("backend") != "postgres" or manifest.get("database") != "woo_analysis":
            raise BackupError("数据库备份不是当前 woo_analysis PostgreSQL 库")
        if file_hash(archive) != checksum.read_text().split()[0]:
            raise BackupError("源 PostgreSQL 备份 SHA256 不一致")
        listed = subprocess.run(["pg_restore", "--list", str(archive)], capture_output=True, text=True, timeout=120)
        if listed.returncode or "TABLE DATA public orders" not in listed.stdout:
            raise BackupError("PostgreSQL 归档目录校验失败")
        return archive
    raise BackupError("没有找到完整的 PostgreSQL 备份")


def build_bundle(config, stage, day):
    stage = Path(stage)
    stage.mkdir(mode=0o700, parents=True, exist_ok=True)
    source = newest_snapshot(config.get("backup_dir", "/www/backups/woo-orders"))
    if shutil.disk_usage(stage).free < source.stat().st_size * 2 + 256 * 1024 * 1024:
        raise BackupError("日备份暂存目录剩余空间不足")
    app_dir = Path(config.get("app_dir", "/www/wwwroot/woo-analysis"))
    commit = subprocess.check_output(["git", "-C", str(app_dir), "rev-parse", "HEAD"], text=True).strip()
    dirty = subprocess.check_output(["git", "-C", str(app_dir), "status", "--porcelain", "--untracked-files=no"], text=True)
    if dirty.strip():
        raise BackupError("生产代码存在未提交修改，不能把提交归档标记为完整代码备份")
    code = stage / "source.tar.gz"
    subprocess.run(["git", "-C", str(app_dir), "archive", "--format=tar.gz", "--output=" + str(code), commit], check=True, capture_output=True, timeout=120)
    os.chmod(code, 0o600)
    configs = stage / "runtime-config.tar.gz"
    with tarfile.open(configs, "w:gz") as tar:
        for name in config.get("runtime_files", []):
            path = Path(name)
            if not path.is_file() or path.is_symlink():
                raise BackupError("必需的运行配置文件缺失或为符号链接")
            tar.add(path, arcname=str(path).lstrip("/"), recursive=False)
    os.chmod(configs, 0o600)
    sources = [(source, "database/" + source.name), (Path(str(source) + ".sha256"), "database/" + source.name + ".sha256"), (Path(str(source) + ".manifest.json"), "database/" + source.name + ".manifest.json"), (code, "app/source.tar.gz"), (configs, "config/runtime-config.tar.gz")]
    manifest = stage / "recovery.json"
    atomic_json(manifest, {"format": MANAGED_BY, "created_at": now_iso(), "backup_day": day, "database": "woo_analysis", "backend": "postgres", "git_commit": commit, "files": [{"path": arcname, "sha256": file_hash(path), "bytes": path.stat().st_size} for path, arcname in sources], "restore": "Verify each SHA256; restore the PostgreSQL dump with pg_restore --exit-on-error --no-owner --no-acl into an EMPTY isolated database first. Review runtime config before restoring services. Do not run a deep sync as a substitute for restoring this database."})
    name = f"woo-analysis_{day}_{source.stem.removeprefix('woo_analysis_')}.tar"
    bundle = stage / name
    temporary = bundle.with_suffix(".partial")
    with tarfile.open(temporary, "w") as tar:
        for path, arcname in sources + [(manifest, "recovery.json")]:
            tar.add(path, arcname=arcname, recursive=False)
    os.chmod(temporary, 0o600)
    os.replace(temporary, bundle)
    # Re-read the packaged bytes, not just the source files, before uploading.
    expected = read_json(manifest)
    with tarfile.open(bundle, "r") as tar:
        for item in expected["files"]:
            digest = hashlib.sha256()
            member = tar.extractfile(item["path"])
            for chunk in iter(lambda: member.read(4 * 1024 * 1024), b""):
                digest.update(chunk)
            if digest.hexdigest() != item["sha256"]:
                raise BackupError("打包内容 SHA256 校验失败")
    return {"day": day, "path": str(bundle), "size": bundle.stat().st_size, "sha256": file_hash(bundle), "md5": file_hash(bundle, "md5"), "git_commit": commit}


def run_backup(config, state_dir):
    today = datetime.now(ZoneInfo("Asia/Shanghai")).date()
    day = today.isoformat()
    receipt_path = state_dir / "receipts" / (day + ".json")
    receipt = read_json(receipt_path)
    client = DriveClient(read_json(config["credentials_file"]))
    folder = config["folder_id"]
    if receipt.get("verified_at"):
        actual = client.verify(receipt, folder)
    else:
        if not receipt:
            receipt = build_bundle(config, state_dir / "staging" / day, day)
            atomic_json(receipt_path, receipt)
        if not Path(receipt["path"]).is_file() or file_hash(receipt["path"]) != receipt["sha256"]:
            raise BackupError("待上传备份文件缺失或校验失败")
        actual = client.upload(receipt["path"], folder, receipt, lambda data: atomic_json(receipt_path, data))
        receipt["verified_at"] = now_iso()
        receipt["web_view_link"] = actual.get("webViewLink")
        receipt.pop("upload_url", None)
        atomic_json(receipt_path, receipt)
    status = read_json(state_dir / "status.json")
    status.update({"last_success_at": receipt["verified_at"], "last_backup_name": actual["name"], "last_backup_bytes": receipt["size"], "last_file_url": actual.get("webViewLink"), "git_commit": receipt["git_commit"], "last_error": None})
    atomic_json(state_dir / "status.json", status)
    client.retain(folder, int(config.get("retention_days", 30)), today=today)
    notified = record_health(state_dir, healthy=True, detail=f"{actual['name']}；大小和内容校验通过。")
    # Only our generated staging directories are removed, after cloud readback.
    staging = (state_dir / "staging").resolve()
    if staging.exists():
        for path in staging.iterdir():
            if path.is_dir() and not path.is_symlink() and re.fullmatch(r"\d{4}-\d{2}-\d{2}", path.name) and path.resolve().parent == staging:
                shutil.rmtree(path)
    log(f"Google Drive 备份已核验：{actual['name']}")
    if not notified:
        log("备份已核验，恢复提醒发送失败；健康检查将重试邮件")


def check_health(state_dir):
    status = read_json(state_dir / "status.json")
    last_success = status.get("last_success_at")
    since = last_success or status.get("configured_at")
    if not since:
        raise BackupError("Google Drive 日备份未完成配置")
    age = time.time() - datetime.fromisoformat(since).timestamp()
    service = subprocess.run(["systemctl", "is-failed", "woo-drive-backup.service"], capture_output=True, timeout=10)
    timer = subprocess.run(["systemctl", "is-active", "woo-drive-backup.timer"], capture_output=True, timeout=10)
    if service.returncode == 0:
        raise BackupError(status.get("last_error") or "Google Drive 备份服务执行失败")
    if timer.returncode != 0:
        raise BackupError("Google Drive 日备份定时器未运行")
    if age > 26 * 3600 or (not last_success and age > 2 * 3600):
        raise BackupError("Google Drive 最近成功备份已超时或首次备份未完成")
    # Preserve a genuine upload failure until a verified run resolves it.
    if status.get("last_error"):
        raise BackupError(status["last_error"])
    if last_success:
        if not record_health(state_dir, healthy=True, detail="Google Drive 日备份按时完成"):
            log("恢复提醒发送失败；健康检查将继续重试")


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=["run", "check", "test-email"])
    args = parser.parse_args(argv)
    os.umask(0o077)
    STATE_DIR.mkdir(mode=0o700, parents=True, exist_ok=True)
    import fcntl  # production uses Linux; pure functions also test on Windows
    with (STATE_DIR / "job.lock").open("a") as lock:
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            log("已有备份任务运行，本次跳过")
            return 0
        try:
            if args.action == "test-email":
                send_email("[订单系统] 备份失败提醒通道测试", "这是订单系统备份提醒通道的配置测试，不代表发生真实故障。\n以后 Google Drive 日备份重试失败或超过 26 小时没有成功备份时，会发送失败提醒；核验恢复后会发送恢复通知。")
                log("测试邮件已由 SMTP 接受")
                return 0
            if args.action == "check":
                check_health(STATE_DIR)
            else:
                config = read_json(CONFIG_PATH)
                status = read_json(STATE_DIR / "status.json")
                status["last_attempt_at"] = now_iso()
                atomic_json(STATE_DIR / "status.json", status)
                run_backup(config, STATE_DIR)
            return 0
        except Exception as exc:
            detail = str(exc) if isinstance(exc, BackupError) else f"备份执行异常：{type(exc).__name__}"
            log(detail)
            if args.action == "test-email":
                return 1
            try:
                status = read_json(STATE_DIR / "status.json")
            except (OSError, ValueError):
                status = {}
            status["last_error"] = detail
            try:
                atomic_json(STATE_DIR / "status.json", status)
            except OSError:
                log("备份状态写入失败，继续尝试发送邮件")
            if not record_health(STATE_DIR, healthy=False, detail=detail):
                log("失败提醒发送失败；健康检查将继续重试")
            return 1


if __name__ == "__main__":
    raise SystemExit(main())
