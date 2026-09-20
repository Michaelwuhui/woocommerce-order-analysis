import hashlib
import json
import subprocess
import tarfile
from datetime import date, datetime, timedelta, timezone
from types import SimpleNamespace

import pytest
import requests

import backup_alerts
import backup_drive as backup


class Response:
    def __init__(self, status=200, data=None, headers=None):
        self.status_code = status
        self.data = data or {}
        self.headers = headers or {}

    def json(self):
        return self.data


class Session:
    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = []

    def request(self, method, url, **kwargs):
        self.calls.append((method, url, kwargs))
        response = self.responses.pop(0)
        if isinstance(response, Exception):
            raise response
        return response


def client(responses, attempts=3):
    session = Session(responses)
    sleeps = []
    instance = backup.DriveClient({}, session=session, sleep=sleeps.append, attempts=attempts)
    instance.token = "test-token"
    return instance, session, sleeps


def metadata(receipt):
    return {"id": "file-1", "name": "bundle.tar", "size": str(receipt["size"]), "md5Checksum": receipt["md5"], "sha256Checksum": receipt["sha256"], "parents": ["folder-1"], "trashed": False}


def make_receipt(path):
    path.write_bytes(b"abcdefgh")
    return {"day": "2026-09-20", "size": 8, "md5": hashlib.md5(b"abcdefgh").hexdigest(), "sha256": hashlib.sha256(b"abcdefgh").hexdigest()}


def test_quota_403_retries_then_succeeds():
    c, session, sleeps = client([Response(403, {"error": {"errors": [{"reason": "rateLimitExceeded"}]}}), Response(200, {"id": "ok"})])
    assert c.request("GET", backup.DRIVE_API).json()["id"] == "ok"
    assert len(session.calls) == 2 and len(sleeps) == 1


def test_auth_denial_does_not_retry_or_expose_provider_details():
    c, session, sleeps = client([Response(403, {"error": {"message": "secret-value", "errors": [{"reason": "insufficientPermissions"}]}})])
    with pytest.raises(backup.BackupError) as error:
        c.request("GET", backup.DRIVE_API)
    assert "secret-value" not in str(error.value)
    assert not sleeps and len(session.calls) == 1


def test_retry_budget_is_bounded():
    c, session, sleeps = client([Response(503)] * 3)
    with pytest.raises(backup.BackupError):
        c.request("GET", backup.DRIVE_API)
    assert len(session.calls) == 3 and len(sleeps) == 2


def test_upload_recovers_committed_range_after_lost_response(tmp_path, monkeypatch):
    monkeypatch.setattr(backup, "CHUNK_SIZE", 4)
    path = tmp_path / "bundle.tar"
    receipt = make_receipt(path)
    c, session, _ = client([
        Response(200, {"ids": ["file-1"]}), Response(404),
        Response(200, headers={"Location": backup.UPLOAD_API + "?upload_id=private"}),
        Response(308), requests.Timeout("lost after commit"),
        Response(308, headers={"Range": "bytes=0-3"}),
        Response(200), Response(200, metadata(receipt)),
    ])
    persisted = []
    result = c.upload(path, "folder-1", receipt, lambda r: persisted.append(dict(r)))
    assert result["id"] == "file-1"
    uploads = [call[2]["data"] for call in session.calls if call[0] == "PUT" and call[2].get("data")]
    assert uploads == [b"abcd", b"efgh"]
    assert persisted[0]["file_id"] == "file-1"
    assert "upload_url" in persisted[1]


def test_completed_upload_is_read_back_without_creating_duplicate(tmp_path):
    path = tmp_path / "bundle.tar"
    receipt = make_receipt(path)
    receipt["file_id"] = "file-1"
    c, session, _ = client([Response(200, metadata(receipt)), Response(200, metadata(receipt))])
    assert c.upload(path, "folder-1", receipt, lambda _: None)["id"] == "file-1"
    assert all(call[0] == "GET" for call in session.calls)


@pytest.mark.parametrize("change", [{"md5Checksum": "wrong"}, {"size": "7"}, {"sha256Checksum": "wrong"}, {"trashed": True}, {"parents": ["wrong-folder"]}])
def test_remote_mismatch_fails(tmp_path, change):
    receipt = make_receipt(tmp_path / "bundle.tar")
    receipt["file_id"] = "file-1"
    actual = metadata(receipt)
    actual.update(change)
    c, _, _ = client([Response(200, actual)])
    with pytest.raises(backup.BackupError):
        c.verify(receipt, "folder-1")


def test_resumable_url_cannot_exfiltrate_authorization():
    c, session, _ = client([])
    with pytest.raises(backup.BackupError):
        c.request("PUT", "https://not-google.example/upload")
    assert not session.calls


def test_retention_only_targets_tagged_folder_objects():
    c, session, _ = client([Response(200, {"files": [
        {"id": "old", "appProperties": {"backup_day": "2026-08-01"}},
        {"id": "keep", "appProperties": {"backup_day": "2026-09-01"}},
        {"id": "unknown", "appProperties": {}},
    ]}), Response(200)])
    c.retain("folder-1", 30, today=date(2026, 9, 20))
    query = session.calls[0][2]["params"]["q"]
    assert "'folder-1' in parents" in query and backup.MANAGED_BY in query
    assert len(session.calls) == 2
    assert session.calls[1][1].endswith("/old")
    assert session.calls[1][2]["json"] == {"trashed": True}


def test_zero_day_retention_refuses_any_cloud_delete():
    c, session, _ = client([])
    with pytest.raises(backup.BackupError):
        c.retain("folder-1", 0, today=date(2026, 9, 20))
    assert not session.calls


def write_snapshot(tmp_path, *, age=0, checksum=True):
    archive = tmp_path / "woo_analysis_20260920_180037.dump"
    archive.write_bytes(b"database")
    (tmp_path / (archive.name + ".sha256")).write_text((backup.file_hash(archive) if checksum else "wrong") + "  " + archive.name)
    backup_alerts.atomic_json(str(archive) + ".manifest.json", {"backend": "postgres", "database": "woo_analysis", "created_at": (datetime.now(timezone.utc) - timedelta(seconds=age)).isoformat()})
    return archive


def test_snapshot_skips_incomplete_published_dump(tmp_path, monkeypatch):
    archive = write_snapshot(tmp_path)
    (tmp_path / "woo_analysis_20260920_190000.dump").write_bytes(b"publishing")
    monkeypatch.setattr(backup.subprocess, "run", lambda *a, **k: SimpleNamespace(returncode=0, stdout="TABLE DATA public orders owner"))
    assert backup.newest_snapshot(tmp_path) == archive


@pytest.mark.parametrize("kwargs", [{"age": 7201}, {"checksum": False}])
def test_stale_or_corrupt_snapshot_rejected(tmp_path, kwargs):
    write_snapshot(tmp_path, **kwargs)
    with pytest.raises(backup.BackupError):
        backup.newest_snapshot(tmp_path)


def test_pg_restore_failure_rejected(tmp_path, monkeypatch):
    write_snapshot(tmp_path)
    monkeypatch.setattr(backup.subprocess, "run", lambda *a, **k: SimpleNamespace(returncode=1, stdout=""))
    with pytest.raises(backup.BackupError):
        backup.newest_snapshot(tmp_path)


def test_failed_email_retried_and_successful_failure_deduplicated(tmp_path):
    sent = []
    def failed(*args):
        raise RuntimeError("do-not-log-password")
    assert backup_alerts.record_health(tmp_path, healthy=False, detail="HTTP 403", sender=failed, clock=lambda: 1000) is False
    state = backup_alerts.read_json(tmp_path / "health.json")
    assert state["status"] == "failed" and "failure_mail_sent_at" not in state
    assert "do-not-log-password" not in json.dumps(state)
    sender = lambda *args: sent.append(args)
    assert backup_alerts.record_health(tmp_path, healthy=False, detail="HTTP 403", sender=sender, clock=lambda: 1001)
    backup_alerts.record_health(tmp_path, healthy=False, detail="HTTP 403", sender=sender, clock=lambda: 1100)
    assert len(sent) == 1
    backup_alerts.record_health(tmp_path, healthy=True, detail="verified", sender=sender, clock=lambda: 1200)
    backup_alerts.record_health(tmp_path, healthy=True, detail="verified", sender=sender, clock=lambda: 1300)
    assert len(sent) == 2 and "已恢复" in sent[1][0]


def test_recovery_email_failure_is_retried(tmp_path):
    backup_alerts.record_health(tmp_path, healthy=False, detail="failed", sender=lambda *a: None, clock=lambda: 1000)
    def reject(*args):
        raise RuntimeError("smtp refused")
    assert not backup_alerts.record_health(tmp_path, healthy=True, detail="verified", sender=reject, clock=lambda: 1100)
    assert backup_alerts.read_json(tmp_path / "health.json")["failure_mail_sent_at"] == 1000
    sent = []
    backup_alerts.record_health(tmp_path, healthy=True, detail="verified", sender=lambda *a: sent.append(a), clock=lambda: 1200)
    assert len(sent) == 1 and "已恢复" in sent[0][0]


def test_corrupt_health_record_still_sends_failure(tmp_path):
    (tmp_path / "health.json").write_text("broken-json")
    sent = []
    backup_alerts.record_health(tmp_path, healthy=False, detail="failed", sender=lambda *a: sent.append(a))
    assert len(sent) == 1


def test_full_state_disk_does_not_block_email(tmp_path, monkeypatch):
    def full(*args):
        raise OSError("disk full")
    monkeypatch.setattr(backup_alerts, "atomic_json", full)
    sent = []
    assert backup_alerts.record_health(tmp_path, healthy=False, detail="disk full", sender=lambda *a: sent.append(a))
    assert len(sent) == 1


def test_public_status_excludes_private_upload_receipts(tmp_path):
    backup_alerts.atomic_json(tmp_path / "status.json", {"configured_at": "now", "last_success_at": "yesterday", "upload_url": "secret", "credentials": "secret"})
    visible = backup_alerts.public_status(tmp_path)
    assert visible["enabled"] and "secret" not in json.dumps(visible)


def test_watchdog_detects_stale_backup_and_disabled_timer(tmp_path, monkeypatch):
    backup_alerts.atomic_json(tmp_path / "status.json", {"last_success_at": (datetime.now(timezone.utc) - timedelta(hours=27)).isoformat()})
    monkeypatch.setattr(backup.subprocess, "run", lambda cmd, **kw: SimpleNamespace(returncode=1 if "is-failed" in cmd else 0))
    with pytest.raises(backup.BackupError, match="超时"):
        backup.check_health(tmp_path)
    backup_alerts.atomic_json(tmp_path / "status.json", {"last_success_at": datetime.now(timezone.utc).isoformat()})
    monkeypatch.setattr(backup.subprocess, "run", lambda *a, **kw: SimpleNamespace(returncode=1))
    with pytest.raises(backup.BackupError, match="定时器"):
        backup.check_health(tmp_path)


def test_smtp_recipient_cannot_expand_to_unapproved_addresses():
    with pytest.raises(ValueError):
        backup_alerts.send_email("test", "test", config={"host": "smtp", "port": 465, "username": "user", "password": "secret", "from": "sender@example.com", "to": "owner@example.com,extra@example.com"})


def test_bundle_contains_database_source_config_and_verifiable_manifest(tmp_path, monkeypatch):
    app = tmp_path / "app"
    app.mkdir()
    subprocess.run(["git", "init", str(app)], check=True, capture_output=True)
    (app / "app.py").write_text("print('example')\n")
    subprocess.run(["git", "-C", str(app), "add", "app.py"], check=True, capture_output=True)
    subprocess.run(["git", "-C", str(app), "-c", "user.email=test@example.com", "-c", "user.name=BackupTest", "commit", "-m", "fixture"], check=True, capture_output=True)
    snapshots = tmp_path / "snapshots"
    snapshots.mkdir()
    source = write_snapshot(snapshots)
    runtime = tmp_path / "runtime.env"
    runtime.write_text("EXAMPLE=test-only\n")
    monkeypatch.setattr(backup, "newest_snapshot", lambda directory: source)
    receipt = backup.build_bundle({"app_dir": str(app), "backup_dir": str(snapshots), "runtime_files": [str(runtime)]}, tmp_path / "stage", "2026-09-20")
    with tarfile.open(receipt["path"]) as tar:
        manifest = json.load(tar.extractfile("recovery.json"))
        assert manifest["backend"] == "postgres" and len(manifest["git_commit"]) == 40
        for item in manifest["files"]:
            content = tar.extractfile(item["path"]).read()
            assert hashlib.sha256(content).hexdigest() == item["sha256"]
        assert "app/source.tar.gz" in tar.getnames()
        assert "config/runtime-config.tar.gz" in tar.getnames()
    assert backup.file_hash(receipt["path"]) == receipt["sha256"]
    (app / "app.py").write_text("uncommitted")
    with pytest.raises(backup.BackupError, match="未提交"):
        backup.build_bundle({"app_dir": str(app), "runtime_files": []}, tmp_path / "stage2", "2026-09-20")
