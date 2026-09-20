"""Backup email alerts independent of Flask, PostgreSQL and the mail centre."""
from __future__ import annotations

import json
import os
import smtplib
import ssl
import time
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path


STATE_DIR = Path(os.getenv("WOO_DRIVE_BACKUP_STATE_DIR", "/var/lib/woo-analysis-backup"))
ALERT_CONFIG = Path(os.getenv("WOO_BACKUP_ALERT_CONFIG", "/etc/woo-analysis/backup-alerts.json"))


def now_iso():
    return datetime.now(timezone.utc).isoformat()


def read_json(path, default=None):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {} if default is None else default


def atomic_json(path, value):
    path = Path(path)
    path.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    with temporary.open("w", encoding="utf-8") as handle:
        os.chmod(temporary, 0o600)
        json.dump(value, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temporary, path)


def send_email(subject, body, *, config=None):
    config = config or read_json(ALERT_CONFIG)
    required = ("host", "port", "username", "password", "from", "to")
    if any(not config.get(key) for key in required):
        raise RuntimeError("backup email is not configured")
    # Exactly one explicit recipient; do not inherit panel mailing lists.
    recipient = str(config["to"])
    if any(char in recipient for char in "\r\n,;") or "@" not in recipient:
        raise ValueError("invalid backup alert recipient")
    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = config["from"]
    message["To"] = recipient
    message.set_content(body)
    context = ssl.create_default_context()
    if config.get("ssl", True):
        transport = smtplib.SMTP_SSL(config["host"], int(config["port"]), timeout=30, context=context)
    else:
        transport = smtplib.SMTP(config["host"], int(config["port"]), timeout=30)
        transport.starttls(context=context)
    with transport as smtp:
        smtp.login(config["username"], config["password"])
        rejected = smtp.send_message(message, from_addr=config["from"], to_addrs=[recipient])
        if rejected:
            raise RuntimeError("backup alert recipient rejected")


def record_health(state_dir, *, healthy, detail, sender=send_email, clock=time.time):
    """Persist failures before sending; retry rejected email and deduplicate alerts.

    Caller owns the process lock. No email error can turn a failed backup green.
    An incident is closed only after any required recovery mail is accepted.
    """
    state_dir = Path(state_dir)
    path = state_dir / "health.json"
    try:
        state = read_json(path)
    except (OSError, ValueError):
        state = {}
    now = clock()
    previously_failed = state.get("status") == "failed"
    state["checked_at"] = now_iso()
    state["status"] = "healthy" if healthy else "failed"
    state["detail"] = detail
    if not healthy:
        if not previously_failed:
            state["incident_at"] = now_iso()
        state["last_failure_at"] = now_iso()
        state["last_failure_detail"] = detail
    atomic_json(path, state)
    kind = None
    if not healthy and (not state.get("failure_mail_sent_at") or now - state["failure_mail_sent_at"] >= 6 * 3600):
        kind = "failure"
        subject = "[订单系统] Google Drive 日备份失败"
        body = f"Google Drive 日备份未完成。\n时间（UTC）：{now_iso()}\n原因：{detail}\n\n请查看订单系统：系统设置 → 数据备份与灾备。\n服务器：cangfu_hk；服务：woo-drive-backup.service。\n本邮件不代表 R2 小时备份也发生故障。"
    elif healthy and state.get("failure_mail_sent_at"):
        kind = "recovery"
        subject = "[订单系统] Google Drive 日备份已恢复"
        body = f"Google Drive 日备份已恢复，并已核验云端文件。\n时间（UTC）：{now_iso()}\n{detail}"
    if kind:
        try:
            sender(subject, body)
        except Exception as exc:
            state["email_status"] = "failed"
            # Exception messages may contain credentials or SMTP responses.
            state["email_error"] = type(exc).__name__
            atomic_json(path, state)
            return False
        state["email_status"] = "accepted_by_smtp"
        state["email_error"] = None
        state["last_email_at"] = now_iso()
        state["last_email_kind"] = kind
        if kind == "failure":
            state["failure_mail_sent_at"] = now
        else:
            state.pop("failure_mail_sent_at", None)
            state.pop("incident_at", None)
        atomic_json(path, state)
    if not healthy and (not previously_failed or kind):
        with (state_dir / "events.jsonl").open("a", encoding="utf-8") as handle:
            os.chmod(handle.name, 0o600)
            handle.write(json.dumps({"at": now_iso(), "status": "failed", "detail": detail, "email_status": state.get("email_status")}, ensure_ascii=False) + "\n")
    return True


def public_status(state_dir=None):
    """Expose an allowlist only. Upload URLs and credentials are never public."""
    root = Path(state_dir or STATE_DIR)
    try:
        health = read_json(root / "health.json")
        status = read_json(root / "status.json")
        result = {key: status.get(key) for key in (
            "configured_at", "last_attempt_at", "last_success_at", "last_backup_name",
            "last_backup_bytes", "drive_folder_url", "last_file_url", "schedule",
            "alert_recipient", "retention_days", "git_commit", "last_error",
        )}
        result.update({"enabled": bool(status.get("configured_at")), "health": health.get("status", "unknown"), "detail": health.get("detail"), "last_failure_at": health.get("last_failure_at"), "last_failure_detail": health.get("last_failure_detail"), "email_status": health.get("email_status"), "last_email_at": health.get("last_email_at")})
        return result
    except (OSError, ValueError):
        return {"enabled": False, "health": "error", "detail": "备份状态文件无法读取"}
