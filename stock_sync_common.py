"""Shared stock-sync contracts. Never imports the Flask application."""
from datetime import datetime, timezone, timedelta
import hashlib
import json
import os
import uuid

import db_backend as db


class SyncError(ValueError):
    def __init__(self, code, message=None, status=409):
        self.code, self.status = code, status
        super().__init__(message or code)


def connect():
    c = db.connect(os.getenv('INV_DB_FILE', os.getenv('WOO_SQLITE_PATH', 'woocommerce_orders.db')))
    c.row_factory = db.Row
    return c


def now():
    return datetime.now(timezone.utc)


def stamp(offset=0):
    return (now() + timedelta(seconds=offset)).isoformat()


def parse_time(value):
    if isinstance(value, datetime):
        return value.replace(tzinfo=value.tzinfo or timezone.utc)
    if not value:
        return None
    parsed = datetime.fromisoformat(str(value).replace('Z', '+00:00'))
    return parsed.replace(tzinfo=parsed.tzinfo or timezone.utc).astimezone(timezone.utc)


def uid():
    return uuid.uuid4().hex


def dumps(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(',', ':'), default=str)


def loads(value, default=None):
    return json.loads(value) if value else ({} if default is None else default)


def digest(value):
    return hashlib.sha256(dumps(value).encode()).hexdigest()


def exists(c, table):
    if hasattr(c, '_raw'):
        return bool(c.execute('SELECT 1 FROM information_schema.tables WHERE table_schema=current_schema() AND table_name=?', (table,)).fetchone())
    return bool(c.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone())


def rows(c, sql, args=()):
    return [dict(r) for r in c.execute(sql, args).fetchall()]


def one(c, sql, args=()):
    r = c.execute(sql, args).fetchone()
    return dict(r) if r else None


def begin(c):
    c.commit()
    c.execute('BEGIN' if hasattr(c, '_raw') else 'BEGIN IMMEDIATE')


def lock_clause(c):
    return ' FOR UPDATE' if hasattr(c, '_raw') else ''


def event(c, kind, actor_id, object_id, detail):
    c.execute('INSERT INTO stock_sync_events(id,kind,actor_id,object_id,detail_json,created_at) VALUES(?,?,?,?,?,?)',
              (uid(), kind, actor_id, object_id, dumps(detail), stamp()))


def enabled():
    return os.getenv('STOCK_SYNC_ENABLED', '0') == '1'


def positive_ids(value, field='ids'):
    if not isinstance(value, list) or any(type(x) is not int or x < 1 for x in value):
        raise SyncError('INVALID_INPUT', f'{field} 必须为正整数列表', 400)
    return sorted(set(value))
