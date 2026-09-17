"""Expiring warehouse stocktake grants, separate from global inventory privileges."""
import datetime

from inv_common import _table_exists


def utc_now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat(timespec='microseconds')


def active_grants(conn, user_id, warehouse_id=None, lock=False):
    if not _table_exists(conn, 'inv_temporary_stock_grants'):
        return []
    sql = '''SELECT * FROM inv_temporary_stock_grants
        WHERE user_id=? AND revoked_at IS NULL AND expires_at>?'''
    params = [user_id, utc_now()]
    if warehouse_id is not None:
        sql += ' AND warehouse_id=?'
        params.append(warehouse_id)
    sql += ' ORDER BY id DESC'
    if lock and hasattr(conn, '_raw'):
        sql += ' FOR UPDATE'
    rows = conn.execute(sql, params).fetchall()
    # A request may wait on a concurrent revocation or on another stock adjustment.
    return [dict(row) for row in rows if not row['revoked_at'] and row['expires_at'] > utc_now()]
