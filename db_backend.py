"""Database backend selection and a narrow SQLite-to-PostgreSQL adapter.

The legacy application has a large, well-tested sqlite3 call surface.  This
module keeps that connection/cursor shape while moving persistence to psycopg.
It is deliberately not a blind SQL string replacement:

* placeholders are rewritten by a quote/comment-aware lexer;
* SQLite REPLACE is compiled into a conflict-targeted PostgreSQL UPSERT using
  the real PostgreSQL unique indexes;
* INSERT ids are obtained with RETURNING, never last_insert_rowid();
* PRAGMA table_info is answered from PostgreSQL catalogs;
* SQLite date/json helpers are versioned database functions installed by the
  migration.

Temporary databases and explicit non-production SQLite paths continue to use
stdlib sqlite3, which keeps the existing unit-test suite and emergency rollback
path available.
"""

from __future__ import annotations

import datetime as _datetime
import json
import os
import re
import sqlite3 as _sqlite3
import threading
import uuid as _uuid
from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass
from decimal import Decimal
from typing import Any

try:
    import psycopg
    from psycopg import errors as _pg_errors
    from psycopg.pq import TransactionStatus
    from psycopg_pool import ConnectionPool
except ImportError:  # SQLite rollback/tests do not require PostgreSQL packages.
    psycopg = None
    _pg_errors = None
    TransactionStatus = None
    ConnectionPool = None


SQLITE_DB_BASENAME = "woocommerce_orders.db"
_REPLACE_CONFLICT_KEYS: dict[str, tuple[str, ...]] = {
    "settings": ("key",),
    "exchange_rates": ("year_month", "currency"),
    "user_preferences": ("user_id", "preference_key"),
    "report_cache": ("cache_key",),
    "inv_schema_migrations": ("version",),
    "reconciliation_statements": ("partner_id", "period_year", "period_month"),
}

_FUNCTION_REWRITES = {
    "datetime": "sqlite_datetime",
    "date": "sqlite_date",
    "strftime": "sqlite_strftime",
    "julianday": "sqlite_julianday",
}

_POOL_LOCK = threading.Lock()
_POOLS: dict[tuple[int, str, int], ConnectionPool] = {}


class CompatRow(Mapping[str, Any], Sequence[Any]):
    """sqlite3.Row-compatible result supporting integer and named access."""

    __slots__ = ("_names", "_values", "_index")

    def __init__(self, names: Sequence[str], values: Sequence[Any]):
        self._names = tuple(names)
        self._values = tuple(values)
        self._index: dict[str, int] = {}
        for position, name in enumerate(self._names):
            self._index.setdefault(name, position)

    def __getitem__(self, key):
        if isinstance(key, slice):
            return self._values[key]
        if isinstance(key, int):
            return self._values[key]
        return self._values[self._index[str(key)]]

    def __iter__(self) -> Iterator[Any]:
        return iter(self._values)

    def __len__(self) -> int:
        return len(self._values)

    def keys(self):
        return list(self._names)

    def values(self):
        return list(self._values)

    def items(self):
        return [(name, self._values[pos]) for pos, name in enumerate(self._names)]

    def get(self, key, default=None):
        try:
            return self[key]
        except (KeyError, IndexError):
            return default

    def __repr__(self) -> str:
        return repr(dict(self.items()))


def _legacy_result_value(value):
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, dict) or isinstance(value, list):
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    if isinstance(value, _uuid.UUID):
        return str(value)
    if isinstance(value, _datetime.datetime):
        if value.tzinfo is not None:
            return value.isoformat()
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, _datetime.date):
        return value.isoformat()
    if isinstance(value, memoryview):
        return bytes(value)
    return value


def _compat_row_factory(cursor):
    names = [column.name for column in (cursor.description or ())]
    return lambda values: CompatRow(
        names, [_legacy_result_value(value) for value in values]
    )


def backend_name() -> str:
    return str(os.getenv("WOO_DB_BACKEND", "sqlite")).strip().lower()


def is_postgres_backend() -> bool:
    return backend_name() in {"postgres", "postgresql", "pg"}


def is_sqlite_backend() -> bool:
    return not is_postgres_backend()


def _is_legacy_database_path(database: Any) -> bool:
    if database is None:
        return True
    raw = os.fspath(database) if isinstance(database, os.PathLike) else str(database)
    if raw == ":memory:" or raw.startswith("file:"):
        return False
    configured = os.getenv("WOO_SQLITE_PATH", "")
    if configured and os.path.abspath(raw) == os.path.abspath(configured):
        return True
    return os.path.basename(raw) == SQLITE_DB_BASENAME


def _postgres_kwargs() -> dict[str, Any]:
    required = {
        "host": os.getenv("WOO_DB_HOST", "127.0.0.1"),
        "port": int(os.getenv("WOO_DB_PORT", "5432")),
        "dbname": os.getenv("WOO_DB_NAME_OVERRIDE") or os.getenv("WOO_DB_NAME", "woo_analysis"),
        "user": os.getenv("WOO_DB_USER", "woo_analysis_app"),
        "password": os.getenv("WOO_DB_PASSWORD"),
    }
    missing = [key for key in ("dbname", "user", "password") if not required.get(key)]
    if missing:
        raise RuntimeError("PostgreSQL configuration is incomplete: " + ", ".join(missing))
    required.update(
        {
            "connect_timeout": int(os.getenv("WOO_DB_CONNECT_TIMEOUT", "5")),
            "application_name": os.getenv("WOO_DB_APPLICATION_NAME", "woo-analysis"),
            "row_factory": _compat_row_factory,
            "options": "-c timezone=UTC -c idle_in_transaction_session_timeout=60000",
        }
    )
    return required


def _pool(max_size: int | None = None):
    if psycopg is None or ConnectionPool is None:
        raise RuntimeError("psycopg and psycopg_pool are required for PostgreSQL")
    kwargs = _postgres_kwargs()
    pool_max = max_size or int(os.getenv("WOO_DB_POOL_MAX", "2"))
    key = (os.getpid(), str(kwargs["dbname"]), pool_max)
    with _POOL_LOCK:
        pool = _POOLS.get(key)
        if pool is None:
            pool = ConnectionPool(
                conninfo="",
                kwargs=kwargs,
                min_size=int(os.getenv("WOO_DB_POOL_MIN", "0")),
                max_size=pool_max,
                timeout=float(os.getenv("WOO_DB_POOL_TIMEOUT", "10")),
                max_idle=float(os.getenv("WOO_DB_POOL_MAX_IDLE", "300")),
                max_lifetime=float(os.getenv("WOO_DB_POOL_MAX_LIFETIME", "1800")),
                open=True,
                name="woo-analysis-" + str(os.getpid()),
            )
            _POOLS[key] = pool
    return pool


def close_pools() -> None:
    with _POOL_LOCK:
        pools = list(_POOLS.values())
        _POOLS.clear()
    for pool in pools:
        try:
            pool.close()
        except Exception:
            pass


def _quote_ident(identifier: str) -> str:
    return '"' + identifier.replace('"', '""') + '"'


def _unquote_ident(identifier: str) -> str:
    value = identifier.strip()
    if value.startswith('"') and value.endswith('"'):
        return value[1:-1].replace('""', '"')
    return value.split(".")[-1]


def _lex_sql(sql: str) -> str:
    """Rewrite placeholders/functions only outside SQL strings and comments."""

    out: list[str] = []
    length = len(sql)
    index = 0
    while index < length:
        char = sql[index]
        if char == "'":
            start = index
            index += 1
            while index < length:
                if sql[index] == "'" and index + 1 < length and sql[index + 1] == "'":
                    index += 2
                    continue
                if sql[index] == "'":
                    index += 1
                    break
                index += 1
            out.append(sql[start:index])
            continue
        if char == '"':
            if index + 1 < length and sql[index + 1] == '"':
                out.append("''")
                index += 2
                continue
            start = index
            index += 1
            while index < length:
                if sql[index] == '"' and index + 1 < length and sql[index + 1] == '"':
                    index += 2
                    continue
                if sql[index] == '"':
                    index += 1
                    break
                index += 1
            out.append(sql[start:index])
            continue
        if sql.startswith("--", index):
            end = sql.find("\n", index)
            if end < 0:
                out.append(sql[index:])
                break
            out.append(sql[index:end])
            index = end
            continue
        if sql.startswith("/*", index):
            end = sql.find("*/", index + 2)
            if end < 0:
                out.append(sql[index:])
                break
            end += 2
            out.append(sql[index:end])
            index = end
            continue
        if char == "?":
            out.append("%s")
            index += 1
            continue
        if char.isalpha() or char == "_":
            end = index + 1
            while end < length and (sql[end].isalnum() or sql[end] in "_$"):
                end += 1
            token = sql[index:end]
            look = end
            while look < length and sql[look].isspace():
                look += 1
            replacement = _FUNCTION_REWRITES.get(token.lower()) if look < length and sql[look] == "(" else None
            out.append(replacement or token)
            index = end
            continue
        out.append(char)
        index += 1
    return "".join(out)


def _code_spans(sql: str):
    """Yield SQL ranges outside single-quoted strings and comments."""

    start = 0
    index = 0
    length = len(sql)
    while index < length:
        if sql[index] == "'":
            if start < index:
                yield start, index
            index += 1
            while index < length:
                if sql[index] == "'" and index + 1 < length and sql[index + 1] == "'":
                    index += 2
                    continue
                if sql[index] == "'":
                    index += 1
                    break
                index += 1
            start = index
            continue
        if sql.startswith("--", index):
            if start < index:
                yield start, index
            end = sql.find("\n", index)
            index = length if end < 0 else end
            start = index
            continue
        if sql.startswith("/*", index):
            if start < index:
                yield start, index
            end = sql.find("*/", index + 2)
            index = length if end < 0 else end + 2
            start = index
            continue
        index += 1
    if start < length:
        yield start, length


def _transform_code(sql: str, transform) -> str:
    pieces = []
    position = 0
    for start, end in _code_spans(sql):
        pieces.append(sql[position:start])
        pieces.append(transform(sql[start:end]))
        position = end
    pieces.append(sql[position:])
    return "".join(pieces)


def _placeholder_offsets(sql: str) -> list[int]:
    result = []
    for start, end in _code_spans(sql):
        position = start
        while True:
            found = sql.find("%s", position, end)
            if found < 0:
                break
            result.append(found)
            position = found + 2
    return result


def _escape_psycopg_percents(sql: str) -> str:
    """Escape literal percent signs while preserving generated placeholders.

    Psycopg applies its placeholder parser to the complete query text whenever
    parameters are supplied, including percent signs inside SQL string
    literals such as ``strftime('%Y-%m', value)`` and ``LIKE '%term%'``.
    Qmark placeholders have already become ``%s`` at this point; only those
    code-span offsets remain single-percent tokens.
    """

    placeholders = set(_placeholder_offsets(sql))
    pieces: list[str] = []
    index = 0
    while index < len(sql):
        if index in placeholders:
            pieces.append("%s")
            index += 2
        elif sql[index] == "%":
            pieces.append("%%")
            index += 1
        else:
            pieces.append(sql[index])
            index += 1
    return "".join(pieces)


def _matching_parenthesis(sql: str, opening: int) -> int | None:
    depth = 0
    quote = None
    index = opening
    while index < len(sql):
        char = sql[index]
        if quote:
            if char == quote and index + 1 < len(sql) and sql[index + 1] == quote:
                index += 2
                continue
            if char == quote:
                quote = None
        elif char in {"'", '"'}:
            quote = char
        elif char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
            if depth == 0:
                return index
        index += 1
    return None


def _split_expressions(sql: str, start: int, end: int):
    spans = []
    depth = 0
    quote = None
    item_start = start
    index = start
    while index < end:
        char = sql[index]
        if quote:
            if char == quote and index + 1 < end and sql[index + 1] == quote:
                index += 2
                continue
            if char == quote:
                quote = None
        elif char in {"'", '"'}:
            quote = char
        elif char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        elif char == "," and depth == 0:
            spans.append((item_start, index))
            item_start = index + 1
        index += 1
    spans.append((item_start, end))
    return spans


def _insert_value_spans(sql: str):
    match = re.search(r"\bVALUES\s*\(", sql, flags=re.IGNORECASE)
    if not match:
        return []
    opening = sql.find("(", match.start())
    closing = _matching_parenthesis(sql, opening)
    if closing is None:
        return []
    return _split_expressions(sql, opening + 1, closing)


def _boolean_value(value):
    if value is None or isinstance(value, bool):
        return value
    if value in (0, 1):
        return bool(value)
    if isinstance(value, str) and value.strip().lower() in {
        "0", "1", "false", "true", "off", "on", "no", "yes"
    }:
        return value.strip().lower() in {"1", "true", "on", "yes"}
    return value


def _convert_boolean_params(params, positions: set[int]):
    if not positions or params is None or isinstance(params, Mapping):
        return params
    values = list(params)
    for position in positions:
        if 0 <= position < len(values):
            values[position] = _boolean_value(values[position])
    return tuple(values)


def _insert_before_returning(sql: str, clause: str) -> str:
    body = sql.rstrip()
    semicolon = ""
    if body.endswith(";"):
        body, semicolon = body[:-1].rstrip(), ";"
    match = re.search(r"\bRETURNING\b", body, flags=re.IGNORECASE)
    if match:
        return body[: match.start()].rstrip() + " " + clause + " " + body[match.start() :] + semicolon
    return body + " " + clause + semicolon


_INSERT_PREFIX = re.compile(
    r"^(?P<lead>\s*)INSERT\s+OR\s+(?P<mode>IGNORE|REPLACE)\s+INTO\s+"
    r"(?P<table>(?:\"(?:\"\"|[^\"])+\"|[A-Za-z_][\w$]*)(?:\.(?:\"(?:\"\"|[^\"])+\"|[A-Za-z_][\w$]*))?)"
    r"\s*\((?P<columns>[^)]+)\)",
    flags=re.IGNORECASE | re.DOTALL,
)


@dataclass
class CompiledSQL:
    sql: str
    params: Any
    table: str | None = None
    inserted_columns: tuple[str, ...] = ()
    user_returning: bool = False
    boolean_param_positions: frozenset[int] = frozenset()


class PgCompatConnection:
    def __init__(self, pool, raw):
        self._pool = pool
        self._raw = raw
        self._closed = False
        self._last_insert_id = None
        self._last_changes = 0
        self._total_changes = 0
        self._row_factory = CompatRow
        self._metadata_cache: dict[tuple[str, str], Any] = {}

    @property
    def row_factory(self):
        return self._row_factory

    @row_factory.setter
    def row_factory(self, value):
        self._row_factory = value

    @property
    def total_changes(self) -> int:
        return self._total_changes

    @property
    def in_transaction(self) -> bool:
        return self._raw.info.transaction_status != TransactionStatus.IDLE

    def cursor(self):
        self._ensure_open()
        return PgCompatCursor(self, self._raw.cursor())

    def execute(self, sql, params=()):
        return self.cursor().execute(sql, params)

    def executemany(self, sql, seq_of_params):
        return self.cursor().executemany(sql, seq_of_params)

    def executescript(self, _script):
        raise RuntimeError("SQLite schema scripts are disabled on PostgreSQL; run versioned migrations")

    def commit(self):
        self._ensure_open()
        self._raw.commit()

    def rollback(self):
        if not self._closed:
            self._raw.rollback()

    def close(self):
        if self._closed:
            return
        try:
            if self._raw.info.transaction_status != TransactionStatus.IDLE:
                self._raw.rollback()
        finally:
            self._pool.putconn(self._raw)
            self._closed = True

    def __del__(self):
        # Backstop for legacy call sites that raise before their explicit
        # close(). Flask request teardown is the primary deterministic guard.
        try:
            self.close()
        except Exception:
            pass

    def __enter__(self):
        self._ensure_open()
        return self

    def __exit__(self, exc_type, exc, _tb):
        if exc_type is None:
            self.commit()
        else:
            self.rollback()
        return False

    def _ensure_open(self):
        if self._closed:
            raise RuntimeError("database connection is closed")

    def _unique_keys(self, table: str) -> list[tuple[tuple[str, ...], bool]]:
        key = ("unique", table)
        cached = self._metadata_cache.get(key)
        if cached is not None:
            return cached
        query = """
            SELECT array_agg(att.attname ORDER BY ord.n), idx.indisprimary
            FROM pg_index idx
            JOIN pg_class cls ON cls.oid=idx.indrelid
            JOIN pg_namespace ns ON ns.oid=cls.relnamespace
            JOIN LATERAL unnest(idx.indkey) WITH ORDINALITY ord(attnum,n) ON TRUE
            JOIN pg_attribute att ON att.attrelid=cls.oid AND att.attnum=ord.attnum
            WHERE ns.nspname=current_schema() AND cls.relname=%s
              AND idx.indisunique AND ord.attnum > 0
            GROUP BY idx.indexrelid, idx.indisprimary
        """
        with self._raw.cursor(row_factory=None) as cursor:
            cursor.execute(query, (table,))
            result = [(tuple(row[0]), bool(row[1])) for row in cursor.fetchall()]
        self._metadata_cache[key] = result
        return result

    def conflict_keys(self, table: str, inserted: Sequence[str]) -> tuple[str, ...]:
        inserted_set = set(inserted)
        configured = _REPLACE_CONFLICT_KEYS.get(table)
        if configured and set(configured).issubset(inserted_set):
            return configured
        candidates = [
            (columns, primary)
            for columns, primary in self._unique_keys(table)
            if set(columns).issubset(inserted_set)
        ]
        if not candidates:
            raise RuntimeError(
                "INSERT OR REPLACE has no matching PostgreSQL conflict key for " + table
            )
        candidates.sort(key=lambda item: (item[1], len(item[0])))
        return candidates[0][0]

    def integer_primary_key(self, table: str) -> str | None:
        key = ("integer_pk", table)
        if key in self._metadata_cache:
            return self._metadata_cache[key]
        query = """
            SELECT att.attname, typ.typname
            FROM pg_index idx
            JOIN pg_class cls ON cls.oid=idx.indrelid
            JOIN pg_namespace ns ON ns.oid=cls.relnamespace
            JOIN LATERAL unnest(idx.indkey) WITH ORDINALITY ord(attnum,n) ON TRUE
            JOIN pg_attribute att ON att.attrelid=cls.oid AND att.attnum=ord.attnum
            JOIN pg_type typ ON typ.oid=att.atttypid
            WHERE ns.nspname=current_schema() AND cls.relname=%s
              AND idx.indisprimary
            ORDER BY ord.n
        """
        with self._raw.cursor(row_factory=None) as cursor:
            cursor.execute(query, (table,))
            rows = cursor.fetchall()
        result = rows[0][0] if len(rows) == 1 and rows[0][1] in {"int2", "int4", "int8"} else None
        self._metadata_cache[key] = result
        return result

    def boolean_columns(self, table: str | None = None) -> set[str]:
        cache_key = ("boolean_columns", table or "*")
        cached = self._metadata_cache.get(cache_key)
        if cached is not None:
            return cached
        query = """
            SELECT table_name,column_name
            FROM information_schema.columns
            WHERE table_schema=current_schema() AND data_type='boolean'
        """
        values = ()
        if table:
            query += " AND table_name=%s"
            values = (table,)
        with self._raw.cursor(row_factory=None) as cursor:
            cursor.execute(query, values)
            rows = cursor.fetchall()
        result = {str(row[1]) for row in rows}
        self._metadata_cache[cache_key] = result
        return result

    def _rewrite_boolean_literals(
        self, sql: str, table: str | None, columns: Sequence[str]
    ) -> str:
        names = self.boolean_columns()
        if not names:
            return sql
        name_pattern = "|".join(
            re.escape(name) for name in sorted(names, key=len, reverse=True)
        )
        identifier = (
            r'(?:(?:"[A-Za-z_][\w$]*"|[A-Za-z_][\w$]*)\s*\.\s*)?'
            r'"?(?:' + name_pattern + r')"?'
        )

        def rewrite(fragment: str) -> str:
            fragment = re.sub(
                r"(" + identifier + r")\s+IS\s+(NOT\s+)?([01])\b",
                lambda match: (
                    match.group(1)
                    + " IS "
                    + (match.group(2) or "")
                    + ("true" if match.group(3) == "1" else "false")
                ),
                fragment,
                flags=re.IGNORECASE,
            )
            fragment = re.sub(
                r"(" + identifier + r")\s*(=|==|<>|!=)\s*([01])\b",
                lambda match: (
                    match.group(1)
                    + " "
                    + ("=" if match.group(2) == "==" else match.group(2))
                    + " "
                    + ("true" if match.group(3) == "1" else "false")
                ),
                fragment,
                flags=re.IGNORECASE,
            )
            fragment = re.sub(
                r"\b([01])\s*(=|==|<>|!=)\s*(" + identifier + r")",
                lambda match: (
                    ("true" if match.group(1) == "1" else "false")
                    + " "
                    + ("=" if match.group(2) == "==" else match.group(2))
                    + " "
                    + match.group(3)
                ),
                fragment,
                flags=re.IGNORECASE,
            )
            fragment = re.sub(
                r"(" + identifier + r")\s+IN\s*\(\s*0\s*,\s*1\s*\)",
                lambda match: match.group(1) + " IN (false,true)",
                fragment,
                flags=re.IGNORECASE,
            )
            fragment = re.sub(
                r"COALESCE\s*\(\s*(" + identifier + r")\s*,\s*([01])\s*\)",
                lambda match: (
                    "COALESCE("
                    + match.group(1)
                    + ","
                    + ("true" if match.group(2) == "1" else "false")
                    + ")"
                ),
                fragment,
                flags=re.IGNORECASE,
            )
            fragment = re.sub(
                r"(COALESCE\s*\(\s*" + identifier
                + r"\s*,\s*(?:true|false)\s*\))\s*(=|==|<>|!=)\s*([01])\b",
                lambda match: (
                    match.group(1)
                    + " "
                    + ("=" if match.group(2) == "==" else match.group(2))
                    + " "
                    + ("true" if match.group(3) == "1" else "false")
                ),
                fragment,
                flags=re.IGNORECASE,
            )
            return fragment

        rewritten = _transform_code(sql, rewrite)
        if table and columns:
            boolean_columns = self.boolean_columns(table)
            spans = _insert_value_spans(rewritten)
            replacements = []
            for column, span in zip(columns, spans):
                if column not in boolean_columns:
                    continue
                raw = rewritten[span[0]:span[1]]
                stripped = raw.strip()
                if stripped in {"0", "1"}:
                    lead = len(raw) - len(raw.lstrip())
                    trail = len(raw) - len(raw.rstrip())
                    value = "false" if stripped == "0" else "true"
                    replacements.append(
                        (span[0], span[1], " " * lead + value + " " * trail)
                    )
            for start, end, value in reversed(replacements):
                rewritten = rewritten[:start] + value + rewritten[end:]
        return rewritten

    def _boolean_parameter_positions(
        self, sql: str, table: str | None, columns: Sequence[str]
    ) -> frozenset[int]:
        offsets = _placeholder_offsets(sql)
        if not offsets:
            return frozenset()
        offset_to_position = {offset: index for index, offset in enumerate(offsets)}
        result: set[int] = set()
        if table and columns:
            booleans = self.boolean_columns(table)
            for column, (start, end) in zip(columns, _insert_value_spans(sql)):
                if column not in booleans:
                    continue
                for offset in offsets:
                    if start <= offset < end:
                        result.add(offset_to_position[offset])

        names = self.boolean_columns()
        if not names:
            return frozenset(result)
        name_pattern = "|".join(
            re.escape(name) for name in sorted(names, key=len, reverse=True)
        )
        identifier = (
            r'(?:(?:"[A-Za-z_][\w$]*"|[A-Za-z_][\w$]*)\s*\.\s*)?'
            r'"?(?:' + name_pattern + r')"?'
        )
        patterns = [
            r"(?:" + identifier + r")\s*(?:=|==|<>|!=)\s*(%s)",
            r"(%s)\s*(?:=|==|<>|!=)\s*(?:" + identifier + r")",
            r"COALESCE\s*\(\s*(?:" + identifier + r")\s*,\s*(%s)\s*\)",
        ]
        for start, end in _code_spans(sql):
            fragment = sql[start:end]
            for pattern in patterns:
                for match in re.finditer(pattern, fragment, flags=re.IGNORECASE):
                    offset = start + match.start(1)
                    if offset in offset_to_position:
                        result.add(offset_to_position[offset])
            for match in re.finditer(
                r"(?:" + identifier + r")\s+IN\s*\(([^)]*)\)",
                fragment,
                flags=re.IGNORECASE,
            ):
                inner_start = start + match.start(1)
                inner_end = start + match.end(1)
                for offset in offsets:
                    if inner_start <= offset < inner_end:
                        result.add(offset_to_position[offset])
        return frozenset(result)

    def compile(self, sql: str, params=()) -> CompiledSQL:
        original = str(sql)
        user_returning = bool(re.search(r"\bRETURNING\b", original, re.IGNORECASE))
        match = _INSERT_PREFIX.match(original)
        table = None
        columns: tuple[str, ...] = ()
        if match:
            mode = match.group("mode").upper()
            table_token = match.group("table")
            table = _unquote_ident(table_token)
            columns = tuple(_unquote_ident(item) for item in match.group("columns").split(","))
            replacement = match.group("lead") + "INSERT INTO " + table_token + " (" + match.group("columns") + ")"
            original = replacement + original[match.end() :]
            if mode == "IGNORE":
                original = _insert_before_returning(original, "ON CONFLICT DO NOTHING")
            else:
                keys = self.conflict_keys(table, columns)
                updates = [column for column in columns if column not in keys]
                if updates:
                    target = ", ".join(_quote_ident(column) for column in keys)
                    assignments = ", ".join(
                        _quote_ident(column) + "=EXCLUDED." + _quote_ident(column)
                        for column in updates
                    )
                    clause = "ON CONFLICT (" + target + ") DO UPDATE SET " + assignments
                else:
                    clause = "ON CONFLICT DO NOTHING"
                original = _insert_before_returning(original, clause)
        compiled = _lex_sql(original)
        if table is None:
            insert = re.match(
                r"^\s*INSERT\s+INTO\s+(?P<table>(?:\"(?:\"\"|[^\"])+\"|[A-Za-z_][\w$]*))"
                r"\s*\((?P<columns>[^)]+)\)",
                compiled,
                flags=re.IGNORECASE | re.DOTALL,
            )
            if insert:
                table = _unquote_ident(insert.group("table"))
                columns = tuple(_unquote_ident(item) for item in insert.group("columns").split(","))
        compiled = self._rewrite_boolean_literals(compiled, table, columns)
        boolean_positions = self._boolean_parameter_positions(
            compiled, table, columns
        )
        return CompiledSQL(
            compiled,
            _convert_boolean_params(params, set(boolean_positions)),
            table,
            columns,
            user_returning,
            boolean_positions,
        )


class PgCompatCursor:
    def __init__(self, connection: PgCompatConnection, raw_cursor):
        self.connection = connection
        self._raw = raw_cursor
        self._synthetic: list[Any] | None = None
        self._synthetic_position = 0
        self.lastrowid = None

    @property
    def rowcount(self):
        if self._synthetic is not None:
            return len(self._synthetic)
        return self._raw.rowcount

    @property
    def description(self):
        return self._raw.description

    def _set_synthetic(self, rows):
        self._synthetic = list(rows)
        self._synthetic_position = 0
        return self

    def _pragma(self, sql: str):
        compact = sql.strip().rstrip(";")
        table_match = re.match(
            r"PRAGMA\s+table_(?:info|xinfo)\s*\(\s*['\"]?([^)'\"]+)['\"]?\s*\)",
            compact,
            flags=re.IGNORECASE,
        )
        if table_match:
            table = table_match.group(1)
            query = """
                SELECT (col.ordinal_position - 1)::int AS cid,
                       col.column_name AS name,
                       upper(col.data_type) AS type,
                       CASE WHEN col.is_nullable='NO' THEN 1 ELSE 0 END AS notnull,
                       col.column_default AS dflt_value,
                       COALESCE(pk.position,0)::int AS pk
                FROM information_schema.columns col
                LEFT JOIN (
                    SELECT att.attname, ord.n AS position
                    FROM pg_index idx
                    JOIN pg_class cls ON cls.oid=idx.indrelid
                    JOIN pg_namespace ns ON ns.oid=cls.relnamespace
                    JOIN LATERAL unnest(idx.indkey) WITH ORDINALITY ord(attnum,n) ON TRUE
                    JOIN pg_attribute att ON att.attrelid=cls.oid AND att.attnum=ord.attnum
                    WHERE idx.indisprimary AND ns.nspname=current_schema()
                      AND cls.relname=%s
                ) pk ON pk.attname=col.column_name
                WHERE col.table_schema=current_schema() AND col.table_name=%s
                ORDER BY col.ordinal_position
            """
            self._raw.execute(query, (table, table))
            return self
        if re.match(
            r"PRAGMA\s+(foreign_keys|busy_timeout|journal_mode|query_only|synchronous)\b",
            compact,
            flags=re.IGNORECASE,
        ):
            return self._set_synthetic([])
        if re.match(r"PRAGMA\s+integrity_check", compact, flags=re.IGNORECASE):
            raise RuntimeError("PostgreSQL integrity is verified with migration checks, not PRAGMA")
        raise RuntimeError("Unsupported PostgreSQL PRAGMA compatibility query: " + compact[:120])

    def execute(self, sql, params=()):
        self._synthetic = None
        text = str(sql)
        if re.match(r"^\s*PRAGMA\b", text, re.IGNORECASE):
            return self._pragma(text)
        if re.match(r"^\s*SELECT\s+last_insert_rowid\s*\(\s*\)", text, re.IGNORECASE):
            return self._set_synthetic([(self.connection._last_insert_id,)])
        if re.match(r"^\s*SELECT\s+changes\s*\(\s*\)", text, re.IGNORECASE):
            return self._set_synthetic([(self.connection._last_changes,)])
        if re.match(
            r"^\s*BEGIN(?:\s+(?:DEFERRED|IMMEDIATE|EXCLUSIVE))?"
            r"(?:\s+TRANSACTION)?\s*;?\s*$",
            text,
            re.IGNORECASE,
        ):
            # Compile-time catalog lookups start a psycopg transaction. Handle
            # transaction control before compile() so a legacy explicit BEGIN
            # does not become a nested BEGIN and flood PostgreSQL with warnings.
            # psycopg starts a transaction automatically on the next SQL
            # statement. Sending BEGIN through a non-autocommit connection
            # would itself cause an implicit BEGIN followed by a redundant
            # nested BEGIN warning. A no-op here preserves the transaction
            # boundary for the immediately following legacy statement.
            self.connection._last_changes = 0
            return self
        compiled = self.connection.compile(text, params)
        automatic_returning = False
        primary_key = None
        if compiled.table and not compiled.user_returning:
            primary_key = self.connection.integer_primary_key(compiled.table)
            if primary_key:
                compiled.sql = _insert_before_returning(
                    compiled.sql, "RETURNING " + _quote_ident(primary_key)
                )
                automatic_returning = True
        # Passing an empty sequence still enables psycopg's client-side
        # placeholder parser.  That makes otherwise literal percent signs
        # (for example ``LIKE 'pytest:%'``) look like malformed placeholders.
        # Match sqlite3's parameterless execute behaviour when there are no
        # values to bind.
        if compiled.params is None or (
            isinstance(compiled.params, (tuple, list, dict))
            and not compiled.params
        ):
            self._raw.execute(compiled.sql)
        else:
            self._raw.execute(
                _escape_psycopg_percents(compiled.sql), compiled.params
            )
        if automatic_returning:
            row = self._raw.fetchone()
            if row is not None:
                self.lastrowid = row[0]
                self.connection._last_insert_id = row[0]
        if self._raw.rowcount and self._raw.rowcount > 0:
            self.connection._total_changes += self._raw.rowcount
        self.connection._last_changes = max(0, int(self._raw.rowcount or 0))
        return self

    def executemany(self, sql, seq_of_params):
        self._synthetic = None
        compiled = self.connection.compile(str(sql), ())
        converted = (
            _convert_boolean_params(params, set(compiled.boolean_param_positions))
            for params in seq_of_params
        )
        self._raw.executemany(_escape_psycopg_percents(compiled.sql), converted)
        if self._raw.rowcount and self._raw.rowcount > 0:
            self.connection._total_changes += self._raw.rowcount
        self.connection._last_changes = max(0, int(self._raw.rowcount or 0))
        return self

    def fetchone(self):
        if self._synthetic is None:
            return self._raw.fetchone()
        if self._synthetic_position >= len(self._synthetic):
            return None
        row = self._synthetic[self._synthetic_position]
        self._synthetic_position += 1
        return row

    def fetchmany(self, size=None):
        if self._synthetic is None:
            return self._raw.fetchmany(size)
        count = size or 1
        start = self._synthetic_position
        self._synthetic_position = min(len(self._synthetic), start + count)
        return self._synthetic[start : self._synthetic_position]

    def fetchall(self):
        if self._synthetic is None:
            return self._raw.fetchall()
        rows = self._synthetic[self._synthetic_position :]
        self._synthetic_position = len(self._synthetic)
        return rows

    def close(self):
        self._raw.close()

    def __iter__(self):
        while True:
            row = self.fetchone()
            if row is None:
                break
            yield row

    def __enter__(self):
        return self

    def __exit__(self, _exc_type, _exc, _tb):
        self.close()
        return False

    def __getattr__(self, name):
        return getattr(self._raw, name)


def connect(database=None, timeout=30, *args, **kwargs):
    if not is_postgres_backend() or not _is_legacy_database_path(database):
        return _sqlite3.connect(database, timeout=timeout, *args, **kwargs)
    pool = _pool()
    raw = pool.getconn(timeout=float(timeout))
    return PgCompatConnection(pool, raw)


# sqlite3-compatible public names used by existing exception handlers and type
# annotations.  A tuple is valid in an except clause.
Row = _sqlite3.Row
Connection = _sqlite3.Connection
Cursor = _sqlite3.Cursor
Binary = _sqlite3.Binary
sqlite_version = _sqlite3.sqlite_version
Error = (_sqlite3.Error, psycopg.Error) if psycopg else _sqlite3.Error
DatabaseError = (_sqlite3.DatabaseError, psycopg.DatabaseError) if psycopg else _sqlite3.DatabaseError
IntegrityError = (_sqlite3.IntegrityError, psycopg.IntegrityError) if psycopg else _sqlite3.IntegrityError
OperationalError = (_sqlite3.OperationalError, psycopg.OperationalError) if psycopg else _sqlite3.OperationalError
ProgrammingError = (_sqlite3.ProgrammingError, psycopg.ProgrammingError) if psycopg else _sqlite3.ProgrammingError
