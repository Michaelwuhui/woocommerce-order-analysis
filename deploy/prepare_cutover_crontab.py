#!/usr/bin/env python3
"""Create exact maintenance and PostgreSQL root crontabs for cutover."""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import re


LEGACY_MARKERS = {
    "deep_and_clean": ("1.wooorders_sqlite.py", 2),
    "automatic_sync": (
        "/run/lock/woo-analysis-auto-sync.cron.lock",
        1,
    ),
    "sqlite_backup": (
        "/www/wwwroot/woo-analysis/backup_db.py",
        1,
    ),
}

DATABASE_WRITER_MARKERS = {
    "inventory_push": (
        "/www/wwwroot/woo-analysis/inv_push_cron.py",
        1,
    ),
    "resolve_outcomes": (
        "/www/wwwroot/woo-analysis/resolve_outcomes.py --live",
        1,
    ),
}

ENVIRONMENT_PREFIX = "set -a; . /etc/woo-analysis/woo-analysis.env; set +a; "
CRON_COMMAND_RE = re.compile(
    r"^(?P<schedule>\s*\S+(?:[ \t]+\S+){4}[ \t]+)(?P<command>.*?)(?P<newline>\r?\n)?$"
)


def _postgres_line(line: str) -> str:
    match = CRON_COMMAND_RE.match(line)
    if not match or not match.group("command").strip():
        raise ValueError("database writer does not have a standard five-field cron line")
    command = match.group("command")
    if ENVIRONMENT_PREFIX in command:
        raise ValueError("database writer already has the PostgreSQL environment prefix")
    return (
        match.group("schedule")
        + ENVIRONMENT_PREFIX
        + command
        + (match.group("newline") or "")
    )


def filtered_crontab(source: str, mode: str) -> tuple[str, dict[str, int]]:
    if mode not in {"maintenance", "postgres"}:
        raise ValueError(f"unsupported cutover crontab mode: {mode}")

    markers = {**LEGACY_MARKERS, **DATABASE_WRITER_MARKERS}
    counts = {name: 0 for name in markers}
    preserved = []
    for line in source.splitlines(keepends=True):
        stripped = line.lstrip()
        matches = [] if stripped.startswith("#") else [
            name
            for name, (marker, _expected) in markers.items()
            if marker in line
        ]
        if len(matches) > 1:
            raise ValueError("cron line matched more than one managed category")
        if matches:
            name = matches[0]
            counts[name] += 1
            if name in DATABASE_WRITER_MARKERS and mode == "postgres":
                preserved.append(_postgres_line(line))
        else:
            preserved.append(line)

    expected = {name: value[1] for name, value in markers.items()}
    if counts != expected:
        raise ValueError(f"managed cron counts changed: expected={expected} actual={counts}")
    return "".join(preserved), counts


def main(argv=None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", required=True, choices=("maintenance", "postgres"))
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args(argv)
    if args.output.exists():
        raise SystemExit(f"refusing to overwrite {args.output}")

    content, counts = filtered_crontab(
        args.input.read_text(encoding="utf-8"),
        args.mode,
    )
    descriptor = os.open(
        args.output,
        os.O_WRONLY | os.O_CREAT | os.O_EXCL,
        0o600,
    )
    with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
        handle.write(content)
    print(
        f"prepared {args.mode} crontab: "
        + " ".join(f"{name}={count}" for name, count in counts.items())
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
