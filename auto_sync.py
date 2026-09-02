#!/usr/bin/env python3
"""Compatibility entry point that submits one durable automatic sync run.

Celery Beat is the production scheduler.  This command remains useful for a
manual, idempotent submission and returns the active run when another mode is
already running.
"""

from __future__ import annotations

import json

from sync_service import start_sync


def main() -> int:
    status, created = start_sync(
        mode="auto",
        created_by="auto-sync-cli",
        params={
            "per_page": 50,
            "incremental_overlap_minutes": 10,
            "notes_per_page": 10,
        },
    )
    print(
        json.dumps(
            {
                "run_id": status["run_id"],
                "status": status["status"],
                "created": created,
            },
            ensure_ascii=False,
        ),
        flush=True,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
