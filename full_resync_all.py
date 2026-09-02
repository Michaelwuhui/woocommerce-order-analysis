#!/usr/bin/env python3
"""Submit a globally-exclusive durable deep synchronization run."""

from __future__ import annotations

import argparse
import json

from sync_service import start_sync


def main(argv=None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--status-id",
        help="accepted for old callers; durable runs always use a UUID",
    )
    parser.parse_args(argv)
    status, created = start_sync(
        mode="deep",
        created_by="deep-sync-cli",
        params={"per_page": 50, "notes_per_page": 0},
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
