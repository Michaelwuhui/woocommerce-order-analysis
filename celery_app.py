"""Celery configuration for isolated fetch and serial write workers."""

from __future__ import annotations

import os

from celery import Celery
from kombu import Exchange, Queue


BROKER_URL = (
    os.getenv("CELERY_BROKER_URL")
    or os.getenv("WOO_CELERY_BROKER_URL")
    or "redis://127.0.0.1:6379/0"
)

celery_app = Celery(
    "woo_analysis",
    broker=BROKER_URL,
    include=["sync_tasks"],
)

celery_app.conf.update(
    accept_content=["json"],
    task_serializer="json",
    result_serializer="json",
    result_backend=None,
    task_ignore_result=True,
    task_store_errors_even_if_ignored=False,
    timezone="Asia/Hong_Kong",
    enable_utc=True,
    worker_prefetch_multiplier=1,
    task_acks_late=True,
    task_acks_on_failure_or_timeout=True,
    task_reject_on_worker_lost=True,
    task_track_started=False,
    broker_connection_retry_on_startup=True,
    broker_connection_max_retries=None,
    broker_transport_options={
        "visibility_timeout": int(
            os.getenv("WOO_CELERY_VISIBILITY_TIMEOUT", "3600")
        ),
        "global_keyprefix": "woo-analysis:",
    },
    task_default_delivery_mode=2,
    task_default_exchange="woo_sync",
    task_default_exchange_type="direct",
    task_default_queue="sync_fetch",
    task_queues=(
        Queue(
            "sync_fetch",
            Exchange("woo_sync", type="direct", durable=True),
            routing_key="sync_fetch",
            durable=True,
        ),
        Queue(
            "sync_write",
            Exchange("woo_sync", type="direct", durable=True),
            routing_key="sync_write",
            durable=True,
        ),
    ),
    task_routes={
        "woo_sync.fetch_page": {"queue": "sync_fetch", "routing_key": "sync_fetch"},
        "woo_sync.write_page": {"queue": "sync_write", "routing_key": "sync_write"},
        "woo_sync.post_commit_page": {"queue": "sync_write", "routing_key": "sync_write"},
        "woo_sync.maintenance": {"queue": "sync_write", "routing_key": "sync_write"},
        "woo_sync.schedule_auto": {"queue": "sync_write", "routing_key": "sync_write"},
        "woo_sync.schedule_deep": {"queue": "sync_write", "routing_key": "sync_write"},
    },
    beat_schedule={
        "sync-recovery-and-outbox": {
            "task": "woo_sync.maintenance",
            "schedule": 30.0,
        },
        "automatic-sync-due-check": {
            "task": "woo_sync.schedule_auto",
            "schedule": 60.0,
        },
        "deep-sync-due-check": {
            "task": "woo_sync.schedule_deep",
            "schedule": 60.0,
        },
    },
)
