"""Durable parcel verification and bounded restoration outside Gunicorn."""
import logging

from celery_app import celery_app
from shipment_reconciliation import due_operations, reconcile_operation

LOG = logging.getLogger(__name__)


def enqueue_orders(order_ids=None, *, manual=False):
    ids = due_operations(order_ids, manual=manual)
    for operation_id in ids:
        options = {'kwargs': {'manual': True}} if manual else {}
        reconcile_shipment.apply_async(args=[operation_id], expires=300, **options)
    return len(ids)


@celery_app.task(name='woo_sync.scan_shipment_reconciliation', acks_late=True,
                 reject_on_worker_lost=True)
def scan_shipments():
    return {'queued': enqueue_orders()}


@celery_app.task(name='woo_sync.reconcile_shipment', acks_late=True,
                 reject_on_worker_lost=True)
def reconcile_shipment(operation_id, manual=False):
    return reconcile_operation(operation_id, manual=manual)


def enqueue_after_uncertain(operation_id):
    if not operation_id:
        return
    try:
        reconcile_shipment.apply_async(args=[operation_id], countdown=310, expires=900)
    except Exception:
        # The persisted ledger and periodic sweep recover broker outages.
        LOG.exception('Unable to queue shipment reconciliation; periodic scan will recover')
