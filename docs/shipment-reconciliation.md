# Legacy shipment result reconciliation

The Celery Beat `shipment-result-reconciliation` entry scans every 60 seconds.
It selects `ship_order` operations in `pending`, `external_success`, or
`reconciliation_required` that have been unchanged for at least five minutes.
Source GET requests run individually on `sync_fetch`, outside a database
transaction. Ordinary order sync also queues eligible operations after commit;
an uncertain shipping request schedules a delayed verification. The database
ledger remains discoverable if publication or a worker fails.

Verification requires the original site/order identity, request hash, carrier,
tracking, product/variation IDs, order line IDs and full quantities. AST product
lists, VillaTheme line tracking, and custom line tracking are supported. Partial,
reship, OMS, return, conflicting and incomplete evidence requires review.

After a verified GET, the worker rechecks local state while holding row locks.
A per-order PostgreSQL advisory lock also coordinates new shipment operations.
The order mirror, existing/missing shipping log, operation ledger and internal
audit note commit atomically. Original shipment time and completed/delivered
states are preserved. Repeated delivery of a task is harmless.

This path never POSTs/PUTs to WooCommerce, sends a customer email, adjusts stock,
or automatically retries a shipment mutation. A source that lacks tracking
remains visibly unresolved and is rechecked after 5, 10, 20, 40, then 60 minutes.
The shipping modal disables duplicate shipment actions during reconciliation;
its check button respects the same grace period and backoff. Definitively failed
operations retain the existing explicit original-parcel retry flow.

The order detail exposes the last verification result. The pending queue shows
the reason and a check action, including when a worker died before writing a
shipping log. Site and shipping permissions are unchanged.

Tests: `tests/test_shipment_reconciliation.py`,
`tests/test_postgres_shipment_reconciliation.py` (requires an isolated database
whose name begins `woo_reconcile_test_`), and
`node tests/test_shipment_reconciliation_ui.js`.
