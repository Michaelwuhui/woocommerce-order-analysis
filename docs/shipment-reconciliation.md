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

When a source lacks tracking, two matching absence checks at least five minutes
apart can restore the original tracking through the existing VillaTheme/custom
line-item order PUT integration. Eligibility additionally requires one original
pending log, unchanged full-order items, a processing/on-hold order, no delivered
or return flags, no competing parcel, a recorded supported format and an original
on-hold target. Damaged tracking metadata is never overwritten. AST/unknown
integrations remain review-only because their shipment POST can append parcels
and trigger notifications.

Restoration copies the original carrier, tracking and shipment date; it changes
only tracking metadata and the originally requested order status. A durable
attempt is recorded before the PUT, with a maximum of three automatic attempts.
Every PUT, including timeouts and HTTP 200 responses, is followed by a separate
GET and the full parcel validation. A worker interrupted after saving the remote
parcel recovers on its next read without writing again. The job never calls
customer-email, note, inventory or shipment-creation endpoints. Source-side order
status hooks still apply as they do for the original shipment action.

Unconfirmed results stay pending and are rechecked after 5, 10, 20, 40, then 60
minutes. The shipping modal disables duplicate shipment actions; its check button
can bypass the hourly backoff but retains the five-minute grace period. Source
edits outside this system are detected on read-back, not protected by the local
advisory lock. Conflicting results stay unresolved. Definitively failed operations
retain the existing explicit original-parcel retry flow.

The order detail exposes the last verification result. The pending queue shows
the reason and a check action, including when a worker died before writing a
shipping log. Site and shipping permissions are unchanged.

Tests: `tests/test_shipment_repair.py`, `tests/test_shipment_reconciliation.py`,
`tests/test_postgres_shipment_reconciliation.py` (requires an isolated database
whose name begins `woo_reconcile_test_`), and
`node tests/test_shipment_reconciliation_ui.js`.
