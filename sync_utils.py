import json
import hashlib
import time
import db_backend as sqlite3
import threading
import concurrent.futures
from datetime import datetime, timedelta
from woocommerce import API
from oid_utils import make_oid, site_id_for_source, woo_post_id  # cross-site-safe surrogate order id

# Database configuration
DB_FILE = 'woocommerce_orders.db'

# Proxy configuration (optional)
PROXY_CONFIG = {}

DEFAULT_INCREMENTAL_OVERLAP_MINUTES = 10
DEFAULT_NOTES_ACTIVE_LIMIT = 25
DEFAULT_NOTES_REFRESH_INTERVAL_HOURS = 24
DEFAULT_NOTE_WORKERS = 1


def _subtract_minutes_from_iso(value, minutes):
    """Return an ISO timestamp moved backwards by ``minutes``.

    WooCommerce returns site-local ISO timestamps, sometimes with an explicit
    offset and sometimes without one.  Preserve that shape so the next API
    request is interpreted in the same site timezone.  A small overlap avoids
    missing orders that changed at the exact checkpoint timestamp.
    """
    if not value:
        return None
    try:
        raw = str(value).strip()
        parsed = datetime.fromisoformat(raw.replace("Z", "+00:00"))
        shifted = parsed - timedelta(minutes=max(0, int(minutes or 0)))
        result = shifted.isoformat(timespec="seconds")
        if raw.endswith("Z") and result.endswith("+00:00"):
            result = result[:-6] + "Z"
        return result
    except (TypeError, ValueError):
        return value


def _flag_enabled(conn, key):
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return bool(row and str(row[0]).strip().lower() in {'1', 'true', 'yes', 'on'})


LEGACY_TERMINAL_ORDER_STATUSES = {
    'shipped', 'completed', 'cancelled', 'refunded', 'failed', 'trash'
}


def reconcile_legacy_terminal_shortages(conn, candidates):
    """Clear stale shortage flags after a legacy/manual order has terminated.

    The legacy shipping/Woo status path predates ``oms_fulfillments``. A
    managed-product order could therefore be marked short first, then shipped
    manually, leaving the order list with contradictory badges. Only orders
    without a live OMS fulfillment are reconciled; started multi-warehouse
    work remains authoritative and is never hidden by a Woo status update.
    """

    table_names = {
        row[0] for row in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        ).fetchall()
    }
    required = {'oms_order_fulfillment_state', 'oms_order_items', 'oms_fulfillments'}
    if not required.issubset(table_names):
        return 0
    reconciled = 0
    for item in candidates or []:
        status = str(item.get('status') or '').strip().lower()
        order_id = item.get('order_id')
        if status not in LEGACY_TERMINAL_ORDER_STATUSES or not order_id:
            continue
        state = conn.execute(
            '''SELECT aggregate_status,has_shortage,manual_review,manual_reason
               FROM oms_order_fulfillment_state
               WHERE order_id=? AND aggregate_status='stock_shortage'
                 AND has_shortage=1
                 AND NOT EXISTS (
                   SELECT 1 FROM oms_fulfillments f
                   WHERE f.order_id=?
                     AND f.status NOT IN ('superseded','cancelled')
                 )''',
            (order_id, order_id),
        ).fetchone()
        if not state:
            continue
        target = 'delivered' if status == 'completed' else (
            'cancelled' if status in {'cancelled', 'refunded', 'failed', 'trash'}
            else 'shipped'
        )
        conn.execute(
            '''UPDATE oms_order_fulfillment_state
               SET aggregate_status=?,has_shortage=0,manual_review=0,
                   manual_reason=NULL,updated_at=CURRENT_TIMESTAMP
               WHERE order_id=?''',
            (target, order_id),
        )
        conn.execute(
            '''UPDATE oms_order_items
               SET shortage_qty=0,updated_at=CURRENT_TIMESTAMP
               WHERE order_id=?''',
            (order_id,),
        )
        if 'oms_domain_events' in table_names:
            conn.execute(
                '''INSERT INTO oms_domain_events
                   (aggregate_type,aggregate_id,event_type,from_status,to_status,
                    actor_type,reason,payload_json)
                   VALUES ('order',?,'legacy_terminal_shortage_reconciled',?,?,
                           'system',?,?)''',
                (
                    order_id,
                    state[0],
                    target,
                    '订单已通过旧手工流程发货/终止，清理遗留缺货标记',
                    json.dumps({'woo_status': status}, ensure_ascii=False),
                ),
            )
        reconciled += 1
    return reconciled


def _enqueue_fulfillment_plans(candidates, *, raise_on_error=False):
    """Queue planning after the Woo order transaction has committed.

    A separate row-factory connection isolates fulfillment failures from the
    core sync.  The dark-launch flags prevent historical orders being planned
    until operations explicitly enables the workflow.
    """
    if not candidates:
        return
    try:
        from fulfillment_common import get_conn
        from fulfillment_service import enqueue_job, order_contains_managed_product

        conn = get_conn()
        try:
            fulfillment_enabled = _flag_enabled(conn, 'oms_fulfillment_enabled')
            auto_plan_enabled = _flag_enabled(conn, 'oms_auto_plan_enabled')
            managed_isolation_enabled = _flag_enabled(
                conn, 'oms_managed_product_isolation_enabled'
            )
            if not fulfillment_enabled or not (auto_plan_enabled or managed_isolation_enabled):
                return
            for item in candidates:
                if item['status'] not in {'processing', 'offline', 'on-hold', 'partial-shipped', 'shipped'}:
                    continue
                site = conn.execute(
                    "SELECT country FROM sites WHERE url=?", (item['source'],)
                ).fetchone()
                if not site or str(site['country'] or '').upper() not in {'PL', 'CZ', 'HU'}:
                    continue
                # Keep the legacy manual-shipping workflow unchanged for
                # ordinary products while the full auto-plan rollout remains
                # off. Orders containing one of the reserved WMS families are
                # still planned immediately so those lines cannot leak to the
                # legacy Poland partner queue.
                if not auto_plan_enabled and not order_contains_managed_product(
                    conn, item['order_id']
                ):
                    continue
                # Do not retrofit already-shipped legacy orders into the new
                # fulfillment domain merely because a historical line belongs
                # to one of the four families. Their parcel history remains
                # authoritative; isolation starts before shipment only.
                if not auto_plan_enabled and item['status'] == 'shipped':
                    continue
                version = hashlib.sha256(
                    f"{item['order_id']}|{item.get('date_modified') or ''}|{item['status']}".encode('utf-8')
                ).hexdigest()[:20]
                enqueue_job(
                    conn, 'PLAN_ORDER', 'order', item['order_id'],
                    f"plan:{item['order_id']}:{version}", {'order_id': item['order_id']},
                )
            conn.commit()
        finally:
            conn.close()
    except Exception as exc:
        if raise_on_error:
            raise
        # Order sync remains authoritative and must not fail because the
        # optional fulfillment queue is unavailable during rollout.
        print(f"[fulfillment] enqueue skipped: {exc}")


def _enqueue_order_notifications(candidates, *, raise_on_error=False):
    """Dark-launched order-card notifications after the authoritative commit."""
    try:
        from order_notification_service import enqueue_synced_orders

        enqueue_synced_orders(candidates, raise_on_error=raise_on_error)
    except Exception as exc:
        if raise_on_error:
            raise
        # Notification availability must never break WooCommerce synchronization.
        print(f"[order-notification] enqueue skipped: {type(exc).__name__}")

# 线程局部存储，用于数据库连接复用
_thread_local = threading.local()

def get_thread_db_connection():
    """获取当前线程的数据库连接（复用，启用 WAL 模式支持并发）"""
    if not hasattr(_thread_local, 'connection') or _thread_local.connection is None:
        conn = sqlite3.connect(DB_FILE, timeout=30)
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA busy_timeout=30000")
        _thread_local.connection = conn
    return _thread_local.connection

def close_thread_db_connection():
    """关闭当前线程的数据库连接"""
    if hasattr(_thread_local, 'connection') and _thread_local.connection is not None:
        try:
            _thread_local.connection.close()
        except:
            pass
        _thread_local.connection = None

def create_robust_wcapi(url, consumer_key, consumer_secret, proxy_config=None):
    """Create robust WooCommerce API client"""
    try:
        wcapi = API(
            url=url,
            consumer_key=consumer_key,
            consumer_secret=consumer_secret,
            version="wc/v3",
            timeout=60
        )
        return wcapi
    except Exception as e:
        print(f"Error creating WooCommerce API client: {e}")
        return None

def create_database_connection():
    """Create SQLite database connection (兼容旧接口，新代码建议使用 get_thread_db_connection)"""
    try:
        connection = sqlite3.connect(DB_FILE, timeout=30)
        connection.execute("PRAGMA busy_timeout=30000")
        return connection
    except Exception as e:
        print(f"Error creating database connection: {e}")
        return None

def get_last_order_date_from_db(site_url):
    """Get last order date for a site from DB"""
    connection = create_database_connection()
    if not connection:
        return None
    
    try:
        cursor = connection.cursor()
        query = "SELECT MAX(date_created) FROM orders WHERE source = ?"
        cursor.execute(query, (site_url,))
        result = cursor.fetchone()
        return result[0] if result[0] else None
    except Exception as e:
        print(f"Error getting last order date: {e}")
        return None
    finally:
        if connection:
            connection.close()

def get_last_modified_date_from_db(site_url):
    """Get last modified date for a site from DB"""
    connection = create_database_connection()
    if not connection:
        return None
    try:
        cursor = connection.cursor()
        query = "SELECT MAX(date_modified) FROM orders WHERE source = ?"
        cursor.execute(query, (site_url,))
        result = cursor.fetchone()
        return result[0] if result[0] else None
    except Exception as e:
        return None
    finally:
        if connection:
            connection.close()

WC_ORDER_FIELDS = [
    'id', 'parent_id', 'number', 'order_key', 'created_via', 'version',
    'status', 'currency', 'date_created', 'date_created_gmt', 'date_modified',
    'date_modified_gmt', 'discount_total', 'discount_tax', 'shipping_total',
    'shipping_tax', 'cart_tax', 'total', 'total_tax', 'prices_include_tax',
    'customer_id', 'customer_ip_address', 'customer_user_agent',
    'customer_note', 'billing', 'shipping', 'payment_method',
    'payment_method_title', 'transaction_id', 'date_paid', 'date_paid_gmt',
    'date_completed', 'date_completed_gmt', 'cart_hash', 'meta_data',
    'line_items', 'tax_lines', 'shipping_lines', 'fee_lines', 'coupon_lines',
    'refunds', 'set_paid', 'source',
]


def upsert_orders_in_transaction(orders_data, connection):
    """Write one bounded page without committing or swallowing failures."""

    filtered = [
        order for order in (orders_data or [])
        if order.get('status') != 'checkout-draft'
    ]
    all_columns = WC_ORDER_FIELDS + ['woo_id', 'updated_at']
    placeholders = ', '.join(['?'] * len(all_columns))
    update_set = ', '.join(
        f'{column} = excluded.{column}'
        for column in all_columns
        if column != 'id'
    )
    insert_query = f"""
        INSERT INTO orders ({', '.join(all_columns)})
        VALUES ({placeholders})
        ON CONFLICT(id) DO UPDATE SET {update_set}
    """

    processed_orders = []
    planning_candidates = []
    identity_versions = {}
    now = datetime.now().isoformat()
    for order in filtered:
        woo_id = order.get('id')
        site_id = site_id_for_source(connection, order.get('source'))
        if site_id is None:
            raise ValueError(
                f"unknown source for WooCommerce order {woo_id}"
            )
        oid = make_oid(site_id, woo_id)
        values = []
        for field in WC_ORDER_FIELDS:
            value = order.get(field)
            if field == 'id':
                values.append(oid)
            elif field == 'set_paid':
                values.append(
                    0 if isinstance(value, dict) or value is None
                    else (1 if value else 0)
                )
            elif field == 'prices_include_tax':
                values.append(bool(value))
            elif isinstance(value, (dict, list)):
                values.append(json.dumps(value, ensure_ascii=False))
            else:
                values.append(value)
        values.extend((woo_id, now))
        processed_orders.append(tuple(values))
        identity_versions[str(oid)] = (
            str(order.get('date_modified') or ''),
            str(order.get('status') or ''),
        )
        planning_candidates.append({
            'order_id': oid,
            'status': order.get('status'),
            'date_modified': order.get('date_modified'),
            'source': order.get('source'),
        })

    existing = {}
    ids = list(identity_versions)
    for start in range(0, len(ids), 500):
        chunk = ids[start:start + 500]
        marks = ','.join('?' for _ in chunk)
        for row in connection.execute(
            f"SELECT id,date_modified,status FROM orders WHERE id IN ({marks})",
            tuple(chunk),
        ).fetchall():
            existing[str(row['id'])] = (
                str(row['date_modified'] or ''),
                str(row['status'] or ''),
            )

    if processed_orders:
        connection.executemany(insert_query, processed_orders)
        reconcile_legacy_terminal_shortages(connection, planning_candidates)
    inserted = sum(1 for oid in identity_versions if oid not in existing)
    changed = sum(
        1 for oid, version in identity_versions.items()
        if oid not in existing or existing[oid] != version
    )
    return {
        'written': len(processed_orders),
        'inserted': inserted,
        'changed': changed,
        'planning_candidates': planning_candidates,
    }


def upsert_order_notes_in_transaction(notes_data, connection):
    """Upsert fetched order notes in the caller's page transaction."""

    rows = []
    for note in notes_data or []:
        order_id = note.get('_local_order_id')
        if order_id is None or note.get('id') is None:
            continue
        rows.append((
            note.get('id'),
            order_id,
            note.get('note', ''),
            note.get('date_created', ''),
            bool(note.get('customer_note', False)),
            note.get('author', ''),
            bool(note.get('added_by_user', False)),
        ))
    if not rows:
        return 0
    connection.executemany(
        """
        INSERT INTO order_notes (
            wc_note_id,order_id,note,date_created,customer_note,author,added_by_user
        ) VALUES (?,?,?,?,?,?,?)
        ON CONFLICT(order_id,wc_note_id) DO UPDATE SET
            note=excluded.note,
            date_created=excluded.date_created,
            customer_note=excluded.customer_note,
            author=CASE
                WHEN order_notes.added_by_user IS TRUE
                     AND COALESCE(order_notes.author,'') NOT IN ('','WooCommerce')
                THEN order_notes.author
                ELSE excluded.author
            END,
            added_by_user=COALESCE(order_notes.added_by_user,FALSE)
                          OR excluded.added_by_user
        """,
        rows,
    )
    return len(rows)


def run_post_commit_sync_actions(planning_candidates, *, strict=False):
    """Run existing idempotent local queue hooks only after page commit."""

    _enqueue_fulfillment_plans(planning_candidates, raise_on_error=strict)
    _enqueue_order_notifications(planning_candidates, raise_on_error=strict)


def save_orders_to_db(orders_data, connection=None):
    """Legacy wrapper; Celery uses upsert_orders_in_transaction directly."""

    if not orders_data:
        return {'written': 0, 'inserted': 0, 'changed': 0}
    own_connection = connection is None
    if own_connection:
        connection = create_database_connection()
        if not connection:
            raise RuntimeError("database connection unavailable")
    try:
        result = upsert_orders_in_transaction(orders_data, connection)
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        if own_connection and connection:
            connection.close()
    run_post_commit_sync_actions(result['planning_candidates'])
    return result

def fetch_orders_incrementally(wcapi, site_url, last_order_date=None, progress_callback=None, connection=None):
    """Fetch orders incrementally"""
    orders = []
    page = 1
    per_page = 100
    max_retries = 3
    retry_count = 0

    params = {
        "per_page": per_page,
        "page": page,
        "expand": "line_items,shipping_lines,tax_lines,fee_lines,coupon_lines,refunds"
    }

    if last_order_date:
        params['after'] = last_order_date
        if progress_callback: progress_callback(f"Fetching orders after {last_order_date}...")

    while True:
        try:
            if progress_callback: progress_callback(f"Fetching page {page}...")
            response = wcapi.get("orders", params=params)

            if response.status_code != 200:
                if retry_count < max_retries:
                    retry_count += 1
                    if progress_callback: progress_callback(f"Error {response.status_code}, retrying ({retry_count}/{max_retries})...")
                    time.sleep(2)
                    continue
                else:
                    if progress_callback: progress_callback(f"Failed after max retries.")
                    break

            data = response.json()
            if not data:
                if progress_callback: progress_callback(f"No more orders found.")
                break

            for order in data:
                order['source'] = site_url

            save_orders_to_db(data, connection=connection)
            orders.extend(data)
            
            if progress_callback: progress_callback(f"Saved {len(data)} orders from page {page}.")
            
            page += 1
            params['page'] = page
            retry_count = 0

        except Exception as e:
            if progress_callback: progress_callback(f"Error: {str(e)}")
            if retry_count < max_retries:
                retry_count += 1
                time.sleep(2)
                continue
            else:
                break

    return orders

def fetch_orders_modified_after(wcapi, site_url, modified_after=None, progress_callback=None, connection=None):
    """Fetch modified orders"""
    orders = []
    page = 1
    per_page = 100
    max_retries = 3
    retry_count = 0
    
    params = {
        "per_page": per_page,
        "page": page,
        "expand": "line_items,shipping_lines,tax_lines,fee_lines,coupon_lines,refunds"
    }
    
    if modified_after:
        params['modified_after'] = modified_after
        if progress_callback: progress_callback(f"Checking for updates after {modified_after}...")
        
    while True:
        try:
            if progress_callback: progress_callback(f"Fetching updates page {page}...")
            response = wcapi.get("orders", params=params)
            
            if response.status_code != 200:
                if retry_count < max_retries:
                    retry_count += 1
                    time.sleep(2)
                    continue
                else:
                    break
                    
            data = response.json()
            if not data:
                break
                
            for order in data:
                order['source'] = site_url
                
            save_orders_to_db(data, connection=connection)
            orders.extend(data)
            
            if progress_callback: progress_callback(f"Updated {len(data)} orders from page {page}.")
            
            page += 1
            params['page'] = page
            retry_count = 0
            
        except Exception as e:
            if progress_callback: progress_callback(f"Error: {str(e)}")
            if retry_count < max_retries:
                retry_count += 1
                time.sleep(2)
                continue
            else:
                break
                
    return orders

def sync_order_notes(
    wcapi,
    site_url,
    connection=None,
    changed_woo_ids=None,
    active_limit=DEFAULT_NOTES_ACTIVE_LIMIT,
    active_refresh_interval_hours=DEFAULT_NOTES_REFRESH_INTERVAL_HOURS,
    max_workers=DEFAULT_NOTE_WORKERS,
    progress_callback=None,
):
    """Refresh notes for changed orders plus a bounded active-order rotation.

    The old hourly path fetched notes for every active order, using five note
    workers inside each of four concurrent site workers. This bounded rotation
    keeps changed orders fresh without flooding small WooCommerce servers.
    """
    result = {"candidates": 0, "synced": 0, "failed": 0, "notes": 0}
    if not connection:
        return result

    try:
        cursor = connection.cursor()
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS order_note_sync_state (
                order_id TEXT PRIMARY KEY,
                last_synced_at TEXT NOT NULL
            )
            """
        )
        connection.commit()

        candidates = {}
        changed_ids = []
        for value in changed_woo_ids or []:
            try:
                changed_ids.append(int(value))
            except (TypeError, ValueError):
                continue

        # SQLite commonly limits a statement to 999 bound parameters.
        for start in range(0, len(changed_ids), 500):
            chunk = changed_ids[start:start + 500]
            placeholders = ",".join("?" for _ in chunk)
            cursor.execute(
                f"""
                SELECT id, woo_id
                FROM orders
                WHERE source = ? AND woo_id IN ({placeholders})
                """,
                (site_url, *chunk),
            )
            for oid, woo_id in cursor.fetchall():
                candidates[str(oid)] = (oid, woo_id)

        stale_before = (
            datetime.now()
            - timedelta(hours=max(0, int(active_refresh_interval_hours or 0)))
        ).isoformat(timespec="seconds")

        if active_limit is None:
            limit_clause = ""
            params = (site_url, stale_before)
        else:
            limit_clause = "LIMIT ?"
            params = (site_url, stale_before, max(0, int(active_limit)))

        if active_limit is None or int(active_limit) > 0:
            cursor.execute(
                f"""
                SELECT o.id, o.woo_id
                FROM orders AS o
                LEFT JOIN order_note_sync_state AS s
                  ON s.order_id = CAST(o.id AS TEXT)
                WHERE o.source = ?
                  AND o.status IN ('processing', 'offline', 'on-hold')
                  AND (
                    s.last_synced_at IS NULL
                    OR s.last_synced_at <= ?
                  )
                ORDER BY
                  CASE WHEN s.last_synced_at IS NULL THEN 0 ELSE 1 END,
                  COALESCE(s.last_synced_at, '') ASC,
                  COALESCE(o.date_modified, '') DESC
                {limit_clause}
                """,
                params,
            )
            for oid, woo_id in cursor.fetchall():
                candidates.setdefault(str(oid), (oid, woo_id))

        selected = list(candidates.values())
        result["candidates"] = len(selected)
        if not selected:
            return result

        if progress_callback:
            progress_callback(
                f"Refreshing notes for {len(selected)} orders "
                f"(changed={len(changed_ids)}, active_cap={active_limit})..."
            )

        def fetch_notes_for_order(oid, woo_id):
            wc_pid = woo_id if woo_id is not None else woo_post_id(oid)
            for attempt in range(3):
                try:
                    response = wcapi.get(f"orders/{wc_pid}/notes")
                    if response.status_code == 200:
                        notes_data = response.json()
                        for note in notes_data:
                            note["_local_order_id"] = oid
                        return oid, notes_data, True
                    if response.status_code not in {429, 502, 503, 504}:
                        break
                except Exception as exc:
                    if attempt == 2:
                        print(f"Error fetching notes for order {oid}: {exc}")
                if attempt < 2:
                    time.sleep(2 ** attempt)
            return oid, [], False

        fetched = []
        workers = max(1, int(max_workers or 1))
        if workers == 1:
            for oid, woo_id in selected:
                fetched.append(fetch_notes_for_order(oid, woo_id))
        else:
            with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
                futures = [
                    executor.submit(fetch_notes_for_order, oid, woo_id)
                    for oid, woo_id in selected
                ]
                for future in concurrent.futures.as_completed(futures):
                    fetched.append(future.result())

        all_notes = []
        successful_order_ids = []
        for oid, notes, success in fetched:
            if success:
                successful_order_ids.append(
                    (str(oid), datetime.now().isoformat(timespec="seconds"))
                )
                result["synced"] += 1
                all_notes.extend(notes)
            else:
                result["failed"] += 1

        if all_notes:
            # Keep the legacy direct-sync path on the same PostgreSQL-safe,
            # idempotent upsert used by the durable page writer.
            result["notes"] = upsert_order_notes_in_transaction(
                all_notes, connection
            )

        if successful_order_ids:
            cursor.executemany(
                """
                INSERT INTO order_note_sync_state (order_id, last_synced_at)
                VALUES (?, ?)
                ON CONFLICT(order_id) DO UPDATE SET
                    last_synced_at = excluded.last_synced_at
                """,
                successful_order_ids,
            )
        connection.commit()
        return result

    except Exception as exc:
        try:
            connection.rollback()
        except Exception:
            pass
        print(f"Error syncing order notes: {exc}")
        result["failed"] += 1
        return result

def sync_site(
    url,
    consumer_key,
    consumer_secret,
    progress_callback=None,
    sync_days=0,
    full_history=False,
    incremental_overlap_minutes=DEFAULT_INCREMENTAL_OVERLAP_MINUTES,
    notes_active_limit=DEFAULT_NOTES_ACTIVE_LIMIT,
    notes_refresh_interval_hours=DEFAULT_NOTES_REFRESH_INTERVAL_HOURS,
    note_workers=DEFAULT_NOTE_WORKERS,
):
    """Sync a single site

    Args:
        sync_days: Optional legacy lookback in days. The default uses the last
                   stored modification timestamp with a small safety overlap.
        full_history: When True, ignore all local cutoffs and fetch every order
                      page from the WooCommerce API. Use this for first-time
                      sync of a site or when local DB is missing historical data.
        incremental_overlap_minutes: Re-read this many minutes before the last
                                     stored modification timestamp.
        notes_active_limit: Maximum stale active orders to rotate per run, in
                            addition to orders changed in the current run.
        notes_refresh_interval_hours: Minimum age before an unchanged active
                                      order is eligible for note rotation.
        note_workers: Per-site order-note request concurrency.
    """
    if progress_callback: progress_callback(f"Connecting to {url}...")

    wcapi = create_robust_wcapi(url, consumer_key, consumer_secret, PROXY_CONFIG)
    if not wcapi:
        return {"status": "error", "message": "Failed to create API client"}

    # 使用线程局部数据库连接，整个同步过程复用
    conn = get_thread_db_connection()

    try:
        new_orders = []
        updated_orders = []

        if full_history:
            if progress_callback:
                progress_callback("Full history sync (no date filter)...")
            new_orders = fetch_orders_incrementally(
                wcapi,
                url,
                None,
                progress_callback,
                connection=conn,
            )
        elif sync_days and sync_days > 0:
            cutoff_date = datetime.now() - timedelta(days=sync_days)
            modified_after = cutoff_date.strftime("%Y-%m-%dT00:00:00")
            if progress_callback:
                progress_callback(f"Syncing orders modified in last {sync_days} days...")
            updated_orders = fetch_orders_modified_after(
                wcapi,
                url,
                modified_after,
                progress_callback,
                connection=conn,
            )
        else:
            checkpoint = get_last_modified_date_from_db(url)
            modified_after = _subtract_minutes_from_iso(
                checkpoint,
                incremental_overlap_minutes,
            )
            if progress_callback:
                if modified_after:
                    progress_callback(
                        f"Incremental sync from {modified_after} "
                        f"({incremental_overlap_minutes}m overlap)..."
                    )
                else:
                    progress_callback("No local checkpoint; fetching full history once...")
            updated_orders = fetch_orders_modified_after(
                wcapi,
                url,
                modified_after,
                progress_callback,
                connection=conn,
            )

        changed_woo_ids = [
            order.get("id")
            for order in (new_orders + updated_orders)
            if order.get("id") is not None
        ]
        note_result = sync_order_notes(
            wcapi,
            url,
            connection=conn,
            changed_woo_ids=changed_woo_ids,
            active_limit=notes_active_limit,
            active_refresh_interval_hours=notes_refresh_interval_hours,
            max_workers=note_workers,
            progress_callback=progress_callback,
        )

        msg = (
            f"Sync complete. New: {len(new_orders)}, Updated: {len(updated_orders)}, "
            f"Notes checked: {note_result['synced']}, note failures: {note_result['failed']}"
        )
        if progress_callback: progress_callback(msg)
        
        return {
            "status": "success",
            "new_orders": len(new_orders),
            "updated_orders": len(updated_orders),
            "notes_checked": note_result["synced"],
            "note_failures": note_result["failed"],
        }
    except Exception as e:
        if progress_callback: progress_callback(f"Critical Error: {str(e)}")
        return {"status": "error", "message": str(e)}
    finally:
        close_thread_db_connection()
