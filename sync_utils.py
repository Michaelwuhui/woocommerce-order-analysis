import json
import hashlib
import time
import sqlite3
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


def _enqueue_fulfillment_plans(candidates):
    """Queue planning after the Woo order transaction has committed.

    A separate row-factory connection isolates fulfillment failures from the
    core sync.  The dark-launch flags prevent historical orders being planned
    until operations explicitly enables the workflow.
    """
    if not candidates:
        return
    try:
        from fulfillment_common import get_conn
        from fulfillment_service import enqueue_job

        conn = get_conn()
        try:
            if not _flag_enabled(conn, 'oms_fulfillment_enabled') or not _flag_enabled(conn, 'oms_auto_plan_enabled'):
                return
            for item in candidates:
                if item['status'] not in {'processing', 'offline', 'on-hold', 'partial-shipped', 'shipped'}:
                    continue
                site = conn.execute(
                    "SELECT country FROM sites WHERE url=?", (item['source'],)
                ).fetchone()
                if not site or str(site['country'] or '').upper() not in {'PL', 'CZ', 'HU'}:
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
        # Order sync remains authoritative and must not fail because the
        # optional fulfillment queue is unavailable during rollout.
        print(f"[fulfillment] enqueue skipped: {exc}")

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
        connection = sqlite3.connect(DB_FILE)
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

def save_orders_to_db(orders_data, connection=None):
    """Save orders to SQLite database"""
    if not orders_data:
        return
    
    own_connection = connection is None
    if own_connection:
        connection = create_database_connection()
        if not connection:
            return
    
    try:
        cursor = connection.cursor()

        # WC-managed columns. Local-only columns (is_undelivered,
        # shipping_loss_amount, undelivered_*, is_problem_return,
        # problem_return_*, product_loss_amount) are deliberately NOT listed —
        # they're set by /api/order/<id>/mark-* and must survive every sync.
        # Previously we used INSERT OR REPLACE which DELETEs the row first,
        # wiping those flags on every refresh; UPSERT only touches the columns
        # we name, so the local markings stay intact.
        wc_fields = [
            'id', 'parent_id', 'number', 'order_key', 'created_via', 'version', 'status', 'currency',
            'date_created', 'date_created_gmt', 'date_modified', 'date_modified_gmt',
            'discount_total', 'discount_tax', 'shipping_total', 'shipping_tax', 'cart_tax',
            'total', 'total_tax', 'prices_include_tax', 'customer_id', 'customer_ip_address',
            'customer_user_agent', 'customer_note', 'billing', 'shipping', 'payment_method',
            'payment_method_title', 'transaction_id', 'date_paid', 'date_paid_gmt',
            'date_completed', 'date_completed_gmt', 'cart_hash', 'meta_data', 'line_items',
            'tax_lines', 'shipping_lines', 'fee_lines', 'coupon_lines', 'refunds', 'set_paid', 'source'
        ]
        # woo_id keeps the raw per-site WC post id; id is the cross-site-safe
        # surrogate "<sites.id>-<woo_id>" so same-numbered orders from different
        # stores no longer collide under ON CONFLICT(id). See oid_utils.py.
        all_columns = wc_fields + ['woo_id', 'updated_at']
        placeholders = ', '.join(['?'] * len(all_columns))
        # On UPDATE, set every column EXCEPT id (the conflict key).
        update_set = ', '.join(f'{c} = excluded.{c}' for c in all_columns if c != 'id')
        insert_query = f"""
        INSERT INTO orders ({', '.join(all_columns)})
        VALUES ({placeholders})
        ON CONFLICT(id) DO UPDATE SET {update_set}
        """

        # Filter out checkout-draft orders - they should not be synced
        orders_data = [o for o in orders_data if o.get('status') != 'checkout-draft']

        if not orders_data:
            return

        processed_orders = []
        planning_candidates = []
        for order in orders_data:
            woo_id = order.get('id')
            site_id = site_id_for_source(connection, order.get('source'))
            if site_id is None:
                # Unknown source -> cannot build a safe surrogate; skip rather
                # than mis-key. (Should not happen: every synced site is in `sites`.)
                print(f"[save_orders_to_db] skip order {woo_id}: unknown source {order.get('source')!r}")
                continue
            oid = make_oid(site_id, woo_id)
            processed_order = []
            for field in wc_fields:
                value = order.get(field)
                if field == 'id':
                    processed_order.append(oid)
                elif field == 'set_paid':
                    if isinstance(value, dict) or value is None:
                        processed_order.append(0)
                    else:
                        processed_order.append(1 if value else 0)
                elif field == 'prices_include_tax':
                    processed_order.append(1 if value else 0)
                elif isinstance(value, (dict, list)):
                    processed_order.append(json.dumps(value, ensure_ascii=False))
                else:
                    processed_order.append(value)

            processed_order.append(woo_id)                      # woo_id
            processed_order.append(datetime.now().isoformat())  # updated_at
            processed_orders.append(tuple(processed_order))
            planning_candidates.append({
                'order_id': oid,
                'status': order.get('status'),
                'date_modified': order.get('date_modified'),
                'source': order.get('source'),
            })

        cursor.executemany(insert_query, processed_orders)
        connection.commit()
        _enqueue_fulfillment_plans(planning_candidates)
        
    except Exception as e:
        print(f"Error saving orders: {e}")
    finally:
        if own_connection and connection:
            connection.close()

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
            insert_query = """
            INSERT INTO order_notes (
                wc_note_id, order_id, note, date_created, customer_note, author, added_by_user
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(order_id, wc_note_id) DO UPDATE SET
                note = excluded.note,
                date_created = excluded.date_created,
                customer_note = excluded.customer_note,
                author = CASE
                    WHEN order_notes.added_by_user = 1
                         AND COALESCE(order_notes.author, '') NOT IN ('', 'WooCommerce')
                    THEN order_notes.author
                    ELSE excluded.author
                END,
                added_by_user = CASE
                    WHEN order_notes.added_by_user = 1 THEN 1
                    ELSE excluded.added_by_user
                END
            """
            processed_notes = []
            for note in all_notes:
                processed_notes.append((
                    note.get("id"),
                    note["_local_order_id"],
                    note.get("note", ""),
                    note.get("date_created", ""),
                    1 if note.get("customer_note", False) else 0,
                    note.get("author", ""),
                    1 if note.get("added_by_user", False) else 0,
                ))
            cursor.executemany(insert_query, processed_notes)
            result["notes"] = len(processed_notes)

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
