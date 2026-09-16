"""Read-back checks before explicitly retrying a rejected legacy shipment."""

from order_shipments import extract_tracking_candidates
from shipment_split import order_products


class ShipmentRetryError(ValueError):
    pass


def classify_retry_preflight(remote, woo_order_id, local_items, tracking_number):
    """Return saved/retry only for the same unchanged order and parcel."""
    if not isinstance(remote, dict) or str(remote.get('id')) != str(woo_order_id):
        raise ShipmentRetryError('来源站点未返回对应订单，请先核对同步状态')
    status = str(remote.get('status') or '')
    if status not in {'pending', 'processing', 'on-hold', 'offline', 'shipped', 'completed'}:
        raise ShipmentRetryError('来源订单状态已变化，请先核对订单，不能直接重试发货')
    expected = sorted(order_products(local_items), key=lambda item: item['item_id'])
    actual = sorted(order_products(remote.get('line_items')), key=lambda item: item['item_id'])
    if not expected or expected != actual:
        raise ShipmentRetryError('来源订单商品或数量已变化，请先同步并核对订单')
    numbers = {
        item['tracking_number'] for item in extract_tracking_candidates(
            remote.get('meta_data'), remote.get('line_items'), remote.get('shipping_lines'), []
        )
    }
    if numbers == {tracking_number}:
        return 'saved'
    if numbers or status in {'shipped', 'completed'}:
        raise ShipmentRetryError('来源订单已有其他运单或已完成发货，请先核对包裹')
    return 'retry'
