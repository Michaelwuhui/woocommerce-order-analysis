"""Pure helpers for item-level split-shipment validation.

The shipping UI and WooCommerce adapter both use Woo line-item IDs as the
stable key.  Keeping this logic outside the Flask route makes the quantity
rules easy to test without touching production data or sending requests.
"""

from __future__ import annotations


class ShipmentItemError(ValueError):
    """Raised when a parcel contains invalid or over-shipped quantities."""


def order_products(line_items):
    """Return canonical AST products for the order's shippable line items."""
    products = []
    for item in line_items or []:
        if not isinstance(item, dict) or item.get("id") is None:
            continue
        try:
            quantity = int(item.get("quantity") or 0)
        except (TypeError, ValueError):
            quantity = 0
        if quantity <= 0:
            continue
        products.append({
            # AST's email template compares this value with the concrete order
            # item product.  For variations that must be the variation ID, not
            # the variable parent ID.
            "product": str(item.get("variation_id") or item.get("product_id") or ""),
            "item_id": str(item["id"]),
            "qty": str(quantity),
        })
    return products


def normalize_batch_items(line_items, requested, *, require_explicit=False):
    """Validate a parcel selection and return canonical AST products.

    ``requested`` is a list of ``{item_id, qty}``.  A normal single-parcel
    shipment may omit it and defaults to the whole order.  Split shipments
    require an explicit selection so a tracking number can never silently be
    attached to every order item.
    """
    ordered = {p["item_id"]: p for p in order_products(line_items)}
    if not requested:
        if require_explicit:
            raise ShipmentItemError("分批发货必须选择本批商品和数量")
        return list(ordered.values())
    if not isinstance(requested, list):
        raise ShipmentItemError("本批商品格式错误")

    quantities = {}
    for row in requested:
        if not isinstance(row, dict):
            raise ShipmentItemError("本批商品格式错误")
        item_id = str(row.get("item_id") or "").strip()
        if item_id not in ordered:
            raise ShipmentItemError(f"订单商品行 {item_id or '?'} 不存在")
        try:
            qty = int(row.get("qty"))
        except (TypeError, ValueError):
            raise ShipmentItemError(f"订单商品行 {item_id} 的数量无效")
        if qty <= 0:
            raise ShipmentItemError(f"订单商品行 {item_id} 的数量必须大于 0")
        quantities[item_id] = quantities.get(item_id, 0) + qty

    products = []
    for item_id, qty in quantities.items():
        maximum = int(ordered[item_id]["qty"])
        if qty > maximum:
            raise ShipmentItemError(
                f"订单商品行 {item_id} 本批数量 {qty} 超过下单数量 {maximum}"
            )
        products.append({**ordered[item_id], "qty": str(qty)})
    return products


def remaining_after(line_items, prior_parcels, current_products):
    """Return remaining quantities after applying the proposed parcel.

    Every prior parcel must carry item-level metadata.  Refusing an ambiguous
    legacy row is safer than allowing a continuation that over-ships goods.
    """
    ordered = {p["item_id"]: int(p["qty"]) for p in order_products(line_items)}
    shipped = {item_id: 0 for item_id in ordered}

    for parcel in prior_parcels or []:
        products = parcel.get("products_list")
        if not products:
            raise ShipmentItemError(
                "已有包裹缺少商品数量记录，请先由管理员补齐后再继续分批发货"
            )
        for product in products:
            item_id = str(product.get("item_id") or "")
            if item_id not in ordered:
                raise ShipmentItemError(f"历史包裹包含未知订单商品行 {item_id or '?'}")
            try:
                qty = int(product.get("qty") or 0)
            except (TypeError, ValueError):
                raise ShipmentItemError(f"历史包裹商品行 {item_id} 数量无效")
            if qty <= 0:
                raise ShipmentItemError(f"历史包裹商品行 {item_id} 数量无效")
            shipped[item_id] += qty

    for product in current_products or []:
        item_id = str(product.get("item_id") or "")
        if item_id not in ordered:
            raise ShipmentItemError(f"本批包含未知订单商品行 {item_id or '?'}")
        shipped[item_id] += int(product.get("qty") or 0)

    remaining = {}
    for item_id, ordered_qty in ordered.items():
        if shipped[item_id] > ordered_qty:
            raise ShipmentItemError(
                f"订单商品行 {item_id} 累计发货 {shipped[item_id]} 超过下单数量 {ordered_qty}"
            )
        remaining[item_id] = ordered_qty - shipped[item_id]
    return remaining
