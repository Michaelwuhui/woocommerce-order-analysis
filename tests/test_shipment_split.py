import unittest

from shipment_split import (
    ShipmentItemError,
    normalize_batch_items,
    order_products,
    remaining_after,
)


LINES = [
    {"id": 4034, "product_id": 13792, "variation_id": 13802, "quantity": 10},
    {"id": 4035, "product_id": 13792, "variation_id": 13794, "quantity": 10},
    {"id": 4036, "product_id": 13792, "variation_id": 13796, "quantity": 10},
]


class ShipmentSplitTests(unittest.TestCase):
    def test_single_parcel_defaults_to_all_items(self):
        self.assertEqual(3, len(normalize_batch_items(LINES, None)))

    def test_split_requires_explicit_items(self):
        with self.assertRaisesRegex(ShipmentItemError, "必须选择"):
            normalize_batch_items(LINES, [], require_explicit=True)

    def test_watermelon_batch_is_item_level(self):
        products = normalize_batch_items(
            LINES, [{"item_id": 4035, "qty": 10}], require_explicit=True
        )
        self.assertEqual(
            [{"product": "13794", "item_id": "4035", "qty": "10"}],
            products,
        )
        self.assertEqual(
            {"4034": 10, "4035": 0, "4036": 10},
            remaining_after(LINES, [], products),
        )

    def test_cumulative_over_ship_is_rejected(self):
        prior = [{"products_list": [{"item_id": "4035", "qty": "10"}]}]
        current = [{"item_id": "4035", "qty": "1"}]
        with self.assertRaisesRegex(ShipmentItemError, "超过下单数量"):
            remaining_after(LINES, prior, current)

    def test_legacy_parcel_without_items_blocks_continuation(self):
        with self.assertRaisesRegex(ShipmentItemError, "缺少商品数量"):
            remaining_after(LINES, [{}], order_products(LINES))


if __name__ == "__main__":
    unittest.main()
