"""Canonical SKU catalog for the Jin Yi/Jin Gu temporary transit stock.

The quantities originate from the approved Fumot-0808 transfer workbook.
This module intentionally contains product identity only; stock movements are
created by migration 017 so every quantity change remains auditable.
"""

from __future__ import annotations


JYJG_TRANSIT_WAREHOUSE_CODE = "PL-JYJG-TRANSIT"
JYJG_TRANSIT_ROUTING_POLICY = "managed_transfer_stock"
JYJG_TRANSIT_ROUTING_SETTING = "oms_jyjg_transit_routing_enabled"
JYJG_TRANSFER_SOURCE = "Fumot-0808调拨量 的副本.xlsx"


# family, product name, puff count, flavor, warehouse SKU, transferred qty
JYJG_TRANSFER_CATALOG = (
    ("fumot-eco-4in1-80k", "Fumot Eco 4in1 80k", 80000, "Magic Love & Kiwi Passion Fruit Guava & Dragonfruit Raspberry & Lemon Lime", "80K-MKDL", 10),
    ("fumot-eco-4in1-80k", "Fumot Eco 4in1 80k", 80000, "Peach Ice & Mixed Berries & Lemon Peach Passionfruit & Black Ice Dragonfruit Strawberry", "80K-PMLB", 10),
    ("fumot-eco-4in1-80k", "Fumot Eco 4in1 80k", 80000, "Strawberry Banana & Strawberry Kiwi & Strawberry Watermelon & Strawberry Ice", "80K-SSSS", 10),
    ("fumot-eco-4in1-80k", "Fumot Eco 4in1 80k", 80000, "Fresh Menthol Mojito & Pink Lemonade & Cool Mint & Mint Watermelon", "80K-FPCM", 10),
    ("fumot-eco-4in1-80k", "Fumot Eco 4in1 80k", 80000, "Grape Burst & Blueberry Sour Raspberry & Strawberry Raspberry Cherry Ice & Cherry Ice", "80K-GBSC", 10),
    ("fumot-eco-4in1-80k", "Fumot Eco 4in1 80k", 80000, "Watermelon Ice & Peach Mango & Lady Killer & Strawberry Red Bull", "80K-WPLS", 10),
    ("fumot-eco-4in1-80k", "Fumot Eco 4in1 80k", 80000, "Blueberry On Ice & Pineapple Lemonade & Peach Mango Watermelon & Blue Razz Cherry", "80K-BPPB", 10),
    ("fumot-eco-4in1-80k", "Fumot Eco 4in1 80k", 80000, "Grape Gummy Bear & Blueberry Hubba Bubba & Watermelon Bubblegum & Skittles", "80K-GBWS", 10),
    ("fumot-eco-4in1-80k", "Fumot Eco 4in1 80k", 80000, "Peach Berry & Lemon Peach & Passionfruit Grapefruit Orange & Blueberry Cherry Cranberry", "80K-PLPB", 10),
    ("fumot-eco-4in1-80k", "Fumot Eco 4in1 80k", 80000, "Cola Lime & Cherry Cola & Dr Blue & Strawberry Grape", "80K-CCDS", 10),
    ("fumot-leopard-40k", "Fumot Leopard 40K", 40000, "Mixed Berries", "40K-MB", 10),
    ("fumot-leopard-40k", "Fumot Leopard 40K", 40000, "Blueberry Raspberry", "40K-BR", 10),
    ("fumot-leopard-40k", "Fumot Leopard 40K", 40000, "Passion Fruit Kiwi", "40K-PFK", 10),
    ("fumot-leopard-40k", "Fumot Leopard 40K", 40000, "Mango Passion Fruit", "40K-MPF", 10),
    ("fumot-leopard-40k", "Fumot Leopard 40K", 40000, "Blueberry On Ice", "40K-BOI", 10),
    ("fumot-leopard-40k", "Fumot Leopard 40K", 40000, "Peach Berry", "40K-PB", 10),
    ("fumot-leopard-40k", "Fumot Leopard 40K", 40000, "Pink Lemonade", "40K-PL", 10),
    ("fumot-leopard-40k", "Fumot Leopard 40K", 40000, "Watermelon Bubblegum", "40K-WB", 10),
    ("fumot-leopard-40k", "Fumot Leopard 40K", 40000, "Blue Razz Gummy Bear", "40K-BRGB", 10),
    ("fumot-leopard-40k", "Fumot Leopard 40K", 40000, "Lemon & Lime", "40K-LL", 10),
    ("fumot-leopard-40k", "Fumot Leopard 40K", 40000, "Strawberry Watermelon", "40K-SW", 10),
    ("fumot-leopard-40k", "Fumot Leopard 40K", 40000, "Strawberry Ice", "40K-SI", 10),
    ("fumot-leopard-40k", "Fumot Leopard 40K", 40000, "Watermelon Ice", "40K-WI", 10),
    ("fumot-leopard-40k", "Fumot Leopard 40K", 40000, "Kiwi Passion Fruit Guava", "40K-KPFG", 10),
    ("fumot-leopard-40k", "Fumot Leopard 40K", 40000, "Cola Ice", "40K-CI", 10),
    ("fumot-leopard-40k", "Fumot Leopard 40K", 40000, "Pineapple Ice", "40K-PI", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Blueberry Raspberry", "9K-BR", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Strawberry Watermelon", "9K-SW", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Watermelon Bubble Gum", "9K-WBG", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Blue Razz Ice", "9K-BRI", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Kiwi Passion Fruit Guava", "9K-KPFG", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Mixed Berries", "9K-MB", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Cool Mint", "9K-CM", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Peach Berry", "9K-PB", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Strawberry Cream", "9K-SC", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Grape", "9K-G", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Strawberry Ice", "9K-SI", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Blueberry Ice", "9K-BI", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Orange Soda", "9K-OS", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Pink Lemonade", "9K-PL", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Watermelon Ice", "9K-WI", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Straw Mango", "9K-SM", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Peachy Mango Pineapple", "9K-PMP", 10),
    ("fumot-randm-tornado-9000", "Fumot Tornado 9K", 9000, "Lemon & Lime", "9K-LL", 10),
    ("fumot-randm-tornado-15000", "Fumot Tornado 15000", 15000, "Blueberry On Ice", "15K-BOI", 10),
    ("fumot-randm-tornado-15000", "Fumot Tornado 15000", 15000, "Strawberry Ice", "15K-SI", 10),
    ("fumot-randm-tornado-15000", "Fumot Tornado 15000", 15000, "Grape Ice", "15K-GI", 10),
    ("fumot-randm-tornado-15000", "Fumot Tornado 15000", 15000, "Watermelon Ice", "15K-WI", 10),
    ("fumot-randm-tornado-15000", "Fumot Tornado 15000", 15000, "Strawberry Watermelon", "15K-SW", 10),
    ("fumot-randm-tornado-15000", "Fumot Tornado 15000", 15000, "Pink Lemonade", "15K-PL", 10),
    ("fumot-randm-tornado-15000", "Fumot Tornado 15000", 15000, "Mixed Berry", "15K-MB", 10),
    ("fumot-randm-tornado-15000", "Fumot Tornado 15000", 15000, "Strawberry Kiwi", "15K-SK", 10),
    ("fumot-randm-tornado-15000", "Fumot Tornado 15000", 15000, "Triple Berry", "15K-TB", 10),
    ("fumot-randm-tornado-15000", "Fumot Tornado 15000", 15000, "Red Energy Ice", "15K-REI", 10),
    ("fumot-randm-tornado-15000", "Fumot Tornado 15000", 15000, "Peach Ice", "15K-PI", 10),
    ("fumot-randm-tornado-15000", "Fumot Tornado 15000", 15000, "Peach Mango", "15K-PM", 10),
    ("fumot-randm-tornado-15000", "Fumot Tornado 15000", 15000, "Mr Blue", "15K-MBL", 10),
    ("fumot-randm-tornado-15000", "Fumot Tornado 15000", 15000, "Ice Pop", "15K-IP", 10),
    ("fumot-randm-tornado-15000", "Fumot Tornado 15000", 15000, "Blue Razz Gummy Bear", "15K-BRGB", 10),
)


def catalog_rows() -> list[dict]:
    """Return mutable dictionaries for migration and validation callers."""

    return [
        {
            "family": family,
            "product_name": product_name,
            "puff_count": puff_count,
            "flavor": flavor,
            "sku_code": sku_code,
            "quantity": quantity,
        }
        for family, product_name, puff_count, flavor, sku_code, quantity
        in JYJG_TRANSFER_CATALOG
    ]
