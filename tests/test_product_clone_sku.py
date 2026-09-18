from datetime import datetime

import pytest

from product_clone_sku import (
    MAX_WC_SKU_LENGTH,
    build_clone_sku,
    make_clone_suffix,
    normalize_clone_suffix,
)


def test_make_clone_suffix_is_safe_and_stable_with_explicit_inputs():
    suffix = make_clone_suffix(datetime(2026, 8, 21, 12, 0), "a1b2c3")
    assert suffix == "NEW-20260821-A1B2C3"
    assert normalize_clone_suffix(" new / 20260821 / a1b2c3 ") == "NEW-20260821-A1B2C3"


def test_build_clone_sku_namespaces_parent_and_variation():
    suffix = "NEW-20260821-A1B2C3"
    assert build_clone_sku("PARENT", suffix, fallback="unused") == f"PARENT-{suffix}"
    assert build_clone_sku("", suffix, fallback="VAR-10-20") == f"VAR-10-20-{suffix}"


def test_build_clone_sku_truncates_with_hash_and_never_exceeds_wc_limit():
    sku = build_clone_sku("X" * 150, "NEW-20260821-A1B2C3", fallback="unused")
    assert len(sku) == MAX_WC_SKU_LENGTH
    assert sku.endswith("-NEW-20260821-A1B2C3")
    assert "-" in sku[: -len("NEW-20260821-A1B2C3")]


def test_build_clone_sku_requires_suffix():
    with pytest.raises(ValueError):
        build_clone_sku("SKU", "", fallback="unused")
