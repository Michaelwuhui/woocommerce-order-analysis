from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
RECON_TEMPLATE = (ROOT / "templates" / "partner_reconciliation.html").read_text(
    encoding="utf-8"
)
COSTS_TEMPLATE = (ROOT / "templates" / "product_costs.html").read_text(
    encoding="utf-8"
)


def test_reconciliation_unmapped_link_carries_exact_context():
    assert 'id="reconUnmappedCostLink"' in RECON_TEMPLATE
    for context_field in ("unmapped: '1'", "year_month: period", "partner_id: String(partnerId)"):
        assert context_field in RECON_TEMPLATE


def test_cost_page_opens_same_partner_unmapped_dataset():
    assert "params.get('unmapped') === '1'" in COSTS_TEMPLATE
    assert "document.getElementById('unmappedMonth').value = deepMonth" in COSTS_TEMPLATE
    assert "window.setTimeout(showUnmapped, 0)" in COSTS_TEMPLATE
    assert "/api/reconciliation/unmapped-products?partner_id=${reconUnmappedPartnerId}" in COSTS_TEMPLATE
    assert "来自合伙人对账的同一口径" in COSTS_TEMPLATE
