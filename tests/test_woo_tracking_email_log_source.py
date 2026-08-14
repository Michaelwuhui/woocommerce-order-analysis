from pathlib import Path


SOURCE = (
    Path(__file__).resolve().parents[1]
    / "templates"
    / "woo-orders-tracking-rest-api.php"
)


def test_fluent_smtp_order_logs_always_probe_table_and_deduplicate():
    source = SOURCE.read_text(encoding="utf-8")
    assert "Version: 1.2.1" in source
    assert "$seen_log_ids = array();" in source
    assert source.count("isset($seen_log_ids[$log_id])") == 2
    assert source.count("$seen_log_ids[$log_id] = true") == 2
    assert "if ($detected_plugin) {\n            $adapters = array();" not in source
    assert "Always probe the raw table as well" in source
