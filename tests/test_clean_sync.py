"""Fail-closed WooCommerce evidence checks for PostgreSQL clean sync."""

import pytest

import clean_sync


class Response:
    def __init__(self, status=200, body=None, total=None, pages=None):
        self.status_code = status
        self.body = body
        self.headers = {}
        if total is not None:
            self.headers["X-WP-Total"] = str(total)
        if pages is not None:
            self.headers["X-WP-TotalPages"] = str(pages)

    def json(self):
        return self.body


class API:
    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = []

    def get(self, path, params=None):
        self.calls.append((path, params))
        response = self.responses.pop(0)
        if isinstance(response, Exception):
            raise response
        return response


def test_complete_stable_scan_includes_every_page():
    api = API([
        Response(body=[{"id": 1}, {"id": 2}], total=3, pages=2),
        Response(body=[{"id": 3}], total=3, pages=2),
        Response(body=[{"id": 1}, {"id": 2}], total=3, pages=2),
    ])
    progress = []
    assert clean_sync.scan_remote_ids(
        api, lambda page, pages, count: progress.append((page, pages, count))
    ) == ({1, 2, 3}, 3, 2)
    assert progress == [(1, 2, 2), (2, 2, 3)]
    assert all(call[1]["status"] == "any" for call in api.calls)


@pytest.mark.parametrize("responses", [
    [Response(body=[{"id": 1}], total=2, pages=1),
     Response(body=[{"id": 1}], total=2, pages=1)],
    [Response(body=[{"id": 1}], total=1, pages=1),
     Response(body=[{"id": 2}], total=1, pages=1)],
    [Response(body={"id": 1}, total=1, pages=1)],
    [Response(body=[], total=None, pages=None)],
])
def test_incomplete_changing_or_malformed_scan_never_produces_candidates(responses):
    with pytest.raises(clean_sync.CleanSyncError):
        clean_sync.scan_remote_ids(API(responses))


def test_verified_empty_shop_is_distinct_from_failed_fetch():
    assert clean_sync.scan_remote_ids(API([
        Response(body=[], total=0, pages=0),
        Response(body=[], total=0, pages=0),
    ])) == (set(), 0, 0)
    api = API([Response(status=525)] * 3)
    with pytest.raises(clean_sync.CleanSyncError, match="525"):
        clean_sync.scan_remote_ids(api)


def test_only_exact_404_and_empty_trash_prove_remote_deletion():
    invalid = {"code": "woocommerce_rest_shop_order_invalid_id"}
    assert clean_sync.confirm_remote_deleted(API([
        Response(status=404, body=invalid),
        Response(body=[]),
    ]), 123)
    assert not clean_sync.confirm_remote_deleted(API([
        Response(status=200, body={"id": 123, "status": "trash"}),
    ]), 123)
    assert not clean_sync.confirm_remote_deleted(API([
        Response(status=404, body=invalid),
        Response(body=[{"id": 123}]),
    ]), 123)
    with pytest.raises(clean_sync.CleanSyncError):
        clean_sync.confirm_remote_deleted(API([
            Response(status=404, body={"code": "rest_forbidden"}),
        ]), 123)
