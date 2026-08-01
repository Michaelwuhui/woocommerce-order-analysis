from pathlib import Path
import unittest
from unittest.mock import Mock

import carrier_tracking as tracking


class CarrierClassificationTests(unittest.TestCase):
    def test_packeta_provider_names_are_explicit(self):
        for provider in ("packeta", "Packeta CZ", "zasilkovna", "Zásilkovna"):
            with self.subTest(provider=provider):
                self.assertEqual(
                    "packeta",
                    tracking.classify_carrier(provider, "Z1465635854"),
                )

    def test_packeta_packet_id_is_detected_without_provider(self):
        self.assertEqual(
            "packeta",
            tracking.classify_carrier("", "Z1465635854"),
        )

    def test_non_packeta_z_number_is_not_misclassified(self):
        self.assertEqual(
            "unknown",
            tracking.classify_carrier("", "Z12345"),
        )

    def test_packeta_has_fixed_track718_code(self):
        self.assertEqual(
            tracking.TRACK718_PACKETA,
            tracking.track718_code_for("packeta"),
        )

    def test_hungary_expressone_number_uses_dedicated_adapter(self):
        self.assertEqual(
            "expressone_hu",
            tracking.classify_carrier(
                "packeta-hu", "671555557697000013601086", "HU"
            ),
        )
        self.assertEqual(
            "packeta",
            tracking.classify_carrier("packeta-hu", "Z1465635854", "HU"),
        )
        self.assertEqual(
            "expressone_hu",
            tracking.classify_carrier(
                "packeta", "671555557697000013601086", "HU"
            ),
        )

    def test_hungary_packeta_official_url_depends_on_number_type(self):
        self.assertEqual(
            "https://tracking.packeta.com/en/Z1465635854",
            tracking.packeta_hu_official_url("Z1465635854"),
        )
        self.assertEqual(
            "https://tracking.expressone.hu/?plc_number=671555557697000013601086",
            tracking.packeta_hu_official_url("671555557697000013601086"),
        )
        self.assertEqual(
            "expressone_hu",
            tracking.classify_carrier(
                "custom", "671555557697000013601086", "HU"
            ),
        )


class ExpressOneHungaryTests(unittest.TestCase):
    def test_official_html_history_is_parsed(self):
        response = Mock(status_code=200)
        response.text = """
        <div class="headline"><h5><span>2026-07-30</span></h5></div>
        <table><tbody><tr><td>10:48:00</td><td>PRE</td>
        <td>Szállítási adatok regisztrálva, a küldemény még nem érkezett be.</td>
        </tr></tbody></table>
        """
        session = Mock()
        session.get.return_value = response

        result = tracking.expressone_hu_detail(
            "671555557697000013601086", session=session
        )

        self.assertTrue(result["ok"])
        self.assertEqual("in_transit", result["outcome"])
        self.assertEqual("PRE", result["events"][0]["code"])
        self.assertEqual("2026-07-30 10:48:00", result["events"][0]["time"])
        self.assertIn("plc_number=671555557697000013601086", result["official_url"])


class PacketaTrack718Tests(unittest.TestCase):
    def test_detail_forces_packeta_code_for_add_and_query(self):
        add_response = Mock()
        add_response.json.return_value = {"data": {"added": 1}}
        query_response = Mock()
        query_response.json.return_value = {
            "data": {
                "list": [
                    {
                        "trackNum": "Z1465635854",
                        "code": "packeta",
                        "result": 40,
                        "toDetail": [
                            {
                                "date": "2026-07-26T12:00:00",
                                "status": "delivered",
                                "address": "Praha",
                            }
                        ],
                        "fromDetail": [],
                    }
                ]
            }
        }
        session = Mock()
        session.post.side_effect = [add_response, query_response]

        result = tracking.track718_detail(
            "Z1465635854",
            "test-key",
            code=tracking.TRACK718_PACKETA,
            poll=1,
            session=session,
        )

        self.assertTrue(result["ok"])
        self.assertEqual("Packeta".lower(), result["carrier"])
        self.assertEqual("delivered", result["outcome"])
        self.assertEqual(
            {"trackNum": "Z1465635854", "code": "packeta"},
            session.post.call_args_list[0].kwargs["json"][0],
        )
        self.assertEqual(
            {"trackNum": "Z1465635854", "code": "packeta"},
            session.post.call_args_list[1].kwargs["json"][0],
        )

    def test_frontend_uses_packeta_official_url_and_translation_table(self):
        base = (
            Path(__file__).resolve().parents[1] / "templates" / "base.html"
        ).read_text(encoding="utf-8")

        self.assertIn("https://tracking.packeta.com/en/", base)
        self.assertIn("PACKETA_EVENT_PHRASES", base)
        self.assertIn(
            "we are aware of your parcel and are waiting for the sender",
            base,
        )
        self.assertIn("rejected by recipient", base)
        self.assertIn("https://tracking.expressone.hu/?plc_number=", base)
        self.assertIn("EXPRESSONE_HU_EVENT_PHRASES", base)
        self.assertIn("const packetaHu", base)
        self.assertIn("/^\\d{20,}$/", base)


if __name__ == "__main__":
    unittest.main()
