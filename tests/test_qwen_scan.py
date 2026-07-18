import json
import base64
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from invoice_system.config import Settings
from invoice_system.qwen_scan import QWEN_IMAGE_MAX_BYTES, QwenScanRecognizer, _extract_message_content, _qwen_image_url


class QwenScanTests(unittest.TestCase):
    def test_qwen_image_uses_1600px_for_normal_crop(self):
        from PIL import Image

        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "normal.jpg"
            Image.new("RGB", (2400, 1800), "white").save(path)

            encoded = _qwen_image_url(path)
            with Image.open(BytesIO(base64.b64decode(encoded.split(",", 1)[1]))) as image:
                self.assertEqual(max(image.size), 1600)

    def test_qwen_image_uses_2000px_for_long_receipt(self):
        from PIL import Image

        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "long.jpg"
            Image.new("RGB", (1000, 3000), "white").save(path)

            encoded = _qwen_image_url(path)
            with Image.open(BytesIO(base64.b64decode(encoded.split(",", 1)[1]))) as image:
                self.assertEqual(max(image.size), 2000)

    def test_qwen_image_stays_under_900kb(self):
        from PIL import Image

        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "noise.jpg"
            Image.effect_noise((2400, 2400), 100).convert("RGB").save(path, quality=100)

            payload = base64.b64decode(_qwen_image_url(path).split(",", 1)[1])

            self.assertLessEqual(len(payload), QWEN_IMAGE_MAX_BYTES)

    def test_extract_message_content_from_qwen_response(self):
        data = {"choices": [{"message": {"content": "{\"seller\":\"Cafe\"}"}}]}

        self.assertEqual(_extract_message_content(data), "{\"seller\":\"Cafe\"}")

    def test_recognize_posts_image_and_parses_invoice_json(self):
        response = {
            "choices": [
                {
                    "message": {
                        "content": json.dumps(
                            {
                                "invoice_date": "2026-06-12",
                                "raw_date": "12/06/2026",
                                "expense_category": "餐饮",
                                "contents": "Comida",
                                "currency": "MXN",
                                "total_amount": 126.0,
                                "expense_amount": 108.62,
                                "vat_amount": 17.38,
                                "sales_tax": 0,
                                "tips": 0,
                                "seller": "Cafe Xuan",
                                "remarks": "Qwen parsed",
                                "rotate_degrees": 180,
                                "orientation_confidence": 0.97,
                            }
                        )
                    }
                }
            ]
        }
        with tempfile.TemporaryDirectory() as temp:
            image = Path(temp) / "receipt.jpg"
            _write_test_image(image)
            recognizer = QwenScanRecognizer(
                Settings(
                    qwen_api_key="token",
                    qwen_scan_enabled=True,
                    qwen_base_url="https://example.test/chat/completions",
                )
            )

            with patch("invoice_system.qwen_scan.urllib.request.urlopen", return_value=FakeResponse(response)):
                result = recognizer.recognize(image)

        self.assertEqual(result.engine, "qwen_scan")
        self.assertEqual(result.error, "")
        self.assertEqual(result.parsed_invoice.seller, "Cafe Xuan")
        self.assertEqual(result.parsed_invoice.total_amount, 126.0)
        self.assertTrue(result.parsed_invoice.report_components)
        self.assertEqual(result.rotate_degrees, 180)
        self.assertEqual(result.orientation_confidence, 0.97)

    def test_qwen_recognize_uses_raw_date_for_ambiguous_mx_receipt(self):
        response = {
            "choices": [
                {
                    "message": {
                        "content": json.dumps(
                            {
                                "invoice_date": "2026-10-05",
                                "raw_date": "10/05/2026",
                                "currency": "MXN",
                                "total_amount": 126,
                                "seller": "Restaurante Mexico",
                            }
                        )
                    }
                }
            ]
        }
        with tempfile.TemporaryDirectory() as temp:
            image = Path(temp) / "receipt.jpg"
            _write_test_image(image)
            recognizer = QwenScanRecognizer(
                Settings(
                    qwen_api_key="token",
                    qwen_scan_enabled=True,
                    qwen_base_url="https://example.test/chat/completions",
                )
            )

            with patch("invoice_system.qwen_scan.urllib.request.urlopen", return_value=FakeResponse(response)):
                result = recognizer.recognize(image)

        self.assertEqual(result.error, "")
        self.assertEqual(result.parsed_invoice.invoice_date, "2026-05-10")

    def test_qwen_default_remark_does_not_claim_codex_was_used(self):
        response = {
            "choices": [
                {
                    "message": {
                        "content": json.dumps(
                            {
                                "invoice_date": "2026-06-12",
                                "currency": "MXN",
                                "total_amount": 126,
                                "seller": "Cafe Xuan",
                            }
                        )
                    }
                }
            ]
        }
        with tempfile.TemporaryDirectory() as temp:
            image = Path(temp) / "receipt.jpg"
            _write_test_image(image)
            recognizer = QwenScanRecognizer(
                Settings(
                    qwen_api_key="token",
                    qwen_scan_enabled=True,
                    qwen_base_url="https://example.test/chat/completions",
                )
            )

            with patch("invoice_system.qwen_scan.urllib.request.urlopen", return_value=FakeResponse(response)):
                result = recognizer.recognize(image)

        self.assertEqual(result.parsed_invoice.remarks, "Qwen Scan used")
        self.assertNotIn("Codex", result.parsed_invoice.remarks)


class FakeResponse:
    def __init__(self, payload: dict) -> None:
        self.payload = payload

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def read(self) -> bytes:
        return json.dumps(self.payload).encode("utf-8")


def _write_test_image(path: Path) -> None:
    from PIL import Image

    image = Image.new("RGB", (80, 80), "white")
    image.save(path)


if __name__ == "__main__":
    unittest.main()
