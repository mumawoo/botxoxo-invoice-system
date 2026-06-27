import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from invoice_system.config import Settings
from invoice_system.qwen_scan import QwenScanRecognizer, _extract_message_content


class QwenScanTests(unittest.TestCase):
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
        self.assertEqual(result.rotate_degrees, 180)
        self.assertEqual(result.orientation_confidence, 0.97)


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
