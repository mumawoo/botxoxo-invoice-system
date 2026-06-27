import unittest
from pathlib import Path

from invoice_system.config import Settings
from invoice_system.dual_ocr import DualOCRResolver
from invoice_system.models import InvoiceRecord, OCRResult


class RecordingRecognizer:
    def __init__(self, engine, calls):
        self.engine = engine
        self.calls = calls

    def recognize(self, image_path):
        self.calls.append(self.engine)
        return OCRResult(
            self.engine,
            parsed_invoice=InvoiceRecord(invoice_date="2026-06-12", seller="Cafe", total_amount=10),
            confidence=0.9,
        )


class RuntimeOrderTests(unittest.TestCase):
    def test_easyocr_runs_before_paddleocr_for_windows_dll_order(self):
        calls = []
        resolver = DualOCRResolver(
            RecordingRecognizer("paddleocr", calls),
            RecordingRecognizer("easyocr", calls),
            RecordingRecognizer("codex_scan", calls),
            Settings(),
        )
        resolver.scan(Path("dummy.jpg"))
        self.assertEqual(calls[:2], ["easyocr", "paddleocr"])


if __name__ == "__main__":
    unittest.main()
