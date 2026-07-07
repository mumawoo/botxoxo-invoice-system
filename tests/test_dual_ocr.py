import tempfile
import unittest
from pathlib import Path

from invoice_system.config import Settings
from invoice_system.dual_ocr import DualOCRResolver
from invoice_system.models import InvoiceRecord, OCRResult, OCRTextLine
from invoice_system.recognizers import StaticRecognizer


def result(
    engine,
    date="2026-06-12",
    seller="Cafe Xuan",
    amount=120.0,
    confidence=0.9,
    currency="MXN",
    vat=0.0,
    sales_tax=0.0,
    tips=0.0,
):
    record = InvoiceRecord(
        invoice_date=date,
        seller=seller,
        currency=currency,
        total_amount=amount,
        expense_amount=amount,
        vat_amount=vat,
        sales_tax=sales_tax,
        tips=tips,
    )
    return OCRResult(engine, [OCRTextLine("sample", confidence)], record, confidence)


class DualOCRTests(unittest.TestCase):
    def test_runs_both_local_ocr_engines_before_accepting_agreement(self):
        calls = []
        resolver = DualOCRResolver(
            RecordingRecognizer(result("paddleocr"), calls),
            RecordingRecognizer(result("easyocr"), calls),
            RecordingRecognizer(result("codex_scan", amount=999), calls),
            Settings(),
        )

        scan = resolver.scan(Path("dummy.jpg"))

        self.assertFalse(scan.used_codex)
        self.assertEqual(calls, ["easyocr", "paddleocr"])

    def test_codex_fallback_is_removed_even_when_configured(self):
        calls = []
        resolver = DualOCRResolver(
            RecordingRecognizer(result("paddleocr", amount=120), calls),
            RecordingRecognizer(result("easyocr", amount=130), calls),
            RecordingRecognizer(result("codex_scan", amount=125), calls),
            _codex_enabled_settings(),
        )

        scan = resolver.scan(Path("dummy.jpg"))

        self.assertFalse(scan.used_codex)
        self.assertEqual(calls, ["easyocr", "paddleocr"])
        self.assertEqual(scan.record.total_amount, 120)
        self.assertIn("Codex Scan disabled", scan.record.remarks)

    def test_uses_best_local_result_when_codex_scan_is_disabled(self):
        calls = []
        resolver = DualOCRResolver(
            RecordingRecognizer(result("paddleocr", amount=120, confidence=0.8), calls),
            RecordingRecognizer(result("easyocr", amount=130, confidence=0.9), calls),
            RecordingRecognizer(result("codex_scan", amount=125), calls),
            Settings(codex_scan_enabled=False),
        )

        scan = resolver.scan(Path("dummy.jpg"))

        self.assertFalse(scan.used_codex)
        self.assertEqual(scan.record.total_amount, 130)
        self.assertEqual(calls, ["easyocr", "paddleocr"])
        self.assertIn("Codex Scan disabled", scan.record.remarks)

    def test_uses_qwen_fallback_when_qwen_scan_is_enabled(self):
        calls = []
        resolver = DualOCRResolver(
            RecordingRecognizer(result("paddleocr", amount=120), calls),
            RecordingRecognizer(result("easyocr", amount=130), calls),
            RecordingRecognizer(result("qwen_scan", amount=125), calls),
            Settings(qwen_scan_enabled=True),
        )

        scan = resolver.scan(Path("dummy.jpg"))

        self.assertTrue(scan.used_codex)
        self.assertEqual(calls, ["easyocr", "paddleocr", "qwen_scan"])
        self.assertEqual(scan.record.total_amount, 125)
        self.assertIn("Qwen Scan used", scan.record.remarks)

    def test_accepts_matching_local_ocr(self):
        resolver = DualOCRResolver(
            StaticRecognizer(result("paddleocr")),
            StaticRecognizer(result("easyocr")),
            StaticRecognizer(result("codex_scan", amount=999)),
            Settings(),
        )
        scan = resolver.scan(Path("dummy.jpg"))
        self.assertFalse(scan.used_codex)
        self.assertEqual(scan.record.remarks, "PaddleOCR agreed with EasyOCR")

    def test_accepts_local_ocr_with_normalized_date_currency_and_minor_tax_difference(self):
        resolver = DualOCRResolver(
            StaticRecognizer(result("paddleocr", date="2026-06-12", currency="MXN", vat=16.0, tips=10.0)),
            StaticRecognizer(result("easyocr", date="12/06/2026", currency="M.N.", vat=16.49, tips=10.5)),
            StaticRecognizer(result("codex_scan", amount=999)),
            Settings(),
        )

        scan = resolver.scan(Path("dummy.jpg"))

        self.assertFalse(scan.used_codex)

    def test_does_not_use_codex_when_amounts_disagree(self):
        resolver = DualOCRResolver(
            StaticRecognizer(result("paddleocr", amount=120)),
            StaticRecognizer(result("easyocr", amount=130)),
            StaticRecognizer(result("codex_scan", amount=125)),
            _codex_enabled_settings(),
        )
        scan = resolver.scan(Path("dummy.jpg"))
        self.assertFalse(scan.used_codex)
        self.assertEqual(scan.record.total_amount, 120)
        self.assertIn("Codex Scan disabled", scan.record.remarks)

    def test_does_not_use_codex_when_currency_disagrees(self):
        resolver = DualOCRResolver(
            StaticRecognizer(result("paddleocr", currency="MXN")),
            StaticRecognizer(result("easyocr", currency="USD")),
            StaticRecognizer(result("codex_scan", amount=125)),
            _codex_enabled_settings(),
        )

        scan = resolver.scan(Path("dummy.jpg"))

        self.assertFalse(scan.used_codex)
        self.assertEqual(scan.reason, "OCR mismatch")

    def test_does_not_use_codex_when_tax_or_tips_disagree(self):
        resolver = DualOCRResolver(
            StaticRecognizer(result("paddleocr", vat=16.0, tips=10.0)),
            StaticRecognizer(result("easyocr", vat=20.0, tips=10.0)),
            StaticRecognizer(result("codex_scan", amount=125)),
            _codex_enabled_settings(),
        )

        scan = resolver.scan(Path("dummy.jpg"))

        self.assertFalse(scan.used_codex)
        self.assertEqual(scan.reason, "OCR mismatch")

    def test_does_not_use_codex_when_confidence_low(self):
        resolver = DualOCRResolver(
            StaticRecognizer(result("paddleocr", confidence=0.3)),
            StaticRecognizer(result("easyocr", confidence=0.9)),
            StaticRecognizer(result("codex_scan", amount=125)),
            _codex_enabled_settings(),
        )
        scan = resolver.scan(Path("dummy.jpg"))
        self.assertFalse(scan.used_codex)

    def test_does_not_call_codex_even_if_injected_result_would_be_invalid(self):
        codex_text = "CAFE XUAN total 126, but not JSON"
        resolver = DualOCRResolver(
            StaticRecognizer(result("paddleocr", amount=120, confidence=0.8)),
            StaticRecognizer(result("easyocr", amount=130, confidence=0.9)),
            StaticRecognizer(
                OCRResult(
                    "codex_scan",
                    [OCRTextLine(codex_text, 1.0)],
                    error="Codex Scan JSON failed validation: seller",
                )
            ),
            _codex_enabled_settings(),
        )

        scan = resolver.scan(Path("dummy.jpg"))

        self.assertFalse(scan.used_codex)
        self.assertEqual(scan.record.total_amount, 130)
        self.assertIn("Codex Scan disabled", scan.record.remarks)
        self.assertIsNone(scan.codex)

    def test_does_not_use_codex_when_image_quality_is_poor_even_if_local_ocr_agrees(self):
        with tempfile.TemporaryDirectory() as temp:
            image_path = Path(temp) / "flat.jpg"
            _write_flat_image(image_path)
            resolver = DualOCRResolver(
                StaticRecognizer(result("paddleocr")),
                StaticRecognizer(result("easyocr")),
                StaticRecognizer(result("codex_scan", amount=125)),
                _codex_enabled_settings(),
            )

            scan = resolver.scan(image_path)

            self.assertFalse(scan.used_codex)
            self.assertEqual(scan.reason, "poor image quality")
            self.assertEqual(scan.record.total_amount, 120)


def _write_flat_image(path: Path) -> None:
    import cv2
    import numpy as np

    image = np.full((120, 120), 128, dtype=np.uint8)
    ok, encoded = cv2.imencode(".jpg", image)
    if not ok:
        raise RuntimeError("could not encode test image")
    encoded.tofile(str(path))


def _codex_enabled_settings() -> Settings:
    return Settings(codex_scan_enabled=True)


class RecordingRecognizer:
    def __init__(self, ocr_result: OCRResult, calls: list[str]) -> None:
        self.result = ocr_result
        self.engine = ocr_result.engine
        self.calls = calls

    def recognize(self, image_path: Path) -> OCRResult:
        self.calls.append(self.engine)
        return self.result


if __name__ == "__main__":
    unittest.main()
