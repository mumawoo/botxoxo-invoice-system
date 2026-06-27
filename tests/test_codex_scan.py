import unittest

from invoice_system.codex_scan import _extract_json_object, _ocr_result_from_response_text, _record_from_json


JSON_TEXT = '{"invoice_date":"2026-06-12","seller":"CAFE XUAN","total_amount":126,"vat_amount":16,"tips":10}'


class CodexScanTests(unittest.TestCase):
    def test_record_from_direct_json(self):
        record = _record_from_json(JSON_TEXT)
        self.assertEqual(record.invoice_date, "2026-06-12")
        self.assertEqual(record.seller, "CAFE XUAN")
        self.assertEqual(record.total_amount, 126)
        self.assertEqual(record.expense_category, "Food")
        self.assertEqual(record.expense_amount, 110)

    def test_record_normalizes_english_category(self):
        record = _record_from_json(
            '{"invoice_date":"2026-06-12","seller":"Office Depot","expense_category":"office supplies","total_amount":500}'
        )

        self.assertEqual(record.expense_category, "Office supplies")

    def test_record_normalizes_codex_date(self):
        record = _record_from_json('{"invoice_date":"12/06/2026","seller":"CAFE XUAN","total_amount":126}')

        self.assertEqual(record.invoice_date, "2026-06-12")

    def test_record_rejects_missing_required_fields(self):
        with self.assertRaisesRegex(ValueError, "seller"):
            _record_from_json('{"invoice_date":"2026-06-12","total_amount":126}')

    def test_record_rejects_zero_total(self):
        with self.assertRaisesRegex(ValueError, "total_amount"):
            _record_from_json('{"invoice_date":"2026-06-12","seller":"CAFE XUAN","total_amount":0}')

    def test_record_rejects_invalid_date(self):
        with self.assertRaisesRegex(ValueError, "invoice_date"):
            _record_from_json('{"invoice_date":"not a date","seller":"CAFE XUAN","total_amount":126}')

    def test_extracts_fenced_json(self):
        wrapped = f"```json\n{JSON_TEXT}\n```"
        self.assertEqual(_extract_json_object(wrapped), JSON_TEXT)

    def test_extracts_json_from_surrounding_text(self):
        wrapped = f"Here is the result:\n{JSON_TEXT}\nThanks"
        self.assertEqual(_extract_json_object(wrapped), JSON_TEXT)

    def test_raises_for_missing_json(self):
        with self.assertRaises(ValueError):
            _extract_json_object("no json here")

    def test_ocr_result_from_response_text_parses_valid_json(self):
        result = _ocr_result_from_response_text(JSON_TEXT)

        self.assertIsNotNone(result.parsed_invoice)
        self.assertEqual(result.confidence, 1.0)
        self.assertEqual(result.error, "")
        self.assertEqual(result.text, JSON_TEXT)

    def test_ocr_result_from_response_text_preserves_invalid_text_for_audit(self):
        text = "I can read CAFE XUAN, but this is not JSON."

        result = _ocr_result_from_response_text(text)

        self.assertIsNone(result.parsed_invoice)
        self.assertEqual(result.confidence, 0.0)
        self.assertIn("valid JSON object", result.error)
        self.assertEqual(result.text, text)


if __name__ == "__main__":
    unittest.main()
