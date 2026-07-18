import unittest
from pathlib import Path

from invoice_system.codex_scan import CodexScanRecognizer, _extract_json_object, _ocr_result_from_response_text, _record_from_json
from invoice_system.config import Settings


JSON_TEXT = '{"invoice_date":"2026-06-12","seller":"CAFE XUAN","total_amount":126,"vat_amount":16,"tips":10}'


class CodexScanTests(unittest.TestCase):
    def test_codex_scan_recognizer_is_removed(self):
        result = CodexScanRecognizer(Settings(openai_api_key="key", codex_scan_enabled=True)).recognize(Path("receipt.jpg"))

        self.assertIsNone(result.parsed_invoice)
        self.assertIn("removed", result.error)

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

    def test_record_uses_raw_date_to_fix_ambiguous_mx_date(self):
        record = _record_from_json(
            '{"invoice_date":"2026-10-05","raw_date":"10/05/2026","currency":"MXN","seller":"Restaurante Mexico","total_amount":126}'
        )

        self.assertEqual(record.invoice_date, "2026-05-10")

    def test_record_keeps_ambiguous_us_date_month_first(self):
        record = _record_from_json(
            '{"invoice_date":"2026-10-05","raw_date":"10/05/2026","currency":"USD","seller":"Walmart","contents":"sales tax","total_amount":12}'
        )

        self.assertEqual(record.invoice_date, "2026-10-05")

    def test_company_profile_overrides_qwen_category_suggestion(self):
        record = _record_from_json(
            '{"invoice_date":"2026-10-05","currency":"USD","seller":"Walmart Supercenter","contents":"groceries","expense_category":"Other","total_amount":12}'
        )

        self.assertEqual(record.expense_category, "Food")

    def test_record_allows_missing_date_to_remain_blank(self):
        record = _record_from_json('{"seller":"CAFE XUAN","total_amount":126}')

        self.assertEqual(record.invoice_date, "")

    def test_unlabeled_handwritten_line_is_not_treated_as_tip(self):
        record = _record_from_json(
            '{"seller":"Nota De Cuenta","currency":"MXN","contents":"2 papas, 2 refrescos",'
            '"total_amount":170,"expense_amount":140,"tips":30}'
        )

        self.assertEqual(record.total_amount, 140)
        self.assertEqual(record.tips, 0)

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
