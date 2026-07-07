import unittest

from invoice_system.models import OCRTextLine
from invoice_system.parsing import extract_amounts, normalize_date, normalize_receipt_date, parse_amount, parse_invoice_from_lines
from invoice_system.recognizers import _paddle_lines


class ParsingTests(unittest.TestCase):
    def test_normalize_date(self):
        self.assertEqual(normalize_date("Fecha 12/06/2026"), "2026-06-12")

    def test_normalize_receipt_date_uses_mx_day_first_for_ambiguous_dates(self):
        self.assertEqual(
            normalize_receipt_date("2026-10-05", raw_date="10/05/2026", currency="MXN", context="Fecha IVA Mexico"),
            "2026-05-10",
        )

    def test_normalize_receipt_date_uses_us_month_first_when_context_is_clearly_us(self):
        self.assertEqual(
            normalize_receipt_date("10/05/2026", currency="USD", context="Walmart sales tax USA"),
            "2026-10-05",
        )

    def test_parse_amount_supports_common_mx_formats(self):
        self.assertEqual(parse_amount("$1,234.50 MXN"), 1234.5)
        self.assertEqual(parse_amount("M.N. 1.234,50"), 1234.5)
        self.assertEqual(parse_amount("1,234"), 1234.0)
        self.assertEqual(parse_amount("1.234"), 1234.0)
        self.assertEqual(parse_amount("TOTAL M.N. 1.260,00"), 1260.0)
        self.assertEqual(parse_amount("Importe: $1,260.00 MXN"), 1260.0)

    def test_extract_amounts_supports_thousands_and_decimal_variants(self):
        self.assertEqual(extract_amounts("Subtotal 1,000 IVA 160.00 TOTAL 1,160.00"), [1000.0, 160.0, 1160.0])
        self.assertEqual(extract_amounts("Subtotal 1.000 IVA 160,00 TOTAL 1.160,00"), [1000.0, 160.0, 1160.0])

    def test_parse_invoice_fields(self):
        lines = [
            OCRTextLine("Cafe Xuan", 0.9),
            OCRTextLine("Fecha 2026-06-12", 0.9),
            OCRTextLine("IVA 16.00", 0.9),
            OCRTextLine("TOTAL $116.00", 0.9),
        ]
        record = parse_invoice_from_lines(lines, "test")
        self.assertEqual(record.seller, "Cafe Xuan")
        self.assertEqual(record.invoice_date, "2026-06-12")
        self.assertEqual(record.total_amount, 116.0)
        self.assertEqual(record.vat_amount, 16.0)

    def test_parse_invoice_fields_with_mx_thousands_format(self):
        lines = [
            OCRTextLine("Restaurante Centro", 0.9),
            OCRTextLine("Fecha 12/06/2026", 0.9),
            OCRTextLine("IVA 160,00", 0.9),
            OCRTextLine("Propina 100,00", 0.9),
            OCRTextLine("TOTAL M.N. 1.260,00", 0.9),
        ]

        record = parse_invoice_from_lines(lines, "test")

        self.assertEqual(record.total_amount, 1260.0)
        self.assertEqual(record.vat_amount, 160.0)
        self.assertEqual(record.tips, 100.0)

    def test_missing_total_does_not_use_date_year_as_amount(self):
        lines = [
            OCRTextLine("Cafe Xuan", 0.9),
            OCRTextLine("Fecha 2026-06-12", 0.9),
            OCRTextLine("Mesa cuatro", 0.9),
        ]

        record = parse_invoice_from_lines(lines, "test")

        self.assertEqual(record.total_amount, 0.0)
        self.assertIn("incomplete", record.remarks)

    def test_missing_total_falls_back_to_non_date_amounts(self):
        lines = [
            OCRTextLine("Cafe Xuan", 0.9),
            OCRTextLine("Fecha 2026-06-12", 0.9),
            OCRTextLine("Consumo alimentos 180.00", 0.9),
            OCRTextLine("IVA 28.80", 0.9),
        ]

        record = parse_invoice_from_lines(lines, "test")

        self.assertEqual(record.total_amount, 180.0)

    def test_missing_total_ignores_metadata_tax_and_tip_amounts(self):
        lines = [
            OCRTextLine("Cafe Xuan", 0.9),
            OCRTextLine("RFC XUA260612AB1", 0.9),
            OCRTextLine("Folio 987654", 0.9),
            OCRTextLine("Ticket 12345", 0.9),
            OCRTextLine("Mesa 4", 0.9),
            OCRTextLine("IVA 28.80", 0.9),
            OCRTextLine("Propina 20.00", 0.9),
        ]

        record = parse_invoice_from_lines(lines, "test")

        self.assertEqual(record.total_amount, 0.0)
        self.assertIn("incomplete", record.remarks)

    def test_missing_total_keeps_real_purchase_amount_over_metadata(self):
        lines = [
            OCRTextLine("Cafe Xuan", 0.9),
            OCRTextLine("RFC XUA260612AB1", 0.9),
            OCRTextLine("Ticket 12345", 0.9),
            OCRTextLine("Consumo alimentos 180.00", 0.9),
            OCRTextLine("IVA 28.80", 0.9),
        ]

        record = parse_invoice_from_lines(lines, "test")

        self.assertEqual(record.total_amount, 180.0)

    def test_parse_sales_tax_and_expense_amount(self):
        lines = [
            OCRTextLine("Hotel Centro", 0.9),
            OCRTextLine("Fecha 2026-06-12", 0.9),
            OCRTextLine("IVA 16.00", 0.9),
            OCRTextLine("ISH 4.00", 0.9),
            OCRTextLine("TOTAL $120.00", 0.9),
        ]

        record = parse_invoice_from_lines(lines, "test")

        self.assertEqual(record.vat_amount, 16.0)
        self.assertEqual(record.sales_tax, 4.0)
        self.assertEqual(record.expense_amount, 100.0)

    def test_sales_tax_keyword_does_not_become_seller(self):
        lines = [
            OCRTextLine("Sales Tax 4.00", 0.9),
            OCRTextLine("Cafe Xuan", 0.9),
            OCRTextLine("Fecha 2026-06-12", 0.9),
            OCRTextLine("TOTAL $104.00", 0.9),
        ]

        record = parse_invoice_from_lines(lines, "test")

        self.assertEqual(record.seller, "Cafe Xuan")
        self.assertEqual(record.sales_tax, 4.0)

    def test_parse_paddle_v3_result_shape(self):
        lines = _paddle_lines([{"rec_texts": ["CAFE XUAN", "TOTAL $126.00"], "rec_scores": [0.9, 0.8]}])
        self.assertEqual([line.text for line in lines], ["CAFE XUAN", "TOTAL $126.00"])
        self.assertEqual(lines[1].confidence, 0.8)


if __name__ == "__main__":
    unittest.main()
