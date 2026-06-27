import tempfile
import unittest
import json
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from invoice_system.fx_rates import ExchangeRate
from invoice_system.models import InvoiceRecord
from invoice_system.reimbursement_excel import (
    CHECKED_WORKBOOK_NAME,
    FOOD_EXP_SHEET,
    INVOICE_EXP_SHEET,
    OTHER_EXP_SHEET,
    REIMBURSEMENT_WORKBOOK_NAME,
    ReimbursementWorkbook,
    assign_available_line_numbers,
    build_checked_outputs,
    change_reimbursement_record,
    clear_generated_crops,
    corrected_crop_names,
    focus_reimbursement_workbook,
    load_reimbursement_records,
)


class ReimbursementExcelTests(unittest.TestCase):
    def test_reimbursement_workbook_name_is_real_chinese(self):
        self.assertEqual(REIMBURSEMENT_WORKBOOK_NAME, "报销明细_2026_xlsx.xlsx")

    def test_usd_receipt_writes_chinese_type_english_accounting_and_original_currency_fields(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / REIMBURSEMENT_WORKBOOK_NAME
            store = ReimbursementWorkbook(
                path,
                fetch_rates=lambda: [ExchangeRate(date(2026, 6, 12), usd_cny_per_100=700, mxn_per_100_cny=250)],
            )
            record = InvoiceRecord(
                line_no=1,
                invoice_date="2026-06-12",
                expense_category="Food",
                currency="USD",
                total_amount=10,
                seller="Cafe",
            )

            store.write_records([record])

            wb = load_workbook(path, data_only=True)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                self.assertEqual(ws.cell(2, 4).value, 175)
                self.assertEqual(ws.cell(2, 5).value, "餐饮")
                self.assertEqual(ws.cell(2, 6).value, "USD")
                self.assertEqual(ws.cell(2, 7).value, 10)
                self.assertEqual(ws.cell(2, 8).value, 17.5)
                self.assertEqual(ws.cell(2, 11).value, "Food")
                rates = wb["exchange rate"]
                self.assertEqual(rates.cell(1, 1).value, "日期")
                self.assertEqual(rates.cell(1, 2).value, "USD")
                self.assertEqual(rates.cell(1, 25).value, "MXN")
                self.assertEqual(rates.cell(2, 2).value, 700)
                self.assertEqual(rates.cell(2, 25).value, 250)
            finally:
                wb.close()

    def test_reimbursement_workbook_keeps_manual_rows_in_one_sheet(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / REIMBURSEMENT_WORKBOOK_NAME
            store = ReimbursementWorkbook(path)

            store.write_records(
                [
                    InvoiceRecord(line_no=1, invoice_date="2026-06-12", expense_category="Food", total_amount=100, seller="Cafe"),
                    InvoiceRecord(line_no=2, invoice_date="2026-06-13", expense_category="Gas", total_amount=200, seller="Pemex"),
                ]
            )

            wb = load_workbook(path, data_only=True)
            try:
                self.assertEqual(wb[INVOICE_EXP_SHEET].cell(2, 9).value, "Cafe")
                self.assertEqual(wb[INVOICE_EXP_SHEET].cell(3, 9).value, "Pemex")
                self.assertNotIn(FOOD_EXP_SHEET, wb.sheetnames)
                self.assertNotIn(OTHER_EXP_SHEET, wb.sheetnames)
            finally:
                wb.close()

    def test_usd_receipt_before_rate_table_uses_nearest_available_rate(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / REIMBURSEMENT_WORKBOOK_NAME
            store = ReimbursementWorkbook(
                path,
                fetch_rates=lambda: [ExchangeRate(date(2026, 1, 2), usd_cny_per_100=700, mxn_per_100_cny=250)],
            )

            store.write_records(
                [
                    InvoiceRecord(
                        line_no=1,
                        invoice_date="2025-12-11",
                        expense_category="Other",
                        currency="USD",
                        total_amount=40,
                        seller="Target",
                    )
                ]
            )

            wb = load_workbook(path, data_only=True)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                self.assertEqual(ws.cell(2, 4).value, 700)
                self.assertEqual(ws.cell(2, 6).value, "USD")
                self.assertEqual(ws.cell(2, 7).value, 40)
                self.assertEqual(ws.cell(2, 8).value, 17.5)
            finally:
                wb.close()

    def test_eur_receipt_uses_safe_currency_table_when_available(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / REIMBURSEMENT_WORKBOOK_NAME
            store = ReimbursementWorkbook(
                path,
                fetch_rates=lambda: [
                    ExchangeRate(
                        date(2026, 6, 12),
                        usd_cny_per_100=700,
                        mxn_per_100_cny=250,
                        rates={"USD": 700, "EUR": 800, "MXN": 250},
                    )
                ],
            )

            store.write_records(
                [
                    InvoiceRecord(
                        line_no=1,
                        invoice_date="2026-06-12",
                        expense_category="Other",
                        currency="EUR",
                        total_amount=10,
                        seller="Hotel",
                    )
                ]
            )

            wb = load_workbook(path, data_only=True)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                self.assertEqual(ws.cell(2, 4).value, 200)
                self.assertEqual(ws.cell(2, 6).value, "EUR")
                self.assertEqual(ws.cell(2, 8).value, 20)
            finally:
                wb.close()

    def test_corrected_duplicate_record_is_not_written_again(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / REIMBURSEMENT_WORKBOOK_NAME
            store = ReimbursementWorkbook(path)
            record = InvoiceRecord(
                line_no=1,
                invoice_date="2026-06-12",
                expense_category="Food",
                currency="USD",
                total_amount=10,
                seller="Cafe",
            )
            store.write_records([record])
            wb = load_workbook(path)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                ws.cell(2, 12).value = "corrected"
                wb.save(path)
            finally:
                wb.close()

            store.write_records([record])

            wb = load_workbook(path, data_only=True)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                rows = [
                    [ws.cell(row, col).value for col in range(1, 12)]
                    for row in range(2, ws.max_row + 1)
                    if any(ws.cell(row, col).value not in (None, "") for col in range(1, 12))
                ]
                self.assertEqual(len(rows), 1)
                self.assertEqual(rows[0][0], 1)
            finally:
                wb.close()

    def test_deleted_row_is_not_loaded_but_does_not_block_reupload(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / REIMBURSEMENT_WORKBOOK_NAME
            store = ReimbursementWorkbook(path)
            store.write_records([InvoiceRecord(line_no=2, invoice_date="2026-06-12", expense_category="Food", total_amount=100, seller="Cafe")])
            wb = load_workbook(path)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                ws.cell(2, 12).value = "deleted"
                wb.save(path)
            finally:
                wb.close()

            self.assertEqual([], load_reimbursement_records(path))

            store.write_records([InvoiceRecord(line_no=3, invoice_date="2026-06-12", expense_category="Food", total_amount=100, seller="Cafe")])

            wb = load_workbook(path, data_only=True)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                rows = [
                    [ws.cell(row, col).value for col in range(1, 13)]
                    for row in range(2, ws.max_row + 1)
                    if any(ws.cell(row, col).value not in (None, "") for col in range(1, 13))
                ]
                self.assertEqual(len(rows), 2)
                self.assertEqual(rows[0][11], "deleted")
                self.assertEqual(rows[1][8], "Cafe")
                self.assertIsNone(rows[1][11])
            finally:
                wb.close()

    def test_ok_and_correct_markers_lock_rows_and_delete_takes_precedence(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / REIMBURSEMENT_WORKBOOK_NAME
            store = ReimbursementWorkbook(path)
            records = [
                InvoiceRecord(line_no=1, invoice_date="2026-06-12", expense_category="Food", total_amount=100, seller="Cafe"),
                InvoiceRecord(line_no=2, invoice_date="2026-06-13", expense_category="Gas", total_amount=200, seller="Pemex"),
            ]
            store.write_records(records)
            wb = load_workbook(path)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                ws.cell(2, 12).value = "ok"
                ws.cell(3, 12).value = "ok deleted"
                wb.save(path)
            finally:
                wb.close()

            self.assertEqual(store.locked_numbers(), {1, 2})
            loaded = load_reimbursement_records(path)
            self.assertEqual(len(loaded), 1)
            self.assertEqual(loaded[0].seller, "Cafe")

    def test_build_checked_outputs_filters_deleted_and_reindexes_crops(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / REIMBURSEMENT_WORKBOOK_NAME
            review = root / "review_crops"
            review.mkdir()
            crop1 = review / "001_2026-06-12_MXN_100.00_Cafe.jpg"
            crop2 = review / "002_2026-06-13_MXN_200.00_Pemex.jpg"
            crop3 = review / "003_2026-06-14_MXN_50.00_Store.jpg"
            for crop in (crop1, crop2, crop3):
                crop.write_bytes(b"jpg")
            ReimbursementWorkbook(path).write_records(
                [
                    InvoiceRecord(line_no=1, invoice_date="2026-06-12", expense_category="Food", total_amount=100, seller="Cafe", crop_image=str(crop1)),
                    InvoiceRecord(line_no=2, invoice_date="2026-06-13", expense_category="Gas", total_amount=200, seller="Pemex", crop_image=str(crop2)),
                    InvoiceRecord(line_no=3, invoice_date="2026-06-14", expense_category="Other", total_amount=50, seller="Store", crop_image=str(crop3)),
                ]
            )
            wb = load_workbook(path)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                ws.cell(3, 12).value = "deleted"
                wb.save(path)
            finally:
                wb.close()

            result = build_checked_outputs(root)

            self.assertEqual(result.records_written, 2)
            self.assertEqual(result.crops_written, 2)
            self.assertEqual(result.workbook_path.name, CHECKED_WORKBOOK_NAME)
            final_crops = sorted((root / "final_crops").glob("*.jpg"))
            self.assertEqual([path.name for path in final_crops], ["001_2026-06-12_MXN_100.00_Cafe.jpg", "002_2026-06-14_MXN_50.00_Store.jpg"])
            checked = load_workbook(root / CHECKED_WORKBOOK_NAME, data_only=True)
            try:
                self.assertIn(FOOD_EXP_SHEET, checked.sheetnames)
                self.assertIn(OTHER_EXP_SHEET, checked.sheetnames)
                self.assertEqual(checked[FOOD_EXP_SHEET].cell(2, 8).value, "Cafe")
                self.assertEqual(checked[OTHER_EXP_SHEET].cell(2, 8).value, "Store")
                ws = checked[INVOICE_EXP_SHEET]
                self.assertEqual(ws.max_row, 3)
                self.assertEqual(ws.cell(1, 2).value, "Date")
                self.assertEqual(ws.cell(1, 11).value, "Invoice link")
                self.assertEqual(ws.cell(2, 1).value, 1)
                self.assertEqual(ws.cell(3, 1).value, 2)
                self.assertEqual(ws.cell(3, 8).value, "Store")
                self.assertEqual(ws.cell(3, 11).value, "final_crops/002_2026-06-14_MXN_50.00_Store.jpg")
                self.assertEqual(ws.cell(3, 11).hyperlink.target, "final_crops/002_2026-06-14_MXN_50.00_Store.jpg")
            finally:
                checked.close()

    def test_build_checked_outputs_keeps_combined_supporting_crops_as_ab_files(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / REIMBURSEMENT_WORKBOOK_NAME
            review = root / "review_crops"
            review.mkdir()
            primary = review / "040_2026-06-12_MXN_126.00_Cafe.jpg"
            supporting = review / "041_2026-06-12_MXN_126.00_Cafe_card.jpg"
            primary.write_bytes(b"invoice")
            supporting.write_bytes(b"payment")
            ReimbursementWorkbook(path).write_records(
                [
                    InvoiceRecord(
                        line_no=1,
                        invoice_date="2026-06-12",
                        expense_category="Food",
                        total_amount=126,
                        seller="Cafe",
                        crop_image=str(primary),
                        supporting_crop_images=[str(supporting)],
                    )
                ]
            )

            result = build_checked_outputs(root)

            self.assertEqual(result.records_written, 1)
            self.assertEqual(result.crops_written, 2)
            final_crops = sorted((root / "final_crops").glob("*.jpg"))
            self.assertEqual(
                [path.name for path in final_crops],
                [
                    "001a_2026-06-12_MXN_126.00_Cafe.jpg",
                    "001b_2026-06-12_MXN_126.00_Cafe.jpg",
                ],
            )
            checked = load_workbook(root / CHECKED_WORKBOOK_NAME, data_only=True)
            try:
                ws = checked[INVOICE_EXP_SHEET]
                self.assertEqual(ws.max_row, 2)
                self.assertEqual(ws.cell(2, 11).value, "final_crops/001a_2026-06-12_MXN_126.00_Cafe.jpg")
            finally:
                checked.close()

    def test_change_reimbursement_record_updates_category_and_locks_by_crop_id(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / REIMBURSEMENT_WORKBOOK_NAME
            review = root / "review_crops"
            review.mkdir()
            crop = review / "021_2026-06-12_MXN_100.00_Cafe.jpg"
            crop.write_bytes(b"jpg")
            ReimbursementWorkbook(path).write_records(
                [InvoiceRecord(line_no=1, invoice_date="2026-06-12", expense_category="Food", total_amount=100, seller="Cafe", crop_image=str(crop))]
            )

            result = change_reimbursement_record(root, "021", category="Other")

            self.assertEqual(result.crop_id, "021")
            self.assertEqual(result.status, "ok")
            wb = load_workbook(path, data_only=True)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                columns = {ws.cell(1, col).value: col for col in range(1, ws.max_column + 1)}
                self.assertEqual(ws.cell(2, columns["Accounting Category"]).value, "Other")
                self.assertEqual(ws.cell(2, columns["Manual status"]).value, "ok")
            finally:
                wb.close()

    def test_change_reimbursement_record_updates_usd_amount_and_delete_excludes_checked(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / REIMBURSEMENT_WORKBOOK_NAME
            review = root / "review_crops"
            review.mkdir()
            crop = review / "021_2026-06-12_USD_10.00_Cafe.jpg"
            crop.write_bytes(b"jpg")
            store = ReimbursementWorkbook(
                path,
                fetch_rates=lambda: [ExchangeRate(date(2026, 6, 12), usd_cny_per_100=700, mxn_per_100_cny=250)],
            )
            store.write_records(
                [InvoiceRecord(line_no=1, invoice_date="2026-06-12", expense_category="Food", currency="USD", total_amount=10, seller="Cafe", crop_image=str(crop))]
            )

            change_reimbursement_record(root, "021", amount=20, currency="USD", status="correct")
            wb = load_workbook(path, data_only=True)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                columns = {ws.cell(1, col).value: col for col in range(1, ws.max_column + 1)}
                self.assertEqual(ws.cell(2, columns["原金额"]).value, 20)
                self.assertEqual(ws.cell(2, columns["MXN Amount"]).value, 350)
                self.assertEqual(ws.cell(2, columns["Manual status"]).value, "correct")
            finally:
                wb.close()

            change_reimbursement_record(root, "021", status="delete")
            checked = build_checked_outputs(root)

            self.assertEqual(checked.records_written, 0)
            self.assertEqual(checked.crops_written, 0)

    def test_deleted_row_does_not_block_same_invoice_reupload(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / REIMBURSEMENT_WORKBOOK_NAME
            review = root / "review_crops"
            review.mkdir()
            old_crop = review / "038_2026-06-12_USD_11.00_Store.jpg"
            new_crop = review / "041_2026-06-12_USD_11.00_Store.jpg"
            old_crop.write_bytes(b"old")
            new_crop.write_bytes(b"new")
            store = ReimbursementWorkbook(
                path,
                fetch_rates=lambda: [ExchangeRate(date(2026, 6, 12), usd_cny_per_100=700, mxn_per_100_cny=250)],
            )
            store.write_records(
                [InvoiceRecord(line_no=38, invoice_date="2026-06-12", expense_category="Other", currency="USD", total_amount=11, seller="Store", crop_image=str(old_crop))]
            )
            change_reimbursement_record(root, "038", status="delete")

            store.write_records(
                [
                    InvoiceRecord(line_no=38, invoice_date="2026-06-12", expense_category="Other", currency="USD", total_amount=11, seller="Store", crop_image=str(old_crop)),
                    InvoiceRecord(line_no=41, invoice_date="2026-06-12", expense_category="Other", currency="USD", total_amount=11, seller="Store", crop_image=str(new_crop)),
                ]
            )

            wb = load_workbook(path, data_only=True)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                rows = [
                    (ws.cell(row, 2).value, ws.cell(row, 12).value)
                    for row in range(2, ws.max_row + 1)
                    if ws.cell(row, 2).value
                ]
                self.assertEqual(rows, [("review_crops/038_2026-06-12_USD_11.00_Store.jpg", "delete"), ("review_crops/041_2026-06-12_USD_11.00_Store.jpg", None)])
            finally:
                wb.close()

    def test_build_checked_outputs_migrates_legacy_final_crops_to_review(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / REIMBURSEMENT_WORKBOOK_NAME
            legacy_final = root / "final_crops" / "001_2026-06-12_MXN_100.00_Cafe.jpg"
            legacy_final.parent.mkdir()
            legacy_final.write_bytes(b"jpg")
            ReimbursementWorkbook(path).write_records(
                [
                    InvoiceRecord(
                        line_no=1,
                        invoice_date="2026-06-12",
                        expense_category="Food",
                        total_amount=100,
                        seller="Cafe",
                        crop_image=str(legacy_final),
                    )
                ]
            )

            result = build_checked_outputs(root)

            self.assertEqual(result.records_written, 1)
            self.assertTrue((root / "review_crops" / legacy_final.name).exists())
            self.assertTrue((root / "final_crops" / legacy_final.name).exists())
            wb = load_workbook(path, data_only=True)
            try:
                self.assertEqual(wb[INVOICE_EXP_SHEET].cell(2, 2).value, f"review_crops/{legacy_final.name}")
            finally:
                wb.close()

    def test_build_checked_outputs_reconciles_manual_link_to_original_processing_crop_id(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / REIMBURSEMENT_WORKBOOK_NAME
            old_review = root / "review_crops" / "018_2025-12-03_USD_38.33_ALICIA.jpg"
            raw_crop = root / "crops" / "021_2025-12-03_USD_38.33_ALICIA.jpg"
            old_review.parent.mkdir()
            raw_crop.parent.mkdir()
            old_review.write_bytes(b"old")
            raw_crop.write_bytes(b"raw")
            store = ReimbursementWorkbook(
                path,
                fetch_rates=lambda: [ExchangeRate(date(2025, 12, 3), usd_cny_per_100=700, mxn_per_100_cny=250)],
            )
            store.write_records(
                [
                    InvoiceRecord(
                        line_no=18,
                        invoice_date="2025-12-03",
                        expense_category="Food",
                        currency="USD",
                        total_amount=38.33,
                        seller="ALICIA",
                        crop_image=str(old_review),
                    )
                ]
            )
            (root / "processing_state.json").write_text(
                json.dumps(
                    {
                        "records": [
                            {
                                "invoice_date": "2025-12-03",
                                "seller": "ALICIA",
                                "currency": "USD",
                                "total_amount": 38.33,
                                "crop_image": str(raw_crop),
                            }
                        ]
                    }
                ),
                encoding="utf-8",
            )

            build_checked_outputs(root)

            wb = load_workbook(path, data_only=False)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                columns = {ws.cell(1, col).value: col for col in range(1, ws.max_column + 1)}
                link = ws.cell(2, columns["Invoice link"]).hyperlink.target
                self.assertEqual(link, "review_crops/021_2025-12-03_USD_38.33_ALICIA.jpg")
                self.assertEqual(ws.cell(1, columns["Manual status"]).value, "Manual status")
                self.assertTrue((root / "review_crops" / raw_crop.name).exists())
            finally:
                wb.close()

    def test_corrected_row_is_preserved_and_new_rows_skip_locked_number(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / REIMBURSEMENT_WORKBOOK_NAME
            store = ReimbursementWorkbook(path)
            store.write_records(
                [InvoiceRecord(line_no=1, invoice_date="2026-06-12", expense_category="Food", total_amount=100, seller="Cafe")]
            )
            wb = load_workbook(path)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                ws.cell(2, 12).value = "corrected"
                ws.cell(2, 9).value = "Manual Cafe"
                wb.save(path)
            finally:
                wb.close()

            records = [InvoiceRecord(invoice_date="2026-06-13", expense_category="Gas", total_amount=200, seller="Pemex")]
            assign_available_line_numbers(records, store.locked_numbers())
            store.write_records(records)

            wb = load_workbook(path, data_only=True)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                self.assertEqual(ws.cell(2, 1).value, 1)
                self.assertEqual(ws.cell(2, 9).value, "Manual Cafe")
                self.assertEqual(ws.cell(3, 1).value, 2)
                self.assertEqual(ws.cell(3, 5).value, "汽油")
                self.assertEqual(ws.cell(3, 9).value, "Pemex")
                self.assertEqual(ws.cell(3, 11).value, "Gas")
            finally:
                wb.close()

    def test_corrected_crop_names_are_preserved_when_clearing_generated_crops(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / REIMBURSEMENT_WORKBOOK_NAME
            final_dir = root / "final_crops"
            final_dir.mkdir()
            keep = final_dir / "001_2026-06-12_MXN_100.00_Cafe.jpg"
            stale = final_dir / "002_2026-06-12_MXN_50.00_Stale.jpg"
            keep.write_bytes(b"keep")
            stale.write_bytes(b"stale")
            ReimbursementWorkbook(path).write_records(
                [
                    InvoiceRecord(
                        line_no=1,
                        invoice_date="2026-06-12",
                        expense_category="Food",
                        total_amount=100,
                        seller="Cafe",
                        crop_image=str(keep),
                    )
                ]
            )
            wb = load_workbook(path)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                ws.cell(2, 12).value = "corrected"
                wb.save(path)
            finally:
                wb.close()

            clear_generated_crops(final_dir, corrected_crop_names(path))

            self.assertTrue(keep.exists())
            self.assertFalse(stale.exists())

    def test_focus_workbook_selects_first_row_matching_target_day(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / REIMBURSEMENT_WORKBOOK_NAME
            ReimbursementWorkbook(path).write_records(
                [
                    InvoiceRecord(line_no=1, invoice_date="2026-06-20", expense_category="Food", total_amount=100, seller="Cafe"),
                    InvoiceRecord(line_no=2, invoice_date="2026-06-21", expense_category="Gas", total_amount=200, seller="Pemex"),
                ]
            )

            row = focus_reimbursement_workbook(path, date(2026, 6, 21))

            wb = load_workbook(path)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                self.assertEqual(row, 3)
                self.assertEqual(wb.active.title, INVOICE_EXP_SHEET)
                self.assertEqual(ws.sheet_view.selection[0].activeCell, "A3")
            finally:
                wb.close()


if __name__ == "__main__":
    unittest.main()
