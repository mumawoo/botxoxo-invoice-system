import unittest
import tempfile
from pathlib import Path

from invoice_system.config import Settings
from invoice_system.diagnostics import format_checks, prepare_handoff, run_checks
from invoice_system.excel_store import InvoiceWorkbook
from invoice_system.models import InvoiceRecord


class DiagnosticsTests(unittest.TestCase):
    def test_run_checks_reports_core_items(self):
        text = format_checks(run_checks(Settings()))
        self.assertIn("Core pipeline:", text)
        self.assertIn("optional", text)
        self.assertIn("openpyxl", text)
        self.assertIn("paddleocr", text)
        self.assertIn("paddle", text)
        self.assertIn("TELEGRAM_BOT_TOKEN", text)

    def test_prepare_handoff_creates_folders_and_reports_counts(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
            )
            (settings.trial_dir).mkdir(parents=True)
            (settings.trial_dir / "photo.jpg").write_bytes(b"jpg")
            (settings.baseline_dir).mkdir(parents=True)
            (settings.baseline_dir / "Invoice_Output.xlsx").write_bytes(b"xlsx")
            (settings.baseline_dir / "crop.jpg").write_bytes(b"jpg")

            text = prepare_handoff(settings)

            self.assertTrue(settings.inbound_dir.exists())
            self.assertTrue(settings.output_dir.exists())
            self.assertTrue((settings.trial_dir / "README.txt").exists())
            self.assertTrue((settings.baseline_dir / "README.txt").exists())
            self.assertIn("run-trial.ps1", (settings.trial_dir / "README.txt").read_text(encoding="utf-8"))
            self.assertIn("compare-ubuntu.ps1", (settings.baseline_dir / "README.txt").read_text(encoding="utf-8"))
            self.assertIn("Trial photos: 1", text)
            self.assertIn("Ubuntu baseline workbooks: 1", text)
            self.assertIn("Ubuntu baseline crop/source images: 1", text)
            self.assertIn("Existing Windows trial output rows: 0", text)
            self.assertNotIn("may be stale", text)

    def test_prepare_handoff_ignores_non_candidate_baseline_workbooks(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
            )
            settings.baseline_dir.mkdir(parents=True)
            (settings.baseline_dir / "Invoice_Manually_Checked.xlsx").write_bytes(b"xlsx")
            (settings.baseline_dir / "comparison_report.xlsx").write_bytes(b"xlsx")
            (settings.baseline_dir / "~$Invoice_Output.xlsx").write_bytes(b"xlsx")

            text = prepare_handoff(settings)

            self.assertIn("Ubuntu baseline workbooks: 0", text)
            self.assertIn("Add the Ubuntu baseline Excel workbook", text)

    def test_prepare_handoff_notes_do_not_count_as_inputs(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
            )

            text = prepare_handoff(settings)

            self.assertIn("Trial photos: 0", text)
            self.assertIn("Ubuntu baseline workbooks: 0", text)
            self.assertIn("Ubuntu baseline crop/source images: 0", text)

    def test_prepare_handoff_warns_when_trial_output_exists_without_trial_photos(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
            )
            workbook = settings.output_dir / "trial" / "Invoice_Output_Trial.xlsx"
            InvoiceWorkbook(workbook).write_records(
                [InvoiceRecord(invoice_date="2026-06-12", seller="Cafe", total_amount=10)]
            )

            text = prepare_handoff(settings)

            self.assertIn("Existing Windows trial output rows: 1", text)
            self.assertIn("Existing trial output may be stale", text)


if __name__ == "__main__":
    unittest.main()
