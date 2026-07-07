from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class WindowsScriptTests(unittest.TestCase):
    def test_launcher_scripts_exist(self):
        for name in ("invoice.ps1", "check.ps1", "run-trial.ps1", "telegram.ps1", "compare-ubuntu.ps1", "smoke.ps1", "worker.ps1", "reimburse.ps1"):
            with self.subTest(name=name):
                self.assertTrue((ROOT / "scripts" / name).exists())

    def test_invoice_launcher_invokes_package_module(self):
        text = (ROOT / "scripts" / "invoice.ps1").read_text(encoding="utf-8")
        self.assertIn("-m invoice_system", text)
        self.assertIn("INVOICE_SYSTEM_PYTHON", text)
        self.assertIn("exit $LASTEXITCODE", text)

    def test_compare_launcher_uses_expected_handoff_paths(self):
        text = (ROOT / "scripts" / "compare-ubuntu.ps1").read_text(encoding="utf-8")
        self.assertIn("--baseline data/baseline", text)
        self.assertIn("--candidate data/output/trial", text)
        self.assertIn("data/output/ubuntu_comparison_report.xlsx", text)

    def test_compare_launcher_requires_trial_photos_before_compare(self):
        text = (ROOT / "scripts" / "compare-ubuntu.ps1").read_text(encoding="utf-8")
        self.assertIn('Get-ChildItem -Path "data\\trial"', text)
        self.assertIn("No trial photos found", text)
        self.assertIn("run-trial.ps1", text)

    def test_compare_launcher_preflights_baseline_and_candidate_workbooks(self):
        text = (ROOT / "scripts" / "compare-ubuntu.ps1").read_text(encoding="utf-8")
        self.assertIn('Get-ChildItem -Path "data\\baseline"', text)
        self.assertIn("No Ubuntu baseline workbook found", text)
        self.assertIn('Test-Path "data\\output\\trial\\Invoice_Output_Trial.xlsx"', text)
        self.assertIn("No Windows trial workbook found", text)

    def test_check_launcher_prints_single_strict_report(self):
        text = (ROOT / "scripts" / "check.ps1").read_text(encoding="utf-8")
        self.assertEqual(text.count("check"), 1)
        self.assertIn("--create-dirs --strict", text)

    def test_smoke_launcher_keeps_samples_out_of_trial_folder(self):
        text = (ROOT / "scripts" / "smoke.ps1").read_text(encoding="utf-8")
        self.assertIn("data\\samples\\synthetic_receipt.jpg", text)
        self.assertIn("data\\output\\actual_ocr_smoke", text)
        self.assertIn("data\\samples\\synthetic_receipts_multi.jpg", text)
        self.assertNotIn("data\\trial", text)

    def test_workbench_has_current_crop_and_pid_controls(self):
        text = (ROOT / "scripts" / "workbench.ps1").read_text(encoding="utf-8")
        self.assertIn("Open Crops", text)
        self.assertIn("Build Checked + Final Crops", text)
        self.assertIn("Stop Bot PIDs + Close", text)
        self.assertIn("Show Running PIDs", text)
        self.assertIn("Start / Restart Auto Scan", text)
        self.assertIn("Confirm-AndStopExistingInvoiceProcesses", text)
        self.assertIn("-RestartExisting", text)
        self.assertIn("-m invoice_system", text)
        self.assertNotIn("Open Review Crops", text)


if __name__ == "__main__":
    unittest.main()
