import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from invoice_system.audit import audit_requirements, format_audit
from invoice_system.config import Settings


class AuditTests(unittest.TestCase):
    def test_audit_reports_waiting_for_external_credentials_and_inputs(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
            )
            settings.trial_dir.mkdir(parents=True)
            settings.baseline_dir.mkdir(parents=True)

            with patch("invoice_system.audit.run_checks", return_value=[]):
                text = format_audit(audit_requirements(settings))

            self.assertIn("[WAITING] Qwen-only invoice recognition", text)
            self.assertIn("[WAITING] Qwen Scan credentials", text)
            self.assertIn("[REMOVED] OpenAI fallback", text)
            self.assertIn("[WAITING] Telegram polling ingestion", text)
            self.assertIn("[WAITING] Ubuntu comparison run", text)
            self.assertIn("trial photos=0", text)

    def test_audit_reports_comparison_ready_when_required_files_exist(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
                qwen_api_key="qwen-test",
                qwen_scan_enabled=True,
                telegram_bot_token="token",
                telegram_allowed_user_ids=frozenset({123}),
            )
            settings.trial_dir.mkdir(parents=True)
            settings.baseline_dir.mkdir(parents=True)
            (settings.output_dir / "trial").mkdir(parents=True)
            (settings.trial_dir / "photo.jpg").write_bytes(b"jpg")
            (settings.baseline_dir / "Invoice_Output.xlsx").write_bytes(b"xlsx")
            (settings.baseline_dir / "crop.jpg").write_bytes(b"jpg")
            (settings.output_dir / "trial" / "Invoice_Output_Trial.xlsx").write_bytes(b"xlsx")

            with patch("invoice_system.audit.run_checks", return_value=[]):
                text = format_audit(audit_requirements(settings))

            self.assertIn("[READY] Qwen-only invoice recognition", text)
            self.assertIn("[READY] Qwen Scan credentials", text)
            self.assertIn("[READY] Telegram polling ingestion", text)
            self.assertIn("[READY] Ubuntu comparison run", text)
            self.assertIn("trial photos=1", text)


if __name__ == "__main__":
    unittest.main()
