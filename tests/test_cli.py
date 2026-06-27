import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from invoice_system.__main__ import _run_command, main
from invoice_system.config import Settings
from invoice_system.models import PipelineSummary


class CliTests(unittest.TestCase):
    def test_run_returns_error_for_empty_input_folder(self):
        with tempfile.TemporaryDirectory() as temp:
            stderr = io.StringIO()
            with contextlib.redirect_stderr(stderr):
                code = main(["run", "--trial", "--input", temp, "--output", str(Path(temp) / "out")])

            self.assertEqual(code, 1)
            self.assertIn("No invoice photos found", stderr.getvalue())

    def test_compare_returns_error_for_missing_workbooks(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            baseline = root / "baseline"
            candidate = root / "candidate"
            baseline.mkdir()
            candidate.mkdir()
            stderr = io.StringIO()
            with contextlib.redirect_stderr(stderr):
                code = main(["compare", "--baseline", str(baseline), "--candidate", str(candidate)])

            self.assertEqual(code, 1)
            self.assertIn("Comparison input error", stderr.getvalue())

    def test_ab_test_requires_input_or_user_id(self):
        stderr = io.StringIO()
        with contextlib.redirect_stderr(stderr):
            code = main(["ab-test"])

        self.assertEqual(code, 1)
        self.assertIn("pass --input or --user-id", stderr.getvalue())

    def test_prepare_prints_handoff_summary(self):
        stdout = io.StringIO()
        with contextlib.redirect_stdout(stdout):
            code = main(["prepare"])

        self.assertEqual(code, 0)
        self.assertIn("Handoff folders ready.", stdout.getvalue())

    def test_audit_prints_requirement_status(self):
        stdout = io.StringIO()
        with patch("invoice_system.audit.run_checks", return_value=[]), contextlib.redirect_stdout(stdout):
            code = main(["audit"])

        self.assertEqual(code, 0)
        self.assertIn("V2 Windows rewrite audit:", stdout.getvalue())

    def test_telegram_check_returns_error_when_config_missing(self):
        stdout = io.StringIO()
        with patch("invoice_system.__main__.Settings.from_env", return_value=Settings()), contextlib.redirect_stdout(stdout):
            code = main(["telegram", "--check"])

        self.assertEqual(code, 1)
        self.assertIn("Status: NOT READY", stdout.getvalue())
        self.assertIn("Polling startup: NOT READY", stdout.getvalue())
        self.assertIn("Photo ingestion: NOT READY", stdout.getvalue())

    def test_telegram_check_reports_polling_ready_before_allowed_ids(self):
        stdout = io.StringIO()
        with patch("invoice_system.__main__.Settings.from_env", return_value=Settings(telegram_bot_token="token")), contextlib.redirect_stdout(stdout):
            code = main(["telegram", "--check"])

        self.assertEqual(code, 1)
        text = stdout.getvalue()
        self.assertIn("Polling startup: READY", text)
        self.assertIn("Photo ingestion: NOT READY", text)
        self.assertIn("Polling can start for /whoami", text)

    def test_telegram_check_returns_ready_when_configured(self):
        settings = Settings(telegram_bot_token="token", telegram_allowed_user_ids=frozenset({123}))
        stdout = io.StringIO()
        with patch("invoice_system.__main__.Settings.from_env", return_value=settings), contextlib.redirect_stdout(stdout):
            code = main(["telegram", "--check", "--process"])

        self.assertEqual(code, 0)
        text = stdout.getvalue()
        self.assertIn("Status: READY", text)
        self.assertIn("Auto process: enabled", text)

    def test_telegram_startup_error_is_user_friendly(self):
        stderr = io.StringIO()
        with (
            patch("invoice_system.__main__.Settings.from_env", return_value=Settings()),
            patch("invoice_system.telegram_bot.run_polling_bot", side_effect=RuntimeError("TELEGRAM_BOT_TOKEN is not configured")),
            contextlib.redirect_stderr(stderr),
        ):
            code = main(["telegram"])

        self.assertEqual(code, 1)
        text = stderr.getvalue()
        self.assertIn("Telegram startup error: TELEGRAM_BOT_TOKEN is not configured", text)
        self.assertIn("telegram --check", text)

    def test_run_command_returns_clear_error_when_pipeline_fails(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            image = root / "photo.jpg"
            image.write_bytes(b"jpg")
            args = _args(input=image, output=root / "out", trial=True)
            stderr = io.StringIO()

            with contextlib.redirect_stderr(stderr):
                code = _run_command(args, Settings(root=root), pipeline_factory=FailingPipeline)

            self.assertEqual(code, 1)
            self.assertIn("Run error: OCR model unavailable", stderr.getvalue())

    def test_run_command_prints_summary_when_pipeline_succeeds(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            image = root / "photo.jpg"
            image.write_bytes(b"jpg")
            args = _args(input=image, output=root / "out", trial=True)
            stdout = io.StringIO()

            with contextlib.redirect_stdout(stdout):
                code = _run_command(args, Settings(root=root), pipeline_factory=SuccessfulPipeline)

            self.assertEqual(code, 0)
            text = stdout.getvalue()
            self.assertIn("Processed 1 source image(s)", text)
            self.assertIn("Created 2 crop(s)", text)
            self.assertIn("Wrote 2 invoice row(s)", text)


def _args(input: Path, output: Path, trial: bool):
    class Args:
        pass

    args = Args()
    args.input = input
    args.output = output
    args.trial = trial
    args.resume = False
    return args


class FailingPipeline:
    def __init__(self, settings: Settings, trial: bool, output_dir: Path | None):
        pass

    def process_path(self, input_path: Path | None = None, resume: bool = False):
        raise RuntimeError("OCR model unavailable")


class SuccessfulPipeline:
    def __init__(self, settings: Settings, trial: bool, output_dir: Path | None):
        self.workbook = (output_dir or Path(".")) / "Invoice_Output_Trial.xlsx"

    def process_path(self, input_path: Path | None = None, resume: bool = False):
        return PipelineSummary(1, 2, 2, self.workbook)


if __name__ == "__main__":
    unittest.main()
