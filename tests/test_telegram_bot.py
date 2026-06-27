import tempfile
import unittest
import json
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from invoice_system.config import Settings
from invoice_system.models import InvoiceRecord, PipelineSummary
from invoice_system.queue_worker import DONE, QueueItem, save_queue_state, telegram_user_queue_path, telegram_user_workbook, QueueState
from invoice_system.reimbursement_excel import INVOICE_EXP_SHEET, ReimbursementWorkbook
from invoice_system.telegram_bot import (
    append_process_status,
    change_message,
    delete_message,
    format_telegram_config,
    is_allowed_user,
    processing_failure_message,
    processing_success_message,
    queued_photo_message,
    recent_message,
    resolve_auto_process,
    review_crop_paths,
    rerun_message,
    saved_photo_message,
    scan_completion_message,
    set_user_language,
    submit_confirmation_message,
    submit_message,
    submit_pending_confirmation,
    telegram_batch_source,
    telegram_command_menu,
    telegram_config_ready,
    telegram_help_message,
    telegram_photo_filename,
    telegram_polling_ready,
    telegram_start_message,
    user_language,
    whoami_message,
)


class TelegramBotTests(unittest.TestCase):
    def test_empty_allowed_list_rejects_photos(self):
        self.assertFalse(is_allowed_user(1, set()))

    def test_allowed_user_filter(self):
        self.assertTrue(is_allowed_user(7, {7, 8}))
        self.assertFalse(is_allowed_user(9, {7, 8}))

    def test_resolve_auto_process_uses_env_setting_by_default(self):
        self.assertTrue(resolve_auto_process(Settings(telegram_auto_process=True)))
        self.assertFalse(resolve_auto_process(Settings(telegram_auto_process=False)))

    def test_resolve_auto_process_allows_cli_override(self):
        self.assertFalse(resolve_auto_process(Settings(telegram_auto_process=True), False))
        self.assertTrue(resolve_auto_process(Settings(telegram_auto_process=False), True))

    def test_telegram_config_ready_requires_token_and_allowed_ids(self):
        self.assertFalse(telegram_polling_ready(Settings()))
        self.assertTrue(telegram_polling_ready(Settings(telegram_bot_token="token")))
        self.assertFalse(telegram_config_ready(Settings(telegram_bot_token="token")))
        self.assertFalse(telegram_config_ready(Settings(telegram_allowed_user_ids=frozenset({123}))))
        self.assertTrue(telegram_config_ready(Settings(telegram_bot_token="token", telegram_allowed_user_ids=frozenset({123}))))

    def test_format_telegram_config_reports_missing_setup(self):
        text = format_telegram_config(Settings())

        self.assertIn("Bot token: missing", text)
        self.assertIn("Telegram package:", text)
        self.assertIn("Allowed user IDs: missing", text)
        self.assertIn("Polling startup: NOT READY", text)
        self.assertIn("Photo ingestion: NOT READY", text)
        self.assertIn("Status: NOT READY", text)

    def test_format_telegram_config_reports_setup_mode(self):
        text = format_telegram_config(Settings(telegram_bot_token="token"))

        self.assertIn("Bot token: configured", text)
        self.assertIn("Allowed user IDs: missing", text)
        self.assertIn("Polling startup: READY", text)
        self.assertIn("Photo ingestion: NOT READY", text)
        self.assertIn("Polling can start for /whoami", text)
        self.assertIn("Status: NOT READY", text)

    def test_format_telegram_config_reports_ready_setup(self):
        text = format_telegram_config(
            Settings(
                telegram_bot_token="token",
                telegram_allowed_user_ids=frozenset({123, 456}),
                telegram_auto_process=True,
                openai_api_key="key",
            )
        )

        self.assertIn("Bot token: configured", text)
        self.assertIn("Allowed user IDs: 2", text)
        self.assertIn("Auto process: enabled", text)
        self.assertIn("Qwen OCR: disabled", text)
        self.assertIn("Codex Scan fallback: disabled", text)
        self.assertIn("Polling startup: READY", text)
        self.assertIn("Photo ingestion: READY", text)
        self.assertIn("Status: READY", text)

    def test_format_telegram_config_reports_enabled_qwen_scan(self):
        text = format_telegram_config(Settings(qwen_scan_enabled=True, qwen_api_key="key"))

        self.assertIn("Qwen OCR: enabled", text)

    def test_format_telegram_config_reports_enabled_codex_scan(self):
        text = format_telegram_config(Settings(codex_scan_enabled=True, openai_api_key="key"))

        self.assertIn("Codex Scan fallback: enabled", text)

    def test_telegram_start_message_reports_setup_mode_without_allowed_ids(self):
        self.assertIn("setup mode", telegram_start_message(Settings()))
        self.assertIn("photos are rejected", telegram_start_message(Settings()))
        self.assertIn("Invoice bot is ready", telegram_start_message(Settings(telegram_allowed_user_ids=frozenset({123}))))

    def test_whoami_message(self):
        self.assertEqual(whoami_message(123, "marco"), "Your Telegram user ID is 123 (@marco).")
        self.assertEqual(whoami_message(123), "Your Telegram user ID is 123.")

    def test_append_process_status_adds_pid(self):
        self.assertEqual(append_process_status("Status\nPending: 0", pid=1234), "Status\nPending: 0\nTelegram bot PID: 1234")

    def test_saved_photo_message(self):
        path = Path("data/inbound/telegram/2026-06-12/photo.jpg")
        self.assertEqual(saved_photo_message(path), "Saved photo")

    def test_queued_photo_message_reports_worker_status(self):
        path = Path("data/inbound/telegram/123/2026-06-12/photo.jpg")

        self.assertIn("Scanner: started", queued_photo_message(path, True))
        self.assertIn("Scanner: already running", queued_photo_message(path, False))

    def test_help_message_lists_safe_mobile_commands(self):
        text = telegram_help_message()

        self.assertIn("/status", text)
        self.assertIn("/restart", text)
        self.assertIn("/excel", text)
        self.assertIn("/crops", text)
        self.assertIn("/change", text)
        self.assertIn("/del", text)
        self.assertNotIn("/today_excel -", text)
        self.assertIn("/submit", text)

    def test_chinese_help_and_language_preference(self):
        with tempfile.TemporaryDirectory() as temp:
            settings = Settings(root=Path(temp), output_dir=Path(temp) / "out", telegram_language="en")

            self.assertEqual(user_language(settings, 123), "en")
            self.assertEqual(set_user_language(settings, 123, "zh"), "zh")
            self.assertEqual(user_language(settings, 123), "zh")

        text = telegram_help_message("zh")
        self.assertIn("命令", text)
        self.assertIn("/lang", text)
        self.assertIn("切换语言", text)

    def test_telegram_command_menu_lists_popup_commands(self):
        commands = dict(telegram_command_menu())

        self.assertIn("change", commands)
        self.assertIn("del", commands)
        self.assertIn("rerun", commands)
        self.assertIn("submit", commands)
        self.assertIn("edit crop", commands["change"])

    def test_empty_change_message_shows_mobile_template(self):
        text = change_message(Settings(), 123, [])

        self.assertIn("/change 021 type Other", text)
        self.assertIn("/change 021 amount 33.35 currency USD", text)
        self.assertIn("/del 021", text)

    def test_empty_del_message_shows_mobile_template(self):
        text = delete_message(Settings(), 123, [])

        self.assertIn("/del 021", text)
        self.assertIn("/del 021 022", text)

    def test_processing_success_message(self):
        summary = PipelineSummary(
            source_images=1,
            crops=2,
            records_written=2,
            workbook_path=Path("data/output/production/报销明细_2026_xlsx.xlsx"),
        )
        self.assertEqual(
            processing_success_message(summary),
            "Saved photo and processed current Telegram batch. Sources: 1. Rows: 2. Excel: data\\output\\production\\报销明细_2026_xlsx.xlsx",
        )

    def test_telegram_batch_source_uses_day_folder(self):
        day = Path("data/inbound/telegram/2026-06-12")

        self.assertEqual(telegram_batch_source(day), day)

    def test_processing_failure_message_includes_saved_path_and_error(self):
        path = Path("data/inbound/telegram/2026-06-12/photo.jpg")
        self.assertEqual(
            processing_failure_message(path, RuntimeError("OCR unavailable")),
            "Saved photo\nProcessing failed: OCR unavailable",
        )

    def test_processing_failure_message_handles_empty_error_text(self):
        path = Path("photo.jpg")
        self.assertEqual(processing_failure_message(path, RuntimeError()), "Saved photo\nProcessing failed: RuntimeError")

    def test_scan_completion_message_reports_crop_lines_and_today_totals(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
                telegram_allowed_user_ids=frozenset({123}),
            )
            records = [
                InvoiceRecord(
                    line_no=1,
                    invoice_date="2026-06-20",
                    expense_category="Food",
                    seller="Cafe",
                    total_amount=150,
                    crop_image=str(root / "crops" / "021_2026-06-20_MXN_150.00_Cafe.jpg"),
                )
            ]
            ReimbursementWorkbook(telegram_user_workbook(settings, 123)).write_records(records)
            item = QueueItem(
                path=str(root / "photo.jpg"),
                status=DONE,
                row_count=1,
                total_amount=150,
                category_totals={"Food": 150},
                updated_at=datetime.now().isoformat(timespec="seconds"),
            )

            text = scan_completion_message(settings, 123, item, records)

            self.assertIn("Scan complete", text)
            self.assertNotIn("Photo:", text)
            self.assertNotIn("photo.jpg", text)
            self.assertIn("This scan", text)
            self.assertIn("Crops:\n- 021 Food: MXN 150.00 | Cafe", text)
            self.assertIn("Original totals:\n- MXN: 150.00", text)
            self.assertIn("Today\nRows: 1\nMXN total: 150.00", text)
            self.assertNotIn("Since last submit", text)
            self.assertIn("- Food: 150.00", text)
            self.assertIn("Queue", text)
            self.assertIn("Review: /excel", text)

    def test_change_message_edits_manual_workbook_by_crop_id(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
                telegram_allowed_user_ids=frozenset({123}),
            )
            output = settings.output_dir / "telegram" / "123"
            review = output / "review_crops"
            review.mkdir(parents=True)
            crop = review / "021_2026-06-20_MXN_150.00_Cafe.jpg"
            crop.write_bytes(b"jpg")
            ReimbursementWorkbook(telegram_user_workbook(settings, 123)).write_records(
                [InvoiceRecord(line_no=1, invoice_date="2026-06-20", expense_category="Food", seller="Cafe", total_amount=150, crop_image=str(crop))]
            )

            text = change_message(settings, 123, ["021", "correct", "type", "Other", "+", "wrong", "OCR"])

            self.assertIn("Change saved", text)
            self.assertIn("Crop: 021", text)
            self.assertIn("Status: correct", text)
            wb = load_workbook(telegram_user_workbook(settings, 123), data_only=True)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                headers = {ws.cell(1, col).value: col for col in range(1, ws.max_column + 1)}
                self.assertEqual(ws.cell(2, headers["Accounting Category"]).value, "Other")
                self.assertEqual(ws.cell(2, headers["Manual status"]).value, "correct")
                self.assertIn("wrong OCR", ws.cell(2, headers["Detail"]).value)
            finally:
                wb.close()

    def test_change_message_rejects_delete_and_del_deletes_crop(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
                telegram_allowed_user_ids=frozenset({123}),
            )
            output = settings.output_dir / "telegram" / "123"
            review = output / "review_crops"
            review.mkdir(parents=True)
            crop = review / "021_2026-06-20_MXN_150.00_Cafe.jpg"
            crop.write_bytes(b"jpg")
            ReimbursementWorkbook(telegram_user_workbook(settings, 123)).write_records(
                [InvoiceRecord(line_no=1, invoice_date="2026-06-20", expense_category="Food", seller="Cafe", total_amount=150, crop_image=str(crop))]
            )

            rejected = change_message(settings, 123, ["021", "delete"])
            deleted = delete_message(settings, 123, ["021"])

            self.assertIn("Use /del 021", rejected)
            self.assertIn("Deleted", deleted)
            wb = load_workbook(telegram_user_workbook(settings, 123), data_only=True)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                headers = {ws.cell(1, col).value: col for col in range(1, ws.max_column + 1)}
                self.assertEqual(ws.cell(2, headers["Manual status"]).value, "delete")
            finally:
                wb.close()

    def test_del_message_deletes_multiple_crops(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
                telegram_allowed_user_ids=frozenset({123}),
            )
            output = settings.output_dir / "telegram" / "123"
            review = output / "review_crops"
            review.mkdir(parents=True)
            crop_38 = review / "038_2026-06-20_MXN_150.00_Cafe.jpg"
            crop_39 = review / "039_2026-06-20_MXN_50.00_Gas.jpg"
            crop_38.write_bytes(b"jpg")
            crop_39.write_bytes(b"jpg")
            ReimbursementWorkbook(telegram_user_workbook(settings, 123)).write_records(
                [
                    InvoiceRecord(line_no=1, invoice_date="2026-06-20", expense_category="Food", seller="Cafe", total_amount=150, crop_image=str(crop_38)),
                    InvoiceRecord(line_no=2, invoice_date="2026-06-20", expense_category="Gas", seller="Gas", total_amount=50, crop_image=str(crop_39)),
                ]
            )

            text = delete_message(settings, 123, ["038", "039"])

            self.assertIn("Deleted", text)
            self.assertIn("Crops: 038, 039", text)
            self.assertIn("Count: 2", text)
            wb = load_workbook(telegram_user_workbook(settings, 123), data_only=True)
            try:
                ws = wb[INVOICE_EXP_SHEET]
                headers = {ws.cell(1, col).value: col for col in range(1, ws.max_column + 1)}
                self.assertEqual(ws.cell(2, headers["Manual status"]).value, "delete")
                self.assertEqual(ws.cell(3, headers["Manual status"]).value, "delete")
            finally:
                wb.close()

    def test_rerun_message_reports_missing_checked_baseline(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
                telegram_allowed_user_ids=frozenset({123}),
            )

            text = rerun_message(settings, 123)

            self.assertIn("Cannot rerun", text)
            self.assertIn("/report", text)

    def test_telegram_photo_filename_uses_timestamp_and_file_id(self):
        stamp = datetime(2026, 6, 12, 8, 9, 10)

        self.assertEqual(telegram_photo_filename(stamp, "AgACAgQAAxkBAAIB"), "080910_AgACAgQAAxkBAAIB.jpg")

    def test_telegram_photo_filename_is_windows_safe(self):
        stamp = datetime(2026, 6, 12, 8, 9, 10)

        self.assertEqual(telegram_photo_filename(stamp, 'file/id:with*bad?chars'), "080910_file_id_with_bad_chars.jpg")

    def test_telegram_photo_filename_falls_back_to_unique_id(self):
        stamp = datetime(2026, 6, 12, 8, 9, 10)

        self.assertEqual(telegram_photo_filename(stamp, "???", "unique-id"), "080910_unique-id.jpg")

    def test_review_crop_paths_returns_review_images(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
            )
            user_dir = settings.output_dir / "telegram" / "123"
            crop = user_dir / "review_crops" / "001_2026-06-12_MXN_100.00_Cafe.jpg"
            crop.parent.mkdir(parents=True)
            crop.write_bytes(b"jpg")

            paths = review_crop_paths(settings, 123)

            self.assertEqual(paths, [crop])

    def test_review_crop_paths_returns_latest_completed_photo_only(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
            )
            user_dir = settings.output_dir / "telegram" / "123"
            review = user_dir / "review_crops"
            review.mkdir(parents=True)
            old_crop = review / "001_2026-06-12_MXN_100.00_Old.jpg"
            new_crop = review / "002_2026-06-13_MXN_200.00_New.jpg"
            old_crop.write_bytes(b"old")
            new_crop.write_bytes(b"new")
            old_photo = root / "old.jpg"
            new_photo = root / "new.jpg"
            old_photo.write_bytes(b"old")
            new_photo.write_bytes(b"new")
            save_queue_state(
                telegram_user_queue_path(settings, 123),
                QueueState(
                    [
                        QueueItem(path=str(old_photo), status=DONE, received_at="2026-06-21T10:00:00", row_count=1),
                        QueueItem(path=str(new_photo), status=DONE, received_at="2026-06-21T11:00:00", row_count=1),
                    ]
                ),
            )
            (user_dir / "processing_state.json").write_text(
                json.dumps(
                    {
                        "records": [
                            {"source_image": str(old_photo), "crop_image": str(old_crop)},
                            {"source_image": str(new_photo), "crop_image": str(new_crop)},
                        ]
                    }
                ),
                encoding="utf-8",
            )

            paths = review_crop_paths(settings, 123)

            self.assertEqual(paths, [new_crop])

    def test_review_crop_paths_falls_back_to_recent_rows_when_state_names_are_stale(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
            )
            user_dir = settings.output_dir / "telegram" / "123"
            review = user_dir / "review_crops"
            review.mkdir(parents=True)
            old_crop = review / "001_2026-06-12_MXN_100.00_Old.jpg"
            new_crop_1 = review / "002_2026-06-13_MXN_200.00_New.jpg"
            new_crop_2 = review / "003_2026-06-13_MXN_300.00_New.jpg"
            old_crop.write_bytes(b"old")
            new_crop_1.write_bytes(b"new1")
            new_crop_2.write_bytes(b"new2")
            photo = root / "new.jpg"
            photo.write_bytes(b"photo")
            save_queue_state(
                telegram_user_queue_path(settings, 123),
                QueueState([QueueItem(path=str(photo), status=DONE, received_at="2026-06-21T11:00:00", row_count=2)]),
            )
            (user_dir / "processing_state.json").write_text(
                json.dumps(
                    {
                        "records": [
                            {"source_image": str(photo), "crop_image": "999_missing.jpg"},
                            {"source_image": str(photo), "crop_image": "998_missing.jpg"},
                        ]
                    }
                ),
                encoding="utf-8",
            )

            paths = review_crop_paths(settings, 123)

            self.assertEqual(paths, [new_crop_1, new_crop_2])

    def test_review_crop_paths_uses_raw_crop_paths_when_review_names_are_missing(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
            )
            user_dir = settings.output_dir / "telegram" / "123"
            review = user_dir / "review_crops"
            raw = user_dir / "crops"
            review.mkdir(parents=True)
            raw.mkdir(parents=True)
            old_review = review / "001_2026-06-12_MXN_100.00_Old.jpg"
            raw_crop_1 = raw / "019_2026-06-13_MXN_200.00_New.jpg"
            raw_crop_2 = raw / "020_2026-06-13_MXN_300.00_New.jpg"
            old_review.write_bytes(b"old")
            raw_crop_1.write_bytes(b"raw1")
            raw_crop_2.write_bytes(b"raw2")
            photo = root / "new.jpg"
            photo.write_bytes(b"photo")
            save_queue_state(
                telegram_user_queue_path(settings, 123),
                QueueState([QueueItem(path=str(photo), status=DONE, received_at="2026-06-21T11:00:00", row_count=2)]),
            )
            (user_dir / "processing_state.json").write_text(
                json.dumps(
                    {
                        "records": [
                            {"source_image": str(photo), "crop_image": str(raw_crop_1)},
                            {"source_image": str(photo), "crop_image": str(raw_crop_2)},
                        ]
                    }
                ),
                encoding="utf-8",
            )

            paths = review_crop_paths(settings, 123)

            self.assertEqual(paths, [raw_crop_1, raw_crop_2])

    def test_recent_message_reports_latest_two_uploads_and_excel_rows(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
            )
            user_dir = settings.output_dir / "telegram" / "123"
            review = user_dir / "review_crops"
            review.mkdir(parents=True)
            crop_40 = review / "040_2026-06-12_MXN_126.00_Cafe.jpg"
            crop_41 = review / "041_2026-06-12_MXN_11.00_Store.jpg"
            crop_39 = review / "039_2026-06-11_MXN_24.00_Old.jpg"
            for crop in (crop_39, crop_40, crop_41):
                crop.write_bytes(b"jpg")
            old_photo = root / "old.jpg"
            new_photo = root / "new.jpg"
            old_photo.write_bytes(b"old")
            new_photo.write_bytes(b"new")
            save_queue_state(
                telegram_user_queue_path(settings, 123),
                QueueState(
                    [
                        QueueItem(path=str(old_photo), status=DONE, received_at="2026-06-21T10:00:00", row_count=1),
                        QueueItem(path=str(new_photo), status=DONE, received_at="2026-06-21T11:00:00", row_count=2),
                    ]
                ),
            )
            ReimbursementWorkbook(telegram_user_workbook(settings, 123)).write_records(
                [
                    InvoiceRecord(line_no=39, invoice_date="2026-06-11", expense_category="Food", seller="Old", total_amount=24, crop_image=str(crop_39)),
                    InvoiceRecord(line_no=40, invoice_date="2026-06-12", expense_category="Food", seller="Cafe", total_amount=126, crop_image=str(crop_40)),
                    InvoiceRecord(line_no=41, invoice_date="2026-06-12", expense_category="Other", seller="Store", total_amount=11, crop_image=str(crop_41)),
                ]
            )
            wb = load_workbook(telegram_user_workbook(settings, 123))
            try:
                ws = wb[INVOICE_EXP_SHEET]
                ws.cell(4, 12).value = "delete"
                wb.save(telegram_user_workbook(settings, 123))
            finally:
                wb.close()
            (user_dir / "processing_state.json").write_text(
                json.dumps(
                    {
                        "records": [
                            {"source_image": str(old_photo), "crop_image": str(crop_39)},
                            {"source_image": str(new_photo), "crop_image": str(crop_40)},
                            {"source_image": str(new_photo), "crop_image": str(crop_41)},
                        ]
                    }
                ),
                encoding="utf-8",
            )

            text = recent_message(settings, 123)

            self.assertIn("Recent uploads", text)
            self.assertIn("Input 1: done | Excel rows 2 | crops 2", text)
            self.assertIn("040 -> Excel row 3, No. 040", text)
            self.assertIn("041 -> Excel row 4, No. 041", text)
            self.assertIn("| delete", text)
            self.assertIn("Input 2: done | Excel rows 1 | crops 1", text)
            self.assertIn("039 -> Excel row 2, No. 039", text)

    def test_submit_message_requires_confirmation(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            settings = Settings(
                root=root,
                inbound_dir=root / "data" / "inbound",
                trial_dir=root / "data" / "trial",
                output_dir=root / "data" / "output",
                baseline_dir=root / "data" / "baseline",
            )
            crop = settings.output_dir / "telegram" / "123" / "review_crops" / "001_2026-06-12_MXN_100.00_Cafe.jpg"
            crop.parent.mkdir(parents=True)
            crop.write_bytes(b"jpg")
            ReimbursementWorkbook(telegram_user_workbook(settings, 123)).write_records(
                [InvoiceRecord(line_no=1, invoice_date="2026-06-12", expense_category="Food", seller="Cafe", total_amount=100, crop_image=str(crop))]
            )

            preview = submit_message(settings, 123, [])

            self.assertIn("Submit preview", preview)
            self.assertIn("Reply confirm", preview)
            self.assertTrue(telegram_user_workbook(settings, 123).exists())
            self.assertTrue(submit_pending_confirmation(settings, 123))

            slash_confirm = submit_message(settings, 123, ["confirm"])

            self.assertIn("Use /submit first", slash_confirm)
            self.assertTrue(telegram_user_workbook(settings, 123).exists())

            handled, pending = submit_confirmation_message(settings, 123, "maybe")

            self.assertFalse(handled)
            self.assertIn("Submit pending", pending)

            handled, submitted = submit_confirmation_message(settings, 123, "confirm")

            self.assertTrue(handled)
            self.assertIn("Submit /", submitted)
            self.assertIn("Batch:", submitted)
            self.assertIn("Records: 1", submitted)

    def test_submit_confirmation_can_cancel(self):
        handled, text = submit_confirmation_message(Settings(), 123, "cancel")

        self.assertTrue(handled)
        self.assertIn("cancelled", text)


if __name__ == "__main__":
    unittest.main()
