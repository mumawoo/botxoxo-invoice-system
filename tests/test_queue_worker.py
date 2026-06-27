import tempfile
import unittest
from pathlib import Path

from invoice_system.config import Settings
from invoice_system.models import PipelineSummary
from invoice_system.queue_worker import (
    DONE,
    FAILED_RETRYABLE,
    PENDING,
    discover_and_enqueue,
    enqueue_photo,
    load_queue_state,
    process_user_queue_once,
    reset_active_user_workspace,
    retry_failed,
    telegram_user_day_dir,
    telegram_user_output_dir,
    telegram_user_queue_path,
)


class QueueWorkerTests(unittest.TestCase):
    def test_enqueue_photo_writes_per_user_queue(self):
        with tempfile.TemporaryDirectory() as temp:
            settings = _settings(Path(temp))
            photo = telegram_user_day_dir(settings, 123) / "photo.jpg"
            photo.parent.mkdir(parents=True)
            photo.write_bytes(b"jpg")

            summary = enqueue_photo(settings, 123, photo)

            self.assertEqual(summary.pending, 1)
            self.assertTrue(telegram_user_queue_path(settings, 123).exists())
            self.assertFalse(telegram_user_queue_path(settings, 456).exists())

    def test_discover_and_enqueue_scans_user_root_only(self):
        with tempfile.TemporaryDirectory() as temp:
            settings = _settings(Path(temp))
            own = telegram_user_day_dir(settings, 123) / "own.jpg"
            other = telegram_user_day_dir(settings, 456) / "other.jpg"
            own.parent.mkdir(parents=True)
            other.parent.mkdir(parents=True)
            own.write_bytes(b"jpg")
            other.write_bytes(b"jpg")

            summary = discover_and_enqueue(settings, 123)

            self.assertEqual(summary.pending, 1)
            state = load_queue_state(telegram_user_queue_path(settings, 123))
            self.assertEqual(Path(state.items[0].path).name, "own.jpg")

    def test_worker_marks_failure_retryable_and_continues(self):
        with tempfile.TemporaryDirectory() as temp:
            settings = _settings(Path(temp))
            good = telegram_user_day_dir(settings, 123) / "good.jpg"
            bad = telegram_user_day_dir(settings, 123) / "bad.jpg"
            good.parent.mkdir(parents=True)
            good.write_bytes(b"jpg")
            bad.write_bytes(b"jpg")

            summary = process_user_queue_once(settings, 123, pipeline_factory=FakeQueuePipeline)

            self.assertEqual(summary.done, 1)
            self.assertEqual(summary.failed, 1)
            state = load_queue_state(telegram_user_queue_path(settings, 123))
            statuses = {Path(item.path).name: item.status for item in state.items}
            self.assertEqual(statuses["good.jpg"], DONE)
            self.assertEqual(statuses["bad.jpg"], FAILED_RETRYABLE)

    def test_retry_failed_moves_only_failed_to_pending(self):
        with tempfile.TemporaryDirectory() as temp:
            settings = _settings(Path(temp))
            good = telegram_user_day_dir(settings, 123) / "good.jpg"
            bad = telegram_user_day_dir(settings, 123) / "bad.jpg"
            good.parent.mkdir(parents=True)
            good.write_bytes(b"jpg")
            bad.write_bytes(b"jpg")
            process_user_queue_once(settings, 123, pipeline_factory=FakeQueuePipeline)

            count, summary = retry_failed(settings, 123)

            self.assertEqual(count, 1)
            self.assertEqual(summary.pending, 1)
            state = load_queue_state(telegram_user_queue_path(settings, 123))
            statuses = {Path(item.path).name: item.status for item in state.items}
            self.assertEqual(statuses["good.jpg"], DONE)
            self.assertEqual(statuses["bad.jpg"], PENDING)

    def test_worker_notifies_after_each_item(self):
        with tempfile.TemporaryDirectory() as temp:
            settings = _settings(Path(temp))
            good = telegram_user_day_dir(settings, 123) / "good.jpg"
            bad = telegram_user_day_dir(settings, 123) / "bad.jpg"
            good.parent.mkdir(parents=True)
            good.write_bytes(b"jpg")
            bad.write_bytes(b"jpg")
            events = []

            process_user_queue_once(
                settings,
                123,
                pipeline_factory=FakeQueuePipeline,
                item_callback=lambda user_id, item, records: events.append((user_id, Path(item.path).name, item.status)),
            )

            self.assertEqual(set(events), {(123, "good.jpg", DONE), (123, "bad.jpg", FAILED_RETRYABLE)})

    def test_reset_active_user_workspace_archives_inbound_output_and_clears_queue(self):
        with tempfile.TemporaryDirectory() as temp:
            settings = _settings(Path(temp))
            photo = telegram_user_day_dir(settings, 123) / "photo.jpg"
            photo.parent.mkdir(parents=True)
            photo.write_bytes(b"jpg")
            output_dir = telegram_user_output_dir(settings, 123)
            final_crop = output_dir / "final_crops" / "001.jpg"
            final_crop.parent.mkdir(parents=True)
            final_crop.write_bytes(b"crop")
            enqueue_photo(settings, 123, photo)

            summary = reset_active_user_workspace(settings, 123)

            self.assertTrue(summary.archive_dir.exists())
            self.assertFalse(photo.exists())
            self.assertFalse(final_crop.exists())
            self.assertEqual(load_queue_state(telegram_user_queue_path(settings, 123)).items, [])
            archived_files = list(summary.archive_dir.rglob("*"))
            self.assertTrue(any(path.name == "photo.jpg" for path in archived_files))
            self.assertTrue(any(path.name == "001.jpg" for path in archived_files))


class FakeQueuePipeline:
    def __init__(self, settings: Settings, trial: bool, output_dir: Path | None):
        self.output_dir = output_dir or Path(".")

    def process_path(self, input_path: Path | None = None, resume: bool = False):
        if input_path and input_path.name == "bad.jpg":
            raise RuntimeError("simulated OCR failure")
        workbook = self.output_dir / "Invoice_Output.xlsx"
        workbook.parent.mkdir(parents=True, exist_ok=True)
        workbook.write_bytes(b"xlsx")
        return PipelineSummary(1, 1, 1, workbook)


def _settings(root: Path) -> Settings:
    return Settings(
        root=root,
        inbound_dir=root / "data" / "inbound",
        trial_dir=root / "data" / "trial",
        output_dir=root / "data" / "output",
        baseline_dir=root / "data" / "baseline",
        telegram_allowed_user_ids=frozenset({123, 456}),
    )


if __name__ == "__main__":
    unittest.main()
