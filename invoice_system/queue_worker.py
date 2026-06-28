from __future__ import annotations

import json
import shutil
import threading
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable

from .config import Settings
from .expense_categories import normalize_expense_category
from .excel_store import load_invoice_records
from .image_splitter import iter_images
from .models import InvoiceRecord
from .pipeline import InvoicePipeline
from .reimbursement_excel import reimbursement_workbook_path

PENDING = "pending"
PROCESSING = "processing"
DONE = "done"
FAILED_RETRYABLE = "failed_retryable"

_WORKER_LOCK = threading.Lock()
_WORKER_THREAD: threading.Thread | None = None


@dataclass
class QueueItem:
    path: str
    status: str = PENDING
    received_at: str = ""
    updated_at: str = ""
    error: str = ""
    crop_count: int = 0
    row_count: int = 0
    total_amount: float = 0.0
    category_totals: dict[str, float] | None = None
    workbook_path: str = ""


@dataclass
class QueueState:
    items: list[QueueItem]
    worker_status: str = "stopped"
    current_photo: str = ""
    last_error: str = ""
    updated_at: str = ""


@dataclass(frozen=True)
class QueueSummary:
    user_id: int
    pending: int
    processing: int
    done: int
    failed: int
    worker_status: str
    current_photo: str
    excel_path: Path
    queue_path: Path
    last_error: str = ""


@dataclass(frozen=True)
class QueueTotals:
    record_count: int
    total_amount: float
    category_totals: dict[str, float]


@dataclass(frozen=True)
class ResetSummary:
    user_id: int
    archive_dir: Path
    moved_paths: int
    removed_queue: bool


def telegram_user_inbound_root(settings: Settings, user_id: int) -> Path:
    return settings.inbound_dir / "telegram" / str(user_id)


def telegram_user_day_dir(settings: Settings, user_id: int, when: datetime | None = None) -> Path:
    stamp = when or datetime.now()
    return telegram_user_inbound_root(settings, user_id) / stamp.strftime("%Y-%m-%d")


def telegram_user_output_dir(settings: Settings, user_id: int) -> Path:
    return settings.output_dir / "telegram" / str(user_id)


def telegram_user_workbook(settings: Settings, user_id: int) -> Path:
    return reimbursement_workbook_path(telegram_user_output_dir(settings, user_id))


def telegram_user_queue_path(settings: Settings, user_id: int) -> Path:
    return telegram_user_output_dir(settings, user_id) / "queue_state.json"


def reset_active_user_workspace(settings: Settings, user_id: int) -> ResetSummary:
    output_dir = telegram_user_output_dir(settings, user_id)
    inbound_dir = telegram_user_inbound_root(settings, user_id)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    archive_dir = output_dir / "reset_archive" / stamp
    archive_dir.mkdir(parents=True, exist_ok=True)
    moved = 0

    if inbound_dir.exists():
        inbound_archive = archive_dir / "inbound"
        inbound_archive.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(inbound_dir), str(inbound_archive))
        moved += 1
        inbound_dir.mkdir(parents=True, exist_ok=True)

    output_dir.mkdir(parents=True, exist_ok=True)
    for path in list(output_dir.iterdir()):
        if path.name in {"reset_archive", "submitted"}:
            continue
        target = archive_dir / "output" / path.name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(path), str(target))
        moved += 1

    save_queue_state(telegram_user_queue_path(settings, user_id), QueueState([]))
    return ResetSummary(user_id, archive_dir, moved, True)


def enqueue_photo(settings: Settings, user_id: int, photo_path: Path, received_at: datetime | None = None) -> QueueSummary:
    queue_path = telegram_user_queue_path(settings, user_id)
    state = load_queue_state(queue_path)
    key = _path_key(photo_path)
    if not any(_path_key(Path(item.path)) == key for item in state.items):
        now = _timestamp(received_at)
        state.items.append(QueueItem(path=str(photo_path), received_at=now, updated_at=now))
        save_queue_state(queue_path, state)
    return summarize_queue(settings, user_id)


def discover_and_enqueue(settings: Settings, user_id: int) -> QueueSummary:
    root = telegram_user_inbound_root(settings, user_id)
    queue_path = telegram_user_queue_path(settings, user_id)
    state = load_queue_state(queue_path)
    known = {_path_key(Path(item.path)) for item in state.items}
    now = _timestamp()
    for image in iter_images(root):
        key = _path_key(image)
        if key in known:
            continue
        state.items.append(QueueItem(path=str(image), received_at=now, updated_at=now))
        known.add(key)
    save_queue_state(queue_path, state)
    return summarize_queue(settings, user_id)


def retry_failed(settings: Settings, user_id: int) -> tuple[int, QueueSummary]:
    queue_path = telegram_user_queue_path(settings, user_id)
    state = load_queue_state(queue_path)
    count = 0
    now = _timestamp()
    for item in state.items:
        if item.status == FAILED_RETRYABLE:
            item.status = PENDING
            item.error = ""
            item.updated_at = now
            count += 1
    save_queue_state(queue_path, state)
    return count, summarize_queue(settings, user_id)


QueueItemCallback = Callable[[int, QueueItem, list[InvoiceRecord]], None]


def process_user_queue_once(
    settings: Settings,
    user_id: int,
    pipeline_factory=InvoicePipeline,
    item_callback: QueueItemCallback | None = None,
) -> QueueSummary:
    discover_and_enqueue(settings, user_id)
    queue_path = telegram_user_queue_path(settings, user_id)
    output_dir = telegram_user_output_dir(settings, user_id)
    state = load_queue_state(queue_path)
    state.worker_status = "running"
    state.updated_at = _timestamp()
    save_queue_state(queue_path, state)

    try:
        for item in state.items:
            if item.status != PENDING:
                continue
            item.status = PROCESSING
            item.error = ""
            item.updated_at = _timestamp()
            state.current_photo = item.path
            state.worker_status = "running"
            save_queue_state(queue_path, state)
            new_records: list[InvoiceRecord] = []
            try:
                before = _completed_source_count(output_dir)
                before_records = _safe_load_invoice_records(telegram_user_workbook(settings, user_id))
                summary = pipeline_factory(settings=settings, trial=False, output_dir=output_dir).process_path(
                    Path(item.path),
                    resume=True,
                )
                after = _completed_source_count(output_dir)
                after_records = _safe_load_invoice_records(summary.workbook_path)
                new_records = after_records[len(before_records) :]
                item.status = DONE
                item.crop_count = max(summary.crops, item.crop_count)
                item.row_count = len(new_records)
                item.total_amount = round(sum(record.total_amount for record in new_records), 2)
                item.category_totals = _category_totals(new_records)
                item.workbook_path = str(summary.workbook_path)
                if after <= before and before > 0:
                    item.error = "Already completed in processing checkpoint"
            except Exception as exc:
                item.status = FAILED_RETRYABLE
                item.error = str(exc).strip() or exc.__class__.__name__
                state.last_error = item.error
            item.updated_at = _timestamp()
            state.current_photo = ""
            save_queue_state(queue_path, state)
            _notify_item_callback(item_callback, user_id, item, new_records)
    finally:
        state.worker_status = "stopped"
        state.current_photo = ""
        state.updated_at = _timestamp()
        save_queue_state(queue_path, state)
    return summarize_queue(settings, user_id)


def process_all_queues_once(settings: Settings, item_callback: QueueItemCallback | None = None) -> list[QueueSummary]:
    summaries: list[QueueSummary] = []
    for user_id in sorted(settings.telegram_allowed_user_ids):
        summaries.append(process_user_queue_once(settings, user_id, item_callback=item_callback))
    return summaries


def start_background_worker(
    settings: Settings,
    user_id: int | None = None,
    item_callback: QueueItemCallback | None = None,
) -> bool:
    global _WORKER_THREAD
    with _WORKER_LOCK:
        if _WORKER_THREAD and _WORKER_THREAD.is_alive():
            return False
        if settings.telegram_allowed_user_ids:
            target = process_all_queues_once
            args = (settings, item_callback)
        elif user_id is not None:
            target = process_user_queue_once
            args = (settings, user_id, InvoicePipeline, item_callback)
        else:
            return False
        _WORKER_THREAD = threading.Thread(target=target, args=args, daemon=True)
        _WORKER_THREAD.start()
        return True


def summarize_queue(settings: Settings, user_id: int) -> QueueSummary:
    queue_path = telegram_user_queue_path(settings, user_id)
    state = load_queue_state(queue_path)
    return QueueSummary(
        user_id=user_id,
        pending=sum(1 for item in state.items if item.status == PENDING),
        processing=sum(1 for item in state.items if item.status == PROCESSING),
        done=sum(1 for item in state.items if item.status == DONE),
        failed=sum(1 for item in state.items if item.status == FAILED_RETRYABLE),
        worker_status=state.worker_status,
        current_photo=state.current_photo,
        excel_path=telegram_user_workbook(settings, user_id),
        queue_path=queue_path,
        last_error=state.last_error,
    )


def format_status(summary: QueueSummary) -> str:
    lines = [
        "Status / 状态",
        f"User: {summary.user_id}",
        f"Scanner: {summary.worker_status}",
        f"Pending: {summary.pending}",
        f"Processing: {summary.processing}",
        f"Done: {summary.done}",
        f"Failed: {summary.failed}",
        f"Excel: {summary.excel_path}",
    ]
    if summary.current_photo:
        lines.append(f"Current: {summary.current_photo}")
    if summary.last_error:
        lines.append(f"Last error: {summary.last_error}")
    if summary.failed:
        lines.append("Use /restart to retry failed photos.")
    return "\n".join(lines)


def format_reset_summary(summary: ResetSummary) -> str:
    return "\n".join(
        [
            "Reset active batch / 重置当前批次",
            f"User: {summary.user_id}",
            f"Moved active paths: {summary.moved_paths}",
            f"Archive: {summary.archive_dir}",
            "Queue is empty. New Telegram photos will start from 001.",
        ]
    )


def queue_totals_for_day(settings: Settings, user_id: int, day: str | None = None) -> QueueTotals:
    output_dir = telegram_user_output_dir(settings, user_id)
    target_day = day or datetime.now().date().isoformat()
    state = load_queue_state(telegram_user_queue_path(settings, user_id))
    today_items = [item for item in state.items if item.status == DONE and str(item.updated_at or "").startswith(target_day)]
    if not today_items:
        return QueueTotals(0, 0.0, {})

    crop_ids = _crop_ids_for_sources(output_dir, [item.path for item in today_items])
    if crop_ids:
        records_from_excel = _safe_load_invoice_records(reimbursement_workbook_path(output_dir))
        records_from_excel = [record for record in records_from_excel if _record_has_crop_id(record, crop_ids)]
        if records_from_excel:
            return QueueTotals(len(records_from_excel), *_record_totals(records_from_excel))

    total = 0.0
    records = 0
    category_totals: dict[str, float] = {}
    for item in today_items:
        records += int(item.row_count or 0)
        total = round(total + float(item.total_amount or 0), 2)
        for category, amount in (item.category_totals or {}).items():
            normalized = normalize_expense_category(category)
            category_totals[normalized] = round(category_totals.get(normalized, 0.0) + float(amount or 0), 2)
    return QueueTotals(records, total, category_totals)


def _crop_ids_for_sources(output_dir: Path, source_images: list[str]) -> set[str]:
    state_path = output_dir / "processing_state.json"
    if not state_path.exists():
        return set()
    try:
        data = json.loads(state_path.read_text(encoding="utf-8"))
    except Exception:
        return set()
    source_keys = {_path_key(Path(source)) for source in source_images}
    crop_ids: set[str] = set()
    for record in data.get("records", []):
        if not isinstance(record, dict):
            continue
        if _path_key(Path(str(record.get("source_image") or ""))) not in source_keys:
            continue
        for crop_text in [str(record.get("crop_image") or ""), *[str(item or "") for item in record.get("supporting_crop_images", []) or []]]:
            crop_id = _crop_id_from_path_text(crop_text)
            if crop_id:
                crop_ids.add(crop_id)
    return crop_ids


def _record_has_crop_id(record: InvoiceRecord, crop_ids: set[str]) -> bool:
    for crop_text in [record.crop_image, *list(getattr(record, "supporting_crop_images", []) or [])]:
        crop_id = _crop_id_from_path_text(str(crop_text or ""))
        if crop_id and crop_id in crop_ids:
            return True
    return False


def _crop_id_from_path_text(value: str) -> str:
    import re

    name = Path(str(value or "")).name
    match = re.match(r"(\d{3,})_", name)
    return match.group(1) if match else ""


def load_queue_state(path: Path) -> QueueState:
    if not path.exists():
        return QueueState([])
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        items = [QueueItem(**item) for item in data.get("items", []) if isinstance(item, dict)]
        return QueueState(
            items=items,
            worker_status=str(data.get("worker_status") or "stopped"),
            current_photo=str(data.get("current_photo") or ""),
            last_error=str(data.get("last_error") or ""),
            updated_at=str(data.get("updated_at") or ""),
        )
    except Exception:
        return QueueState([])


def save_queue_state(path: Path, state: QueueState) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    state = _preserve_concurrent_items(path, state)
    data = {
        "version": 1,
        "items": [asdict(item) for item in state.items],
        "worker_status": state.worker_status,
        "current_photo": state.current_photo,
        "last_error": state.last_error,
        "updated_at": state.updated_at or _timestamp(),
    }
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(path)


def _preserve_concurrent_items(path: Path, state: QueueState) -> QueueState:
    """Keep photos enqueued by Telegram while a worker has an older state copy."""
    if not path.exists():
        return state
    try:
        existing = load_queue_state(path)
    except Exception:
        return state
    known = {_path_key(Path(item.path)) for item in state.items}
    for item in existing.items:
        key = _path_key(Path(item.path))
        if key not in known:
            state.items.append(item)
            known.add(key)
    return state


def _completed_source_count(output_dir: Path) -> int:
    state_path = output_dir / "processing_state.json"
    if not state_path.exists():
        return 0
    try:
        data = json.loads(state_path.read_text(encoding="utf-8"))
        return len(data.get("completed_sources", []))
    except Exception:
        return 0


def _path_key(path: Path) -> str:
    try:
        return str(path.resolve()).casefold()
    except OSError:
        return str(path).casefold()


def _timestamp(value: datetime | None = None) -> str:
    return (value or datetime.now()).isoformat(timespec="seconds")


def _category_totals(records: list[InvoiceRecord]) -> dict[str, float]:
    _, totals = _record_totals(records)
    return totals


def _record_totals(records: list[InvoiceRecord]) -> tuple[float, dict[str, float]]:
    totals: dict[str, float] = {}
    total = 0.0
    for record in records:
        category = normalize_expense_category(record.expense_category)
        amount = round(float(record.total_amount or 0), 2)
        total = round(total + amount, 2)
        totals[category] = round(totals.get(category, 0.0) + amount, 2)
    return total, totals


def _notify_item_callback(
    callback: QueueItemCallback | None,
    user_id: int,
    item: QueueItem,
    records: list[InvoiceRecord],
) -> None:
    if callback is None:
        return
    try:
        callback(user_id, item, records)
    except Exception:
        return


def _safe_load_invoice_records(path: Path) -> list[InvoiceRecord]:
    try:
        return load_invoice_records(path)
    except Exception:
        return []
