from __future__ import annotations

import json
import shutil
import threading
from dataclasses import replace
from dataclasses import asdict, dataclass
from datetime import datetime
from hashlib import sha256
from pathlib import Path
from typing import Callable

from .config import Settings
from .expense_categories import normalize_expense_category
from .excel_store import load_invoice_records
from .image_splitter import iter_images
from .grouping import clear_forced_group_source, is_forced_group_source
from .models import InvoiceRecord
from .pipeline import InvoicePipeline
from .reimbursement_excel import (
    apply_reimbursement_group,
    next_manual_trace_id,
    reimbursement_workbook_path,
    remove_reimbursement_rows_by_trace_ids,
)

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
    image_width: int = 0
    image_height: int = 0
    file_size: int = 0
    upload_quality: str = "unknown"
    detected_receipt_count: int = 0


@dataclass
class QueueState:
    items: list[QueueItem]
    worker_status: str = "stopped"
    current_photo: str = ""
    last_error: str = ""
    updated_at: str = ""
    rollback_blocked: bool = False
    rollback_block_reason: str = ""


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


@dataclass(frozen=True)
class RollbackSummary:
    user_id: int
    photo_path: Path
    status: str
    trace_ids: tuple[str, ...]
    manual_rows_deleted: int
    crop_files_deleted: int
    queue_items_removed: int
    checked_rows: int
    checked_crops: int


@dataclass(frozen=True)
class RescanSummary:
    user_id: int
    photo_path: Path
    trace_ids: tuple[str, ...]
    manual_rows_deleted: int
    crop_files_deleted: int
    queued: bool


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


def enqueue_photo(
    settings: Settings,
    user_id: int,
    photo_path: Path,
    received_at: datetime | None = None,
    *,
    image_width: int = 0,
    image_height: int = 0,
    file_size: int = 0,
    upload_quality: str = "unknown",
) -> QueueSummary:
    queue_path = telegram_user_queue_path(settings, user_id)
    state = load_queue_state(queue_path)
    key = _path_key(photo_path)
    if not any(_path_key(Path(item.path)) == key for item in state.items):
        now = _timestamp(received_at)
        state.items.append(
            QueueItem(
                path=str(photo_path),
                received_at=now,
                updated_at=now,
                image_width=max(int(image_width or 0), 0),
                image_height=max(int(image_height or 0), 0),
                file_size=max(int(file_size or 0), 0),
                upload_quality=normalize_upload_quality(upload_quality),
            )
        )
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


def block_rollback_for_manual_edit(settings: Settings, user_id: int, reason: str) -> None:
    queue_path = telegram_user_queue_path(settings, user_id)
    state = load_queue_state(queue_path)
    state.rollback_blocked = True
    state.rollback_block_reason = reason
    state.updated_at = _timestamp()
    save_queue_state(queue_path, state)


def rollback_last_photo(settings: Settings, user_id: int) -> RollbackSummary:
    queue_path = telegram_user_queue_path(settings, user_id)
    output_dir = telegram_user_output_dir(settings, user_id)
    state = load_queue_state(queue_path)
    if not state.items:
        raise LookupError("No Telegram photos to rollback")
    index, item = len(state.items) - 1, state.items[-1]
    photo_path = Path(item.path)
    status = item.status
    if status == PROCESSING:
        raise RuntimeError("Last photo is still processing. Wait for scanning to finish, then send /rollback again.")

    if status in {PENDING, FAILED_RETRYABLE}:
        removed = _remove_queue_item_at(state, index)
        if not any(existing.status == FAILED_RETRYABLE for existing in state.items):
            state.last_error = ""
        save_queue_state(queue_path, state, preserve_concurrent=False)
        crop_files_deleted = _delete_paths([photo_path])
        return RollbackSummary(user_id, photo_path, status, (), 0, crop_files_deleted, removed, 0, 0)

    if status != DONE:
        raise RuntimeError(f"Cannot rollback last photo with status: {status}")

    _assert_writable(telegram_user_workbook(settings, user_id))
    processing_path = output_dir / "processing_state.json"
    processing_data = _load_processing_data(processing_path)
    source_key = _path_key(photo_path)
    records = [record for record in processing_data.get("records", []) if isinstance(record, dict)]
    target_records = [record for record in records if _path_key(Path(str(record.get("source_image") or ""))) == source_key]
    if not target_records:
        raise LookupError("Cannot find processing records for the last photo")
    trace_ids = sorted({crop_id for record in target_records for crop_id in _crop_ids_from_processing_record(record)})
    if not trace_ids:
        raise LookupError("Cannot find crop Trace IDs for the last photo")

    paths_to_delete = _rollback_paths_for_source(output_dir, photo_path, processing_data, target_records, trace_ids)
    photo_hash = _file_sha256(photo_path)
    _remove_source_from_processing_data(processing_data, photo_path, photo_hash)
    manual_rows_deleted = remove_reimbursement_rows_by_trace_ids(output_dir, set(trace_ids))
    processing_data["next_crop_id"] = _next_trace_after_rollback(output_dir, processing_data)
    _save_processing_data(processing_path, processing_data)
    removed = _remove_queue_item_at(state, index)
    if not any(existing.status == FAILED_RETRYABLE for existing in state.items):
        state.last_error = ""
    save_queue_state(queue_path, state, preserve_concurrent=False)
    crop_files_deleted = _delete_paths(paths_to_delete)
    return RollbackSummary(
        user_id,
        photo_path,
        status,
        tuple(trace_ids),
        manual_rows_deleted,
        crop_files_deleted,
        removed,
        0,
        0,
    )


def prepare_last_photo_rescan(settings: Settings, user_id: int) -> RescanSummary:
    queue_path = telegram_user_queue_path(settings, user_id)
    output_dir = telegram_user_output_dir(settings, user_id)
    state = load_queue_state(queue_path)
    if not state.items:
        raise LookupError("No Telegram photos to rescan")
    if state.worker_status == "running" or state.current_photo or any(item.status == PROCESSING for item in state.items):
        raise RuntimeError("Queue is processing. Wait for scanning to finish before rescanning the last photo.")
    item = state.items[-1]
    photo_path = Path(item.path)
    if item.status == PENDING:
        return RescanSummary(user_id, photo_path, (), 0, 0, True)
    if item.status not in {DONE, FAILED_RETRYABLE}:
        raise RuntimeError(f"Cannot rescan last photo with status: {item.status}")

    _assert_writable(telegram_user_workbook(settings, user_id))
    processing_path = output_dir / "processing_state.json"
    processing_data = _load_processing_data(processing_path)
    source_key = _path_key(photo_path)
    records = [record for record in processing_data.get("records", []) if isinstance(record, dict)]
    target_records = [record for record in records if _path_key(Path(str(record.get("source_image") or ""))) == source_key]
    trace_ids = sorted({crop_id for record in target_records for crop_id in _crop_ids_from_processing_record(record)})

    paths_to_delete: list[Path] = []
    manual_rows_deleted = 0
    if target_records and trace_ids:
        paths_to_delete = [
            path for path in _rollback_paths_for_source(output_dir, photo_path, processing_data, target_records, trace_ids) if _path_key(path) != source_key
        ]
        photo_hash = _file_sha256(photo_path)
        _remove_source_from_processing_data(processing_data, photo_path, photo_hash)
        manual_rows_deleted = remove_reimbursement_rows_by_trace_ids(output_dir, set(trace_ids))
        processing_data["next_crop_id"] = _next_trace_after_rollback(output_dir, processing_data)
        _save_processing_data(processing_path, processing_data)

    item.status = PENDING
    item.error = ""
    item.crop_count = 0
    item.row_count = 0
    item.total_amount = 0.0
    item.category_totals = {}
    item.workbook_path = ""
    item.detected_receipt_count = 0
    item.updated_at = _timestamp()
    state.current_photo = ""
    state.worker_status = "stopped"
    if not any(existing.status == FAILED_RETRYABLE for existing in state.items):
        state.last_error = ""
    save_queue_state(queue_path, state, preserve_concurrent=False)
    crop_files_deleted = _delete_paths(paths_to_delete)
    return RescanSummary(user_id, photo_path, tuple(trace_ids), manual_rows_deleted, crop_files_deleted, True)


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
                forced_group = is_forced_group_source(output_dir, Path(item.path))
                pipeline_settings = replace(settings, pairing_mode="review") if forced_group else settings
                summary = pipeline_factory(settings=pipeline_settings, trial=False, output_dir=output_dir).process_path(
                    Path(item.path),
                    resume=True,
                )
                after = _completed_source_count(output_dir)
                after_records = _safe_load_invoice_records(summary.workbook_path)
                new_records = _workbook_records_for_source(output_dir, Path(item.path), after_records)
                if forced_group:
                    group_ids = sorted(_crop_ids_for_sources(output_dir, [item.path]), key=lambda value: int(value))
                    if len(group_ids) >= 2:
                        apply_reimbursement_group(output_dir, group_ids)
                        clear_forced_group_source(output_dir, Path(item.path))
                        after_records = _safe_load_invoice_records(summary.workbook_path)
                        new_records = _workbook_records_for_source(output_dir, Path(item.path), after_records)
                    else:
                        item.error = "Group requested, but fewer than two crop records were found"
                        clear_forced_group_source(output_dir, Path(item.path))
                item.status = DONE
                item.crop_count = max(summary.crops, item.crop_count)
                item.row_count = len(new_records)
                item.total_amount = round(sum(record.total_amount for record in new_records), 2)
                item.category_totals = _category_totals(new_records)
                item.workbook_path = str(summary.workbook_path)
                item.detected_receipt_count = _detected_receipt_count(output_dir, Path(item.path))
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


def telegram_photo_quality(width: int, height: int) -> str:
    longest_edge = max(int(width or 0), int(height or 0))
    if longest_edge <= 0:
        return "unknown"
    return "sd" if longest_edge <= 1280 else "hd"


def normalize_upload_quality(value: object) -> str:
    normalized = str(value or "unknown").strip().casefold()
    return normalized if normalized in {"sd", "hd", "original"} else "unknown"


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
    match = re.match(r"(\d{3,})[a-zA-Z]?_", name)
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
            rollback_blocked=bool(data.get("rollback_blocked") or False),
            rollback_block_reason=str(data.get("rollback_block_reason") or ""),
        )
    except Exception:
        return QueueState([])


def save_queue_state(path: Path, state: QueueState, preserve_concurrent: bool = True) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if preserve_concurrent:
        state = _preserve_concurrent_items(path, state)
    data = {
        "version": 1,
        "items": [asdict(item) for item in state.items],
        "worker_status": state.worker_status,
        "current_photo": state.current_photo,
        "last_error": state.last_error,
        "updated_at": state.updated_at or _timestamp(),
        "rollback_blocked": state.rollback_blocked,
        "rollback_block_reason": state.rollback_block_reason,
    }
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(path)


def _load_processing_data(path: Path) -> dict:
    if not path.exists():
        return {"version": 1, "completed_sources": [], "records": [], "audits": [], "source_qas": [], "source_hashes": [], "next_crop_id": 1}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        data = {}
    if not isinstance(data, dict):
        data = {}
    data.setdefault("version", 1)
    data.setdefault("completed_sources", [])
    data.setdefault("records", [])
    data.setdefault("audits", [])
    data.setdefault("source_qas", [])
    data.setdefault("source_hashes", [])
    data.setdefault("next_crop_id", 1)
    return data


def _save_processing_data(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(path)


def _remove_source_from_processing_data(data: dict, photo_path: Path, photo_hash: str) -> None:
    source_key = _path_key(photo_path)
    data["completed_sources"] = [
        value for value in data.get("completed_sources", []) if _path_key(Path(str(value or ""))) != source_key
    ]
    if photo_hash:
        data["source_hashes"] = [value for value in data.get("source_hashes", []) if str(value or "") != photo_hash]
    for key in ("records", "audits", "source_qas"):
        data[key] = [
            item
            for item in data.get(key, [])
            if not isinstance(item, dict) or _path_key(Path(str(item.get("source_image") or ""))) != source_key
        ]


def _rollback_paths_for_source(output_dir: Path, photo_path: Path, data: dict, records: list[dict], trace_ids: list[str]) -> list[Path]:
    paths: list[Path] = [photo_path]
    for record in records:
        for crop_text in [str(record.get("crop_image") or ""), *[str(item or "") for item in record.get("supporting_crop_images", []) or []]]:
            _append_existing_path(paths, Path(crop_text))
    source_key = _path_key(photo_path)
    for audit in data.get("audits", []):
        if not isinstance(audit, dict) or _path_key(Path(str(audit.get("source_image") or ""))) != source_key:
            continue
        _append_existing_path(paths, Path(str(audit.get("crop_image") or "")))
    for folder in ("crops", "review_crops", "final_crops"):
        directory = output_dir / folder
        if not directory.exists():
            continue
        for trace_id in trace_ids:
            for path in directory.rglob(f"*{trace_id}*.jpg"):
                if _path_has_trace_id(path, trace_id):
                    _append_existing_path(paths, path)
    return paths


def _append_existing_path(paths: list[Path], path: Path) -> None:
    if not str(path) or str(path) == ".":
        return
    if path.exists() and path not in paths:
        paths.append(path)


def _crop_ids_from_processing_record(record: dict) -> set[str]:
    ids: set[str] = set()
    for crop_text in [str(record.get("crop_image") or ""), *[str(item or "") for item in record.get("supporting_crop_images", []) or []]]:
        crop_id = _crop_id_from_path_text(crop_text)
        if crop_id:
            ids.add(crop_id)
    return ids


def _path_has_trace_id(path: Path, trace_id: str) -> bool:
    import re

    name = path.name
    return bool(re.match(rf"^{re.escape(trace_id)}[a-zA-Z]?_", name) or f"_trace{trace_id}_" in name)


def _next_trace_after_rollback(output_dir: Path, data: dict) -> int:
    max_trace = 0
    for record in data.get("records", []):
        if not isinstance(record, dict):
            continue
        for crop_id in _crop_ids_from_processing_record(record):
            try:
                max_trace = max(max_trace, int(crop_id))
            except ValueError:
                continue
    try:
        max_trace = max(max_trace, next_manual_trace_id(reimbursement_workbook_path(output_dir)) - 1)
    except Exception:
        pass
    return max_trace + 1 if max_trace > 0 else 1


def _remove_queue_item_at(state: QueueState, index: int) -> int:
    del state.items[index]
    state.updated_at = _timestamp()
    return 1


def _delete_paths(paths: list[Path]) -> int:
    deleted = 0
    for path in paths:
        try:
            if path.exists() and path.is_file():
                path.unlink()
                deleted += 1
        except OSError:
            continue
    return deleted


def _assert_writable(path: Path) -> None:
    if not path.exists():
        return
    with path.open("a+b"):
        pass


def _file_sha256(path: Path) -> str:
    try:
        digest = sha256()
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()
    except OSError:
        return ""


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


def _detected_receipt_count(output_dir: Path, source_path: Path) -> int:
    state_path = output_dir / "processing_state.json"
    if not state_path.exists():
        return 0
    try:
        data = json.loads(state_path.read_text(encoding="utf-8"))
    except Exception:
        return 0
    source_key = _path_key(source_path)
    for item in reversed(data.get("source_qas", [])):
        if not isinstance(item, dict):
            continue
        if _path_key(Path(str(item.get("source_image") or ""))) != source_key:
            continue
        counts = [int(item.get("opencv_crop_count") or 0)]
        if item.get("ai_visual_count") not in (None, ""):
            try:
                counts.append(int(item["ai_visual_count"]))
            except (TypeError, ValueError):
                pass
        return max(counts, default=0)
    return 0


def _workbook_records_for_source(output_dir: Path, source_path: Path, workbook_records: list[InvoiceRecord]) -> list[InvoiceRecord]:
    crop_ids = _source_crop_ids(output_dir, source_path)
    if not crop_ids:
        return []
    records: list[InvoiceRecord] = []
    for record in workbook_records:
        record_ids = _crop_ids_from_invoice_record(record)
        if record_ids & crop_ids:
            records.append(record)
    return records


def _source_crop_ids(output_dir: Path, source_path: Path) -> set[str]:
    state_path = output_dir / "processing_state.json"
    if not state_path.exists():
        return set()
    try:
        data = json.loads(state_path.read_text(encoding="utf-8"))
    except Exception:
        return set()
    source_key = _path_key(source_path)
    crop_ids: set[str] = set()
    for record in data.get("records", []):
        if not isinstance(record, dict):
            continue
        if _path_key(Path(str(record.get("source_image") or ""))) != source_key:
            continue
        crop_ids.update(_crop_ids_from_processing_record(record))
    return crop_ids


def _crop_ids_from_invoice_record(record: InvoiceRecord) -> set[str]:
    ids: set[str] = set()
    for crop_text in [str(record.crop_image or ""), *[str(item or "") for item in getattr(record, "supporting_crop_images", []) or []]]:
        crop_id = _crop_id_from_path_text(crop_text)
        if crop_id:
            ids.add(crop_id)
    return ids


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
