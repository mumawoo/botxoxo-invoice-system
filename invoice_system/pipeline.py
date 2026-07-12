from __future__ import annotations

import json
import re
import shutil
from copy import deepcopy
from dataclasses import asdict
from datetime import date, datetime
from hashlib import sha256
from pathlib import Path

from .config import Settings
from .dual_ocr import DualOCRResolver, ResolvedScan
from .expense_categories import normalize_expense_category
from .image_splitter import OpenCVInvoiceSplitter, iter_images
from .models import CropResult, InvoiceRecord, OCRAuditRow, OCRResult, PipelineSummary, SourceQARecord
from .pairing import pair_invoice_payment_slips
from .quality import should_delete_failed_crop
from .qwen_scan import QwenScanRecognizer
from .reimbursement_excel import (
    ReimbursementWorkbook,
    assign_available_line_numbers,
    clear_generated_crops,
    next_manual_trace_id,
    reimbursement_workbook_path,
)
from .visual_count import AIVisualCounter, VisualCountResult


SEQUENCE_START = 1
QWEN_ORIENTATION_CONFIDENCE_THRESHOLD = 0.75


class InvoicePipeline:
    def __init__(
        self,
        settings: Settings,
        trial: bool = False,
        output_dir: Path | None = None,
        resolver: DualOCRResolver | None = None,
        visual_counter: AIVisualCounter | None = None,
    ) -> None:
        self.settings = settings
        self.trial = trial
        self.input_dir = settings.trial_dir if trial else settings.inbound_dir
        self.output_dir = output_dir or (settings.output_dir / ("trial" if trial else "production"))
        self.crops_dir = self.output_dir / "crops"
        self.resolver = resolver or QwenOnlyResolver(settings)
        self.visual_counter = visual_counter or AIVisualCounter(settings)

    def process_path(self, input_path: Path | None = None, resume: bool = False) -> PipelineSummary:
        source = input_path or self.input_dir
        source_images = iter_images(source)
        state_path = self.output_dir / "processing_state.json"
        state = _load_state(state_path) if resume else _empty_state()
        if not resume:
            _reset_generated_dir(self.crops_dir)
            _remove_state(state_path)
        else:
            self.crops_dir.mkdir(parents=True, exist_ok=True)
        splitter = OpenCVInvoiceSplitter(self.crops_dir)
        records = list(state.records)
        audits = list(state.audits)
        source_qas = list(state.source_qas)
        crop_count = state.crop_count
        next_crop_id = max(state.next_crop_id, next_manual_trace_id(reimbursement_workbook_path(self.output_dir)))
        completed = set(state.completed_sources)
        source_hashes = set(state.source_hashes)

        for image in source_images:
            image_key = _source_key(image)
            if resume and image_key in completed:
                continue
            image_hash = _file_sha256(image)
            if image_hash and image_hash in source_hashes:
                source_qas.append(_duplicate_source_qa_record(image))
                completed.add(image_key)
                if resume:
                    _save_state(
                        state_path,
                        _PipelineState(completed, crop_count, records, audits, source_qas, source_hashes, next_crop_id),
                    )
                    self._write_outputs(records, audits, source_qas)
                continue
            crops = splitter.split(image)
            visual_count = self._visual_count_for_image(image, len(crops))
            image_records: list[InvoiceRecord] = []
            image_audits: list[OCRAuditRow] = []
            for crop in crops:
                crop_count += 1
                resolved = self.resolver.scan(crop.crop_path)
                record = resolved.record
                record.source_image = str(crop.source_path)
                record.crop_image = str(crop.crop_path)
                codex = resolved.codex
                decision = resolved.reason
                if should_delete_failed_crop(
                    record,
                    crop.crop_path,
                    ocr_text=codex.text if codex else "",
                    ocr_error=codex.error if codex else "",
                ):
                    decision = f"{decision}; noise filtered; crop deleted"
                    _delete_file(crop.crop_path)
                else:
                    record.line_no = next_crop_id
                    next_crop_id += 1
                    image_records.append(record)
                image_audits.append(
                    OCRAuditRow(
                        source_image=str(crop.source_path),
                        crop_image=str(crop.crop_path),
                        decision=decision,
                        used_codex=resolved.used_codex,
                        paddle_confidence=resolved.paddle.confidence,
                        easy_confidence=resolved.easy.confidence,
                        codex_confidence=codex.confidence if codex else 0.0,
                        paddle_error=resolved.paddle.error,
                        easy_error=resolved.easy.error,
                        codex_error=codex.error if codex else "",
                        paddle_text=resolved.paddle.text,
                        easy_text=resolved.easy.text,
                        codex_text=codex.text if codex else "",
                    )
                )
            _fill_missing_dates_from_neighbors(image_records)
            _rename_valid_crops(image_records, image_audits)
            records.extend(image_records)
            audits.extend(image_audits)
            source_qas.append(_source_qa_record(image, visual_count, len(crops), len(image_records)))
            completed.add(image_key)
            if image_hash:
                source_hashes.add(image_hash)
            if resume:
                _save_state(
                    state_path,
                    _PipelineState(completed, crop_count, records, audits, source_qas, source_hashes, next_crop_id),
                )
                self._write_outputs(records, audits, source_qas)

        written, workbook_path = self._write_outputs(records, audits, source_qas)
        if resume:
            _save_state(state_path, _PipelineState(completed, crop_count, records, audits, source_qas, source_hashes, next_crop_id))
        return PipelineSummary(len(source_images), crop_count, written, workbook_path)

    def _visual_count_for_image(self, image: Path, opencv_crop_count: int) -> VisualCountResult:
        threshold = max(self.settings.ai_visual_count_min_opencv_crops, 1)
        if opencv_crop_count < threshold:
            return VisualCountResult(
                None,
                reason=f"Skipped AI visual count: OpenCV crops {opencv_crop_count} < {threshold}",
            )
        return self.visual_counter.count(image)

    def _write_outputs(
        self,
        records: list[InvoiceRecord],
        audits: list[OCRAuditRow],
        source_qas: list[SourceQARecord],
    ) -> tuple[int, Path]:
        store = ReimbursementWorkbook(reimbursement_workbook_path(self.output_dir))
        _fill_missing_dates_from_neighbors(records)
        output_records = pair_invoice_payment_slips(deepcopy(records), mode=self.settings.pairing_mode)
        output_records = store.unlocked_records(output_records)
        assign_available_line_numbers(output_records, store.locked_numbers())
        result = store.write_records(output_records)
        return result.rows_written, result.workbook_path


class _PipelineState:
    def __init__(
        self,
        completed_sources: set[str],
        crop_count: int,
        records: list[InvoiceRecord],
        audits: list[OCRAuditRow],
        source_qas: list[SourceQARecord],
        source_hashes: set[str],
        next_crop_id: int | None = None,
    ) -> None:
        self.completed_sources = completed_sources
        self.crop_count = crop_count
        self.records = records
        self.audits = audits
        self.source_qas = source_qas
        self.source_hashes = source_hashes
        self.next_crop_id = next_crop_id or SEQUENCE_START


def _empty_state() -> _PipelineState:
    return _PipelineState(set(), 0, [], [], [], set(), SEQUENCE_START)


def _source_key(path: Path) -> str:
    try:
        return str(path.resolve())
    except OSError:
        return str(path)


def _load_state(path: Path) -> _PipelineState:
    if not path.exists():
        return _empty_state()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        completed = set(str(item) for item in data.get("completed_sources", []))
        crop_count = int(data.get("crop_count") or 0)
        records = [InvoiceRecord(**item) for item in data.get("records", []) if isinstance(item, dict)]
        next_crop_id = _normalize_next_sequence(data.get("next_crop_id"), fallback=len(records) + 1)
        audits = [OCRAuditRow(**item) for item in data.get("audits", []) if isinstance(item, dict)]
        source_qas = [SourceQARecord(**item) for item in data.get("source_qas", []) if isinstance(item, dict)]
        source_hashes = set(str(item) for item in data.get("source_hashes", []) if item)
        return _PipelineState(completed, crop_count, records, audits, source_qas, source_hashes, next_crop_id)
    except Exception:
        return _empty_state()


def _save_state(path: Path, state: _PipelineState) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    data = {
        "version": 1,
        "completed_sources": sorted(state.completed_sources),
        "crop_count": state.crop_count,
        "records": [asdict(record) for record in state.records],
        "audits": [asdict(audit) for audit in state.audits],
        "source_qas": [asdict(source_qa) for source_qa in state.source_qas],
        "source_hashes": sorted(state.source_hashes),
        "next_crop_id": state.next_crop_id,
    }
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(path)


def _file_sha256(path: Path) -> str:
    try:
        digest = sha256()
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()
    except OSError:
        return ""


def _duplicate_source_qa_record(image: Path) -> SourceQARecord:
    return SourceQARecord(
        source_image=str(image),
        ai_visual_count=None,
        opencv_crop_count=0,
        final_invoice_rows=0,
        needs_human_review=True,
        reason="Exact duplicate source photo skipped by file hash",
    )


def _source_qa_record(
    image: Path,
    visual_count: VisualCountResult,
    opencv_crop_count: int,
    final_invoice_rows: int,
) -> SourceQARecord:
    needs_review = False
    reasons: list[str] = []
    if visual_count.count is not None:
        if visual_count.count != opencv_crop_count:
            needs_review = True
            reasons.append(f"AI count {visual_count.count} != OpenCV crops {opencv_crop_count}")
        if visual_count.confidence and visual_count.confidence < 0.5:
            needs_review = True
            reasons.append("AI visual count low confidence")
    elif visual_count.error:
        needs_review = True
        reasons.append("AI visual count failed")
    elif visual_count.reason:
        reasons.append(visual_count.reason)

    if final_invoice_rows == 0 and opencv_crop_count > 0:
        needs_review = True
        reasons.append("OpenCV crops produced no invoice rows")

    return SourceQARecord(
        source_image=str(image),
        ai_visual_count=visual_count.count,
        opencv_crop_count=opencv_crop_count,
        final_invoice_rows=final_invoice_rows,
        needs_human_review=needs_review,
        reason="; ".join(reasons),
        ai_confidence=visual_count.confidence,
        ai_error=visual_count.error,
    )


def _remove_state(path: Path) -> None:
    if path.exists():
        path.unlink()


def _normalize_next_sequence(value: object, fallback: int = SEQUENCE_START) -> int:
    try:
        sequence = int(value or 0)
    except (TypeError, ValueError):
        sequence = 0
    if sequence < SEQUENCE_START or sequence > 9999:
        return max(fallback, SEQUENCE_START)
    return sequence


def _copy_final_crops(records: list[InvoiceRecord], final_dir: Path, preserve_names: set[str] | None = None) -> list[Path]:
    preserve = set(preserve_names or set())
    for record in records:
        preserve.update({path.name for path in _record_crop_paths(record)})
    clear_generated_crops(final_dir, preserve)
    copied: list[Path] = []
    for record in records:
        if not record.line_no:
            continue
        copied_paths: list[Path] = []
        for source in _record_crop_paths(record):
            if not source.exists():
                continue
            target = final_dir / source.name
            if source.resolve() != target.resolve():
                shutil.copy2(source, target)
            copied_paths.append(target)
            copied.append(target)
        if copied_paths:
            record.crop_image = str(copied_paths[0])
            record.supporting_crop_images = [str(path) for path in copied_paths[1:]]
    return copied


def _record_crop_paths(record: InvoiceRecord) -> list[Path]:
    paths: list[Path] = []
    for crop_image in [record.crop_image, *list(getattr(record, "supporting_crop_images", []) or [])]:
        crop_text = str(crop_image or "").strip()
        if not crop_text:
            continue
        path = Path(crop_text)
        if path not in paths:
            paths.append(path)
    return paths


def _final_crop_name(record: InvoiceRecord, crop_path: Path) -> str:
    line_no = record.line_no or 0
    invoice_date = _safe_filename_part(record.invoice_date or "unknown-date")
    currency = _safe_filename_part(record.currency or "MXN")
    amount = f"{record.total_amount:.2f}"
    seller = _safe_filename_part(record.seller or "Unknown")[:80]
    return f"{line_no:03d}_{invoice_date}_{currency}_{amount}_{seller}.jpg"


def _rename_valid_crops(records: list[InvoiceRecord], audits: list[OCRAuditRow]) -> None:
    audit_by_crop = {audit.crop_image: audit for audit in audits}
    for record in records:
        if not record.crop_image:
            continue
        source = Path(record.crop_image)
        if not source.exists():
            continue
        target = source.with_name(_final_crop_name(record, source))
        if source.resolve() != target.resolve():
            if target.exists():
                target.unlink()
            source.replace(target)
        old_path = record.crop_image
        record.crop_image = str(target)
        audit = audit_by_crop.get(old_path)
        if audit:
            audit.crop_image = str(target)


def _crop_batch_date(image: Path) -> str:
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", image.parent.name):
        return image.parent.name
    try:
        return datetime.fromtimestamp(image.stat().st_mtime).date().isoformat()
    except OSError:
        return date.today().isoformat()


def _fill_missing_dates_from_neighbors(records: list[InvoiceRecord]) -> None:
    dated_indexes = [index for index, record in enumerate(records) if record.invoice_date.strip()]
    if not dated_indexes:
        for record in records:
            fallback_date = _record_source_date(record)
            if not fallback_date:
                continue
            record.invoice_date = fallback_date
            note = "Missing date filled from source photo date"
            record.remarks = f"{record.remarks}; {note}" if record.remarks else note
        return
    for index, record in enumerate(records):
        if record.invoice_date.strip():
            continue
        nearest = min(dated_indexes, key=lambda dated_index: (abs(dated_index - index), dated_index > index))
        neighbor_date = records[nearest].invoice_date.strip()
        if not neighbor_date:
            continue
        record.invoice_date = neighbor_date
        note = "Missing date filled from nearby receipt"
        record.remarks = f"{record.remarks}; {note}" if record.remarks else note


def _record_source_date(record: InvoiceRecord) -> str:
    source = str(record.source_image or record.crop_image or "").strip()
    if source:
        return _crop_batch_date(Path(source))
    return date.today().isoformat()


def _crop_index(path: Path) -> int:
    match = re.search(r"_d(\d+)", path.stem)
    return int(match.group(1)) if match else 1


def _safe_filename_part(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", value.strip())
    cleaned = re.sub(r"\s+", "_", cleaned).strip(" ._")
    return cleaned or "unknown"


def _reset_generated_dir(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path)


def _delete_file(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except OSError:
        pass


class QwenOnlyResolver:
    def __init__(self, settings: Settings) -> None:
        self.qwen = QwenScanRecognizer(settings)

    def scan(self, image_path: Path) -> ResolvedScan:
        result = self.qwen.recognize(image_path)
        empty = OCRResult("local_ocr_disabled")
        if result.parsed_invoice is not None:
            orientation_note = _apply_qwen_orientation(image_path, result)
            record = result.parsed_invoice
            record.expense_category = normalize_expense_category(
                record.expense_category,
                f"{record.seller} {record.contents}",
                company_profile=self.qwen.settings.company_profile,
                root=self.qwen.settings.root,
            )
            record.remarks = record.remarks or "Qwen Scan used"
            reason = "qwen scan only"
            if orientation_note:
                reason = f"{reason}; {orientation_note}"
            return ResolvedScan(record, empty, empty, result, True, reason)
        record = InvoiceRecord(remarks=f"Qwen Scan failed: {result.error}; needs human review")
        return ResolvedScan(record, empty, empty, result, False, "qwen scan failed")


def _apply_qwen_orientation(image_path: Path, result: OCRResult) -> str:
    rotation = int(result.rotate_degrees or 0) % 360
    confidence = float(result.orientation_confidence or 0.0)
    if rotation == 0:
        return f"qwen orientation upright {confidence:.2f}" if confidence else ""
    if rotation not in {90, 180, 270}:
        return ""
    if confidence < QWEN_ORIENTATION_CONFIDENCE_THRESHOLD:
        return f"qwen orientation uncertain {rotation}deg {confidence:.2f}"
    try:
        from PIL import Image, ImageOps

        with Image.open(image_path) as image:
            image = ImageOps.exif_transpose(image).convert("RGB")
            # Pillow uses positive angles for counterclockwise rotation. Qwen,
            # Tesseract, and OpenCV all report the clockwise correction needed.
            image = image.rotate(-rotation, expand=True)
            image.save(image_path, quality=98)
        return f"qwen rotated crop {rotation}deg clockwise {confidence:.2f}"
    except Exception as exc:
        result.error = f"{result.error}; orientation rotate failed: {exc}" if result.error else f"orientation rotate failed: {exc}"
        return f"qwen orientation rotate failed {rotation}deg"
