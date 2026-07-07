from __future__ import annotations

import json
import time
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook

from .ab_test import telegram_ab_input
from .config import Settings
from .image_splitter import OpenCVInvoiceSplitter, iter_images
from .models import CropResult, InvoiceRecord, OCRResult
from .pipeline import QWEN_ORIENTATION_CONFIDENCE_THRESHOLD, _apply_qwen_orientation
from .qwen_scan import QwenScanRecognizer
from .recognizers import Recognizer


@dataclass(frozen=True)
class OrientationABSummary:
    source_images: int
    local_crops: int
    qwen_crops: int
    report_path: Path
    output_dir: Path


@dataclass(frozen=True)
class _TimedCrop:
    crop: CropResult
    split_seconds: float


@dataclass(frozen=True)
class _ScanOutcome:
    result: OCRResult
    seconds: float
    orientation_note: str


def run_orientation_ab_test(
    settings: Settings,
    input_path: Path,
    output_dir: Path,
    *,
    limit: int = 30,
    qwen_recognizer: Recognizer | None = None,
    source_images: list[Path] | None = None,
) -> OrientationABSummary:
    source_images = source_images if source_images is not None else iter_images(input_path)
    output_dir.mkdir(parents=True, exist_ok=True)
    local_crops_dir = output_dir / "local_orientation" / "crops"
    qwen_crops_dir = output_dir / "qwen_orientation" / "crops"
    local_splitter = OpenCVInvoiceSplitter(local_crops_dir, local_orientation=True)
    qwen_splitter = OpenCVInvoiceSplitter(qwen_crops_dir, local_orientation=False)

    local_crops: list[_TimedCrop] = []
    qwen_crops: list[_TimedCrop] = []
    for image in source_images:
        if limit > 0 and max(len(local_crops), len(qwen_crops)) >= limit:
            break
        local_crops.extend(_split_timed(local_splitter, image))
        qwen_crops.extend(_split_timed(qwen_splitter, image))

    if limit > 0:
        local_crops = local_crops[:limit]
        qwen_crops = qwen_crops[:limit]

    qwen = qwen_recognizer or QwenScanRecognizer(settings)
    rows = []
    for index in range(max(len(local_crops), len(qwen_crops))):
        local = local_crops[index] if index < len(local_crops) else None
        qwen_crop = qwen_crops[index] if index < len(qwen_crops) else None
        local_scan = _scan_with_orientation(qwen, local.crop.crop_path) if local else None
        qwen_scan = _scan_with_orientation(qwen, qwen_crop.crop.crop_path) if qwen_crop else None
        rows.append(_comparison_row(index + 1, local, qwen_crop, local_scan, qwen_scan))

    report_path = output_dir / "Orientation_AB_Report.xlsx"
    _write_report(report_path, rows)
    return OrientationABSummary(len(source_images), len(local_crops), len(qwen_crops), report_path, output_dir)


def orientation_ab_input(settings: Settings, user_id: int, day: str | None = None) -> Path:
    return telegram_ab_input(settings, user_id, day)


def prioritized_orientation_inputs(settings: Settings, user_id: int, day: str | None = None) -> list[Path]:
    selected: list[Path] = []
    seen: set[str] = set()
    for path in _orientation_problem_sources(settings, user_id):
        _append_unique_image(selected, seen, path)
    for path in iter_images(orientation_ab_input(settings, user_id, day)):
        _append_unique_image(selected, seen, path)
    return selected


def _split_timed(splitter: OpenCVInvoiceSplitter, image: Path) -> list[_TimedCrop]:
    start = time.perf_counter()
    crops = splitter.split(image)
    elapsed = time.perf_counter() - start
    per_crop = elapsed / max(len(crops), 1)
    return [_TimedCrop(crop, per_crop) for crop in crops]


def _scan_with_orientation(recognizer: Recognizer, image_path: Path) -> _ScanOutcome:
    start = time.perf_counter()
    try:
        result = recognizer.recognize(image_path)
        note = _apply_qwen_orientation(image_path, result)
    except Exception as exc:
        result = OCRResult(getattr(recognizer, "engine", "qwen_scan"), error=str(exc))
        note = ""
    return _ScanOutcome(result, time.perf_counter() - start, note)


def _comparison_row(
    index: int,
    local: _TimedCrop | None,
    qwen_crop: _TimedCrop | None,
    local_scan: _ScanOutcome | None,
    qwen_scan: _ScanOutcome | None,
) -> dict[str, object]:
    local_result = local_scan.result if local_scan else OCRResult("qwen_scan", error="missing local crop")
    qwen_result = qwen_scan.result if qwen_scan else OCRResult("qwen_scan", error="missing qwen crop")
    left = local_result.parsed_invoice
    right = qwen_result.parsed_invoice
    field_matches = {
        "seller": _same_value(_field(left, "seller"), _field(right, "seller")),
        "date": _same_value(_field(left, "invoice_date"), _field(right, "invoice_date")),
        "currency": _same_value(_field(left, "currency"), _field(right, "currency")),
        "amount": _same_value(_field(left, "total_amount"), _field(right, "total_amount")),
        "tips": _same_value(_field(left, "tips"), _field(right, "tips")),
        "category": _same_value(_field(left, "expense_category"), _field(right, "expense_category")),
    }
    orientation_match = int(local_result.rotate_degrees or 0) % 360 == int(qwen_result.rotate_degrees or 0) % 360
    needs_review = (
        not orientation_match
        or any(not value for value in field_matches.values())
        or bool(local_result.error)
        or bool(qwen_result.error)
        or float(qwen_result.orientation_confidence or 0.0) < QWEN_ORIENTATION_CONFIDENCE_THRESHOLD
    )
    source_photo = local.crop.source_path if local else (qwen_crop.crop.source_path if qwen_crop else Path(""))
    return {
        "Crop No.": index,
        "Source Photo": str(source_photo),
        "A Crop Path": str(local.crop.crop_path) if local else "",
        "B Crop Path": str(qwen_crop.crop.crop_path) if qwen_crop else "",
        "A Split/Orientation Seconds": round(local.split_seconds, 3) if local else "",
        "B Split/Orientation Seconds": round(qwen_crop.split_seconds, 3) if qwen_crop else "",
        "A Qwen OCR Seconds": round(local_scan.seconds, 3) if local_scan else "",
        "B Qwen OCR Seconds": round(qwen_scan.seconds, 3) if qwen_scan else "",
        "A Qwen Rotate": int(local_result.rotate_degrees or 0),
        "A Orientation Confidence": round(float(local_result.orientation_confidence or 0.0), 3),
        "B Qwen Rotate": int(qwen_result.rotate_degrees or 0),
        "B Orientation Confidence": round(float(qwen_result.orientation_confidence or 0.0), 3),
        "A Orientation Note": local_scan.orientation_note if local_scan else "",
        "B Orientation Note": qwen_scan.orientation_note if qwen_scan else "",
        "A Seller": _field(left, "seller"),
        "B Seller": _field(right, "seller"),
        "Seller Match": _yes_no(field_matches["seller"]),
        "A Date": _field(left, "invoice_date"),
        "B Date": _field(right, "invoice_date"),
        "Date Match": _yes_no(field_matches["date"]),
        "A Currency": _field(left, "currency"),
        "B Currency": _field(right, "currency"),
        "Currency Match": _yes_no(field_matches["currency"]),
        "A Amount": _field(left, "total_amount"),
        "B Amount": _field(right, "total_amount"),
        "Amount Match": _yes_no(field_matches["amount"]),
        "A Tips": _field(left, "tips"),
        "B Tips": _field(right, "tips"),
        "Tips Match": _yes_no(field_matches["tips"]),
        "A Category": _field(left, "expense_category"),
        "B Category": _field(right, "expense_category"),
        "Category Match": _yes_no(field_matches["category"]),
        "Readable Direction Match": _yes_no(orientation_match),
        "Needs Human Review": _yes_no(needs_review),
        "A Error": local_result.error,
        "B Error": qwen_result.error,
    }


def _write_report(path: Path, rows: list[dict[str, object]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Orientation_AB_Report"
    headers = list(rows[0].keys()) if rows else _report_headers()
    ws.append(headers)
    for row in rows:
        ws.append([row.get(header, "") for header in headers])
    ws.freeze_panes = "A2"
    for column_cells in ws.columns:
        width = min(max(len(str(cell.value or "")) for cell in column_cells) + 2, 60)
        ws.column_dimensions[column_cells[0].column_letter].width = width
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _report_headers() -> list[str]:
    return [
        "Crop No.",
        "Source Photo",
        "A Crop Path",
        "B Crop Path",
        "A Split/Orientation Seconds",
        "B Split/Orientation Seconds",
        "A Qwen OCR Seconds",
        "B Qwen OCR Seconds",
        "A Qwen Rotate",
        "A Orientation Confidence",
        "B Qwen Rotate",
        "B Orientation Confidence",
        "Readable Direction Match",
        "Needs Human Review",
    ]


def _orientation_problem_sources(settings: Settings, user_id: int) -> list[Path]:
    state_path = settings.output_dir / "telegram" / str(user_id) / "processing_state.json"
    if not state_path.exists():
        return []
    try:
        data = json.loads(state_path.read_text(encoding="utf-8"))
    except Exception:
        return []
    sources: list[Path] = []
    for record in data.get("records", []):
        if not isinstance(record, dict):
            continue
        text = " ".join(str(record.get(key) or "") for key in ("remarks", "crop_image", "source_image"))
        if _looks_orientation_related(text):
            source = Path(str(record.get("source_image") or ""))
            if source.exists():
                sources.append(source)
    for audit in data.get("audits", []):
        if not isinstance(audit, dict):
            continue
        text = " ".join(str(value or "") for value in audit.values())
        if _looks_orientation_related(text):
            source = Path(str(audit.get("source_image") or ""))
            if source.exists():
                sources.append(source)
    return sources


def _looks_orientation_related(text: str) -> bool:
    lowered = text.casefold()
    return "qwen rotated crop" in lowered or "orientation uncertain" in lowered or "rotate_degrees" in lowered or " rotated " in lowered


def _append_unique_image(selected: list[Path], seen: set[str], path: Path) -> None:
    if path.suffix.lower() not in {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}:
        return
    try:
        key = str(path.resolve())
    except OSError:
        key = str(path)
    if key not in seen and path.exists():
        seen.add(key)
        selected.append(path)


def _field(record: InvoiceRecord | None, attr: str) -> object:
    return getattr(record, attr, "") if record else ""


def _same_value(left: object, right: object) -> bool:
    if isinstance(left, (int, float)) or isinstance(right, (int, float)):
        try:
            return abs(float(left or 0) - float(right or 0)) <= 0.50
        except (TypeError, ValueError):
            return False
    return str(left or "").strip().casefold() == str(right or "").strip().casefold()


def _yes_no(value: bool) -> str:
    return "yes" if value else "no"
