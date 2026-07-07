from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .config import Settings
from .models import InvoiceRecord, OCRResult
from .parsing import fuzzy_match, normalize_date, normalize_text
from .quality import poor_image_quality_reason
from .recognizers import Recognizer


@dataclass(frozen=True)
class ResolvedScan:
    record: InvoiceRecord
    paddle: OCRResult
    easy: OCRResult
    codex: OCRResult | None
    used_codex: bool
    reason: str


class DualOCRResolver:
    def __init__(
        self,
        paddle: Recognizer,
        easy: Recognizer,
        codex: Recognizer,
        settings: Settings,
    ) -> None:
        self.paddle = paddle
        self.easy = easy
        self.codex = codex
        self.settings = settings

    def scan(self, image_path: Path) -> ResolvedScan:
        easy_result = self.easy.recognize(image_path)
        paddle_result = self.paddle.recognize(image_path)
        quality_reason = poor_image_quality_reason(image_path)
        ok, reason = self._local_results_agree(paddle_result, easy_result)
        if ok and quality_reason:
            ok = False
            reason = quality_reason
        if ok:
            record = _merge_records(paddle_result.parsed_invoice, easy_result.parsed_invoice)
            record.remarks = "PaddleOCR agreed with EasyOCR"
            return ResolvedScan(record, paddle_result, easy_result, None, False, reason)

        fallback_label = _fallback_label(getattr(self.codex, "engine", "vision_scan"))
        if not _fallback_enabled(self.settings, getattr(self.codex, "engine", "")):
            record = _best_local_record(paddle_result, easy_result)
            record.remarks = f"{fallback_label} disabled; local fallback used: {reason}; needs human review"
            return ResolvedScan(record, paddle_result, easy_result, None, False, reason)

        fallback_result = self.codex.recognize(image_path)
        if fallback_result.parsed_invoice is not None:
            record = fallback_result.parsed_invoice
            record.remarks = f"{fallback_label} used: {reason}"
            return ResolvedScan(record, paddle_result, easy_result, fallback_result, True, reason)

        record = _best_local_record(paddle_result, easy_result)
        record.remarks = f"{fallback_label} unavailable; local fallback used: {reason}; {fallback_result.error}; needs human review"
        return ResolvedScan(record, paddle_result, easy_result, fallback_result, False, reason)

    def _local_results_agree(self, left: OCRResult, right: OCRResult) -> tuple[bool, str]:
        if left.error or right.error:
            return False, "local OCR error"
        a = left.parsed_invoice
        b = right.parsed_invoice
        if a is None or b is None:
            return False, "missing local OCR parse"
        if left.confidence < self.settings.local_confidence_threshold:
            return False, "low confidence"
        if right.confidence < self.settings.local_confidence_threshold:
            return False, "low confidence"
        if not _complete(a) or not _complete(b):
            return False, "missing key fields"
        if not _dates_match(a.invoice_date, b.invoice_date):
            return False, "OCR mismatch"
        if abs(a.total_amount - b.total_amount) > self.settings.amount_tolerance:
            return False, "OCR mismatch"
        if _normalize_currency(a.currency) != _normalize_currency(b.currency):
            return False, "OCR mismatch"
        for field_name in ("vat_amount", "sales_tax", "tips"):
            if abs(getattr(a, field_name) - getattr(b, field_name)) > self.settings.amount_tolerance:
                return False, "OCR mismatch"
        if not fuzzy_match(a.seller, b.seller):
            return False, "OCR mismatch"
        return True, "local OCR agreement"


def _complete(record: InvoiceRecord) -> bool:
    return bool(record.invoice_date) and bool(record.currency.strip()) and record.total_amount > 0 and record.seller != "Unknown"


def _dates_match(left: str, right: str) -> bool:
    normalized_left = normalize_date(left) or (left or "").strip()[:10]
    normalized_right = normalize_date(right) or (right or "").strip()[:10]
    return bool(normalized_left) and normalized_left == normalized_right


def _normalize_currency(value: str) -> str:
    normalized = normalize_text(value or "").casefold()
    if normalized in {"m.n.", "mn", "peso", "pesos"}:
        return "MXN"
    return normalized.upper()


def _merge_records(left: InvoiceRecord | None, right: InvoiceRecord | None) -> InvoiceRecord:
    if left is None:
        return right or InvoiceRecord()
    if right is None:
        return left
    left.expense_amount = left.expense_amount or right.expense_amount
    left.vat_amount = max(left.vat_amount, right.vat_amount)
    left.sales_tax = max(left.sales_tax, right.sales_tax)
    left.tips = max(left.tips, right.tips)
    left.contents = left.contents or right.contents
    return left


def _best_local_record(left: OCRResult, right: OCRResult) -> InvoiceRecord:
    candidates = [item for item in (left, right) if item.parsed_invoice is not None]
    if not candidates:
        return InvoiceRecord(remarks="No OCR result")
    candidates.sort(key=lambda item: item.confidence, reverse=True)
    return candidates[0].parsed_invoice or InvoiceRecord()


def _fallback_enabled(settings: Settings, engine: str) -> bool:
    if engine == "qwen_scan":
        return settings.qwen_scan_enabled
    if engine == "codex_scan":
        return False
    return settings.qwen_scan_enabled


def _fallback_label(engine: str) -> str:
    if engine == "qwen_scan":
        return "Qwen Scan"
    if engine == "codex_scan":
        return "Codex Scan"
    return "Vision Scan"
