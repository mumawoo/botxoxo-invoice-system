from __future__ import annotations

import base64
import json
import mimetypes
import re
from pathlib import Path

from .config import Settings
from .expense_categories import DEFAULT_EXPENSE_CATEGORY, EXPENSE_CATEGORIES, normalize_expense_category
from .models import InvoiceRecord, OCRResult, OCRTextLine
from .parsing import normalize_receipt_date


PROMPT = f"""Extract this receipt/invoice into strict JSON only.
Return keys: invoice_date, raw_date, expense_category, contents, currency, total_amount,
expense_amount, vat_amount, sales_tax, tips, seller, remarks, rotate_degrees,
orientation_confidence.
currency and total_amount are required. Use ISO date YYYY-MM-DD for invoice_date when visible; otherwise return an empty invoice_date string and leave it empty. Also return raw_date exactly as printed on the receipt, such as "10/05/2026". Carefully verify every raw-date digit against the image before normalizing it. For ambiguous numeric dates, use DD/MM/YYYY for Mexico, Spanish-language, or MXN receipts. Use MM/DD/YYYY only when the receipt is clearly US/English/USD. For a generic handwritten restaurant pad headed "Nota De Cuenta", use "Nota De Cuenta" as seller rather than leaving seller empty.
Detect currency from the receipt: use USD for US dollar receipts and MXN for Mexican peso receipts. A "$" symbol alone is ambiguous, so infer from merchant country, tax labels, address, and context before defaulting to MXN. Use 0 for unknown optional amounts.
Set tips only when the receipt explicitly labels the amount as tip, tips, propina, or gratuity. A handwritten food/drink line item is not a tip. total_amount must be the total visibly written or circled on the receipt; do not create a new total by adding inferred components.
Payment confirmations and utility bills are valid reimbursement documents. For Multipagos, CFE/Comision Federal de Electricidad, internet payment receipts, or card payment confirmations, use the beneficiary/service provider as seller when visible, use Utilities when it is an electricity/water/gas/utility payment, and extract total_amount from labels such as Importe, Total, Monto, Cantidad, or amount-in-words like "CINCUENTA Y OCHO PESOS 00/100 MXP". Do not treat a Flap/BBVA branding strip as the seller when the document body names the service provider.
expense_category must be exactly one of: {", ".join(EXPENSE_CATEGORIES)}.
Also judge the crop orientation while reading it. rotate_degrees must be one of
0, 90, 180, or 270 and means how much the saved crop image should be rotated
clockwise to make the receipt text upright. Use 0 when already upright.
orientation_confidence must be a number from 0 to 1.
Do not include markdown."""


class CodexScanRecognizer:
    engine = "codex_scan"

    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def recognize(self, image_path: Path) -> OCRResult:
        return OCRResult(self.engine, error="OpenAI/Codex Scan fallback has been removed; use Qwen Scan")


def _base64(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("ascii")


def _mime_type(path: Path) -> str:
    guessed, _ = mimetypes.guess_type(path.name)
    if guessed and guessed.startswith("image/"):
        return guessed
    return "image/jpeg"


def _collect_output_text(response: object) -> str:
    parts: list[str] = []
    for output in getattr(response, "output", []) or []:
        for content in getattr(output, "content", []) or []:
            text = getattr(content, "text", None)
            if text:
                parts.append(str(text))
    return "\n".join(parts)


def _ocr_result_from_response_text(text: str, engine: str = "codex_scan") -> OCRResult:
    lines = [OCRTextLine(text, 1.0)] if text else []
    try:
        data = _json_data(text)
        default_remarks = "Qwen Scan used" if engine == "qwen_scan" else "Codex Scan used"
        record = _record_from_data(data, default_remarks=default_remarks)
    except Exception as exc:
        return OCRResult(engine, lines, None, 0.0, str(exc))
    return OCRResult(
        engine,
        lines,
        record,
        1.0,
        rotate_degrees=_rotate_degrees_from_data(data),
        orientation_confidence=_orientation_confidence_from_data(data),
    )


def _record_from_json(text: str) -> InvoiceRecord:
    return _record_from_data(_json_data(text))


def _json_data(text: str) -> dict:
    data = json.loads(_extract_json_object(text))
    if not isinstance(data, dict):
        raise ValueError("Codex Scan JSON must be an object")
    return data


def _record_from_data(
    data: dict,
    *,
    validate: bool = True,
    default_remarks: str = "Codex Scan used",
) -> InvoiceRecord:
    contents = str(data.get("contents") or "")
    seller = str(data.get("seller") or "Unknown")
    record = InvoiceRecord(
        invoice_date=_invoice_date_from_data(data),
        expense_category=normalize_expense_category(str(data.get("expense_category") or ""), f"{seller} {contents}"),
        contents=contents,
        currency=str(data.get("currency") or "MXN"),
        total_amount=_float(data.get("total_amount")),
        expense_amount=_float(data.get("expense_amount")),
        vat_amount=_float(data.get("vat_amount")),
        sales_tax=_float(data.get("sales_tax")),
        tips=_float(data.get("tips")),
        seller=seller,
        remarks=str(data.get("remarks") or default_remarks),
        report_components=True,
    )
    if validate:
        _validate_record(record)
    if record.expense_amount <= 0 and record.total_amount > 0:
        record.expense_amount = max(record.total_amount - record.vat_amount - record.sales_tax, 0.0)
    _discard_unlabeled_handwritten_tip(record)
    return record


def _discard_unlabeled_handwritten_tip(record: InvoiceRecord) -> None:
    evidence = f"{record.contents} {record.remarks}".casefold()
    seller = (record.seller or "").strip().casefold()
    has_tip_label = any(label in evidence for label in ("tip", "tips", "propina", "gratuity"))
    if seller != "nota de cuenta" or record.tips <= 0 or has_tip_label or record.expense_amount <= 0:
        return
    if abs(record.total_amount - record.expense_amount - record.tips) > 0.51:
        return
    record.total_amount = record.expense_amount
    record.tips = 0.0


def partial_record_from_response_text(text: str) -> InvoiceRecord:
    """Preserve useful Qwen fields when strict validation rejects one missing field."""

    return _record_from_data(_json_data(text), validate=False)


def _invoice_date_from_data(data: dict) -> str:
    raw = data.get("invoice_date")
    text = str(raw or "").strip()
    raw_date = str(data.get("raw_date") or "").strip()
    if not text and not raw_date:
        return ""
    normalized = normalize_receipt_date(
        text,
        raw_date=raw_date,
        currency=str(data.get("currency") or ""),
        context=" ".join(
            [
                str(data.get("seller") or ""),
                str(data.get("contents") or ""),
                str(data.get("remarks") or ""),
                str(data.get("expense_category") or ""),
            ]
        ),
    )
    if not normalized:
        raise ValueError("Codex Scan JSON failed validation: invoice_date")
    return normalized


def _rotate_degrees_from_data(data: dict) -> int:
    try:
        value = int(float(data.get("rotate_degrees") or 0)) % 360
    except (TypeError, ValueError):
        return 0
    return value if value in {0, 90, 180, 270} else 0


def _orientation_confidence_from_data(data: dict) -> float:
    try:
        value = float(data.get("orientation_confidence") or 0)
    except (TypeError, ValueError):
        return 0.0
    return max(0.0, min(value, 1.0))


def _validate_record(record: InvoiceRecord) -> None:
    missing: list[str] = []
    if not record.currency.strip():
        missing.append("currency")
    if (record.seller or "").strip().casefold() in {"", "unknown"}:
        missing.append("seller")
    if record.total_amount <= 0:
        missing.append("total_amount")
    if record.vat_amount < 0 or record.sales_tax < 0 or record.tips < 0 or record.expense_amount < 0:
        missing.append("non_negative_amounts")
    if missing:
        raise ValueError(f"Codex Scan JSON failed validation: {', '.join(missing)}")


def _extract_json_object(text: str) -> str:
    cleaned = (text or "").strip()
    if not cleaned:
        raise ValueError("Codex Scan returned empty text")

    if _loads(cleaned) is not None:
        return cleaned

    fenced = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", cleaned, flags=re.IGNORECASE | re.DOTALL)
    if fenced and _loads(fenced.group(1)) is not None:
        return fenced.group(1)

    for candidate in _balanced_json_candidates(cleaned):
        if _loads(candidate) is not None:
            return candidate
    raise ValueError("Codex Scan did not return a valid JSON object")


def _balanced_json_candidates(text: str) -> list[str]:
    candidates: list[str] = []
    start = -1
    depth = 0
    in_string = False
    escape = False
    for index, char in enumerate(text):
        if escape:
            escape = False
            continue
        if char == "\\" and in_string:
            escape = True
            continue
        if char == '"':
            in_string = not in_string
            continue
        if in_string:
            continue
        if char == "{":
            if depth == 0:
                start = index
            depth += 1
        elif char == "}" and depth:
            depth -= 1
            if depth == 0 and start >= 0:
                candidates.append(text[start : index + 1])
                start = -1
    return candidates


def _loads(value: str) -> object | None:
    try:
        return json.loads(value)
    except json.JSONDecodeError:
        return None


def _float(value: object) -> float:
    try:
        return round(float(value or 0), 2)
    except (TypeError, ValueError):
        return 0.0
