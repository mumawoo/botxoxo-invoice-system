from __future__ import annotations

import base64
import json
import mimetypes
import re
from pathlib import Path

from .config import Settings
from .expense_categories import DEFAULT_EXPENSE_CATEGORY, EXPENSE_CATEGORIES, normalize_expense_category
from .models import InvoiceRecord, OCRResult, OCRTextLine
from .parsing import normalize_date


PROMPT = f"""Extract this receipt/invoice into strict JSON only.
Return keys: invoice_date, expense_category, contents, currency, total_amount,
expense_amount, vat_amount, sales_tax, tips, seller, remarks, rotate_degrees,
orientation_confidence.
invoice_date, seller, currency, and total_amount are required.
Use ISO date YYYY-MM-DD. Detect currency from the receipt: use USD for US dollar receipts and MXN for Mexican peso receipts. A "$" symbol alone is ambiguous, so infer from merchant country, tax labels, address, and context before defaulting to MXN. Use 0 for unknown optional amounts.
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
        if not self.settings.openai_api_key:
            return OCRResult(self.engine, error="OPENAI_API_KEY is not configured")
        try:
            from openai import OpenAI

            client = OpenAI(api_key=self.settings.openai_api_key)
            image_url = f"data:{_mime_type(image_path)};base64,{_base64(image_path)}"
            response = client.responses.create(
                model=self.settings.openai_model,
                input=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "input_text", "text": PROMPT},
                            {"type": "input_image", "image_url": image_url},
                        ],
                    }
                ],
            )
            text = getattr(response, "output_text", "") or _collect_output_text(response)
            return _ocr_result_from_response_text(text, self.engine)
        except Exception as exc:
            return OCRResult(self.engine, error=str(exc))


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
        record = _record_from_data(data)
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


def _record_from_data(data: dict) -> InvoiceRecord:
    contents = str(data.get("contents") or "")
    seller = str(data.get("seller") or "Unknown")
    record = InvoiceRecord(
        invoice_date=normalize_date(str(data.get("invoice_date") or "")),
        expense_category=normalize_expense_category(str(data.get("expense_category") or ""), f"{seller} {contents}"),
        contents=contents,
        currency=str(data.get("currency") or "MXN"),
        total_amount=_float(data.get("total_amount")),
        expense_amount=_float(data.get("expense_amount")),
        vat_amount=_float(data.get("vat_amount")),
        sales_tax=_float(data.get("sales_tax")),
        tips=_float(data.get("tips")),
        seller=seller,
        remarks=str(data.get("remarks") or "Codex Scan used"),
    )
    _validate_record(record)
    if record.expense_amount <= 0 and record.total_amount > 0:
        record.expense_amount = max(record.total_amount - record.vat_amount - record.sales_tax, 0.0)
    return record


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
    if not record.invoice_date:
        missing.append("invoice_date")
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
