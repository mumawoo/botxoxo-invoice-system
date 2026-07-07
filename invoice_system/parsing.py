from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from difflib import SequenceMatcher

from .expense_categories import normalize_expense_category
from .models import InvoiceRecord, OCRTextLine

DATE_PATTERNS = [
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%d/%m/%Y",
    "%d-%m-%Y",
    "%d/%m/%y",
    "%d-%m-%y",
]

AMOUNT_RE = re.compile(
    r"(?<!\d)(?:\$|MXN|M\.N\.|MN)?\s*([0-9]{1,6}(?:[,.]\d{3})*(?:[,.]\d{2})|[0-9]{1,6}(?:[,.]\d{3})+|[0-9]{1,6})(?!\d)",
    re.I,
)
DATE_RE = re.compile(r"\b(\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4})\b")

TOTAL_KEYWORDS = ("total", "importe", "monto", "pago", "venta")
VAT_KEYWORDS = ("iva", "vat")
SALES_TAX_KEYWORDS = ("sales tax", "tax", "impuesto", "ieps", "ish")
TIP_KEYWORDS = ("propina", "tip", "tips", "servicio")
SKIP_MERCHANT_KEYWORDS = TOTAL_KEYWORDS + VAT_KEYWORDS + SALES_TAX_KEYWORDS + TIP_KEYWORDS + ("fecha", "date", "folio", "ticket", "rfc")
SKIP_FALLBACK_AMOUNT_KEYWORDS = (
    VAT_KEYWORDS
    + SALES_TAX_KEYWORDS
    + TIP_KEYWORDS
    + ("fecha", "date", "folio", "ticket", "rfc", "mesa", "table", "orden", "order", "caja", "terminal", "autorizacion")
)


def normalize_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value or "")
    return "".join(ch for ch in normalized if not unicodedata.combining(ch)).strip()


def normalize_date(value: str) -> str:
    match = DATE_RE.search(value or "")
    if not match:
        return ""
    raw = match.group(1)
    for pattern in DATE_PATTERNS:
        try:
            return datetime.strptime(raw, pattern).date().isoformat()
        except ValueError:
            continue
    return ""


def normalize_receipt_date(value: str, *, raw_date: str = "", currency: str = "", context: str = "") -> str:
    """Normalize receipt dates with a country-aware tie break for ambiguous numeric dates."""
    raw = (raw_date or "").strip()
    if raw:
        normalized = _normalize_contextual_numeric_date(raw, currency=currency, context=context)
        if normalized:
            return normalized
    text = (value or "").strip()
    normalized = _normalize_contextual_numeric_date(text, currency=currency, context=context)
    return normalized or normalize_date(text)


def _normalize_contextual_numeric_date(value: str, *, currency: str = "", context: str = "") -> str:
    match = DATE_RE.search(value or "")
    if not match:
        return ""
    raw = match.group(1)
    if re.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}$", raw):
        return normalize_date(raw)
    parts = re.split(r"[-/]", raw)
    if len(parts) != 3:
        return ""
    first, second, year_text = parts
    try:
        first_num = int(first)
        second_num = int(second)
        year_num = int(year_text)
    except ValueError:
        return ""
    if year_num < 100:
        year_num += 2000 if year_num < 70 else 1900
    if first_num > 12:
        return _date_from_parts(year_num, second_num, first_num)
    if second_num > 12:
        return _date_from_parts(year_num, first_num, second_num)
    if _prefer_us_month_first(currency, context):
        return _date_from_parts(year_num, first_num, second_num)
    return _date_from_parts(year_num, second_num, first_num)


def _prefer_us_month_first(currency: str, context: str) -> bool:
    normalized_currency = normalize_text(currency).casefold()
    normalized_context = normalize_text(context).casefold()
    if normalized_currency in {"mxn", "m.n.", "mn", "peso", "pesos"}:
        return False
    spanish_or_mexico = (
        "fecha",
        "mexico",
        "méxico",
        "m.n.",
        "rfc",
        "iva",
        "propina",
        "factura",
        "ticket",
    )
    if any(keyword in normalized_context for keyword in spanish_or_mexico):
        return False
    us_markers = (
        " usd",
        "dollar",
        "sales tax",
        "subtotal",
        "united states",
        " usa",
        "walmart",
        "target",
        "mcdonald",
        "trader joe",
        "whole foods",
        "starbucks",
    )
    return normalized_currency == "usd" or any(keyword in f" {normalized_context}" for keyword in us_markers)


def _date_from_parts(year: int, month: int, day: int) -> str:
    try:
        return datetime(year, month, day).date().isoformat()
    except ValueError:
        return ""


def parse_amount(value: str) -> float:
    if not value:
        return 0.0
    cleaned = _amount_fragment(value)
    cleaned = re.sub(r"[\s\u00a0]+", "", cleaned)
    cleaned = _normalize_amount_separators(cleaned)
    try:
        return round(float(cleaned), 2)
    except ValueError:
        return 0.0


def _amount_fragment(value: str) -> str:
    cleaned = re.sub(r"(?i)(?:MXN|M\.N\.|MN)", "", value.strip())
    cleaned = cleaned.replace("$", "")
    if re.search(r"[A-Za-z]", cleaned):
        match = re.search(r"\d[\d,.\s\u00a0]*", cleaned)
        return match.group(0) if match else cleaned
    return cleaned


def _normalize_amount_separators(value: str) -> str:
    separators = [index for index, char in enumerate(value) if char in ",."]
    if not separators:
        return value
    last = separators[-1]
    decimals = len(value) - last - 1
    if decimals == 2:
        integer = re.sub(r"[,.]", "", value[:last])
        fraction = re.sub(r"\D", "", value[last + 1 :])
        return f"{integer}.{fraction}"
    if decimals == 3 and len(separators) == 1:
        return value.replace(",", "").replace(".", "")
    if decimals == 3 and all(len(part) == 3 for part in re.split(r"[,.]", value)[1:]):
        return value.replace(",", "").replace(".", "")
    return value.replace(",", "")


def extract_amounts(text: str) -> list[float]:
    return [parse_amount(match.group(1)) for match in AMOUNT_RE.finditer(text or "") if parse_amount(match.group(1)) > 0]


def fuzzy_match(left: str, right: str, threshold: float = 0.78) -> bool:
    a = normalize_text(left).casefold()
    b = normalize_text(right).casefold()
    if not a or not b:
        return False
    return a == b or SequenceMatcher(None, a, b).ratio() >= threshold


def parse_invoice_from_lines(lines: list[OCRTextLine], engine: str = "local") -> InvoiceRecord:
    texts = [line.text.strip() for line in lines if line.text and line.text.strip()]
    joined = "\n".join(texts)
    lowered = normalize_text(joined).casefold()

    invoice_date = ""
    for text in texts:
        invoice_date = normalize_date(text)
        if invoice_date:
            break

    total_amount = _amount_near_keywords(texts, TOTAL_KEYWORDS)
    if total_amount <= 0:
        amounts = _fallback_amounts(texts)
        total_amount = max(amounts) if amounts else 0.0

    vat_amount = _amount_near_keywords(texts, VAT_KEYWORDS)
    sales_tax = _amount_near_keywords(texts, SALES_TAX_KEYWORDS)
    tips = _amount_near_keywords(texts, TIP_KEYWORDS)
    expense_amount = max(total_amount - vat_amount - sales_tax, 0.0)

    seller = _guess_seller(texts)
    contents = _guess_contents(texts)
    if "usd" in lowered or "dolar" in lowered:
        currency = "USD"
    else:
        currency = "MXN"

    confidence = _average_confidence(lines)
    remarks = f"{engine} parsed"
    if total_amount <= 0 or seller == "Unknown":
        remarks = f"{remarks}; incomplete"

    return InvoiceRecord(
        invoice_date=invoice_date,
        expense_category=normalize_expense_category("", joined),
        contents=contents,
        currency=currency,
        total_amount=total_amount,
        expense_amount=expense_amount,
        vat_amount=vat_amount,
        sales_tax=sales_tax,
        tips=tips,
        seller=seller,
        remarks=remarks,
    )


def _average_confidence(lines: list[OCRTextLine]) -> float:
    values = [line.confidence for line in lines if line.confidence > 0]
    return round(sum(values) / len(values), 3) if values else 0.0


def _amount_near_keywords(texts: list[str], keywords: tuple[str, ...]) -> float:
    candidates: list[float] = []
    for index, text in enumerate(texts):
        normalized = normalize_text(text).casefold()
        if any(keyword in normalized for keyword in keywords):
            same_line = extract_amounts(text)
            if same_line:
                candidates.extend(same_line)
                continue
            window = " ".join(texts[index : min(len(texts), index + 2)])
            candidates.extend(extract_amounts(window))
    return max(candidates) if candidates else 0.0


def _fallback_amounts(texts: list[str]) -> list[float]:
    amounts: list[float] = []
    for text in texts:
        normalized = normalize_text(text).casefold()
        if normalize_date(text) or any(keyword in normalized for keyword in SKIP_FALLBACK_AMOUNT_KEYWORDS):
            continue
        amounts.extend(extract_amounts(text))
    return amounts


def _guess_seller(texts: list[str]) -> str:
    for text in texts[:8]:
        normalized = normalize_text(text).casefold()
        if len(text) < 3:
            continue
        if any(keyword in normalized for keyword in SKIP_MERCHANT_KEYWORDS):
            continue
        if extract_amounts(text) or normalize_date(text):
            continue
        letters = sum(ch.isalpha() for ch in text)
        if letters >= 3:
            return text[:80]
    return "Unknown"


def _guess_contents(texts: list[str]) -> str:
    items: list[str] = []
    for text in texts:
        normalized = normalize_text(text).casefold()
        if any(keyword in normalized for keyword in SKIP_MERCHANT_KEYWORDS):
            continue
        if extract_amounts(text) or normalize_date(text):
            continue
        if 3 <= len(text) <= 80:
            items.append(text)
        if len(items) >= 3:
            break
    return "; ".join(items)
