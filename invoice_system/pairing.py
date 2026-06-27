from __future__ import annotations

from difflib import SequenceMatcher

from .models import InvoiceRecord
from .parsing import fuzzy_match, normalize_date, normalize_text

PAYMENT_HINT_KEYWORDS = (
    "tip",
    "tips",
    "propina",
    "gratuity",
    "card",
    "tarjeta",
    "payment",
    "pago",
    "visa",
    "mastercard",
    "amex",
)


def pair_invoice_payment_slips(records: list[InvoiceRecord], mode: str = "auto") -> list[InvoiceRecord]:
    """Merge likely invoice/payment-slip pairs and calculate tip deltas.

    In auto mode the rule keeps the tax-bearing invoice row, merges likely
    payment slips, calculates tip deltas, and removes payment-slip duplicates.
    In review mode it keeps both records, marks possible pairs, and orders the
    pair together for human review.
    """
    if mode == "review":
        return _suggest_pairs_for_review(records)

    paired_payment_indexes: set[int] = set()
    for index, record in enumerate(records):
        if not _looks_like_invoice(record):
            continue
        payment_index, payment = _find_payment_slip(index, record, records, paired_payment_indexes)
        if payment is None:
            continue

        paired_payment_indexes.add(payment_index)
        _merge_payment_slip(record, payment)

    output = [record for index, record in enumerate(records) if index not in paired_payment_indexes]
    return _deduplicate_invoices(output)


def _suggest_pairs_for_review(records: list[InvoiceRecord]) -> list[InvoiceRecord]:
    paired_indexes: set[int] = set()
    groups: list[tuple[int, int]] = []
    group_number = 1
    for invoice_index, invoice in enumerate(records):
        if not _looks_like_invoice(invoice):
            continue
        payment_index, payment = _find_payment_slip(invoice_index, invoice, records, paired_indexes)
        if payment is None:
            continue
        paired_indexes.add(invoice_index)
        paired_indexes.add(payment_index)
        group_id = f"PAIR-{group_number:03d}"
        _mark_possible_pair(invoice, payment, group_id)
        groups.append((invoice_index, payment_index))
        group_number += 1

    output: list[InvoiceRecord] = []
    emitted: set[int] = set()
    for invoice_index, payment_index in groups:
        output.extend([records[invoice_index], records[payment_index]])
        emitted.update({invoice_index, payment_index})
    for index, record in enumerate(records):
        if index not in emitted:
            output.append(record)
    return output


def _find_payment_slip(
    invoice_index: int,
    invoice: InvoiceRecord,
    records: list[InvoiceRecord],
    used: set[int],
) -> tuple[int, InvoiceRecord | None]:
    best_index = -1
    best_record: InvoiceRecord | None = None
    best_delta = float("inf")
    for index, candidate in enumerate(records):
        if index == invoice_index or index in used:
            continue
        if not _same_context(invoice, candidate):
            continue
        if not _looks_like_payment_slip(candidate):
            continue
        delta = abs(candidate.total_amount - invoice.total_amount)
        if _confident_payment_pair(invoice, candidate) and delta < best_delta:
            best_index = index
            best_record = candidate
            best_delta = delta
    return best_index, best_record


def _merge_payment_slip(invoice: InvoiceRecord, payment: InvoiceRecord) -> None:
    amount_delta = round(payment.total_amount - invoice.total_amount, 2)
    _append_supporting_crop(invoice, payment.crop_image)
    for crop_image in getattr(payment, "supporting_crop_images", []) or []:
        _append_supporting_crop(invoice, crop_image)
    if amount_delta > 0.50:
        invoice.tips = max(invoice.tips, amount_delta)
        invoice.total_amount = payment.total_amount
        invoice.expense_amount = max(invoice.total_amount - invoice.vat_amount - invoice.sales_tax, 0.0)
        _append_remark(invoice, f"Combined payment slip; tips calculated as {amount_delta:.2f}; supporting crop kept")
    else:
        _append_remark(invoice, "Combined duplicate payment slip; supporting crop kept")


def _mark_possible_pair(invoice: InvoiceRecord, payment: InvoiceRecord, group_id: str) -> None:
    amount_delta = round(payment.total_amount - invoice.total_amount, 2)
    if amount_delta > 0.50:
        message = f"Possible pair {group_id}: payment slip may include tips {amount_delta:.2f}; human review required"
    else:
        message = f"Possible pair {group_id}: possible duplicate payment slip; human review required"
    _append_remark(invoice, message)
    _append_remark(payment, message)


def _same_context(left: InvoiceRecord, right: InvoiceRecord) -> bool:
    return (
        _same_date(left.invoice_date, right.invoice_date)
        and _same_currency(left.currency, right.currency)
        and _merchant_score(left.seller, right.seller) >= 0.80
    )


def _same_date(left: str, right: str) -> bool:
    normalized_left = normalize_date(left) or (left or "").strip()[:10]
    normalized_right = normalize_date(right) or (right or "").strip()[:10]
    return bool(normalized_left) and normalized_left == normalized_right


def _same_currency(left: str, right: str) -> bool:
    return _normalize_currency(left) == _normalize_currency(right)


def _normalize_currency(value: str) -> str:
    normalized = normalize_text(value or "").casefold()
    if normalized in {"m.n.", "mn", "peso", "pesos"}:
        return "MXN"
    return normalized.upper()


def _looks_like_invoice(record: InvoiceRecord) -> bool:
    return record.total_amount > 0 and record.vat_amount > 0


def _looks_like_payment_slip(record: InvoiceRecord) -> bool:
    return record.total_amount > 0 and record.vat_amount <= 0 and record.sales_tax <= 0


def _confident_payment_pair(invoice: InvoiceRecord, payment: InvoiceRecord) -> bool:
    if invoice.total_amount <= 0 or payment.total_amount <= 0:
        return False
    delta = round(payment.total_amount - invoice.total_amount, 2)
    if abs(delta) <= 0.50:
        return True
    if not _amounts_pair(invoice.total_amount, payment.total_amount):
        return False
    return _has_payment_hint(invoice) or _has_payment_hint(payment)


def _amounts_pair(invoice_total: float, payment_total: float) -> bool:
    if invoice_total <= 0 or payment_total <= 0:
        return False
    delta = round(payment_total - invoice_total, 2)
    if abs(delta) <= 0.50:
        return True
    max_tip = max(50.0, invoice_total * 0.35)
    return 0.50 < delta <= max_tip


def _has_payment_hint(record: InvoiceRecord) -> bool:
    text = normalize_text(" ".join([record.contents, record.remarks, record.seller])).casefold()
    return any(keyword in text for keyword in PAYMENT_HINT_KEYWORDS)


def _merchant_score(left: str, right: str) -> float:
    a = normalize_text(left).casefold()
    b = normalize_text(right).casefold()
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    return SequenceMatcher(None, a, b).ratio()


def _append_remark(record: InvoiceRecord, remark: str) -> None:
    record.remarks = f"{record.remarks}; {remark}" if record.remarks else remark


def _append_supporting_crop(record: InvoiceRecord, crop_image: str) -> None:
    crop_text = str(crop_image or "").strip()
    if not crop_text or crop_text == record.crop_image:
        return
    if crop_text not in record.supporting_crop_images:
        record.supporting_crop_images.append(crop_text)


def _deduplicate_invoices(records: list[InvoiceRecord]) -> list[InvoiceRecord]:
    seen: dict[tuple[object, ...], InvoiceRecord] = {}
    output: list[InvoiceRecord] = []
    for record in records:
        key = _duplicate_key(record)
        existing = seen.get(key)
        if existing is not None and _contents_match_for_duplicate(existing, record):
            _append_remark(existing, "Duplicate invoice photo removed")
            _append_supporting_crop(existing, record.crop_image)
            for crop_image in getattr(record, "supporting_crop_images", []) or []:
                _append_supporting_crop(existing, crop_image)
            continue
        seen[key] = record
        output.append(record)
    return output


def _duplicate_key(record: InvoiceRecord) -> tuple[object, ...]:
    return (
        normalize_date(record.invoice_date) or (record.invoice_date or "").strip()[:10],
        normalize_text(record.seller).casefold(),
        _normalize_currency(record.currency),
        round(record.total_amount, 2),
        round(record.vat_amount, 2),
        round(record.sales_tax, 2),
        round(record.tips, 2),
    )


def _contents_match_for_duplicate(left: InvoiceRecord, right: InvoiceRecord) -> bool:
    left_contents = normalize_text(left.contents).casefold()
    right_contents = normalize_text(right.contents).casefold()
    if not left_contents or not right_contents:
        return False
    return left_contents == right_contents or fuzzy_match(left_contents, right_contents, threshold=0.88)
