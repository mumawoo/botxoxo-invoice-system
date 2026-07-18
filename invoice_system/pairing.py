from __future__ import annotations

import re
from difflib import SequenceMatcher
from pathlib import Path

from .models import InvoiceRecord
from .parsing import normalize_date, normalize_text

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

BANK_OR_PAYMENT_WORDS = {
    "afirme",
    "amex",
    "banamex",
    "banca",
    "banco",
    "bancomer",
    "banorte",
    "bbva",
    "card",
    "credito",
    "credit",
    "debit",
    "debito",
    "hsbc",
    "mastercard",
    "mercado",
    "mifel",
    "pago",
    "payment",
    "propina",
    "santander",
    "tarjeta",
    "visa",
}

MERCHANT_STOP_WORDS = BANK_OR_PAYMENT_WORDS | {
    "and",
    "at",
    "caja",
    "city",
    "con",
    "cv",
    "cumbres",
    "de",
    "del",
    "el",
    "en",
    "est",
    "estado",
    "la",
    "las",
    "los",
    "mex",
    "mexico",
    "monterrey",
    "mx",
    "nl",
    "para",
    "por",
    "restaurante",
    "restaurant",
    "sa",
    "sucursal",
    "the",
    "venta",
    "y",
}


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
    _mark_possible_duplicate_invoices(output)
    return output


def is_confident_invoice_payment_pair(invoice: InvoiceRecord, payment: InvoiceRecord) -> bool:
    """Return whether two records can be safely paired in that direction."""

    return (
        _looks_like_invoice(invoice)
        and _looks_like_payment_slip(payment)
        and _pair_context_matches(invoice, payment)
        and _confident_payment_pair(invoice, payment)
    )


def is_possible_duplicate(left: InvoiceRecord, right: InvoiceRecord) -> bool:
    """Return whether two invoice records require duplicate review."""

    return _possible_duplicate(left, right)


def mark_possible_matches_with_protected(
    records: list[InvoiceRecord], protected_records: list[InvoiceRecord]
) -> list[InvoiceRecord]:
    """Warn about protected matches without changing or absorbing either record."""

    for record in records:
        for protected in protected_records:
            protected_id = _record_trace_label(protected)
            if _same_context(protected, record):
                if _looks_like_invoice(protected) and _looks_like_payment_slip(record) and _confident_payment_pair(protected, record):
                    delta = round(record.total_amount - protected.total_amount, 2)
                    _append_remark(
                        record,
                        f"Possible pair with protected crop {protected_id}: same date, merchant and currency; "
                        f"payment difference {delta:.2f}; human review required",
                    )
                    break
                if _looks_like_invoice(record) and _looks_like_payment_slip(protected) and _confident_payment_pair(record, protected):
                    delta = round(protected.total_amount - record.total_amount, 2)
                    _append_remark(
                        record,
                        f"Possible pair with protected crop {protected_id}: same date, merchant and currency; "
                        f"payment difference {delta:.2f}; human review required",
                    )
                    break
            if is_possible_duplicate(protected, record):
                _append_remark(
                    record,
                    f"Possible duplicate with protected crop {protected_id}; human review required",
                )
                break
    return records


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
        if not _pair_context_matches(invoice, candidate):
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
    supporting_id = _crop_trace_id(payment.crop_image)
    supporting_note = f"supporting crop {supporting_id} kept" if supporting_id else "supporting crop kept"
    invoice.report_components = bool(
        invoice.report_components
        or payment.report_components
        or invoice.vat_amount
        or payment.vat_amount
        or invoice.tips
        or payment.tips
        or amount_delta > 0.50
    )
    _append_supporting_crop(invoice, payment.crop_image)
    for crop_image in getattr(payment, "supporting_crop_images", []) or []:
        _append_supporting_crop(invoice, crop_image)
    if amount_delta > 0.50:
        invoice.tips = max(invoice.tips, amount_delta)
        invoice.total_amount = payment.total_amount
        invoice.expense_amount = max(invoice.total_amount - invoice.vat_amount - invoice.sales_tax, 0.0)
        _append_remark(invoice, f"Combined payment slip; tips calculated as {amount_delta:.2f}; {supporting_note}")
    else:
        _append_remark(invoice, f"Combined duplicate payment slip; {supporting_note}")


def _crop_trace_id(crop_image: str) -> str:
    match = re.match(r"(\d{3,})", Path(crop_image).name)
    return match.group(1) if match else ""


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
        and _same_merchant_context(left, right)
    )


def _pair_context_matches(invoice: InvoiceRecord, payment: InvoiceRecord) -> bool:
    if not _same_date(invoice.invoice_date, payment.invoice_date) or not _same_currency(invoice.currency, payment.currency):
        return False
    return _same_merchant_context(invoice, payment) or _same_source_exact_tip(invoice, payment)


def _same_source_exact_tip(invoice: InvoiceRecord, payment: InvoiceRecord) -> bool:
    invoice_source = str(invoice.source_image or "").strip().casefold()
    payment_source = str(payment.source_image or "").strip().casefold()
    if not invoice_source or invoice_source != payment_source:
        return False
    delta = round(float(payment.total_amount or 0) - float(invoice.total_amount or 0), 2)
    reported_tip = round(float(payment.tips or 0), 2)
    return delta > 0.50 and reported_tip > 0 and abs(delta - reported_tip) <= 0.50


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
    if record.total_amount <= 0:
        return False
    if record.vat_amount > 0 or record.sales_tax > 0:
        return True
    return not _looks_like_strong_payment_slip(record) and bool(_merchant_tokens(record))


def _looks_like_payment_slip(record: InvoiceRecord) -> bool:
    return record.total_amount > 0 and record.vat_amount <= 0 and record.sales_tax <= 0 and _looks_like_strong_payment_slip(record)


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


def _looks_like_strong_payment_slip(record: InvoiceRecord) -> bool:
    if record.total_amount <= 0 or record.vat_amount > 0 or record.sales_tax > 0:
        return False
    text = normalize_text(" ".join([record.seller, record.contents, record.remarks])).casefold()
    return _has_payment_hint(record) or any(word in text for word in BANK_OR_PAYMENT_WORDS)


def _same_merchant_context(left: InvoiceRecord, right: InvoiceRecord) -> bool:
    if _merchant_score(left.seller, right.seller) >= 0.80:
        return True
    if _merchant_score(_merchant_clean_text(left), _merchant_clean_text(right)) >= 0.80:
        return True
    shared = _merchant_tokens(left) & _merchant_tokens(right)
    if len(shared) >= 2:
        return True
    return any(len(token) >= 6 for token in shared)


def _merchant_clean_text(record: InvoiceRecord) -> str:
    tokens = _merchant_tokens(record)
    return " ".join(sorted(tokens))


def _merchant_tokens(record: InvoiceRecord) -> set[str]:
    text = normalize_text(" ".join([record.seller, record.contents, record.remarks])).casefold()
    tokens = set(re.findall(r"[a-z0-9]+", text))
    return {token for token in tokens if len(token) >= 3 and token not in MERCHANT_STOP_WORDS}


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


def _mark_possible_duplicate_invoices(records: list[InvoiceRecord]) -> None:
    for index, record in enumerate(records):
        for existing in records[:index]:
            if not _possible_duplicate(existing, record):
                continue
            existing_id = _record_trace_label(existing)
            _append_remark(record, f"Possible duplicate with {existing_id}; human review required")
            break


def _possible_duplicate(left: InvoiceRecord, right: InvoiceRecord) -> bool:
    if _duplicate_core_key(left) != _duplicate_core_key(right):
        return False
    left_date = normalize_date(left.invoice_date) or (left.invoice_date or "").strip()[:10]
    right_date = normalize_date(right.invoice_date) or (right.invoice_date or "").strip()[:10]
    if left_date and right_date and left_date != right_date:
        return False
    left_id = _record_trace_key(left)
    right_id = _record_trace_key(right)
    return not (left_id and right_id and left_id == right_id)


def _duplicate_key(record: InvoiceRecord) -> tuple[object, ...]:
    return (
        normalize_date(record.invoice_date) or (record.invoice_date or "").strip()[:10],
        *_duplicate_core_key(record),
    )


def _duplicate_core_key(record: InvoiceRecord) -> tuple[object, ...]:
    return (
        normalize_text(record.seller).casefold(),
        _normalize_currency(record.currency),
        round(record.total_amount, 2),
    )


def _record_trace_label(record: InvoiceRecord) -> str:
    trace = _record_trace_key(record)
    return trace or "another crop"


def _record_trace_key(record: InvoiceRecord) -> str:
    trace = _trace_from_path(record.crop_image)
    if trace:
        return trace
    if record.line_no:
        return f"{record.line_no:03d}"
    return ""


def _trace_from_path(value: str) -> str:
    match = re.match(r"(\d{3,})[a-zA-Z]?_", Path(str(value or "")).name)
    return match.group(1) if match else ""
