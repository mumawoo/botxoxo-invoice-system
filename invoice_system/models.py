from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from .expense_categories import DEFAULT_EXPENSE_CATEGORY


EXCEL_HEADERS = [
    "Line No.",
    "Invoice Date",
    "Expense Category",
    "Contents",
    "Currency",
    "Total Amount",
    "Expense Amount",
    "VAT Amount",
    "Sales Tax",
    "Tips",
    "Seller / Service Provider",
    "Remarks",
    "Status",
]

OCR_AUDIT_HEADERS = [
    "Source Image",
    "Crop Image",
    "Decision",
    "Used Codex Scan",
    "Paddle Confidence",
    "Easy Confidence",
    "Codex Confidence",
    "Paddle Error",
    "Easy Error",
    "Codex Error",
    "Paddle Text",
    "Easy Text",
    "Codex Text",
]

SOURCE_QA_HEADERS = [
    "Source Image",
    "AI Visual Count",
    "OpenCV Crop Count",
    "Final Invoice Rows",
    "Needs Human Review",
    "Reason",
    "AI Confidence",
    "AI Error",
]

IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}


@dataclass(frozen=True)
class OCRTextLine:
    text: str
    confidence: float = 0.0


@dataclass
class InvoiceRecord:
    line_no: int | None = None
    invoice_date: str = ""
    expense_category: str = DEFAULT_EXPENSE_CATEGORY
    contents: str = ""
    currency: str = "MXN"
    total_amount: float = 0.0
    expense_amount: float = 0.0
    vat_amount: float = 0.0
    sales_tax: float = 0.0
    tips: float = 0.0
    seller: str = "Unknown"
    remarks: str = ""
    status: str = ""
    source_image: str = ""
    crop_image: str = ""
    supporting_crop_images: list[str] = field(default_factory=list)
    report_components: bool = False

    def to_excel_row(self) -> list[object]:
        return [
            self.line_no,
            self.invoice_date,
            self.expense_category,
            self.contents,
            self.currency,
            round(self.total_amount, 2),
            round(self.expense_amount, 2),
            round(self.vat_amount, 2),
            round(self.sales_tax, 2),
            round(self.tips, 2),
            self.seller,
            self.remarks,
            self.status,
        ]

    @classmethod
    def from_excel_row(cls, values: list[object]) -> "InvoiceRecord":
        padded = list(values[: len(EXCEL_HEADERS)]) + [None] * len(EXCEL_HEADERS)
        return cls(
            line_no=_to_int(padded[0]),
            invoice_date=str(padded[1] or ""),
            expense_category=str(padded[2] or DEFAULT_EXPENSE_CATEGORY),
            contents=str(padded[3] or ""),
            currency=str(padded[4] or "MXN"),
            total_amount=_to_float(padded[5]),
            expense_amount=_to_float(padded[6]),
            vat_amount=_to_float(padded[7]),
            sales_tax=_to_float(padded[8]),
            tips=_to_float(padded[9]),
            seller=str(padded[10] or "Unknown"),
            remarks=str(padded[11] or ""),
            status=str(padded[12] or ""),
        )

    def stable_key(self) -> tuple[str, str, float]:
        return (self.invoice_date.strip(), self.seller.strip().casefold(), round(self.total_amount, 2))


@dataclass
class OCRResult:
    engine: str
    lines: list[OCRTextLine] = field(default_factory=list)
    parsed_invoice: InvoiceRecord | None = None
    confidence: float = 0.0
    error: str = ""
    rotate_degrees: int = 0
    orientation_confidence: float = 0.0

    @property
    def text(self) -> str:
        return "\n".join(line.text for line in self.lines if line.text)


@dataclass(frozen=True)
class CropResult:
    source_path: Path
    crop_path: Path
    index: int


@dataclass(frozen=True)
class PipelineSummary:
    source_images: int
    crops: int
    records_written: int
    workbook_path: Path
    review_warnings: tuple[str, ...] = ()


@dataclass
class OCRAuditRow:
    source_image: str
    crop_image: str
    decision: str
    used_codex: bool
    paddle_confidence: float = 0.0
    easy_confidence: float = 0.0
    codex_confidence: float = 0.0
    paddle_error: str = ""
    easy_error: str = ""
    codex_error: str = ""
    paddle_text: str = ""
    easy_text: str = ""
    codex_text: str = ""

    def to_excel_row(self) -> list[object]:
        return [
            self.source_image,
            self.crop_image,
            self.decision,
            "yes" if self.used_codex else "no",
            round(self.paddle_confidence, 3),
            round(self.easy_confidence, 3),
            round(self.codex_confidence, 3),
            self.paddle_error,
            self.easy_error,
            self.codex_error,
            _truncate_cell(self.paddle_text),
            _truncate_cell(self.easy_text),
            _truncate_cell(self.codex_text),
        ]


def _to_float(value: object) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _to_int(value: object) -> int | None:
    try:
        return int(value) if value is not None and value != "" else None
    except (TypeError, ValueError):
        return None


def _truncate_cell(value: str, limit: int = 32000) -> str:
    return value if len(value) <= limit else value[: limit - 20] + "...[truncated]"


@dataclass
class SourceQARecord:
    source_image: str
    ai_visual_count: int | None
    opencv_crop_count: int
    final_invoice_rows: int
    needs_human_review: bool
    reason: str = ""
    ai_confidence: float = 0.0
    ai_error: str = ""

    def to_excel_row(self) -> list[object]:
        return [
            self.source_image,
            self.ai_visual_count if self.ai_visual_count is not None else "",
            self.opencv_crop_count,
            self.final_invoice_rows,
            "yes" if self.needs_human_review else "no",
            self.reason,
            round(self.ai_confidence, 3),
            self.ai_error,
        ]
