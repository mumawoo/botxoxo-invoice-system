from __future__ import annotations

from pathlib import Path

from .models import OCRResult, OCRTextLine
from .parsing import parse_invoice_from_lines


class PaddleOCRVLRecognizer:
    """PaddleOCR-VL document parser adapted to the recognizer protocol.

    PaddleOCR-VL returns structured document output rather than the invoice JSON
    schema used by Qwen Scan. For A/B testing we extract any readable text from
    the structured result and feed it through the same local invoice parser used
    by the other OCR engines.
    """

    engine = "paddleocr_vl"

    def __init__(self) -> None:
        self._pipeline = None

    def recognize(self, image_path: Path) -> OCRResult:
        try:
            if self._pipeline is None:
                from paddleocr import PaddleOCRVL

                self._pipeline = PaddleOCRVL(device="cpu")
            output = self._pipeline.predict(str(image_path))
            text = _paddle_vl_text(output)
            lines = [OCRTextLine(line, 1.0) for line in text.splitlines() if line.strip()]
            invoice = parse_invoice_from_lines(lines, self.engine) if lines else None
            return OCRResult(self.engine, lines, invoice, 1.0 if lines else 0.0)
        except Exception as exc:
            return OCRResult(self.engine, error=str(exc))


def _paddle_vl_text(output: object) -> str:
    parts: list[str] = []
    _collect_text(output, parts)
    seen: set[str] = set()
    unique: list[str] = []
    for part in parts:
        cleaned = part.strip()
        if not cleaned or cleaned in seen:
            continue
        seen.add(cleaned)
        unique.append(cleaned)
    return "\n".join(unique)


def _collect_text(value: object, parts: list[str]) -> None:
    if value is None:
        return
    if isinstance(value, str):
        parts.append(value)
        return
    if isinstance(value, (int, float, bool)):
        return
    if isinstance(value, dict):
        priority_keys = (
            "markdown",
            "text",
            "content",
            "html",
            "rec_text",
            "transcription",
            "label",
            "words",
        )
        for key in priority_keys:
            if key in value:
                _collect_text(value[key], parts)
        for key, item in value.items():
            if key not in priority_keys:
                _collect_text(item, parts)
        return
    if isinstance(value, (list, tuple, set)):
        for item in value:
            _collect_text(item, parts)
        return
    for attr in ("json", "dict", "markdown", "text", "content"):
        if hasattr(value, attr):
            try:
                item = getattr(value, attr)
                item = item() if callable(item) else item
            except Exception:
                continue
            _collect_text(item, parts)
            return
    try:
        text = str(value)
    except Exception:
        return
    if text and not text.startswith("<"):
        parts.append(text)
