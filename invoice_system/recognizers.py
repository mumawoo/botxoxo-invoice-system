from __future__ import annotations

import os
import subprocess
from pathlib import Path
from typing import Protocol

from .config import Settings
from .models import OCRResult, OCRTextLine
from .parsing import parse_invoice_from_lines


class Recognizer(Protocol):
    engine: str

    def recognize(self, image_path: Path) -> OCRResult:
        ...


class PaddleOCRRecognizer:
    engine = "paddleocr"

    def __init__(self, settings: Settings | None = None) -> None:
        self._ocr = None
        self.lang = (settings.paddleocr_lang if settings else "en") or "en"

    def recognize(self, image_path: Path) -> OCRResult:
        try:
            if self._ocr is None:
                os.environ.setdefault("PADDLE_PDX_ENABLE_MKLDNN_BYDEFAULT", "0")
                from paddleocr import PaddleOCR

                self._ocr = PaddleOCR(
                    lang=self.lang,
                    use_doc_orientation_classify=False,
                    use_doc_unwarping=False,
                    use_textline_orientation=False,
                )
            raw = self._predict(image_path)
            lines = _paddle_lines(raw)
            invoice = parse_invoice_from_lines(lines, self.engine)
            return OCRResult(self.engine, lines, invoice, _avg(lines))
        except Exception as exc:
            return OCRResult(self.engine, error=str(exc))

    def _predict(self, image_path: Path) -> object:
        try:
            return self._ocr.predict(str(image_path))
        except AttributeError:
            return self._ocr.ocr(str(image_path))


class EasyOCRRecognizer:
    engine = "easyocr"

    def __init__(self, settings: Settings | None = None) -> None:
        self._reader = None
        self.langs = list(settings.easyocr_langs if settings else ("es", "en"))

    def recognize(self, image_path: Path) -> OCRResult:
        try:
            if self._reader is None:
                import easyocr

                self._reader = easyocr.Reader(self.langs, gpu=False)
            raw = self._reader.readtext(str(image_path), detail=1, paragraph=False)
            lines = _easy_lines(raw)
            invoice = parse_invoice_from_lines(lines, self.engine)
            return OCRResult(self.engine, lines, invoice, _avg(lines))
        except Exception as exc:
            return OCRResult(self.engine, error=str(exc))


class TesseractRecognizer:
    engine = "tesseract"

    def __init__(self, settings: Settings | None = None) -> None:
        self.cmd = settings.tesseract_cmd if settings else "tesseract"
        self.lang = settings.tesseract_lang if settings else "eng+spa"
        self.psm = settings.tesseract_psm if settings else "6"

    def recognize(self, image_path: Path) -> OCRResult:
        try:
            completed = subprocess.run(
                [self.cmd, str(image_path), "stdout", "-l", self.lang, "--psm", self.psm, "tsv"],
                capture_output=True,
                text=True,
                timeout=120,
                check=False,
            )
            if completed.returncode != 0:
                return OCRResult(self.engine, error=(completed.stderr or completed.stdout).strip())
            lines = _tesseract_lines(completed.stdout)
            invoice = parse_invoice_from_lines(lines, self.engine)
            return OCRResult(self.engine, lines, invoice, _avg(lines))
        except Exception as exc:
            return OCRResult(self.engine, error=str(exc))


class StaticRecognizer:
    """Small test/helper recognizer."""

    def __init__(self, result: OCRResult) -> None:
        self.result = result
        self.engine = result.engine

    def recognize(self, image_path: Path) -> OCRResult:
        return self.result


def _easy_lines(raw: object) -> list[OCRTextLine]:
    lines: list[OCRTextLine] = []
    if not isinstance(raw, list):
        return lines
    for item in raw:
        if isinstance(item, (list, tuple)) and len(item) >= 3:
            lines.append(OCRTextLine(str(item[1]), _confidence(item[2])))
        elif isinstance(item, str):
            lines.append(OCRTextLine(item, 0.0))
    return lines


def _paddle_lines(raw: object) -> list[OCRTextLine]:
    lines: list[OCRTextLine] = []
    if raw is None:
        return lines
    for item in _flatten_paddle(raw):
        mapping = _as_mapping(item)
        if mapping:
            texts = mapping.get("rec_texts") or mapping.get("texts")
            scores = mapping.get("rec_scores") or mapping.get("scores") or []
            if isinstance(texts, list):
                for index, text in enumerate(texts):
                    score = scores[index] if isinstance(scores, list) and index < len(scores) else 0.0
                    lines.append(OCRTextLine(str(text), _confidence(score)))
                continue
            text = mapping.get("text") or mapping.get("rec_text") or mapping.get("transcription")
            score = mapping.get("confidence") or mapping.get("score") or mapping.get("rec_score")
            if text:
                lines.append(OCRTextLine(str(text), _confidence(score)))
                continue
        if isinstance(item, (list, tuple)) and len(item) >= 2:
            candidate = item[-1]
            if isinstance(candidate, (list, tuple)) and len(candidate) >= 2:
                lines.append(OCRTextLine(str(candidate[0]), _confidence(candidate[1])))
            elif isinstance(candidate, str):
                lines.append(OCRTextLine(candidate, 0.0))
    return lines


def _tesseract_lines(raw: str) -> list[OCRTextLine]:
    lines: list[OCRTextLine] = []
    rows = [row.split("\t") for row in (raw or "").splitlines() if row.strip()]
    if not rows:
        return lines
    header = rows[0]
    try:
        conf_index = header.index("conf")
        text_index = header.index("text")
    except ValueError:
        return lines
    for row in rows[1:]:
        if len(row) <= max(conf_index, text_index):
            continue
        text = row[text_index].strip()
        if not text:
            continue
        lines.append(OCRTextLine(text, _confidence(row[conf_index])))
    return lines


def _flatten_paddle(raw: object) -> list[object]:
    if isinstance(raw, list) and len(raw) == 1 and isinstance(raw[0], list):
        return raw[0]
    if isinstance(raw, list):
        return raw
    return [raw]


def _as_mapping(item: object) -> dict | None:
    if isinstance(item, dict):
        return item
    if hasattr(item, "json"):
        try:
            data = item.json
            if isinstance(data, dict):
                return data
        except Exception:
            pass
    try:
        return {key: item[key] for key in item.keys()} if hasattr(item, "keys") else None
    except Exception:
        return None


def _confidence(value: object) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return 0.0
    if number > 1:
        number = number / 100.0
    return max(0.0, min(1.0, number))


def _avg(lines: list[OCRTextLine]) -> float:
    values = [line.confidence for line in lines if line.confidence > 0]
    return round(sum(values) / len(values), 3) if values else 0.0
