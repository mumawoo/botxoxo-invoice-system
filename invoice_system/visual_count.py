from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

from .codex_scan import _extract_json_object
from .config import Settings


COUNT_PROMPT = """Count the separate visible receipt, invoice, and payment-slip papers in this photo.
Return strict JSON only with keys:
visible_document_count: integer,
confidence: number from 0 to 1,
reason: short string.
Count physically separate papers, even if some are payment slips or handwritten restaurant notes.
Do not include markdown."""


@dataclass(frozen=True)
class VisualCountResult:
    count: int | None
    confidence: float = 0.0
    reason: str = ""
    error: str = ""


class AIVisualCounter:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def count(self, image_path: Path) -> VisualCountResult:
        return VisualCountResult(None, reason="OpenAI visual count removed; OpenCV crop count is used")


def _parse_visual_count(text: str) -> VisualCountResult:
    try:
        data = json.loads(_extract_json_object(text))
        raw_count = data.get("visible_document_count")
        count = int(raw_count)
        if count < 0:
            raise ValueError("visible_document_count must be non-negative")
        return VisualCountResult(
            count=count,
            confidence=_float(data.get("confidence")),
            reason=str(data.get("reason") or ""),
        )
    except Exception as exc:
        return VisualCountResult(None, error=str(exc))


def _float(value: object) -> float:
    try:
        return max(0.0, min(float(value or 0), 1.0))
    except (TypeError, ValueError):
        return 0.0
