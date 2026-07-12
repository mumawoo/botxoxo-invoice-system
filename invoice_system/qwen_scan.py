from __future__ import annotations

import json
import urllib.error
import urllib.request
from io import BytesIO
from pathlib import Path

from PIL import Image

from .codex_scan import PROMPT, _mime_type, _ocr_result_from_response_text
from .config import Settings
from .models import OCRResult


QWEN_DEFAULT_MAX_EDGE = 1600
QWEN_LONG_RECEIPT_MAX_EDGE = 2000
QWEN_LONG_RECEIPT_ASPECT_RATIO = 2.2
QWEN_IMAGE_MAX_BYTES = 900 * 1024
QWEN_JPEG_QUALITIES = (88, 82, 76, 70)


class QwenScanRecognizer:
    engine = "qwen_scan"

    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def recognize(self, image_path: Path) -> OCRResult:
        if not self.settings.qwen_api_key:
            return OCRResult(self.engine, error="QWEN_API_KEY is not configured")
        try:
            image_url = _qwen_image_url(image_path)
            payload = {
                "model": self.settings.qwen_model,
                "messages": [
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": PROMPT},
                            {"type": "image_url", "image_url": {"url": image_url}},
                        ],
                    }
                ],
                "temperature": 0,
                "enable_thinking": False,
            }
            request = urllib.request.Request(
                self.settings.qwen_base_url,
                data=json.dumps(payload).encode("utf-8"),
                headers={
                    "Authorization": f"Bearer {self.settings.qwen_api_key}",
                    "Content-Type": "application/json",
                },
                method="POST",
            )
            with urllib.request.urlopen(request, timeout=120) as response:
                body = response.read().decode("utf-8")
            text = _extract_message_content(json.loads(body))
            result = _ocr_result_from_response_text(text, self.engine)
            if result.error:
                result.error = result.error.replace("Codex Scan", "Qwen Scan")
            return result
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            return OCRResult(self.engine, error=f"Qwen HTTP {exc.code}: {detail}")
        except Exception as exc:
            return OCRResult(self.engine, error=str(exc))


def _extract_message_content(data: object) -> str:
    if not isinstance(data, dict):
        return ""
    choices = data.get("choices")
    if not isinstance(choices, list) or not choices:
        return ""
    first = choices[0]
    if not isinstance(first, dict):
        return ""
    message = first.get("message")
    if not isinstance(message, dict):
        return ""
    content = message.get("content")
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts: list[str] = []
        for item in content:
            if isinstance(item, dict) and isinstance(item.get("text"), str):
                parts.append(item["text"])
        return "\n".join(parts)
    return ""


def _qwen_image_url(path: Path) -> str:
    import base64

    return f"data:image/jpeg;base64,{base64.b64encode(_qwen_image_bytes(path)).decode('ascii')}"


def _qwen_image_bytes(path: Path) -> bytes:
    with Image.open(path) as image:
        image = image.convert("RGB")
        width, height = image.size
        short_edge = max(min(width, height), 1)
        aspect_ratio = max(width, height) / short_edge
        max_edge = QWEN_LONG_RECEIPT_MAX_EDGE if aspect_ratio >= QWEN_LONG_RECEIPT_ASPECT_RATIO else QWEN_DEFAULT_MAX_EDGE
        image.thumbnail((max_edge, max_edge), Image.Resampling.LANCZOS)

        while True:
            encoded = b""
            for quality in QWEN_JPEG_QUALITIES:
                buffer = BytesIO()
                image.save(buffer, format="JPEG", quality=quality, optimize=True)
                encoded = buffer.getvalue()
                if len(encoded) <= QWEN_IMAGE_MAX_BYTES:
                    return encoded
            longest = max(image.size)
            if longest <= 640:
                return encoded
            scale = max(640 / longest, 0.85)
            resized = (max(int(image.width * scale), 1), max(int(image.height * scale), 1))
            image = image.resize(resized, Image.Resampling.LANCZOS)
