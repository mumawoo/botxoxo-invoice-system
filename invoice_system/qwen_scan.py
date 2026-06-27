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

    with Image.open(path) as image:
        image = image.convert("RGB")
        image.thumbnail((1000, 1000))
        buffer = BytesIO()
        image.save(buffer, format="JPEG", quality=85, optimize=True)
    return f"data:image/jpeg;base64,{base64.b64encode(buffer.getvalue()).decode('ascii')}"
