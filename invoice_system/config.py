from __future__ import annotations

import os
import re
from dataclasses import dataclass
from pathlib import Path


def load_dotenv(path: Path = Path(".env")) -> None:
    if not path.exists():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def parse_allowed_user_ids(value: str | None) -> set[int]:
    if not value:
        return set()
    return {int(match) for match in re.findall(r"\d+", value)}


def parse_csv(value: str | None, default: tuple[str, ...]) -> tuple[str, ...]:
    if not value:
        return default
    items = tuple(part.strip() for part in value.replace(";", ",").split(",") if part.strip())
    return items or default


def parse_bool(value: str | None, default: bool = False) -> bool:
    if value is None or not value.strip():
        return default
    normalized = value.strip().casefold()
    if normalized in {"1", "true", "yes", "y", "on"}:
        return True
    if normalized in {"0", "false", "no", "n", "off"}:
        return False
    return default


def parse_float(value: str | None, default: float) -> float:
    if value is None or not value.strip():
        return default
    try:
        return float(value.strip())
    except ValueError:
        return default


def parse_int(value: str | None, default: int) -> int:
    if value is None or not value.strip():
        return default
    try:
        return int(value.strip())
    except ValueError:
        return default


def parse_choice(value: str | None, default: str, choices: set[str]) -> str:
    if value is None or not value.strip():
        return default
    normalized = value.strip().casefold()
    return normalized if normalized in choices else default


@dataclass(frozen=True)
class Settings:
    root: Path = Path(".")
    inbound_dir: Path = Path("data/inbound")
    trial_dir: Path = Path("data/trial")
    output_dir: Path = Path("data/output")
    baseline_dir: Path = Path("data/baseline")
    openai_api_key: str | None = None
    openai_model: str = "gpt-4.1"
    codex_scan_enabled: bool = False
    qwen_api_key: str | None = None
    qwen_model: str = "qwen-vl-max"
    qwen_base_url: str = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
    qwen_scan_enabled: bool = False
    telegram_bot_token: str | None = None
    telegram_allowed_user_ids: frozenset[int] = frozenset()
    telegram_auto_process: bool = False
    telegram_language: str = "en"
    ai_visual_count_enabled: bool = False
    ai_visual_count_min_opencv_crops: int = 4
    pairing_mode: str = "auto"
    local_confidence_threshold: float = 0.62
    amount_tolerance: float = 0.50
    paddleocr_lang: str = "en"
    easyocr_langs: tuple[str, ...] = ("es", "en")
    tesseract_cmd: str = "tesseract"
    tesseract_lang: str = "eng+spa"
    tesseract_psm: str = "6"

    @classmethod
    def from_env(cls, root: Path | None = None) -> "Settings":
        load_dotenv()
        base = (root or Path(".")).resolve()
        allowed = parse_allowed_user_ids(os.getenv("TELEGRAM_ALLOWED_USER_IDS"))
        return cls(
            root=base,
            inbound_dir=base / "data" / "inbound",
            trial_dir=base / "data" / "trial",
            output_dir=base / "data" / "output",
            baseline_dir=base / "data" / "baseline",
            openai_api_key=os.getenv("OPENAI_API_KEY"),
            openai_model=os.getenv("OPENAI_MODEL", "gpt-4.1"),
            codex_scan_enabled=parse_bool(os.getenv("ENABLE_CODEX_SCAN"), False),
            qwen_api_key=os.getenv("QWEN_API_KEY") or os.getenv("DASHSCOPE_API_KEY"),
            qwen_model=os.getenv("QWEN_MODEL", "qwen-vl-max"),
            qwen_base_url=os.getenv("QWEN_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"),
            qwen_scan_enabled=parse_bool(os.getenv("ENABLE_QWEN_SCAN"), False),
            telegram_bot_token=os.getenv("TELEGRAM_BOT_TOKEN"),
            telegram_allowed_user_ids=frozenset(allowed),
            telegram_auto_process=parse_bool(os.getenv("TELEGRAM_AUTO_PROCESS"), False),
            telegram_language=parse_choice(os.getenv("TELEGRAM_LANGUAGE"), "en", {"en", "zh"}),
            ai_visual_count_enabled=parse_bool(os.getenv("ENABLE_AI_VISUAL_COUNT"), False),
            ai_visual_count_min_opencv_crops=parse_int(os.getenv("AI_VISUAL_COUNT_MIN_OPENCV_CROPS"), 4),
            pairing_mode=parse_choice(os.getenv("PAIRING_MODE"), "auto", {"auto", "review"}),
            local_confidence_threshold=parse_float(os.getenv("LOCAL_OCR_CONFIDENCE_THRESHOLD"), 0.62),
            amount_tolerance=parse_float(os.getenv("AMOUNT_TOLERANCE_MXN"), 0.50),
            paddleocr_lang=os.getenv("PADDLEOCR_LANG", "en"),
            easyocr_langs=parse_csv(os.getenv("EASYOCR_LANGS"), ("es", "en")),
            tesseract_cmd=os.getenv("TESSERACT_CMD", "tesseract"),
            tesseract_lang=os.getenv("TESSERACT_LANG", "eng+spa"),
            tesseract_psm=os.getenv("TESSERACT_PSM", "6"),
        )
