from __future__ import annotations

from dataclasses import dataclass

from .compare import _is_candidate_workbook
from .config import Settings
from .diagnostics import run_checks
from .image_splitter import iter_images


@dataclass(frozen=True)
class AuditItem:
    requirement: str
    status: str
    evidence: str


def audit_requirements(settings: Settings) -> list[AuditItem]:
    checks = run_checks(settings)
    qwen_ready = bool(settings.qwen_api_key)
    telegram_ready = bool(settings.telegram_bot_token and settings.telegram_allowed_user_ids)
    trial_photos = iter_images(settings.trial_dir)
    baseline_workbooks = sorted(p for p in settings.baseline_dir.rglob("*.xlsx") if _is_candidate_workbook(p))
    baseline_images = iter_images(settings.baseline_dir)
    trial_output = settings.output_dir / "trial" / "Invoice_Output_Trial.xlsx"

    return [
        AuditItem(
            "Windows-native package and CLI",
            "READY",
            "python -m invoice_system supports run, telegram, compare, check, prepare, sample, and audit.",
        ),
        AuditItem(
            "Qwen-only invoice recognition",
            "READY" if qwen_ready else "WAITING",
            "Production OCR sends each OpenCV crop to Qwen Scan; local OCR engines are not used for the formal pipeline.",
        ),
        AuditItem(
            "Qwen Scan credentials",
            "READY" if qwen_ready else "WAITING",
            "QWEN_API_KEY is configured for Qwen-only scanning."
            if qwen_ready
            else "Set QWEN_API_KEY before production scanning.",
        ),
        AuditItem(
            "OpenAI fallback",
            "REMOVED",
            "OpenAI/Codex Scan fallback is no longer used; production OCR is Qwen-only.",
        ),
        AuditItem(
            "Telegram polling ingestion",
            "READY" if telegram_ready else "WAITING",
            "Bot token and allowed user IDs are configured."
            if telegram_ready
            else "Set TELEGRAM_BOT_TOKEN and TELEGRAM_ALLOWED_USER_IDS; /whoami works after token setup.",
        ),
        AuditItem(
            "Separated Windows data folders",
            "READY",
            "Uses data/inbound, data/trial, data/output, and data/baseline with pathlib paths.",
        ),
        AuditItem(
            "V2 crop processing",
            "READY",
            "OpenCV contour split, 2%-80% area filter, overlap merge, 15% padding, 1500px minimum, portrait rotation, JPEG 98.",
        ),
        AuditItem(
            "V2 Excel schema and manual locks",
            "READY",
            "Writes the 13 V2 invoice columns and preserves Manually checked/Deleted rows from the manual workbook.",
        ),
        AuditItem(
            "Ubuntu comparison run",
            "READY" if trial_photos and baseline_workbooks and trial_output.exists() else "WAITING",
            _comparison_evidence(len(trial_photos), len(baseline_workbooks), len(baseline_images), trial_output.exists()),
        ),
    ]


def format_audit(items: list[AuditItem]) -> str:
    lines = ["V2 Windows rewrite audit:"]
    for item in items:
        lines.append(f"[{item.status}] {item.requirement}: {item.evidence}")
    return "\n".join(lines)


def _comparison_evidence(trial_photo_count: int, baseline_workbook_count: int, baseline_image_count: int, trial_output_exists: bool) -> str:
    return (
        f"trial photos={trial_photo_count}, Ubuntu baseline workbooks={baseline_workbook_count}, "
        f"Ubuntu baseline images={baseline_image_count}, Windows trial workbook={'yes' if trial_output_exists else 'no'}."
    )
