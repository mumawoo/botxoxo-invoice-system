from __future__ import annotations

import importlib.util
import importlib
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path

from .excel_store import load_invoice_records
from .compare import _is_candidate_workbook
from .config import Settings
from .image_splitter import iter_images


@dataclass(frozen=True)
class CheckItem:
    name: str
    ok: bool
    detail: str
    required: bool = True


def run_checks(settings: Settings) -> list[CheckItem]:
    checks = [
        _module_check("openpyxl", "Excel read/write"),
        _module_check("cv2", "OpenCV image splitting"),
        CheckItem("QWEN_API_KEY", bool(settings.qwen_api_key), "required for Qwen-only production OCR"),
        _module_check("PIL", "Pillow sample image generator", required=False),
        _module_check("openai", "Codex Scan compatibility package", required=False),
        _module_check("telegram", "Telegram polling bot package", required=False),
        _module_check("torch", "Optional EasyOCR/PaddleOCR-VL runtime for A/B experiments", import_module=True, required=False),
        _module_check("easyocr", f"Optional EasyOCR local OCR, langs={','.join(settings.easyocr_langs)}", import_module=True, required=False),
        _module_check("paddle", "Optional Paddle runtime for A/B experiments", import_module=True, required=False),
        _module_check("paddleocr", f"Optional PaddleOCR local OCR, lang={settings.paddleocr_lang}", import_module=True, required=False),
        CheckItem("tesseract", bool(shutil.which(settings.tesseract_cmd)), f"Optional Tesseract OCR command: {settings.tesseract_cmd}", required=False),
        CheckItem("OPENAI_API_KEY", bool(settings.openai_api_key), "needed only if Codex Scan fallback is enabled", required=False),
        CheckItem("TELEGRAM_BOT_TOKEN", bool(settings.telegram_bot_token), "needed only for telegram command", required=False),
        CheckItem("TELEGRAM_ALLOWED_USER_IDS", bool(settings.telegram_allowed_user_ids), "needed before Telegram accepts photos", required=False),
    ]
    checks.extend(_folder_checks(settings))
    return checks


def format_checks(checks: list[CheckItem]) -> str:
    required_ok = all(item.ok for item in checks if item.required)
    optional_missing = sum(1 for item in checks if not item.required and not item.ok)
    lines: list[str] = [
        f"Core pipeline: {'READY' if required_ok else 'NOT READY'}",
        f"Optional integrations missing: {optional_missing}",
        "",
    ]
    for item in checks:
        marker = "OK" if item.ok else "MISSING"
        scope = "required" if item.required else "optional"
        lines.append(f"[{marker}] {item.name} ({scope}): {item.detail}")
    return "\n".join(lines)


def prepare_handoff(settings: Settings) -> str:
    for folder in (settings.inbound_dir, settings.trial_dir, settings.output_dir, settings.baseline_dir):
        folder.mkdir(parents=True, exist_ok=True)
    _write_handoff_notes(settings)

    trial_photos = iter_images(settings.trial_dir)
    inbound_photos = iter_images(settings.inbound_dir)
    baseline_workbooks = sorted(p for p in settings.baseline_dir.rglob("*.xlsx") if _is_candidate_workbook(p))
    baseline_images = iter_images(settings.baseline_dir)
    trial_workbook = settings.output_dir / "trial" / "Invoice_Output_Trial.xlsx"
    trial_output_rows = len(load_invoice_records(trial_workbook)) if trial_workbook.exists() else 0
    lines = [
        "Handoff folders ready.",
        f"Trial photos: {len(trial_photos)} ({settings.trial_dir})",
        f"Production inbound photos: {len(inbound_photos)} ({settings.inbound_dir})",
        f"Ubuntu baseline workbooks: {len(baseline_workbooks)} ({settings.baseline_dir})",
        f"Ubuntu baseline crop/source images: {len(baseline_images)} ({settings.baseline_dir})",
        f"Existing Windows trial output rows: {trial_output_rows} ({trial_workbook})",
        "",
        "Next real-data commands:",
        f'"{sys.executable}" -m invoice_system run --trial',
        f'"{sys.executable}" -m invoice_system compare --baseline data/baseline --candidate data/output/trial --output data/output/ubuntu_comparison_report.xlsx',
    ]
    if not trial_photos:
        lines.append("Add receipt photos to data/trial/ before running the trial pipeline.")
        if trial_workbook.exists():
            lines.append("Existing trial output may be stale because data/trial/ has no photos.")
    if not baseline_workbooks:
        lines.append("Add the Ubuntu baseline Excel workbook to data/baseline/ before comparing.")
    return "\n".join(lines)


def _write_handoff_notes(settings: Settings) -> None:
    notes = {
        settings.trial_dir
        / "README.txt": """Trial input folder.

Put the real receipt photos you want to compare with Ubuntu here.
Run: .\\scripts\\run-trial.ps1
Output goes to: data\\output\\trial\\
""",
        settings.baseline_dir
        / "README.txt": """Ubuntu baseline folder.

Put the Ubuntu result workbook and any Ubuntu cropped/source images here.
Ignored workbook examples: Invoice_Manually_Checked.xlsx, comparison_report.xlsx, ~$ temporary Excel files.
Run after trial processing: .\\scripts\\compare-ubuntu.ps1
The compare launcher refuses to run until data\\trial\\ contains real receipt photos.
""",
        settings.inbound_dir
        / "README.txt": """Production input folder.

Telegram photos are saved under data\\inbound\\telegram\\YYYY-MM-DD\\.
Local production photos can also be placed here.
Run: .\\scripts\\invoice.ps1 run
""",
        settings.output_dir
        / "README.txt": """Output folder.

Generated Excel workbooks, raw crops, final V2-style crops, OCR audit sheets, and comparison reports are written here.
Do not place Ubuntu baseline files here; use data\\baseline\\ for those.
If data\\trial\\ has no photos but data\\output\\trial\\ has a workbook, that workbook may be stale from a previous smoke/test run.
""",
    }
    for path, text in notes.items():
        path.write_text(text, encoding="utf-8")


def _module_check(module_name: str, detail: str, import_module: bool = False, required: bool = True) -> CheckItem:
    if importlib.util.find_spec(module_name) is None:
        return CheckItem(module_name, False, detail, required)
    if not import_module:
        return CheckItem(module_name, True, detail, required)
    try:
        importlib.import_module(module_name)
    except Exception as exc:
        return CheckItem(module_name, False, f"{detail}; import failed: {exc}", required)
    return CheckItem(module_name, True, detail, required)


def _folder_checks(settings: Settings) -> list[CheckItem]:
    folders: list[Path] = [
        settings.inbound_dir,
        settings.trial_dir,
        settings.output_dir,
        settings.baseline_dir,
    ]
    return [CheckItem(str(path), path.exists(), "data folder") for path in folders]
