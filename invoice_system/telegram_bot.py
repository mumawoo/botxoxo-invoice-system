from __future__ import annotations

import asyncio
import importlib.util
import json
import os
import re
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from .config import Settings
from .expense_categories import normalize_expense_category
from .models import InvoiceRecord, PipelineSummary
from .queue_worker import (
    DONE,
    FAILED_RETRYABLE,
    QueueItem,
    discover_and_enqueue,
    enqueue_photo,
    format_status,
    load_queue_state,
    queue_totals_for_day,
    retry_failed,
    start_background_worker,
    summarize_queue,
    telegram_user_day_dir,
    telegram_user_output_dir,
    telegram_user_queue_path,
    telegram_user_workbook,
)
from .reimbursement import (
    format_reimbursement_summary,
    format_submit_result,
    refresh_checked_outputs,
    submit_unsubmitted,
    submitted_batches_text,
    unsubmitted_summary,
)
from .reimbursement_excel import INVOICE_EXP_SHEET, REVIEW_CROPS_DIR, focus_reimbursement_workbook
from .reimbursement_excel import available_crop_ids, change_reimbursement_record

SUBMIT_CONFIRM_WORDS = {"confirm", "yes", "ok", "\u786e\u8ba4", "\u63d0\u4ea4"}
SUBMIT_CANCEL_WORDS = {"cancel", "no", "\u53d6\u6d88", "\u4e0d\u63d0\u4ea4"}


def is_allowed_user(user_id: int, allowed_ids: set[int] | frozenset[int]) -> bool:
    return user_id in allowed_ids


def whoami_message(user_id: int, username: str | None = None) -> str:
    suffix = f" (@{username})" if username else ""
    return f"Your Telegram user ID is {user_id}{suffix}."


def saved_photo_message(path: Path) -> str:
    return "Saved photo"


def queued_photo_message(path: Path, worker_started: bool) -> str:
    worker = "started" if worker_started else "already running"
    return f"Saved photo\nScanner: {worker}"


def processing_success_message(summary: PipelineSummary) -> str:
    return f"Saved photo and processed current Telegram batch. Sources: {summary.source_images}. Rows: {summary.records_written}. Excel: {summary.workbook_path}"


def processing_failure_message(path: Path, exc: Exception) -> str:
    detail = str(exc).strip() or exc.__class__.__name__
    return f"Saved photo\nProcessing failed: {detail}"


def scan_completion_message(settings: Settings, user_id: int, item: QueueItem, records: list[InvoiceRecord]) -> str:
    if item.status == FAILED_RETRYABLE:
        return "\n".join(["Scan failed", f"Photo: {Path(item.path).name}", f"Error: {item.error or 'unknown error'}", "Use /restart."])
    if item.status != DONE:
        return ""

    today = queue_totals_for_day(settings, user_id)
    queue = summarize_queue(settings, user_id)
    original_totals = _original_currency_totals(records)
    lines = [
        "Scan complete",
        "",
        "This scan",
        f"Rows: {len(records)}",
        "Crops:",
    ]
    lines.extend(_format_scan_record_lines(records))
    if original_totals:
        lines.append("Original totals:")
        lines.extend(_format_currency_lines(original_totals))
    lines.extend(
        [
            "",
            "Today",
            f"Rows: {today.record_count}",
            f"MXN total: {today.total_amount:.2f}",
        ]
    )
    lines.extend(_format_category_block(today.category_totals))
    lines.extend(
        [
            "",
            "Queue",
            f"Done/Pending/Failed: {queue.done}/{queue.pending}/{queue.failed}",
            "Review: /excel",
        ]
    )
    return "\n".join(lines)


def telegram_photo_filename(timestamp: datetime, file_id: str, file_unique_id: str | None = None) -> str:
    identifier = _safe_filename_part(file_id) or _safe_filename_part(file_unique_id or "") or "telegram_photo"
    return f"{timestamp.strftime('%H%M%S')}_{identifier}.jpg"


def resolve_auto_process(settings: Settings, override: bool | None = None) -> bool:
    return settings.telegram_auto_process if override is None else override


def telegram_batch_source(day_dir: Path) -> Path:
    return day_dir


def telegram_polling_ready(settings: Settings) -> bool:
    return bool(settings.telegram_bot_token and telegram_package_ready())


def telegram_config_ready(settings: Settings) -> bool:
    return telegram_polling_ready(settings) and bool(settings.telegram_allowed_user_ids)


def telegram_package_ready() -> bool:
    return importlib.util.find_spec("telegram") is not None and importlib.util.find_spec("telegram.ext") is not None


def telegram_start_message(settings: Settings) -> str:
    if settings.telegram_allowed_user_ids:
        return "Invoice bot is ready. Send receipt photos. Use /help for commands."
    return "Invoice bot setup mode. Use /whoami to get your user ID; receipt photos are rejected until TELEGRAM_ALLOWED_USER_IDS is configured."


def telegram_help_message() -> str:
    return "\n".join(["Commands", *[f"/{command} - {description}" for command, description in telegram_command_menu()]])


def telegram_command_menu() -> list[tuple[str, str]]:
    return [
        ("change", "edit crop: /change 021 type Other + note"),
        ("del", "delete crop: /del 021"),
        ("excel", "download current reimbursement Excel"),
        ("crops", "send latest review crop images"),
        ("recent", "last 2 uploads and Excel rows"),
        ("report", "unsubmitted reimbursement summary"),
        ("restart", "retry failed photos"),
        ("submit", "preview, then reply confirm/cancel"),
        ("status", "queue status"),
        ("whoami", "show your Telegram user ID"),
        ("help", "show commands"),
    ]


def status_message(settings: Settings, user_id: int) -> str:
    discover_and_enqueue(settings, user_id)
    return append_process_status(format_status(summarize_queue(settings, user_id)))


def append_process_status(text: str, pid: int | None = None) -> str:
    process_id = os.getpid() if pid is None else pid
    return "\n".join([text, f"Telegram bot PID: {process_id}"])


def restart_message(settings: Settings, user_id: int) -> str:
    count, summary = retry_failed(settings, user_id)
    started = start_background_worker(settings, user_id)
    return "\n".join(
        [
            "Restart",
            f"Failed photos retried: {count}",
            f"Scanner: {'started' if started else 'already running'}",
            format_status(summary),
        ]
    )


def report_message(settings: Settings, user_id: int) -> str:
    queue = summarize_queue(settings, user_id)
    summary = unsubmitted_summary(settings, user_id)
    summary = type(summary)(
        record_count=summary.record_count,
        photo_count=queue.done,
        total_amount=summary.total_amount,
        category_totals=summary.category_totals,
        codex_used=summary.codex_used,
        failed_count=queue.failed,
        date_min=summary.date_min,
        date_max=summary.date_max,
    )
    return format_reimbursement_summary(summary)


def submit_message(settings: Settings, user_id: int, args: list[str]) -> str:
    if args:
        return "Use /submit first, then reply with plain text: confirm or cancel."
    summary = unsubmitted_summary(settings, user_id)
    if summary.record_count <= 0:
        return "Submit / 提交\nNo unsubmitted records."
    lines = [
        format_reimbursement_summary(summary, title="Submit preview / 提交预览"),
        "",
        "This will archive the current finance Excel and final_crops, then start a new batch from 001.",
        "Reply confirm to submit, or cancel to stop.",
    ]
    return "\n".join(lines)


def submit_pending_confirmation(settings: Settings, user_id: int) -> bool:
    return unsubmitted_summary(settings, user_id).record_count > 0


def submit_confirmation_message(settings: Settings, user_id: int, text: str) -> tuple[bool, str]:
    normalized = str(text or "").strip().casefold()
    if normalized in SUBMIT_CONFIRM_WORDS:
        return True, format_submit_result(submit_unsubmitted(settings, user_id))
    if normalized in SUBMIT_CANCEL_WORDS:
        return True, "Submit cancelled. Nothing was archived."
    return False, "Submit pending. Reply confirm to submit, or cancel to stop."


def recent_message(settings: Settings, user_id: int, limit: int = 2) -> str:
    output_dir = telegram_user_output_dir(settings, user_id)
    state = load_queue_state(telegram_user_queue_path(settings, user_id))
    done_items = [item for item in state.items if item.status == DONE]
    done_items.sort(key=lambda item: (item.received_at or item.updated_at or "", item.updated_at or ""), reverse=True)
    if not done_items:
        return "No completed uploads yet."

    excel_rows = _excel_rows_by_crop_id(telegram_user_workbook(settings, user_id))
    lines = [f"Recent uploads / 最近 {min(limit, len(done_items))} 次"]
    for input_index, item in enumerate(done_items[:limit], start=1):
        source_name = Path(item.path).name
        crop_groups = _crop_groups_for_source(output_dir, item.path)
        excel_matches = sum(1 for crop_ids in crop_groups if crop_ids and excel_rows.get(crop_ids[0]) is not None)
        crop_total = len([crop for group in crop_groups for crop in group])
        lines.append("")
        lines.append(f"Input {input_index}: {item.status} | Excel rows {excel_matches} | crops {crop_total}")
        lines.append(f"Photo: {source_name[:42]}{'...' if len(source_name) > 42 else ''}")
        if not crop_groups:
            lines.append("- no crop records found")
            continue
        for crop_ids in crop_groups:
            primary = crop_ids[0]
            row = excel_rows.get(primary)
            label = "+".join(crop_ids)
            if row is None:
                lines.append(f"- {label} -> not in Excel")
                continue
            status = str(row.get("Manual status") or "active")
            no = _format_excel_no(row.get("No."))
            excel_row = row.get("_row")
            category = str(row.get("Accounting Category") or row.get("Type") or "Other")
            currency = str(row.get("\u539f\u5e01\u79cd") or "MXN")
            amount = float(row.get("\u539f\u91d1\u989d") or row.get("MXN Amount") or 0)
            merchant = str(row.get("Merchant") or "Unknown")
            combined = " | combined one row" if len(crop_ids) > 1 else ""
            lines.append(f"- {label} -> Excel row {excel_row}, No. {no} | {category} {currency} {amount:.2f} | {merchant} | {status}{combined}")
    return "\n".join(lines)


def change_message(settings: Settings, user_id: int, args: list[str]) -> str:
    output_dir = telegram_user_output_dir(settings, user_id)
    try:
        parsed = _parse_change_args(args)
        result = change_reimbursement_record(output_dir, **parsed)
        checked = refresh_checked_outputs(settings, user_id)
    except LookupError as exc:
        recent = ", ".join(available_crop_ids(output_dir)) or "none"
        return f"{exc}\nRecent crops: {recent}"
    except FileNotFoundError:
        return "No reimbursement Excel yet. Send photos first."
    except PermissionError:
        return "Excel is open/locked. Close it and resend /change."
    except OSError as exc:
        return f"Excel is open/locked or unavailable. Close it and resend /change.\n{exc}"
    except ValueError as exc:
        return f"{exc}\nUsage: /change 021 type Other amount 33.35 currency USD + comment"
    changed = _changed_fields_text(result.before, result.after)
    missing = f"\nMissing crops: {len(checked.missing_crops)}" if checked.missing_crops else ""
    return "\n".join(
        [
            "Change saved",
            f"Crop: {result.crop_id}",
            f"Merchant: {result.merchant}",
            f"Status: {result.status}",
            changed or "Changed: status only",
            f"Checked rows/crops: {checked.records_written}/{checked.crops_written}{missing}",
        ]
    )


def delete_message(settings: Settings, user_id: int, args: list[str]) -> str:
    crop_ids = _delete_crop_ids(args)
    if not crop_ids:
        return "Missing crop id\nUsage: /del 021 or /del 021 022"
    output_dir = telegram_user_output_dir(settings, user_id)
    try:
        results = [change_reimbursement_record(output_dir, crop_id, status="delete") for crop_id in crop_ids]
        checked = refresh_checked_outputs(settings, user_id)
    except LookupError as exc:
        recent = ", ".join(available_crop_ids(output_dir)) or "none"
        return f"{exc}\nRecent crops: {recent}"
    except FileNotFoundError:
        return "No reimbursement Excel yet. Send photos first."
    except PermissionError:
        return "Excel is open/locked. Close it and resend /del."
    except OSError as exc:
        return f"Excel is open/locked or unavailable. Close it and resend /del.\n{exc}"
    except ValueError as exc:
        return f"{exc}\nUsage: /del 021 or /del 021 022"
    lines = ["Deleted", f"Crops: {', '.join(result.crop_id for result in results)}"]
    if len(results) == 1:
        lines.append(f"Merchant: {results[0].merchant}")
    else:
        lines.append(f"Count: {len(results)}")
    lines.append(f"Checked rows/crops: {checked.records_written}/{checked.crops_written}")
    return "\n".join(lines)


def _delete_crop_ids(args: list[str]) -> list[str]:
    crop_ids: list[str] = []
    seen: set[str] = set()
    for arg in args:
        for part in re.split(r"[,;\s]+", arg.strip()):
            crop_id = part.strip()
            if not crop_id or crop_id in seen:
                continue
            crop_ids.append(crop_id)
            seen.add(crop_id)
    return crop_ids


def review_crop_paths(settings: Settings, user_id: int) -> list[Path]:
    refresh_checked_outputs(settings, user_id)
    output_dir = telegram_user_output_dir(settings, user_id)
    crops_dir = output_dir / REVIEW_CROPS_DIR
    crop_paths = _review_crop_files(crops_dir)
    latest = _latest_done_item(settings, user_id)
    if latest is None:
        return crop_paths
    names = _crop_names_for_source(output_dir, latest.path)
    if names:
        by_name = {path.name: path for path in crop_paths}
        matched = [by_name[name] for name in names if name in by_name]
        if matched:
            return matched
    source_paths = _crop_paths_for_source(output_dir, latest.path)
    if source_paths:
        return source_paths
    row_count = int(latest.row_count or 0)
    if row_count > 0:
        return crop_paths[-row_count:]
    return []


def _review_crop_files(crops_dir: Path) -> list[Path]:
    if not crops_dir.exists():
        return []
    return sorted(
        (path for path in crops_dir.glob("*.jpg") if path.is_file()),
        key=lambda path: (path.stat().st_mtime, path.name),
    )


def _parse_change_args(args: list[str]) -> dict[str, object]:
    if len(args) < 2:
        raise ValueError("Missing change details")
    crop_id = args[0]
    status = "ok"
    category: str | None = None
    amount: float | None = None
    currency: str | None = None
    comment: str | None = None
    tokens = args[1:]
    keys = {"type", "category", "cat", "amount", "amt", "currency", "cur"}
    statuses = {"ok", "correct", "corrected"}
    index = 0
    saw_change = False
    while index < len(tokens):
        token = tokens[index].casefold()
        if tokens[index] == "+":
            comment = " ".join(tokens[index + 1 :]).strip()
            saw_change = bool(comment) or saw_change
            break
        if token in {"delete", "deleted"}:
            raise ValueError("Use /del 021 to delete a crop")
        if token in statuses:
            status = "correct" if token in {"correct", "corrected"} else "ok"
            saw_change = True
            index += 1
            continue
        if token not in keys:
            raise ValueError(f"Unknown /change token: {tokens[index]}")
        index += 1
        value_parts: list[str] = []
        while index < len(tokens) and tokens[index] != "+" and tokens[index].casefold() not in keys and tokens[index].casefold() not in statuses:
            value_parts.append(tokens[index])
            index += 1
        if not value_parts:
            raise ValueError(f"Missing value after {token}")
        value = " ".join(value_parts)
        if token in {"type", "category", "cat"}:
            category = value
        elif token in {"amount", "amt"}:
            try:
                amount = float(value.replace(",", ""))
            except ValueError as exc:
                raise ValueError(f"Invalid amount: {value}") from exc
        elif token in {"currency", "cur"}:
            currency = value.upper()
        saw_change = True
    if not saw_change:
        raise ValueError("Missing change details")
    return {"crop_id": crop_id, "category": category, "amount": amount, "currency": currency, "comment": comment, "status": status}


def _changed_fields_text(before: dict[str, object], after: dict[str, object]) -> str:
    labels = ["Type", "Accounting Category", "MXN Amount", "\u539f\u5e01\u79cd", "\u539f\u91d1\u989d", "\u6c47\u7387", "Detail", "Manual status"]
    lines: list[str] = []
    for label in labels:
        old = before.get(label)
        new = after.get(label)
        if old != new:
            lines.append(f"{label}: {_display_value(old)} -> {_display_value(new)}")
    if not lines:
        return ""
    return "Changed:\n" + "\n".join(f"- {line}" for line in lines)


def _display_value(value: object) -> str:
    if value in (None, ""):
        return "(blank)"
    if isinstance(value, float):
        return f"{value:.2f}"
    return str(value)


def format_telegram_config(settings: Settings, auto_process: bool | None = None) -> str:
    allowed_count = len(settings.telegram_allowed_user_ids)
    lines = [
        "Telegram bot config:",
        f"Bot token: {'configured' if settings.telegram_bot_token else 'missing'}",
        f"Telegram package: {'installed' if telegram_package_ready() else 'missing'}",
        f"Allowed user IDs: {allowed_count if allowed_count else 'missing'}",
        f"Auto process: {'enabled' if resolve_auto_process(settings, auto_process) else 'disabled'}",
        f"Inbound photo folder: {settings.inbound_dir / 'telegram' / '<telegram_user_id>' / 'YYYY-MM-DD'}",
        f"Qwen OCR: {'enabled' if settings.qwen_api_key else 'disabled'}",
        f"Codex Scan fallback: {'enabled' if settings.codex_scan_enabled and settings.openai_api_key else 'disabled'}",
        f"Polling startup: {'READY' if telegram_polling_ready(settings) else 'NOT READY'}",
        f"Photo ingestion: {'READY' if telegram_config_ready(settings) else 'NOT READY'}",
    ]
    if not settings.telegram_bot_token:
        lines.append("Set TELEGRAM_BOT_TOKEN in .env before starting polling.")
    if not telegram_package_ready():
        lines.append('Install python-telegram-bot with: python -m pip install "python-telegram-bot>=22.0"')
    if not settings.telegram_allowed_user_ids:
        lines.append("Set TELEGRAM_ALLOWED_USER_IDS before sending photos; empty allow-list rejects all photos.")
    if settings.telegram_bot_token and not settings.telegram_allowed_user_ids:
        lines.append("Polling can start for /whoami, but receipt photos will be rejected until allowed IDs are set.")
    lines.append(f"Status: {'READY' if telegram_config_ready(settings) else 'NOT READY'}")
    return "\n".join(lines)


def run_polling_bot(settings: Settings, auto_process: bool | None = None) -> None:
    if not settings.telegram_bot_token:
        raise RuntimeError("TELEGRAM_BOT_TOKEN is not configured")
    try:
        from telegram import BotCommand, Update
        from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters
    except ImportError as exc:
        raise RuntimeError("Install python-telegram-bot to use Telegram ingestion") from exc

    pending_submit_users: set[int] = set()

    async def post_init(application: Application) -> None:
        await application.bot.set_my_commands([BotCommand(command, description) for command, description in telegram_command_menu()])

    async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if update.effective_message:
            await update.effective_message.reply_text(telegram_start_message(settings))

    def allowed(update: Update) -> bool:
        user = update.effective_user
        return bool(user and is_allowed_user(user.id, settings.telegram_allowed_user_ids))

    async def reply_not_allowed(update: Update) -> bool:
        if allowed(update):
            return False
        if update.effective_message:
            await update.effective_message.reply_text("This Telegram user is not allowed.")
        return True

    async def whoami(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        message = update.effective_message
        user = update.effective_user
        if message is not None and user is not None:
            await message.reply_text(whoami_message(user.id, user.username))

    async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if update.effective_message:
            await update.effective_message.reply_text(telegram_help_message())

    async def status(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if await reply_not_allowed(update):
            return
        if update.effective_message and update.effective_user:
            await update.effective_message.reply_text(status_message(settings, update.effective_user.id))

    async def restart(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if await reply_not_allowed(update):
            return
        if update.effective_message and update.effective_user:
            count, summary = retry_failed(settings, update.effective_user.id)
            started = start_background_worker(
                settings,
                update.effective_user.id,
                item_callback=_telegram_item_notifier(settings, context.bot, asyncio.get_running_loop()),
            )
            await update.effective_message.reply_text(
                "\n".join(["Restart", f"Failed photos retried: {count}", f"Scanner: {'started' if started else 'already running'}", format_status(summary)])
            )

    async def report(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if await reply_not_allowed(update):
            return
        if update.effective_message and update.effective_user:
            await update.effective_message.reply_text(report_message(settings, update.effective_user.id))

    async def recent(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if await reply_not_allowed(update):
            return
        if update.effective_message and update.effective_user:
            await update.effective_message.reply_text(recent_message(settings, update.effective_user.id))

    async def submitted(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if await reply_not_allowed(update):
            return
        if update.effective_message and update.effective_user:
            await update.effective_message.reply_text(submitted_batches_text(settings, update.effective_user.id))

    async def submit(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if await reply_not_allowed(update):
            return
        if update.effective_message and update.effective_user:
            text = submit_message(settings, update.effective_user.id, list(context.args or []))
            if not context.args and submit_pending_confirmation(settings, update.effective_user.id):
                pending_submit_users.add(update.effective_user.id)
            await update.effective_message.reply_text(text)

    async def text_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if await reply_not_allowed(update):
            return
        message = update.effective_message
        user = update.effective_user
        if message is None or user is None or user.id not in pending_submit_users:
            return
        handled, text = submit_confirmation_message(settings, user.id, message.text or "")
        if handled:
            pending_submit_users.discard(user.id)
        await message.reply_text(text)

    async def excel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if await reply_not_allowed(update):
            return
        message = update.effective_message
        user = update.effective_user
        if message is None or user is None:
            return
        workbook = telegram_user_workbook(settings, user.id)
        if not workbook.exists():
            await message.reply_text(status_message(settings, user.id))
            return
        refresh_checked_outputs(settings, user.id)
        try:
            focus_reimbursement_workbook(workbook)
        except OSError:
            await message.reply_text("Excel file is open/locked. Close it if you want the file to open at today's row.")
        with workbook.open("rb") as handle:
            await message.reply_document(document=handle, filename=workbook.name)

    async def crops(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if await reply_not_allowed(update):
            return
        message = update.effective_message
        user = update.effective_user
        if message is None or user is None:
            return
        crop_paths = review_crop_paths(settings, user.id)
        if not crop_paths:
            await message.reply_text("No review crops yet. Send photos first, or wait for scanning to finish.")
            return
        await message.reply_text(f"Latest review crops: {len(crop_paths)}")
        sent = 0
        failed: list[str] = []
        for index, crop_path in enumerate(crop_paths, start=1):
            caption = f"{index}/{len(crop_paths)} {crop_path.name}"
            try:
                with crop_path.open("rb") as handle:
                    await message.reply_photo(photo=handle, caption=caption)
                sent += 1
            except Exception as exc:
                try:
                    with crop_path.open("rb") as handle:
                        await message.reply_document(document=handle, filename=crop_path.name, caption=caption)
                    sent += 1
                except Exception as fallback_exc:
                    failed.append(f"{crop_path.name}: {fallback_exc or exc}")
        if failed:
            await message.reply_text("Crop send warning:\n" + "\n".join(failed[:5]))
        await message.reply_text(f"Crops sent: {sent}/{len(crop_paths)}")

    async def change(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if await reply_not_allowed(update):
            return
        message = update.effective_message
        user = update.effective_user
        if message is None or user is None:
            return
        await message.reply_text(change_message(settings, user.id, list(context.args or [])))

    async def delete_crop(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if await reply_not_allowed(update):
            return
        message = update.effective_message
        user = update.effective_user
        if message is None or user is None:
            return
        await message.reply_text(delete_message(settings, user.id, list(context.args or [])))

    async def photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        message = update.effective_message
        user = update.effective_user
        if message is None or user is None:
            return
        if not is_allowed_user(user.id, settings.telegram_allowed_user_ids):
            await message.reply_text("This Telegram user is not allowed.")
            return
        if not message.photo:
            await message.reply_text("Please send a photo.")
            return

        photo_size = message.photo[-1]
        telegram_file = await context.bot.get_file(photo_size.file_id)
        now = datetime.now()
        day_dir = telegram_user_day_dir(settings, user.id, now)
        day_dir.mkdir(parents=True, exist_ok=True)
        target = day_dir / telegram_photo_filename(now, photo_size.file_id, photo_size.file_unique_id)
        await telegram_file.download_to_drive(custom_path=Path(target))
        enqueue_photo(settings, user.id, target, now)

        should_process = resolve_auto_process(settings, auto_process)
        if should_process:
            started = start_background_worker(
                settings,
                user.id,
                item_callback=_telegram_item_notifier(settings, context.bot, asyncio.get_running_loop()),
            )
            await message.reply_text(queued_photo_message(target, started))
        else:
            await message.reply_text(saved_photo_message(target))

    app = Application.builder().token(settings.telegram_bot_token).post_init(post_init).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("whoami", whoami))
    app.add_handler(CommandHandler("status", status))
    app.add_handler(CommandHandler("restart", restart))
    app.add_handler(CommandHandler("excel", excel))
    app.add_handler(CommandHandler("crops", crops))
    app.add_handler(CommandHandler("change", change))
    app.add_handler(CommandHandler("del", delete_crop))
    app.add_handler(CommandHandler("today_excel", excel))
    app.add_handler(CommandHandler("report", report))
    app.add_handler(CommandHandler("recent", recent))
    app.add_handler(CommandHandler("unsubmitted", report))
    app.add_handler(CommandHandler("submitted", submitted))
    app.add_handler(CommandHandler("submit", submit))
    app.add_handler(MessageHandler(filters.PHOTO, photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_message))
    app.run_polling()


def _telegram_item_notifier(settings: Settings, bot, loop: asyncio.AbstractEventLoop):
    def notify(user_id: int, item: QueueItem, records: list[InvoiceRecord]) -> None:
        text = scan_completion_message(settings, user_id, item, records)
        if not text:
            return
        asyncio.run_coroutine_threadsafe(bot.send_message(chat_id=user_id, text=text), loop)

    return notify


def _category_totals(records: list[InvoiceRecord]) -> dict[str, float]:
    totals: dict[str, float] = {}
    for record in records:
        category = normalize_expense_category(record.expense_category)
        amount = float(record.total_amount or 0)
        totals[category] = round(totals.get(category, 0.0) + amount, 2)
    return {category: amount for category, amount in totals.items() if amount}


def _original_currency_totals(records: list[InvoiceRecord]) -> dict[str, float]:
    totals: dict[str, float] = {}
    for record in records:
        currency = str(getattr(record, "original_currency", record.currency) or "MXN").upper()
        amount = float(getattr(record, "original_amount", record.total_amount) or 0)
        totals[currency] = round(totals.get(currency, 0.0) + amount, 2)
    return {currency: amount for currency, amount in totals.items() if amount}


def _format_currency_lines(totals: dict[str, float]) -> list[str]:
    if not totals:
        return ["- none"]
    return [f"- {currency}: {amount:.2f}" for currency, amount in sorted(totals.items()) if amount]


def _format_category_block(totals: dict[str, float]) -> list[str]:
    lines = ["By category:"]
    amounts = [(category, amount) for category, amount in sorted(totals.items()) if amount]
    if not amounts:
        return lines + ["- none"]
    return lines + [f"- {category}: {amount:.2f}" for category, amount in amounts]


def _format_scan_record_lines(records: list[InvoiceRecord]) -> list[str]:
    if not records:
        return ["- none"]
    lines: list[str] = []
    for position, record in enumerate(records, start=1):
        crop_ids = _crop_ids_from_record(record)
        index = "+".join(crop_ids) if crop_ids else (f"{record.line_no:03d}" if record.line_no else f"{position:03d}")
        category = normalize_expense_category(record.expense_category, f"{record.seller} {record.contents}")
        currency = str(getattr(record, "original_currency", record.currency) or "MXN").upper()
        amount = float(getattr(record, "original_amount", record.total_amount) or 0)
        seller = str(record.seller or "Unknown").strip() or "Unknown"
        lines.append(f"- {index} {category}: {currency} {amount:.2f} | {seller}")
        if len(crop_ids) > 1:
            lines.append(f"  Combined: {' + '.join(crop_ids)} -> one Excel row; crops kept as a/b")
    return lines


def _crop_id_from_record(record: InvoiceRecord) -> str:
    name = Path(str(record.crop_image or "")).name
    match = re.match(r"(\d{3,})_", name)
    return match.group(1) if match else ""


def _crop_ids_from_record(record: InvoiceRecord) -> list[str]:
    ids: list[str] = []
    for crop_image in [record.crop_image, *list(getattr(record, "supporting_crop_images", []) or [])]:
        name = Path(str(crop_image or "")).name
        match = re.match(r"(\d{3,})_", name)
        if match and match.group(1) not in ids:
            ids.append(match.group(1))
    return ids


def _record_index_summary(records: list[InvoiceRecord], fallback_photo: Path) -> str:
    if not records:
        return fallback_photo.name
    labels: list[str] = []
    for record in records[:4]:
        index = f"{record.line_no:03d}" if record.line_no else "unknown"
        invoice_date = record.invoice_date or "unknown-date"
        labels.append(f"{index}_{invoice_date}")
    if len(records) > 4:
        labels.append(f"+{len(records) - 4} more")
    return ", ".join(labels)


def _latest_done_item(settings: Settings, user_id: int) -> QueueItem | None:
    state = load_queue_state(telegram_user_queue_path(settings, user_id))
    done = [item for item in state.items if item.status == DONE]
    if not done:
        return None
    return max(done, key=lambda item: (item.received_at or item.updated_at or "", item.updated_at or ""))


def _crop_names_for_source(output_dir: Path, source_image: str) -> list[str]:
    return [path.name for path in _crop_paths_for_source(output_dir, source_image, existing_only=False)]


def _crop_groups_for_source(output_dir: Path, source_image: str) -> list[list[str]]:
    state_path = output_dir / "processing_state.json"
    if not state_path.exists():
        return []
    try:
        data = json.loads(state_path.read_text(encoding="utf-8"))
    except Exception:
        return []
    source_key = _path_key_text(source_image)
    groups: list[list[str]] = []
    for record in data.get("records", []):
        if not isinstance(record, dict):
            continue
        if _path_key_text(str(record.get("source_image") or "")) != source_key:
            continue
        ids: list[str] = []
        for crop_text in [str(record.get("crop_image") or "").strip(), *[str(item or "").strip() for item in record.get("supporting_crop_images", []) or []]]:
            crop_id = _crop_id_from_path_text(crop_text)
            if crop_id and crop_id not in ids:
                ids.append(crop_id)
        if ids:
            groups.append(ids)
    return groups


def _crop_paths_for_source(output_dir: Path, source_image: str, existing_only: bool = True) -> list[Path]:
    state_path = output_dir / "processing_state.json"
    if not state_path.exists():
        return []
    try:
        data = json.loads(state_path.read_text(encoding="utf-8"))
    except Exception:
        return []
    source_key = _path_key_text(source_image)
    paths: list[Path] = []
    for record in data.get("records", []):
        if not isinstance(record, dict):
            continue
        if _path_key_text(str(record.get("source_image") or "")) != source_key:
            continue
        for crop_text in [str(record.get("crop_image") or "").strip(), *[str(item or "").strip() for item in record.get("supporting_crop_images", []) or []]]:
            if not crop_text:
                continue
            path = Path(crop_text)
            if not path.is_absolute():
                path = output_dir / path
            if path in paths:
                continue
            if not existing_only or path.exists():
                paths.append(path)
    return paths


def _excel_rows_by_crop_id(workbook: Path) -> dict[str, dict[str, object]]:
    if not workbook.exists():
        return {}
    try:
        wb = load_workbook(workbook, data_only=True)
    except Exception:
        return {}
    try:
        if INVOICE_EXP_SHEET not in wb.sheetnames:
            return {}
        ws = wb[INVOICE_EXP_SHEET]
        headers = {str(ws.cell(1, col).value or ""): col for col in range(1, ws.max_column + 1)}
        link_col = headers.get("Invoice link", 2)
        rows: dict[str, dict[str, object]] = {}
        for row_idx in range(2, ws.max_row + 1):
            link_cell = ws.cell(row_idx, link_col)
            link = str(link_cell.hyperlink.target if link_cell.hyperlink else link_cell.value or "")
            crop_id = _crop_id_from_path_text(link)
            if not crop_id:
                continue
            row = {header: ws.cell(row_idx, col).value for header, col in headers.items() if header}
            row["_row"] = row_idx
            rows[crop_id] = row
        return rows
    finally:
        wb.close()


def _crop_id_from_path_text(value: str) -> str:
    name = Path(str(value or "")).name
    match = re.match(r"(\d{3,})_", name)
    return match.group(1) if match else ""


def _format_excel_no(value: object) -> str:
    try:
        return f"{int(value):03d}"
    except (TypeError, ValueError):
        return str(value or "unknown")


def _path_key_text(value: str) -> str:
    try:
        return str(Path(value).resolve()).casefold()
    except OSError:
        return str(value or "").casefold()


def _safe_filename_part(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", value.strip())
    cleaned = re.sub(r"\s+", "_", cleaned).strip(" ._")
    return cleaned[:120]
