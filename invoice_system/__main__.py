from __future__ import annotations

import argparse
import sys
from dataclasses import replace
from pathlib import Path

from .compare import compare_outputs
from .config import Settings
from .diagnostics import format_checks, prepare_handoff, run_checks
from .image_splitter import iter_images
from .pipeline import InvoicePipeline


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="python -m invoice_system")
    sub = parser.add_subparsers(dest="command", required=True)

    run = sub.add_parser("run", help="Process production or trial invoice photos.")
    run.add_argument("--trial", action="store_true", help="Use data/trial and trial output files.")
    run.add_argument("--input", type=Path, default=None, help="Override input folder or image path.")
    run.add_argument("--output", type=Path, default=None, help="Override output folder.")
    run.add_argument("--resume", action="store_true", help="Checkpoint after each source photo and skip completed sources on restart.")
    run.add_argument(
        "--pairing-mode",
        choices=("auto", "review"),
        default=None,
        help="auto merges likely payment slips; review keeps rows and marks possible pairs.",
    )

    telegram = sub.add_parser("telegram", help="Start the Telegram polling bot.")
    telegram_process = telegram.add_mutually_exclusive_group()
    telegram_process.add_argument("--process", dest="auto_process", action="store_true", help="Process photos immediately after download.")
    telegram_process.add_argument("--no-process", dest="auto_process", action="store_false", help="Save photos only, even if TELEGRAM_AUTO_PROCESS=true.")
    telegram.add_argument("--check", action="store_true", help="Check Telegram configuration without starting polling.")
    telegram.set_defaults(auto_process=None)

    worker = sub.add_parser("worker", help="Process Telegram user queues.")
    worker.add_argument("--user-id", type=int, required=True, help="Telegram user ID to process.")
    worker.add_argument("--once", action="store_true", help="Process pending photos once and exit.")
    worker.add_argument("--retry-failed", action="store_true", help="Move failed photos back to pending before processing.")
    worker.add_argument("--reset-active", action="store_true", help="Archive active queue/input/output and start fresh from 001.")

    reimburse = sub.add_parser("reimburse", help="Manage reimbursement summaries and submit state.")
    reimburse_sub = reimburse.add_subparsers(dest="reimburse_command", required=True)
    for name in ("report", "unsubmitted", "submitted", "submit"):
        command = reimburse_sub.add_parser(name)
        command.add_argument("--user-id", type=int, required=True, help="Telegram user ID.")

    fx = sub.add_parser("fx", help="Manage exchange-rate cache.")
    fx_sub = fx.add_subparsers(dest="fx_command", required=True)
    fx_update = fx_sub.add_parser("update", help="Download SAFE exchange rates into reimbursement workbook.")
    fx_update.add_argument("--user-id", type=int, required=True, help="Telegram user ID.")
    fx_update.add_argument("--start-date", default="2025-10-10", help="Start date YYYY-MM-DD.")
    fx_update.add_argument("--end-date", default=None, help="End date YYYY-MM-DD. Defaults to today.")

    compare = sub.add_parser("compare", help="Compare Windows output with Ubuntu baseline output.")
    compare.add_argument("--baseline", type=Path, required=True)
    compare.add_argument("--candidate", type=Path, required=True)
    compare.add_argument("--output", type=Path, default=Path("data/output/comparison_report.xlsx"))

    ab_test = sub.add_parser("ab-test", help="Run A/B OCR comparison: PaddleOCR-VL vs Qwen.")
    ab_test.add_argument("--input", type=Path, default=None, help="Input folder or image path. Overrides --user-id.")
    ab_test.add_argument("--user-id", type=int, default=None, help="Telegram user ID; uses that user's latest/day folder.")
    ab_test.add_argument("--date", default=None, help="Telegram date folder, e.g. 2026-06-16. Defaults to latest.")
    ab_test.add_argument("--output", type=Path, default=Path("data/output/ab_test"), help="A/B output folder.")

    check = sub.add_parser("check", help="Check installed dependencies, config, and data folders.")
    check.add_argument("--create-dirs", action="store_true", help="Create expected data folders before checking.")
    check.add_argument("--strict", action="store_true", help="Exit nonzero if a required check fails.")

    sub.add_parser("prepare", help="Create handoff folders and summarize photo/baseline readiness.")
    sub.add_parser("audit", help="Show current V2 rewrite requirement status.")

    sample = sub.add_parser("sample", help="Create a synthetic receipt image for smoke testing.")
    sample.add_argument("--multi", action="store_true", help="Create one photo containing two receipts.")
    sample.add_argument("--output", type=Path, default=None)

    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    settings = Settings.from_env()

    if args.command == "run":
        return _run_command(args, settings)

    if args.command == "telegram":
        from .telegram_bot import format_telegram_config, run_polling_bot, telegram_config_ready

        if args.check:
            print(format_telegram_config(settings=settings, auto_process=args.auto_process))
            return 0 if telegram_config_ready(settings) else 1
        try:
            run_polling_bot(settings=settings, auto_process=args.auto_process)
        except RuntimeError as exc:
            print(f"Telegram startup error: {exc}", file=sys.stderr)
            print("Run `python -m invoice_system telegram --check` for configuration details.", file=sys.stderr)
            return 1
        return 0

    if args.command == "worker":
        from .queue_worker import format_reset_summary, format_status, process_user_queue_once, reset_active_user_workspace, retry_failed, summarize_queue

        if args.reset_active:
            print(format_reset_summary(reset_active_user_workspace(settings, args.user_id)))
            return 0
        if args.retry_failed:
            count, summary = retry_failed(settings, args.user_id)
            print(f"Retried failed photos: {count}")
        if args.once or args.retry_failed:
            summary = process_user_queue_once(settings, args.user_id)
        else:
            summary = summarize_queue(settings, args.user_id)
        print(format_status(summary))
        return 0

    if args.command == "reimburse":
        from .reimbursement import (
            format_reimbursement_summary,
            format_submit_result,
            submit_unsubmitted,
            submitted_batches_text,
            unsubmitted_summary,
        )

        if args.reimburse_command in {"report", "unsubmitted"}:
            print(format_reimbursement_summary(unsubmitted_summary(settings, args.user_id)))
            return 0
        if args.reimburse_command == "submitted":
            print(submitted_batches_text(settings, args.user_id))
            return 0
        if args.reimburse_command == "submit":
            print(format_submit_result(submit_unsubmitted(settings, args.user_id)))
            return 0

    if args.command == "fx":
        return _fx_command(args, settings)

    if args.command == "compare":
        try:
            report = compare_outputs(args.baseline, args.candidate, args.output)
        except FileNotFoundError as exc:
            print(f"Comparison input error: {exc}", file=sys.stderr)
            return 1
        print(f"Comparison report: {report}")
        return 0

    if args.command == "ab-test":
        return _ab_test_command(args, settings)

    if args.command == "check":
        if args.create_dirs:
            for folder in (settings.inbound_dir, settings.trial_dir, settings.output_dir, settings.baseline_dir):
                folder.mkdir(parents=True, exist_ok=True)
        checks = run_checks(settings)
        print(format_checks(checks))
        if args.strict and any(not item.ok for item in checks if item.required):
            return 1
        return 0

    if args.command == "prepare":
        print(prepare_handoff(settings))
        return 0

    if args.command == "audit":
        from .audit import audit_requirements, format_audit

        print(format_audit(audit_requirements(settings)))
        return 0

    if args.command == "sample":
        from .sample_data import create_synthetic_multi_receipt, create_synthetic_receipt

        output = args.output or Path("data/samples/synthetic_receipts_multi.jpg" if args.multi else "data/samples/synthetic_receipt.jpg")
        path = create_synthetic_multi_receipt(output) if args.multi else create_synthetic_receipt(output)
        print(f"Sample receipt: {path}")
        return 0

    return 2


def _run_command(args: argparse.Namespace, settings: Settings, pipeline_factory=InvoicePipeline) -> int:
    pairing_mode = getattr(args, "pairing_mode", None)
    if pairing_mode:
        settings = replace(settings, pairing_mode=pairing_mode)
    source = args.input or (settings.trial_dir if args.trial else settings.inbound_dir)
    if not iter_images(source):
        print(f"No invoice photos found in {source}. Add photos or pass --input with an image/folder.", file=sys.stderr)
        return 1
    try:
        pipeline = pipeline_factory(settings=settings, trial=args.trial, output_dir=args.output)
        summary = pipeline.process_path(args.input, resume=args.resume)
    except Exception as exc:
        detail = str(exc).strip() or exc.__class__.__name__
        print(f"Run error: {detail}", file=sys.stderr)
        return 1
    print(f"Processed {summary.source_images} source image(s)")
    print(f"Created {summary.crops} crop(s)")
    print(f"Wrote {summary.records_written} invoice row(s)")
    print(f"Excel: {summary.workbook_path}")
    return 0


def _ab_test_command(args: argparse.Namespace, settings: Settings) -> int:
    from .ab_test import run_ab_test, telegram_ab_input

    source = args.input
    if source is None:
        if args.user_id is None:
            print("A/B input error: pass --input or --user-id.", file=sys.stderr)
            return 1
        source = telegram_ab_input(settings, args.user_id, args.date)
    if not iter_images(source):
        print(f"A/B input error: no invoice photos found in {source}.", file=sys.stderr)
        return 1
    try:
        summary = run_ab_test(settings, source, args.output)
    except Exception as exc:
        detail = str(exc).strip() or exc.__class__.__name__
        print(f"A/B run error: {detail}", file=sys.stderr)
        return 1
    print(f"A/B source image(s): {summary.source_images}")
    print(f"A/B crop(s): {summary.crops}")
    print(f"PaddleOCR-VL rows: {summary.paddle_vl_rows}")
    print(f"Qwen rows: {summary.qwen_rows}")
    print(f"Output folder: {summary.output_dir}")
    print(f"Comparison report: {summary.comparison_path}")
    return 0


def _fx_command(args: argparse.Namespace, settings: Settings) -> int:
    from datetime import date

    from .queue_worker import telegram_user_workbook
    from .reimbursement_excel import ReimbursementWorkbook

    try:
        start = date.fromisoformat(args.start_date)
        end = date.fromisoformat(args.end_date) if args.end_date else date.today()
    except ValueError as exc:
        print(f"FX date error: {exc}", file=sys.stderr)
        return 1
    if start > end:
        print("FX date error: start-date must be on or before end-date.", file=sys.stderr)
        return 1
    try:
        result = ReimbursementWorkbook(telegram_user_workbook(settings, args.user_id)).update_exchange_rates(start, end)
    except PermissionError:
        print("FX update error: close the reimbursement Excel file and retry.", file=sys.stderr)
        return 1
    print(f"FX workbook: {result.workbook_path}")
    print(f"Requested: {result.start_date} to {result.end_date}")
    print(f"Existing rows: {result.existing_rows}")
    print(f"Fetched rows: {result.fetched_rows}")
    print(f"Total rows: {result.total_rows}")
    if result.error:
        print(f"FX warning: {result.error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
