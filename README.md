# Windows Dual-OCR Invoice System

Windows-native rewrite of the Xuan invoice processing pipeline from the V2
project summary document.

## Commands

On this Windows PC, the tested runtime is:

```powershell
& 'C:\Users\donxi\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m invoice_system check --strict
```

If `python` already points to the correct Python installation on another Windows
machine, these shorter commands are equivalent:

```powershell
python -m invoice_system check --create-dirs
python -m invoice_system check --strict
python -m invoice_system prepare
python -m invoice_system audit
python -m invoice_system sample
python -m invoice_system sample --multi
python -m invoice_system run
python -m invoice_system run --trial
python -m invoice_system run --resume --input data/inbound/telegram/YYYY-MM-DD --output data/output/production
python -m invoice_system telegram --process
python -m invoice_system worker --user-id 123456789 --once
python -m invoice_system worker --user-id 123456789 --retry-failed
python -m invoice_system reimburse report --user-id 123456789
python -m invoice_system reimburse submit --user-id 123456789
python -m invoice_system ab-test --user-id 123456789 --date 2026-06-16 --output data/output/ab_test
python -m invoice_system compare --baseline data/baseline --candidate data/output/trial --output data/output/ubuntu_comparison_report.xlsx
```

PowerShell launchers are also available. They automatically try
`INVOICE_SYSTEM_PYTHON`, then the bundled Codex Python runtime, then `python`:

```powershell
.\scripts\check.ps1
.\scripts\invoice.ps1 audit
.\scripts\smoke.ps1
.\scripts\smoke.ps1 -Multi
.\scripts\run-trial.ps1
.\scripts\telegram.ps1 --check
.\scripts\telegram.ps1 --process
.\scripts\telegram.ps1 --no-process
.\scripts\worker.ps1 --user-id 123456789 --once
.\scripts\worker.ps1 --user-id 123456789 --retry-failed
.\scripts\reimburse.ps1 report --user-id 123456789
.\scripts\reimburse.ps1 submit --user-id 123456789
.\scripts\compare-ubuntu.ps1
```

For a smoke test without mixing sample data into `data/trial`, run:

```powershell
python -m invoice_system sample
python -m invoice_system run --trial --input data/samples/synthetic_receipt.jpg --output data/output/smoke
python -m invoice_system sample --multi
python -m invoice_system run --trial --input data/samples/synthetic_receipts_multi.jpg --output data/output/smoke_multi
```

The equivalent Windows helper scripts run the actual local OCR engines:

```powershell
.\scripts\smoke.ps1
.\scripts\smoke.ps1 -Multi
```

## Configuration

Copy `.env.example` to `.env` when you are ready to use Telegram or Qwen Scan:

```powershell
Copy-Item .env.example .env
```

```env
TELEGRAM_BOT_TOKEN=123456:token
TELEGRAM_ALLOWED_USER_IDS=123456789
TELEGRAM_AUTO_PROCESS=false
OPENAI_API_KEY=sk-...
OPENAI_MODEL=gpt-4.1
QWEN_API_KEY=sk-...
QWEN_MODEL=qwen-vl-max
ENABLE_QWEN_SCAN=true
PADDLEOCR_LANG=en
EASYOCR_LANGS=es,en
TESSERACT_CMD=tesseract
TESSERACT_LANG=eng+spa
TESSERACT_PSM=6
LOCAL_OCR_CONFIDENCE_THRESHOLD=0.62
AMOUNT_TOLERANCE_MXN=0.50
```

`QWEN_API_KEY` is only needed when Qwen Scan fallback is enabled for local OCR
disagreements, low confidence, missing fields, poor image quality, or A/B tests
against PaddleOCR-VL. Paste the DashScope/Qwen compatible API key on the
`QWEN_API_KEY=` line in `.env`; `DASHSCOPE_API_KEY` is also accepted as an alias.
`TELEGRAM_BOT_TOKEN` is only needed for the Telegram polling bot.

For A/B testing PaddleOCR-VL against Qwen on the same Telegram photos, run
`python -m invoice_system ab-test --user-id 123456789 --date YYYY-MM-DD`.
The output folder contains `paddleocr_vl/Invoice_Output.xlsx`,
`qwen/Invoice_Output.xlsx`, and `AB_Comparison.xlsx`.

## Flow

1. Photos enter through `data/inbound/`, `data/trial/`, or Telegram polling.
2. OpenCV splits multi-receipt photos into cropped invoice images.
3. Each crop is scanned by EasyOCR and PaddleOCR.
4. If both local OCR results agree on normalized date, seller, currency, total,
   VAT/tax, and tips, the local result is written to Excel.
5. If they disagree, PaddleOCR confidence is low, key fields are missing, or the
   crop is poor quality, Qwen Scan is invoked when enabled.
6. Qwen Scan output must validate as strict invoice JSON before it can replace
   the local OCR result.
7. Likely invoice/payment-slip pairs are merged so tip deltas are calculated.
8. Output workbooks keep the original 13 invoice columns, preserve rows marked
   `Manually checked` or `Deleted` in `Invoice_Manually_Checked.xlsx`, and
   refresh that manual-review workbook after each run so it can be edited for
   the next locked pass.

## Telegram Setup

1. In Telegram, open BotFather and create a bot.
2. Copy `.env.example` to `.env`.
3. Paste the bot token into `TELEGRAM_BOT_TOKEN`.
4. Run `.\scripts\telegram.ps1 --check`.
   `Polling startup: READY` means the bot can start for `/whoami`; `Photo ingestion: READY`
   appears only after allowed IDs are configured.
5. Run `.\scripts\telegram.ps1`.
6. Send `/whoami` to the bot and copy the numeric ID into `TELEGRAM_ALLOWED_USER_IDS`.
   Multiple IDs can be separated with commas, semicolons, or spaces.
7. Run `.\scripts\telegram.ps1 --check` again.
8. Restart with `.\scripts\telegram.ps1 --process`, or set
   `TELEGRAM_AUTO_PROCESS=true` in `.env`.

Receipt photos are rejected until `TELEGRAM_ALLOWED_USER_IDS` contains the
sender's numeric ID. `/whoami` remains available during setup so you can get
that ID safely. Mobile Telegram commands are intentionally small and safe:
`/status`, `/restart`, `/today_excel`, `/report`, `/unsubmitted`, `/submitted`,
`/submit`, and `/help`.

Use `.\scripts\telegram.ps1 --no-process` when you want the bot to save photos
without running OCR, even if `.env` enables auto-processing.

The bot uses polling, so it can run on this Windows PC without a public HTTPS
webhook. Incoming photos are saved per user under
`data/inbound/telegram/<telegram_user_id>/YYYY-MM-DD/`; per-user output is under
`data/output/telegram/<telegram_user_id>/`. When processing is enabled, the bot
queues photos and starts a background scanner without blocking later uploads.
The scanner refreshes Excel after each source photo and uses
`processing_state.json` plus `queue_state.json` to resume after interruption.

Use the fuller Windows/finance commands from PowerShell:

```powershell
.\scripts\worker.ps1 --user-id 123456789 --once
.\scripts\worker.ps1 --user-id 123456789 --retry-failed
.\scripts\reimburse.ps1 report --user-id 123456789
.\scripts\reimburse.ps1 unsubmitted --user-id 123456789
.\scripts\reimburse.ps1 submitted --user-id 123456789
.\scripts\reimburse.ps1 submit --user-id 123456789
```

`reimburse submit` marks current unsubmitted rows as submitted, archives the
current `Invoice_Output.xlsx` under `submitted/`, and resets active scan output
so new photos start a fresh Excel workbook.

## Ubuntu Comparison Handoff

Place the Ubuntu output workbook and cropped images under `data/baseline/`, then
run:

```powershell
.\scripts\invoice.ps1 prepare
.\scripts\run-trial.ps1 --resume
.\scripts\compare-ubuntu.ps1
```

`run` exits with a clear error if the input folder has no photos. `compare` exits
with a clear error if either side has no invoice workbook.

The invoice workbook includes an `OCR_Audit` sheet with PaddleOCR, EasyOCR, and
Codex Scan text/error/confidence evidence for each crop. `Run_Summary` separates
`Codex Scan attempted` from `Codex Scan used`, so an OCR mismatch with a missing
API key is still visible as a triggered fallback. The comparison report includes
`Summary`, `Comparison`, and `Crop Counts` sheets.
Crop counts use `OCR_Audit` when available; otherwise they are inferred from
crop filenames such as `photo_a_d01.jpg` or V2 row-numbered crop names.
Comparison categories include exact match, near match, missing in Windows, extra
in Windows, OCR disagreement resolved by Codex Scan, and field-level mismatch.
Date formats, seller accents/case, currency aliases, and small amount differences
are normalized during matching.

Raw OCR crops are kept in `data/output/<mode>/crops/`. Row-numbered copies using
the V2 naming style are written to `data/output/<mode>/final_crops/`.
