# Botxoxo Invoice System

Windows-native receipt and reimbursement automation for small-team expense
workflows. The system receives receipt photos, splits multi-receipt images with
OpenCV, sends each crop to Qwen vision OCR, and writes a traceable manual-review
Excel workbook. Finance exports are generated only when requested, so daily
scanning stays fast and local.

## What It Does

- Ingests receipt photos from Telegram polling or local folders.
- Splits photos containing multiple receipts into individual crop images.
- Sends each crop to Qwen OCR and normalizes seller, date, amount, currency,
  tax, tips, and category.
- Writes a manual reimbursement workbook for human review.
- Supports Telegram correction commands such as `/change`, `/del`, `/group`,
  `/rollback`, `/excel`, `/report`, `/checked`, and `/submit`.
- Keeps company-specific category rules outside the bot code in JSON profiles.
- Builds a clean finance package on demand: checked Excel plus final crop images
  split into `food` and `other` folders.
- Includes a Windows Workbench launcher for non-technical daily operation.

## Privacy Model

Runtime business data stays local. Do not commit real receipts, Telegram photos,
crop images, reimbursement Excel files, logs, archives, bot tokens, or API keys.
The repository `.gitignore` excludes `.env`, `data/*`, Excel files, zip files,
and logs by default. Only reusable sample assets under `data/samples/` are meant
to be tracked.

If a real token or API key was ever pasted into a chat, screenshot, or local
file, rotate it before publishing or sharing the project.

## Quick Start

Use Python 3.11+ on Windows. Create a virtual environment if you are not using
the bundled Codex runtime.

```powershell
python -m pip install -r requirements.txt
Copy-Item .env.example .env
python -m invoice_system check --create-dirs
python -m invoice_system check --strict
```

Edit `.env` and add your own credentials:

```env
TELEGRAM_BOT_TOKEN=
TELEGRAM_ALLOWED_USER_IDS=
TELEGRAM_AUTO_PROCESS=false
TELEGRAM_LANGUAGE=en
COMPANY_PROFILE=default

QWEN_API_KEY=
QWEN_MODEL=qwen-vl-max
QWEN_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions
ENABLE_QWEN_SCAN=false
```

`QWEN_API_KEY` is required for production OCR. `DASHSCOPE_API_KEY` is also
accepted as an alias. OpenAI/Codex fallback is not used in production scanning.

For a bilingual step-by-step setup guide, see
[`INSTALLATION_USAGE_ZH_EN.md`](INSTALLATION_USAGE_ZH_EN.md).

## Windows Workbench

The Workbench is included in the downloadable GitHub source package:

```powershell
.\scripts\workbench.cmd
```

The startup file is `scripts/workbench.cmd`; it launches `scripts/workbench.ps1`.

It opens a Windows Forms control panel for common actions:

- Start or restart the Telegram bot with auto scan.
- Stop running bot PIDs and close the panel.
- Check queue status and running PIDs.
- Process pending photos once.
- Retry failed photos.
- Open the manual Excel workbook, crop folder, output folder, `.env`, and
  checked finance output.
- Build checked Excel/final crops and submit reimbursement batches.

The Workbench uses `INVOICE_SYSTEM_PYTHON` when set, then the bundled Codex
Python runtime if present, then `python`.

## Telegram Setup

1. Create a Telegram bot with BotFather.
2. Copy `.env.example` to `.env`.
3. Paste the bot token into `TELEGRAM_BOT_TOKEN`.
4. Run:

   ```powershell
   python -m invoice_system telegram --check
   ```

5. Start polling:

   ```powershell
   python -m invoice_system telegram --no-process
   ```

6. Send `/whoami` to the bot and put your numeric user ID into
   `TELEGRAM_ALLOWED_USER_IDS`.
7. Restart with auto scan:

   ```powershell
   python -m invoice_system telegram --process
   ```

The bot uses polling, so no public HTTPS webhook is required. Photos are saved
per Telegram user under `data/inbound/telegram/<telegram_user_id>/YYYY-MM-DD/`.
Output is isolated per user under `data/output/telegram/<telegram_user_id>/`.

## Mobile Commands

- `/status` shows queue state and bot PID.
- `/excel` sends the manual-review workbook.
- `/report` summarizes valid unsubmitted records from the manual workbook.
- `/change 021 type Gas` edits a crop record by Trace ID.
- `/change 021 2026-07-01` edits the date.
- `/change 021 + client dinner` adds a note.
- `/del 021 022` marks crop rows deleted.
- `/group` arms the next uploaded photo as one grouped receipt set.
- `/group 044 045` previews grouping existing crops.
- `/group confirm` applies the pending group.
- `/rollback` removes the most recent Telegram photo if it has not been manually
  changed after scanning.
- `/crops` sends recent crop screenshots for review.
- `/checked` builds and sends the finance checked workbook.
- `/submit` previews submission; send `confirm` only after review.
- `/lang zh` or `/lang en` changes reply language.

## Data Flow

1. Telegram or local input saves source photos.
2. OpenCV splits the photo into crop images.
3. Qwen OCR reads each crop and returns strict invoice fields plus orientation.
4. The crop is rotated if needed and renamed with its Trace ID.
5. The manual workbook `报销明细_2026_xlsx.xlsx` is updated.
6. Human review happens in the manual workbook or through Telegram commands.
7. `/checked` or `/submit confirm` generates `报销_checked_2026.xlsx` and
   `final_crops/food|other`.

Daily scanning does not rebuild finance outputs unless requested.

## Company Profiles

Company-specific category rules live in `rules/company_profiles/<profile>.json`.
Set the active profile with:

```env
COMPANY_PROFILE=default
```

The default profile includes examples such as Walmart, Cordia/Cordial, and Codie
being treated as `Food`. Manual changes made with `/change` or in the manual
Excel workbook take priority over automatic rules.

## CLI Reference

```powershell
python -m invoice_system check --create-dirs
python -m invoice_system check --strict
python -m invoice_system sample
python -m invoice_system sample --multi
python -m invoice_system run --trial --input data/samples/synthetic_receipt.jpg --output data/output/smoke
python -m invoice_system telegram --check
python -m invoice_system telegram --process
python -m invoice_system worker --user-id 123456789 --once
python -m invoice_system worker --user-id 123456789 --retry-failed
python -m invoice_system worker --user-id 123456789 --reset-active
python -m invoice_system reimburse report --user-id 123456789
python -m invoice_system reimburse submit --user-id 123456789
python -m invoice_system rebuild-final-crops --user-id 123456789
python -m invoice_system fx update --user-id 123456789 --start-date 2025-10-10
python -m invoice_system ab-test --user-id 123456789 --date YYYY-MM-DD
python -m invoice_system orientation-ab-test --user-id 123456789 --limit 30
```

PowerShell helper scripts are available in `scripts/` for the same workflows.

## Downloading From GitHub

After the v1.0 release is public, users can download the whole project from the
GitHub release page or with **Code -> Download ZIP**. The ZIP includes source
code, tests, rules, sample data, PowerShell helpers, and the Windows Workbench.
It does not include private runtime data.

After unzipping:

```powershell
cd botxoxo-invoice-system
python -m pip install -r requirements.txt
Copy-Item .env.example .env
.\scripts\workbench.cmd
```

## Development Checks

```powershell
python -m unittest discover -s tests
```

Before publishing, also verify:

```powershell
git ls-files .env data *.xlsx *.xls *.zip *.log
```

Only reusable sample files should appear from `data/`.

## License

MIT License. See `LICENSE`.
