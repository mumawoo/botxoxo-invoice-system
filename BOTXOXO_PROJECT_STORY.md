# Botxoxo Reimbursement

## Inspiration

As an expatriate working in Mexico, I need to submit receipts to claim my daily stipend. The process is time-consuming, and when work gets busy, receipts can easily be forgotten or misplaced. Missing invoices often means receiving less reimbursement. I wanted a faster, simpler way to capture every expense as it happens.

## What it does

Botxoxo Reimbursement converts receipt photos sent through Telegram into traceable Excel records. It:

- Separates multiple receipts in one photo.
- Extracts dates, merchants, currencies, amounts, IVA/VAT, and tips.
- Identifies possible duplicates and related payment slips.
- Supports corrections and deletions through Telegram.
- Generates categorized, finance-ready Excel files and receipt images.
- Assigns every receipt a permanent Trace ID.

Foreign expenses are converted using:

$$
\text{MXN Amount} = \text{Original Amount} \times \text{Exchange Rate}
$$

## How we built it

We built a Windows-native Python application using:

- OpenCV for receipt detection and cropping.
- Qwen Vision for OCR, handwriting, and orientation analysis.
- Telegram Bot API for mobile photo intake and commands.
- OpenPyXL for Excel generation and human review.
- JSON checkpoints for resumable processing and audit history.

The review workbook acts as the business database, so machine-generated results cannot silently overwrite human corrections.

## Challenges we ran into

Receipt photos can contain shadows, overlapping papers, rotated text, handwriting, or several receipts on one table. Other challenges included:

- Distinguishing duplicates from separate receipts with identical amounts.
- Combining restaurant receipts with card slips containing tips.
- Handling ambiguous dates and multiple currencies.
- Preserving Trace IDs during corrections and grouping.
- Recovering from Telegram network timeouts.
- Preventing automatic rebuilding from overwriting reviewed Excel cells.

## Accomplishments that we're proud of

We created a complete workflow that combines AI automation with human control. The system supports multiple Telegram users, resumable processing, mobile corrections, duplicate warnings, company-specific accounting rules, and clean finance exports.

## What we learned

Financial automation requires more than accurate OCR. Traceability, reliable recovery, stable identifiers, and human review are equally important. AI should accelerate decisions while making uncertainty visible.

## What's next for Botxoxo Reimbursement

Next steps include a browser-based finance dashboard, support for additional messaging platforms, better receipt splitting, more company profiles, cloud deployment, and integrations with accounting and reimbursement systems.
