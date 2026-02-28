---
name: excel-data-quality
description: Read Excel workbooks and produce a data quality report (missing rate, duplicates, type correction, outliers) plus a cleaned.xlsx using pandas+openpyxl. Use when asked to assess or clean Excel data, generate a quality report, or standardize types in spreadsheets.
---

# Excel Data Quality

## Quick start
- Run `python scripts/run.py --input <path> --output-dir <dir> [--sheet <name_or_index>]`.
- Review outputs: `data_quality_report.md`, `data_quality_report.json`, `cleaned.xlsx`.

## Workflow
1. Confirm the input file path and target sheet (default: first sheet).
2. Run the script; avoid manual edits unless the user asks.
3. Summarize the report and call out columns with high missingness, many duplicates, or outliers.
4. If the user wants imputation or outlier removal, do a second pass and note the changes.

## Outputs
- `cleaned.xlsx`: duplicate rows removed, string columns trimmed, types coerced when safe.
- `data_quality_report.md`: human-readable summary and per-column stats.
- `data_quality_report.json`: machine-readable metrics.

## Notes / thresholds
- Type coercion triggers when at least 90% of non-missing values parse.
- Outliers use IQR (1.5 * IQR) per numeric column.
- Duplicates count uses full-row duplicates; column-level duplicate counts are also reported.
- Missing rate counts nulls and empty strings.

## Script behavior
- Use pandas + openpyxl; preserve column order.
- Do not impute missing values; do not drop outliers.
