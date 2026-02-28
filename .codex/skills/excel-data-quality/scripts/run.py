import argparse
import json
import os
from datetime import datetime

import pandas as pd


def _parse_sheet(value):
    if value is None:
        return None
    if isinstance(value, int):
        return value
    text = str(value)
    if text.isdigit():
        return int(text)
    return text


def _series_missing_mask(series):
    if pd.api.types.is_string_dtype(series) or pd.api.types.is_object_dtype(series):
        stripped = series.astype(str).str.strip()
        return series.isna() | (stripped == "")
    return series.isna()


def _date_like_ratio(series):
    if not (pd.api.types.is_string_dtype(series) or pd.api.types.is_object_dtype(series)):
        return 0.0
    non_missing = series.dropna().astype(str)
    if non_missing.empty:
        return 0.0
    date_like = non_missing.str.contains(r"[-/:]", regex=True)
    return float(date_like.mean())


def _boolean_coercion(series):
    if not (pd.api.types.is_string_dtype(series) or pd.api.types.is_object_dtype(series)):
        return None
    non_missing = series.dropna()
    if non_missing.empty:
        return None
    mapping = {
        "true": True,
        "false": False,
        "yes": True,
        "no": False,
        "y": True,
        "n": False,
        "1": True,
        "0": False,
    }
    lowered = non_missing.astype(str).str.strip().str.lower()
    if not lowered.isin(mapping.keys()).all():
        return None
    coerced = lowered.map(mapping)
    full = series.copy()
    full.loc[non_missing.index] = coerced
    return full.astype("boolean")


def _numeric_coercion(series):
    if pd.api.types.is_numeric_dtype(series):
        return series, 1.0
    if not (pd.api.types.is_string_dtype(series) or pd.api.types.is_object_dtype(series)):
        return None, 0.0
    non_missing = series.dropna()
    if non_missing.empty:
        return None, 0.0
    cleaned = non_missing.astype(str).str.replace(",", "", regex=False)
    numeric = pd.to_numeric(cleaned, errors="coerce")
    rate = float(numeric.notna().mean())
    if rate >= 0.9:
        full = series.copy()
        full.loc[non_missing.index] = numeric
        return pd.to_numeric(full, errors="coerce"), rate
    return None, rate


def _datetime_coercion(series):
    if pd.api.types.is_datetime64_any_dtype(series):
        return series, 1.0
    if not (pd.api.types.is_string_dtype(series) or pd.api.types.is_object_dtype(series)):
        return None, 0.0
    non_missing = series.dropna()
    if non_missing.empty:
        return None, 0.0
    dt = pd.to_datetime(non_missing, errors="coerce")
    rate = float(dt.notna().mean())
    if rate >= 0.9:
        full = series.copy()
        full.loc[non_missing.index] = dt
        return pd.to_datetime(full, errors="coerce"), rate
    return None, rate


def _outliers_iqr(series):
    if not pd.api.types.is_numeric_dtype(series):
        return None
    numeric = series.dropna()
    if numeric.empty:
        return None
    q1 = numeric.quantile(0.25)
    q3 = numeric.quantile(0.75)
    iqr = q3 - q1
    lower = q1 - 1.5 * iqr
    upper = q3 + 1.5 * iqr
    mask = (numeric < lower) | (numeric > upper)
    count = int(mask.sum())
    pct = float(mask.mean()) if len(numeric) else 0.0
    return {
        "count": count,
        "pct": pct,
        "lower": float(lower),
        "upper": float(upper),
    }


def _write_report_md(path, report):
    lines = []
    lines.append("# Data Quality Report")
    lines.append("")
    lines.append("## Summary")
    lines.append(f"- Input: {report['input']['path']}")
    lines.append(f"- Sheet: {report['input']['sheet']}")
    lines.append(f"- Rows: {report['input']['rows']}")
    lines.append(f"- Columns: {report['input']['columns']}")
    lines.append(f"- Duplicate rows: {report['duplicates']['duplicate_rows']} ({report['duplicates']['duplicate_rows_pct']:.2%})")
    lines.append(f"- Cleaned rows: {report['output']['cleaned_rows']}")
    lines.append("")

    lines.append("## Missing Rate by Column")
    lines.append("| Column | Missing | Missing % |")
    lines.append("| --- | --- | --- |")
    for item in report["missing"]["by_column"]:
        lines.append(f"| {item['column']} | {item['missing']} | {item['missing_pct']:.2%} |")
    lines.append("")

    lines.append("## Duplicate Values by Column")
    lines.append("| Column | Duplicate Values | Duplicate % |")
    lines.append("| --- | --- | --- |")
    for item in report["duplicates"]["by_column"]:
        lines.append(f"| {item['column']} | {item['duplicate_values']} | {item['duplicate_values_pct']:.2%} |")
    lines.append("")

    lines.append("## Type Corrections")
    if report["type_corrections"]:
        lines.append("| Column | From | To | Parse Rate |")
        lines.append("| --- | --- | --- | --- |")
        for item in report["type_corrections"]:
            lines.append(f"| {item['column']} | {item['from']} | {item['to']} | {item['parse_rate']:.2%} |")
    else:
        lines.append("No type corrections applied.")
    lines.append("")

    lines.append("## Outliers (IQR)")
    if report["outliers"]:
        lines.append("| Column | Count | % of Non-missing | Lower | Upper |")
        lines.append("| --- | --- | --- | --- | --- |")
        for item in report["outliers"]:
            lines.append(
                f"| {item['column']} | {item['count']} | {item['pct']:.2%} | {item['lower']:.6g} | {item['upper']:.6g} |"
            )
    else:
        lines.append("No numeric columns or no outliers detected.")

    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


def main():
    parser = argparse.ArgumentParser(description="Excel data quality report and cleaning")
    parser.add_argument("--input", required=True, help="Path to input .xlsx file")
    parser.add_argument("--output-dir", required=True, help="Output directory")
    parser.add_argument("--sheet", default=None, help="Sheet name or index (default: first sheet)")
    args = parser.parse_args()

    sheet = _parse_sheet(args.sheet)
    os.makedirs(args.output_dir, exist_ok=True)

    sheet_to_load = 0 if sheet is None else sheet
    df = pd.read_excel(args.input, sheet_name=sheet_to_load, engine="openpyxl")
    original_rows = int(len(df))
    original_cols = int(len(df.columns))

    missing_by_column = []
    total_missing = 0
    total_cells = original_rows * original_cols
    for col in df.columns:
        mask = _series_missing_mask(df[col])
        missing = int(mask.sum())
        total_missing += missing
        missing_by_column.append({
            "column": str(col),
            "missing": missing,
            "missing_pct": float(missing / original_rows) if original_rows else 0.0,
        })

    duplicate_rows = int(df.duplicated().sum())
    duplicate_rows_pct = float(duplicate_rows / original_rows) if original_rows else 0.0
    dup_by_column = []
    for col in df.columns:
        series = df[col]
        duplicates = int(series.duplicated().sum())
        denom = int(series.notna().sum())
        dup_by_column.append({
            "column": str(col),
            "duplicate_values": duplicates,
            "duplicate_values_pct": float(duplicates / denom) if denom else 0.0,
        })

    df_clean = df.copy()
    type_corrections = []
    for col in df.columns:
        series = df_clean[col]
        bool_series = _boolean_coercion(series)
        if bool_series is not None:
            df_clean[col] = bool_series
            type_corrections.append({
                "column": str(col),
                "from": str(series.dtype),
                "to": "boolean",
                "parse_rate": 1.0,
            })
            continue

        date_like = _date_like_ratio(series) >= 0.5
        if date_like:
            dt_series, rate = _datetime_coercion(series)
            if dt_series is not None:
                df_clean[col] = dt_series
                type_corrections.append({
                    "column": str(col),
                    "from": str(series.dtype),
                    "to": "datetime64",
                    "parse_rate": rate,
                })
                continue

        num_series, rate = _numeric_coercion(series)
        if num_series is not None:
            df_clean[col] = num_series
            type_corrections.append({
                "column": str(col),
                "from": str(series.dtype),
                "to": "numeric",
                "parse_rate": rate,
            })
            continue

        if pd.api.types.is_string_dtype(series) or pd.api.types.is_object_dtype(series):
            df_clean[col] = series.astype(str).str.strip()

    # Drop rows with missing values (including empty strings) after trimming/coercion.
    missing_row_mask = None
    for col in df_clean.columns:
        col_missing = _series_missing_mask(df_clean[col])
        missing_row_mask = col_missing if missing_row_mask is None else (missing_row_mask | col_missing)
    dropped_missing_rows = int(missing_row_mask.sum()) if missing_row_mask is not None else 0
    df_clean = df_clean.loc[~missing_row_mask].copy()

    df_clean = df_clean.drop_duplicates()
    cleaned_rows = int(len(df_clean))

    outliers = []
    for col in df_clean.columns:
        info = _outliers_iqr(df_clean[col])
        if info:
            outliers.append({
                "column": str(col),
                "count": info["count"],
                "pct": info["pct"],
                "lower": info["lower"],
                "upper": info["upper"],
            })

    cleaned_path = os.path.join(args.output_dir, "cleaned.xlsx")
    report_md_path = os.path.join(args.output_dir, "data_quality_report.md")
    report_json_path = os.path.join(args.output_dir, "data_quality_report.json")

    report = {
        "input": {
            "path": os.path.abspath(args.input),
            "sheet": sheet if sheet is not None else "(first)",
            "rows": original_rows,
            "columns": original_cols,
        },
        "missing": {
            "total_cells": total_cells,
            "total_missing": total_missing,
            "total_missing_pct": float(total_missing / total_cells) if total_cells else 0.0,
            "by_column": missing_by_column,
        },
        "duplicates": {
            "duplicate_rows": duplicate_rows,
            "duplicate_rows_pct": duplicate_rows_pct,
            "by_column": dup_by_column,
        },
        "type_corrections": type_corrections,
        "outliers": outliers,
        "output": {
            "cleaned_path": os.path.abspath(cleaned_path),
            "report_md": os.path.abspath(report_md_path),
            "report_json": os.path.abspath(report_json_path),
            "cleaned_rows": cleaned_rows,
            "dropped_missing_rows": dropped_missing_rows,
        },
        "generated_at": datetime.utcnow().isoformat() + "Z",
    }

    with pd.ExcelWriter(
        cleaned_path,
        engine="openpyxl",
        date_format="yyyy-mm-dd",
        datetime_format="yyyy-mm-dd",
    ) as writer:
        df_clean.to_excel(writer, index=False)
    _write_report_md(report_md_path, report)
    with open(report_json_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
