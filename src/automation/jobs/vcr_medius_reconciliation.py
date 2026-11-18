r"""Reconcile CORSAIR VCR purchase orders with Medius and Lif2 Excel reports.

The Oracle query always uses the configured CORSAIR database; no --db option is needed.

.\.venv\Scripts\python.exe scripts\run_job.py vcr_medius_reconciliation
.\.venv\Scripts\python.exe scripts\run_job.py vcr_medius_reconciliation --medius-file input\medius.xlsx --lif2-file input\lif2.xlsx
"""

from __future__ import annotations

import argparse
import html
import logging
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from ..config import get_database, get_settings
from ..db import get_connection
from ..io_files import pick_files
from ..processing import clean_column_name
from ..results import JobResult

log = logging.getLogger(__name__)
REQUIRES_DB = False
WRITES = False
VCR_DATABASE = "CORSAIR"
SUPPORTED_EXTENSIONS = {".xls", ".xlsx", ".xlsm"}
MEDIUS_REQUIRED = {
    "item_number",
    "external_system_id",
    "ordered_amount",
}
OUTPUT_COLUMNS = [
    "po_identifier_sq",
    "country_code",
    "chargeCode",
    "PO Header Amount",
    "PO Line Amount",
    "last_update_date",
    "External PO Id",
    "Item Number",
    "Medius PO Amount",
    "Medius PO Line Amount",
    "PO Match Status",
    "PO Amount Match",
    "Line Match Status",
    "Line Amount Match",
    "Comments",
]
EXCEL_MAX_ROWS = 1_048_576


class ReconciliationError(RuntimeError):
    """Raised for a controlled reconciliation failure."""


def add_arguments(p: argparse.ArgumentParser) -> None:
    p.add_argument("--medius-file", type=Path)
    p.add_argument("--lif2-file", type=Path)
    p.add_argument("--lookback-days", type=int, default=1)
    p.add_argument("--max-rows", type=int, default=1_000_000)
    p.add_argument("--db-timeout-seconds", type=int, default=300)
    p.add_argument("--tolerance", type=float, default=0.01)


def visible_text(value: Any) -> Any:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return np.nan
    if not isinstance(value, str):
        return value
    text = re.sub(r"<[^>]*>", " ", html.unescape(value).replace("\xa0", " "))
    text = re.sub(r"\s+", " ", text).strip()
    return text if text else np.nan


def normalized_header(value: Any) -> str:
    cleaned = visible_text(value)
    return "" if pd.isna(cleaned) else clean_column_name(str(cleaned))


def detect_header_row(raw: pd.DataFrame, required: set[str] = MEDIUS_REQUIRED) -> int:
    for index in range(min(100, len(raw))):
        values = [normalized_header(v) for v in raw.iloc[index].tolist()]
        if values and values[0] and required.issubset(set(values)):
            return index
    raise ReconciliationError(
        "Medius header was not found in the first 100 rows. Required columns: "
        + ", ".join(sorted(required))
    )


def clean_excel_frame(
    raw: pd.DataFrame, header_row: int, required: set[str]
) -> pd.DataFrame:
    names = [normalized_header(v) for v in raw.iloc[header_row].tolist()]
    nonblank = [n for n in names if n]
    duplicates = sorted({n for n in nonblank if nonblank.count(n) > 1})
    if duplicates:
        raise ReconciliationError(f"Duplicate normalized columns: {duplicates}")
    names = [n or f"unnamed_{i}" for i, n in enumerate(names)]
    frame = raw.iloc[header_row + 1 :].copy()
    frame.columns = names
    frame = frame.map(visible_text).dropna(how="all").reset_index(drop=True)
    missing = sorted(required - set(frame.columns))
    if missing:
        raise ReconciliationError(
            f"Missing required columns: {missing}. Detected columns: {list(frame.columns)}"
        )
    if frame.empty:
        raise ReconciliationError("Medius input contains no data rows")
    return frame


def read_medius(path: Path) -> pd.DataFrame:
    validate_workbook(path)
    try:
        raw = pd.read_excel(path, header=None, dtype=object)
    except Exception as exc:
        raise ReconciliationError(
            f"Cannot read Medius workbook {path.name}; it may be protected: {exc}"
        ) from exc
    return clean_excel_frame(raw, detect_header_row(raw), MEDIUS_REQUIRED)


def validate_workbook(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(f"Input file does not exist: {path}")
    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ReconciliationError(
            f"Unsupported file extension {path.suffix!r}; supported: .xls, .xlsx, .xlsm"
        )


def normalize_identifier(value: Any) -> Any:
    value = visible_text(value)
    if pd.isna(value):
        return pd.NA
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def extract_po_id(value: Any) -> Any:
    value = normalize_identifier(value)
    if pd.isna(value):
        return pd.NA
    parts = [part.strip() for part in str(value).split(";")]
    return parts[1] if len(parts) > 1 and parts[1] else pd.NA


def parse_amount(value: Any) -> float:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return np.nan
    if isinstance(value, (int, float, np.number)):
        return float(value)
    text = html.unescape(str(value)).replace("\xa0", " ").strip()
    if not text:
        return np.nan
    negative = text.startswith("(") and text.endswith(")")
    text = re.sub(r"[^0-9,\.\-]", "", text)
    if negative:
        text = text.replace("-", "")
    if not text or text in {"-", ".", ","}:
        return np.nan
    if "," in text and "." in text:
        decimal = "," if text.rfind(",") > text.rfind(".") else "."
        thousands = "." if decimal == "," else ","
        text = text.replace(thousands, "").replace(decimal, ".")
    elif "," in text:
        tail = len(text) - text.rfind(",") - 1
        text = text.replace(",", ".") if tail in (1, 2) else text.replace(",", "")
    elif text.count(".") > 1:
        head, tail = text.rsplit(".", 1)
        text = head.replace(".", "") + "." + tail
    try:
        result = float(text)
        return -abs(result) if negative else result
    except ValueError:
        return np.nan


def normalize_medius(frame: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    result = frame.copy()
    dummy_rows = (
        result["external_system_id"]
        .astype("string")
        .str.contains("DUMMY", case=False, na=False)
    )
    excluded = int(dummy_rows.sum())
    if excluded:
        log.info(
            "Excluding %d Medius rows with DUMMY in External System Id",
            excluded,
        )
        result = result.loc[~dummy_rows].copy()
    result["external_po_id"] = result["external_system_id"].map(extract_po_id)
    result["item_number"] = result["item_number"].map(normalize_identifier)
    result["ordered_amount"] = result["ordered_amount"].map(parse_amount)
    result["_medius_order"] = np.arange(len(result))
    invalid = int(result["ordered_amount"].isna().sum())
    log.info("Medius invalid or missing amount values: %d", invalid)
    return result, invalid


def normalize_timestamps(values: pd.Series) -> pd.Series:
    """Normalize source timestamps to comparable, timezone-naive UTC values."""
    return pd.to_datetime(values, errors="coerce", utc=True).dt.tz_convert(None)


def normalize_po_match_id(value: Any) -> Any:
    value = normalize_identifier(value)
    if pd.isna(value):
        return pd.NA
    return str(value).lstrip("0") or "0"


def read_lif2(path: Path) -> dict[str, list[str]]:
    validate_workbook(path)
    messages: dict[str, list[str]] = {}
    try:
        with pd.ExcelFile(path) as book:
            sheet_filters = {
                "INVOICE DETAILS": ("bit_class", "VCRP", "source_id"),
                "TRANSFER DETAILS": (
                    "source_transaction_type",
                    "ZVCRP",
                    "source_transaction_id",
                ),
            }
            for sheet, (type_col, expected_type, po_col) in sheet_filters.items():
                if sheet not in book.sheet_names:
                    raise ReconciliationError(
                        f"Lif2 workbook is missing sheet {sheet!r}"
                    )
                raw = book.parse(sheet, dtype=object)
                raw.columns = [normalized_header(c) for c in raw.columns]
                msg_col = next(
                    (c for c in raw.columns if c in {"message_text", "message"}),
                    None,
                )
                if type_col not in raw.columns or po_col not in raw.columns or not msg_col:
                    raise ReconciliationError(
                        f"Lif2 sheet {sheet!r} needs {type_col}, {po_col}, "
                        "and Message Text columns. "
                        f"Detected columns: {list(raw.columns)}"
                    )
                for po, source_type, message in zip(raw[po_col], raw[type_col], raw[msg_col]):
                    row_type = normalize_identifier(source_type)
                    if pd.isna(row_type) or str(row_type).upper() != expected_type:
                        continue
                    po_id, text = normalize_po_match_id(po), visible_text(message)
                    if pd.notna(po_id) and pd.notna(text):
                        messages.setdefault(str(po_id), []).append(f"{sheet}: {text}")
    except ReconciliationError:
        raise
    except Exception as exc:
        raise ReconciliationError(
            f"Cannot read Lif2 workbook {path.name}: {exc}"
        ) from exc
    return messages


def normalize_oracle(frame: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    frame = frame.rename(columns={c: c.lower() for c in frame.columns}).copy()
    aliases = {"chargecode": "charge_code", "totallineamount": "total_line_amount"}
    frame = frame.rename(columns=aliases)
    needed = {
        "po_identifier_sq",
        "country_code",
        "charge_code",
        "total_line_amount",
        "last_update_date",
    }
    missing = needed - set(frame.columns)
    if missing:
        raise ReconciliationError(f"Oracle result missing columns: {sorted(missing)}")
    if frame.empty:
        raise ReconciliationError("Oracle query returned no rows")
    frame["po_identifier_sq"] = frame["po_identifier_sq"].map(normalize_identifier)
    frame["country_code"] = frame["country_code"].map(normalize_identifier)
    frame["charge_code"] = frame["charge_code"].map(normalize_identifier)
    frame["total_line_amount"] = frame["total_line_amount"].map(parse_amount)
    frame["last_update_date"] = normalize_timestamps(frame["last_update_date"])
    frame["_oracle_order"] = np.arange(len(frame))
    invalid = int(frame["total_line_amount"].isna().sum())
    log.info("Oracle invalid or missing amount values: %d", invalid)
    return frame, invalid


def aggregate_oracle_charge_lines(frame: pd.DataFrame) -> pd.DataFrame:
    """Collapse repeated Oracle rows to one summed amount per PO and charge code."""
    grouped = (
        frame.groupby(
            ["po_identifier_sq", "charge_code"],
            dropna=False,
            sort=False,
            as_index=False,
        )
        .agg(
            country_code=("country_code", "first"),
            total_line_amount=("total_line_amount", lambda amounts: amounts.sum(min_count=1)),
            last_update_date=("last_update_date", "max"),
            _oracle_order=("_oracle_order", "min"),
        )
        .sort_values("_oracle_order", kind="stable")
        .reset_index(drop=True)
    )
    log.info(
        "Oracle PO/charge aggregation: %d raw rows reduced to %d groups",
        len(frame),
        len(grouped),
    )
    return grouped


def amounts_match(left: pd.Series, right: pd.Series, tolerance: float) -> pd.Series:
    return (
        left.notna() & right.notna() & np.isclose(left, right, atol=tolerance, rtol=0)
    )


def _key_text(value: Any) -> str | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    return text if text else None


def reconcile(
    oracle: pd.DataFrame, medius: pd.DataFrame, tolerance: float
) -> tuple[pd.DataFrame, dict[str, int | float]]:
    """Compare PO/charge-code aggregates while preserving one output row per Oracle row."""
    oracle = oracle.reset_index(drop=True).copy()
    medius = medius.reset_index(drop=True).copy()
    oracle["_po_key"] = oracle["po_identifier_sq"].map(_key_text)
    oracle["_charge_key"] = oracle["charge_code"].map(_key_text)
    medius["_po_key"] = medius["external_po_id"].map(_key_text)
    medius["_charge_key"] = medius["item_number"].map(_key_text)

    merged = oracle.copy()
    oracle_charge_amounts = (
        oracle.dropna(subset=["_po_key", "_charge_key"])
        .groupby(["_po_key", "_charge_key"])["total_line_amount"]
        .sum(min_count=1)
    )
    medius_charge_amounts = (
        medius.dropna(subset=["_po_key", "_charge_key"])
        .groupby(["_po_key", "_charge_key"])["ordered_amount"]
        .sum(min_count=1)
    )
    medius_charge_rows = (
        medius.dropna(subset=["_po_key", "_charge_key"])
        .drop_duplicates(["_po_key", "_charge_key"])
        .set_index(["_po_key", "_charge_key"])
    )
    merged["_charge_group_key"] = list(zip(merged["_po_key"], merged["_charge_key"]))
    merged["oracle_charge_amount"] = [
        oracle_charge_amounts.get(key, np.nan) for key in merged["_charge_group_key"]
    ]
    merged["medius_line_amount"] = [
        medius_charge_amounts.get(key, np.nan) for key in merged["_charge_group_key"]
    ]
    merged["external_po_id"] = [
        medius_charge_rows.at[key, "external_po_id"]
        if key in medius_charge_rows.index
        else pd.NA
        for key in merged["_charge_group_key"]
    ]
    merged["item_number"] = [
        medius_charge_rows.at[key, "item_number"]
        if key in medius_charge_rows.index
        else pd.NA
        for key in merged["_charge_group_key"]
    ]

    line_difference = (
        merged["oracle_charge_amount"] - merged["medius_line_amount"]
    ).abs()
    line_values_present = (
        merged["oracle_charge_amount"].notna()
        & merged["medius_line_amount"].notna()
    )
    line_amount_match = line_values_present & (line_difference < 1)
    merged["medius_line_amount"] = merged["medius_line_amount"].where(
        ~line_amount_match,
        merged["oracle_charge_amount"],
    )

    oracle_po = (
        oracle_charge_amounts.groupby(level=0, dropna=False)
        .sum(min_count=1)
        .to_dict()
    )
    adjusted_medius_charge_amounts = medius_charge_amounts.copy()
    for key, value in (
        merged.drop_duplicates("_charge_group_key")
        .set_index("_charge_group_key")["medius_line_amount"]
        .items()
    ):
        if key in adjusted_medius_charge_amounts.index:
            adjusted_medius_charge_amounts.loc[key] = value
    medius_po = (
        adjusted_medius_charge_amounts.groupby(level=0, dropna=False)
        .sum(min_count=1)
        .to_dict()
    )
    merged["oracle_po_amount"] = merged["_po_key"].map(oracle_po)
    merged["medius_po_amount"] = merged["_po_key"].map(medius_po)

    medius_pos = set(medius.dropna(subset=["_po_key"])["_po_key"])
    medius_pairs = set(medius_charge_amounts.index)
    merged["po_exists"] = merged["_po_key"].isin(medius_pos)
    merged["charge_match"] = merged["_charge_group_key"].isin(medius_pairs)
    merged["line_match"] = merged["charge_match"]
    merged["po_match"] = merged["po_exists"] & merged["charge_match"]
    merged["po_amount_match"] = amounts_match(
        merged["oracle_po_amount"], merged["medius_po_amount"], tolerance
    )
    merged["line_amount_match"] = line_amount_match
    codes: list[tuple[str, ...]] = []
    for _, row in merged.iterrows():
        reasons: list[str] = []
        if not row["po_exists"]:
            reasons.append("PO_MISSING_MEDIUS")
        if not row["charge_match"]:
            reasons.append("CHARGE_CODE_MISMATCH")
        if pd.isna(row["oracle_po_amount"]):
            reasons.append("ORACLE_PO_AMOUNT_MISSING")
        if row["po_exists"] and pd.isna(row["medius_po_amount"]):
            reasons.append("MEDIUS_PO_AMOUNT_MISSING")
        if (
            pd.notna(row["oracle_po_amount"])
            and pd.notna(row["medius_po_amount"])
            and not row["po_amount_match"]
        ):
            reasons.append("PO_AMOUNT_MISMATCH")
        if (
            pd.notna(row["oracle_charge_amount"])
            and pd.notna(row["medius_line_amount"])
            and not row["line_amount_match"]
        ):
            reasons.append("LINE_AMOUNT_MISMATCH")
        if not row["line_match"]:
            if row["po_exists"]:
                reasons.append("MEDIUS_LINE_MISSING")
        elif pd.isna(row["oracle_charge_amount"]):
            reasons.append("ORACLE_LINE_AMOUNT_MISSING")
        elif pd.isna(row["medius_line_amount"]):
            reasons.append("MEDIUS_LINE_AMOUNT_MISSING")
        codes.append(tuple(dict.fromkeys(reasons)))
    merged["reason_codes"] = codes

    po_status = merged.groupby("_po_key", dropna=False)["po_match"].all()
    po_amount_status = merged.groupby("_po_key", dropna=False)["po_amount_match"].all()
    po_exists_status = merged.groupby("_po_key", dropna=False)["po_exists"].any()
    metrics: dict[str, int | float] = {
        "total_oracle_pos": int(merged["_po_key"].nunique()),
        "matched_pos": int(po_status.sum()),
        "unmatched_pos": int((~po_status).sum()),
        "total_oracle_lines": len(merged),
        "matched_lines": int(merged["line_match"].sum()),
        "unmatched_lines": int((~merged["line_match"]).sum()),
        "po_amount_mismatches": int((~po_amount_status & po_exists_status).sum()),
        "line_amount_mismatches": int(
            (merged["line_match"] & ~merged["line_amount_match"]).sum()
        ),
        "missing_medius_pos": int((~po_exists_status).sum()),
        "charge_code_mismatches": int((~merged["charge_match"]).sum()),
        "medius_only_po_count": len(medius_pos - set(oracle["_po_key"].dropna())),
    }
    metrics["po_match_percentage"] = (
        100 * metrics["matched_pos"] / metrics["total_oracle_pos"]
        if metrics["total_oracle_pos"]
        else 0
    )
    metrics["line_match_percentage"] = (
        100 * metrics["matched_lines"] / metrics["total_oracle_lines"]
        if metrics["total_oracle_lines"]
        else 0
    )
    return merged, metrics


REASON_TEXT = {
    "PO_MISSING_MEDIUS": "PO is missing in Medius",
    "PO_AMOUNT_MISMATCH": "PO Amount is unmatched",
    "ORACLE_PO_AMOUNT_MISSING": "Oracle PO amount is missing",
    "MEDIUS_PO_AMOUNT_MISSING": "Medius PO amount is missing",
    "CHARGE_CODE_MISMATCH": "Charge Code is unmatched",
    "MEDIUS_LINE_MISSING": "PO line is missing in Medius",
    "ORACLE_LINE_AMOUNT_MISSING": "Oracle PO Line Amount is missing",
    "MEDIUS_LINE_AMOUNT_MISSING": "Medius PO Line Amount is missing",
    "LINE_AMOUNT_MISMATCH": "PO Line Amount is unmatched",
}


def add_comments(frame: pd.DataFrame, lif2: dict[str, list[str]]) -> pd.DataFrame:
    result = frame.copy()

    def build(row: pd.Series) -> str:
        reasons = row["reason_codes"]
        if not reasons:
            return ""
        po_id = normalize_po_match_id(row.po_identifier_sq)
        texts = lif2.get(str(po_id), []).copy() if pd.notna(po_id) else []
        if "PO_MISSING_MEDIUS" in reasons:
            reasons = ("PO_MISSING_MEDIUS",)
        else:
            reasons = tuple(
                code for code in reasons if code != "CHARGE_CODE_MISMATCH"
            )
        texts.extend(f"MEDIUS ERROR: {REASON_TEXT[code]}" for code in reasons)
        return "\n".join(dict.fromkeys(texts))

    result["comments"] = result.apply(build, axis=1)
    return result


def presentation(frame: pd.DataFrame) -> pd.DataFrame:
    output = pd.DataFrame(
        {
            "po_identifier_sq": frame["po_identifier_sq"],
            "country_code": frame["country_code"],
            "chargeCode": frame["charge_code"],
            "PO Header Amount": frame["oracle_po_amount"],
            "PO Line Amount": frame["oracle_charge_amount"],
            "last_update_date": frame["last_update_date"],
            "External PO Id": frame["external_po_id"],
            "Item Number": frame["item_number"],
            "Medius PO Amount": frame["medius_po_amount"],
            "Medius PO Line Amount": frame["medius_line_amount"],
            "PO Match Status": frame["po_match"].astype(bool),
            "PO Amount Match": frame["po_amount_match"].astype(bool),
            "Line Match Status": frame["line_match"].astype(bool),
            "Line Amount Match": frame["line_amount_match"].astype(bool),
            "Comments": frame["comments"],
        },
        columns=OUTPUT_COLUMNS,
    )
    return output.where(pd.notna(output), "")


def fetch_oracle(
    db_name: str, sql: str, lookback_days: int, max_rows: int, timeout_seconds: int
) -> pd.DataFrame:
    log.info("Opening Oracle connection")
    rows: list[dict[str, Any]] = []
    try:
        with get_connection(db_name) as connection:
            connection.call_timeout = timeout_seconds * 1000
            log.info("Oracle connection established")
            with connection.cursor() as cursor:
                cursor.arraysize = min(10_000, max_rows + 1)
                cursor.prefetchrows = cursor.arraysize
                log.info("Executing SQL")
                cursor.execute(sql, {"lookback_days": lookback_days})
                columns = [item[0] for item in cursor.description]
                log.info("Fetching rows")
                while len(rows) <= max_rows:
                    batch = cursor.fetchmany(min(10_000, max_rows + 1 - len(rows)))
                    if not batch:
                        break
                    rows.extend(dict(zip(columns, record)) for record in batch)
                    if len(rows) % 50_000 < len(batch):
                        log.info("Fetched %d Oracle rows", len(rows))
                if len(rows) > max_rows:
                    raise ReconciliationError(
                        f"Oracle result exceeds --max-rows={max_rows}; no report was generated"
                    )
    except ReconciliationError:
        raise
    except Exception as exc:
        oracle_error = re.search(
            r"\b(?:ORA|DPI|DPY)-\d{4,5}:\s*[^\r\n]*", str(exc), re.IGNORECASE
        )
        detail = f": {oracle_error.group(0)}" if oracle_error else ""
        raise ReconciliationError(
            f"Oracle query/connection failed or timed out "
            f"({exc.__class__.__name__}){detail}"
        ) from exc
    log.info("Fetch complete: %d rows", len(rows))
    return pd.DataFrame.from_records(rows)


def resolve_inputs(args: argparse.Namespace) -> tuple[Path, Path | None]:
    medius = args.medius_file
    if medius is None:
        selected = pick_files(title="Please upload Medius VCR data")
        if len(selected) != 1:
            raise ReconciliationError("Select exactly one Medius workbook")
        medius = Path(selected[0])
    else:
        medius = Path(medius)
    lif2 = Path(args.lif2_file) if args.lif2_file else None
    validate_workbook(medius)
    if lif2 is not None:
        validate_workbook(lif2)
    return medius, lif2


def _write_table(ws: Any, frame: pd.DataFrame, grouped: bool) -> None:
    start = 2 if grouped else 1
    if len(frame) + start > EXCEL_MAX_ROWS:
        raise ReconciliationError(
            f"Worksheet {ws.title!r} exceeds Excel's {EXCEL_MAX_ROWS:,}-row limit"
        )
    for col, name in enumerate(frame.columns, 1):
        cell = ws.cell(start, col, name)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center")
    for r, values in enumerate(frame.itertuples(index=False, name=None), start + 1):
        for c, value in enumerate(values, 1):
            ws.cell(r, c, value)
    ws.freeze_panes = f"A{start + 1}"
    ws.auto_filter.ref = f"A{start}:{get_column_letter(len(frame.columns))}{max(start, len(frame) + start)}"
    amount_names = {n for n in frame.columns if "Amount" in n}
    for c, name in enumerate(frame.columns, 1):
        if name in amount_names:
            for cell in ws[get_column_letter(c)][start:]:
                cell.number_format = "#,##0.00"
        if name in {"po_identifier_sq", "chargeCode", "External PO Id", "Item Number"}:
            for cell in ws[get_column_letter(c)][start:]:
                cell.number_format = "@"
        width = min(
            50,
            max(
                len(str(name)) + 2,
                max(
                    (len(str(v)) for v in frame[name].head(500) if pd.notna(v)),
                    default=0,
                )
                + 2,
            ),
        )
        ws.column_dimensions[get_column_letter(c)].width = width
    green, red = (
        PatternFill("solid", fgColor="C6EFCE"),
        PatternFill("solid", fgColor="FFC7CE"),
    )
    for c, name in enumerate(frame.columns, 1):
        if "Status" in name or "Match" in name:
            area = f"{get_column_letter(c)}{start + 1}:{get_column_letter(c)}{len(frame) + start}"
            ws.conditional_formatting.add(
                area, CellIsRule(operator="equal", formula=["TRUE"], fill=green)
            )
            ws.conditional_formatting.add(
                area, CellIsRule(operator="equal", formula=["FALSE"], fill=red)
            )


def write_workbook(
    path: Path, reconciliation: pd.DataFrame, summary: dict[str, Any]
) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "PO Reconciliation"
    ws.merge_cells("A1:F1")
    ws["A1"] = "VCR"
    ws.merge_cells("G1:J1")
    ws["G1"] = "Medius"
    ws.merge_cells("K1:O1")
    ws["K1"] = "Reconciliation"
    for cell in (ws["A1"], ws["G1"], ws["K1"]):
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
    _write_table(ws, reconciliation, True)
    sum_ws = wb.create_sheet("Summary")
    sum_frame = pd.DataFrame({"Metric": list(summary), "Value": list(summary.values())})
    _write_table(sum_ws, sum_frame, False)
    try:
        wb.save(path)
    except PermissionError as exc:
        raise ReconciliationError(
            f"Report cannot be written because it is locked: {path}"
        ) from exc
    except Exception as exc:
        raise ReconciliationError(f"Report write failed: {exc}") from exc
    return path


def run(args: argparse.Namespace) -> JobResult:
    started = time.monotonic()
    timestamp = datetime.now()
    settings = get_settings()
    if args.tolerance < 0 or not np.isfinite(args.tolerance):
        raise ValueError("--tolerance must be finite and non-negative")
    if args.lookback_days <= 0:
        raise ValueError("--lookback-days must be positive")
    if args.max_rows <= 0 or args.db_timeout_seconds <= 0:
        raise ValueError("--max-rows and --db-timeout-seconds must be positive")
    log.info("Starting VCR/Medius reconciliation")
    phase_started = time.monotonic()
    log.info("Resolving Medius VCR and optional Lif2 input files")
    medius_path, lif2_path = resolve_inputs(args)
    log.info(
        "Input files selected: Medius=%s, Lif2=%s (%.1fs)",
        medius_path.name,
        lif2_path.name if lif2_path else "not supplied",
        time.monotonic() - phase_started,
    )
    phase_started = time.monotonic()
    log.info("Reading and validating Medius VCR workbook")
    medius, medius_invalid = normalize_medius(read_medius(medius_path))
    log.info(
        "Medius workbook ready: %d rows (%.1fs)",
        len(medius),
        time.monotonic() - phase_started,
    )
    lif2: dict[str, list[str]] = {}
    if lif2_path is not None:
        phase_started = time.monotonic()
        log.info("Reading and validating Lif2 workbook")
        lif2 = read_lif2(lif2_path)
        log.info(
            "Lif2 workbook ready: %d PO IDs with errors (%.1fs)",
            len(lif2),
            time.monotonic() - phase_started,
        )
    else:
        log.info("No Lif2 report supplied; continuing without Lif2 comments")
    sql_path = (
        Path(__file__).resolve().parents[3] / "sql" / "vcr_medius_reconciliation.sql"
    )
    sql = sql_path.read_text(encoding="utf-8")
    log.info("Starting CORSAIR Oracle query")
    oracle_raw = fetch_oracle(
        VCR_DATABASE, sql, args.lookback_days, args.max_rows, args.db_timeout_seconds
    )
    oracle, oracle_invalid = normalize_oracle(oracle_raw)
    oracle_source_rows = len(oracle)
    oracle = aggregate_oracle_charge_lines(oracle)
    log.info("CORSAIR Oracle data ready: %d rows", len(oracle))
    phase_started = time.monotonic()
    log.info("Reconciling Oracle and Medius records")
    merged, metrics = reconcile(oracle, medius, args.tolerance)
    merged = add_comments(merged, lif2)
    log.info(
        "Reconciliation complete: %d lines, %d matched (%.1fs)",
        len(merged),
        metrics["matched_lines"],
        time.monotonic() - phase_started,
    )
    report_frame = presentation(merged)
    db_config = get_database(VCR_DATABASE)
    summary: dict[str, Any] = {
        "Execution timestamp": timestamp,
        "Medius input filename": medius_path.name,
        "Lif2 input filename": lif2_path.name if lif2_path else "Not supplied",
        "Oracle environment": db_config.env,
        "Lookback days": args.lookback_days,
        "Amount tolerance": args.tolerance,
        "Oracle row count": oracle_source_rows,
        "Oracle PO/charge groups": len(oracle),
        "Medius row count": len(medius),
        **metrics,
        "invalid_medius_amount_count": medius_invalid,
        "invalid_oracle_amount_count": oracle_invalid,
        "Execution duration seconds": round(time.monotonic() - started, 2),
        "Final execution status": "SUCCESS",
    }
    output = (
        settings.output_dir
        / f"VCR_MEDIUS_Reconciliation_Report_{timestamp:%Y%m%d_%H%M%S}.xlsx"
    )
    log.info("Writing reconciliation report: %s", output)
    write_workbook(output, report_frame, summary)
    log.info(
        "VCR/Medius reconciliation complete: %s (total %.1fs)",
        output,
        time.monotonic() - started,
    )
    log.info("SUCCESS")
    return JobResult(summary="SUCCESS", files=[output])
