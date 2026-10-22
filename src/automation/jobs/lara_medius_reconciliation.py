r"""Reconcile a Medius LARATO Excel report with read-only LARA Oracle data.

.\.venv\Scripts\python.exe scripts\run_job.py lara_medius_reconciliation --db LARA
.\.venv\Scripts\python.exe scripts\run_job.py lara_medius_reconciliation --db LARA --pick
.\.venv\Scripts\python.exe scripts\run_job.py lara_medius_reconciliation --db LARA --medius-file input\report.xlsx
"""

from __future__ import annotations

import fnmatch
import html
import logging
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

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
REQUIRES_DB = True
WRITES = False
EXCEL_EXTENSIONS = {".xls", ".xlsx", ".xlsm"}
EXCEL_MAX_ROWS = 1_048_576
FETCH_SIZE = 10_000
CONTAINER_TOKEN_INDEX = 0

MEDIUS_LABELS = [
    "Purchase Order: Order Identifier 2",
    "Item: Number",
    "Purchase Order Line: Confirmed Amount",
    "Item: Additional numbers",
    "Purchase Order: Imported Timestamp",
    "Purchase Order Line: Active",
]
MEDIUS_KEYS = [clean_column_name(x) for x in MEDIUS_LABELS]
M_PO, M_ITEM, M_AMOUNT, M_ADDITIONAL, M_DATE, M_ACTIVE = MEDIUS_KEYS
REQUIRED_MEDIUS_LABELS = MEDIUS_LABELS[:4]
REQUIRED_MEDIUS_KEYS = MEDIUS_KEYS[:4]
ORACLE_RENAME = {
    "TREPA_UID": "trepa_uid",
    "TREPA_TO_NUMBER": "trepa_to_number",
    "LINENUMBER": "lineNumber",
    "CHARGECODE": "chargeCode",
    "MAINCURRENCY": "mainCurrency",
    "LINEAMOUNT": "lineAmount",
    "SELFBILLINGREFERENCE": "selfBillingReference",
    "BLNUMBER": "blNumber",
    "CONTAINERNUMBER": "containerNumber",
    "CREATED_DATE": "created_date",
}
HEADER_COLUMNS = [
    "trepa_to_number",
    "chargeCode",
    "PO Amount",
    "created_date",
    *MEDIUS_LABELS[:2],
    "Medius PO Amount",
    "PO Match Status",
    "Comments",
]
LINE_COLUMNS = [
    "trepa_to_number",
    "chargeCode",
    "lineAmount",
    "selfBillingReference",
    "blNumber",
    "containerNumber",
    "created_date",
    *MEDIUS_LABELS[:3],
    "Medius Container Number",
    "Item Line Status",
    "Comments",
]


def add_arguments(p: Any) -> None:
    p.add_argument("--medius-file")
    p.add_argument("--pick", action="store_true")
    p.add_argument("--input-pattern", default="*medius*larato*")
    p.add_argument("--lookback-days", type=int, default=2)
    p.add_argument("--max-rows", type=int, default=1_000_000)
    p.add_argument("--db-timeout-seconds", type=int, default=300)
    p.add_argument("--tolerance", type=float, default=0.01)
    p.add_argument(
        "--sheet", default=0, help="Medius worksheet name or zero-based index"
    )


def visible_text(value: Any) -> Any:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return np.nan
    if not isinstance(value, str):
        return value
    text = re.sub(r"<[^>]*>", " ", html.unescape(value))
    text = re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()
    return text if text else np.nan


def normalize_identifier(value: Any, *, casefold: bool = False) -> Any:
    value = visible_text(value)
    if pd.isna(value):
        return np.nan
    if isinstance(value, float) and value.is_integer():
        value = str(int(value))
    else:
        value = str(value).strip()
    return value.casefold() if casefold else value


def parse_amount(value: Any) -> float:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return np.nan
    if isinstance(value, (int, float, np.number)):
        return float(value)
    s = str(visible_text(value))
    if not s or s.casefold() in {"nan", "none", "na", "n/a", "-"}:
        return np.nan
    negative = s.startswith("(") and s.endswith(")")
    s = re.sub(r"[^0-9,\.\-]", "", s)
    if s.count("-") > 1 or ("-" in s and not s.startswith("-")):
        return np.nan
    if "," in s and "." in s:
        decimal = "," if s.rfind(",") > s.rfind(".") else "."
        thousands = "." if decimal == "," else ","
        s = s.replace(thousands, "").replace(decimal, ".")
    elif "," in s:
        parts = s.split(",")
        s = (
            "".join(parts)
            if len(parts[-1]) == 3 and len(parts) > 1
            else "".join(parts[:-1]) + "." + parts[-1]
        )
    elif s.count(".") > 1:
        parts = s.split(".")
        s = (
            "".join(parts)
            if len(parts[-1]) == 3
            else "".join(parts[:-1]) + "." + parts[-1]
        )
    try:
        result = float(s)
        return -abs(result) if negative else result
    except ValueError:
        return np.nan


def extract_container_number(
    value: Any, token_index: int = CONTAINER_TOKEN_INDEX
) -> Any:
    text = visible_text(value)
    if pd.isna(text):
        return np.nan
    tokens = [
        str(visible_text(x))
        for x in str(text).split(";")
        if not pd.isna(visible_text(x))
    ]
    return tokens[token_index] if len(tokens) > token_index else np.nan


def _raw_excel(path: Path, sheet: str | int) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)
    except ValueError as exc:
        raise ValueError(
            f"Missing sheet {sheet!r} or unreadable workbook {path.name}: {exc}"
        ) from exc
    except Exception as exc:
        raise ValueError(
            f"Unreadable or password-protected workbook {path.name}: {exc}"
        ) from exc


def clean_medius_raw(raw: pd.DataFrame) -> pd.DataFrame:
    required = set(REQUIRED_MEDIUS_KEYS)
    header_at = None
    normalized_header: list[str] = []
    best_detected: list[str] = []
    best_match_count = -1
    for idx in range(min(100, len(raw))):
        normalized_header = [
            clean_column_name(str(visible_text(v)))
            if not pd.isna(visible_text(v))
            else ""
            for v in raw.iloc[idx].tolist()
        ]
        detected = [name for name in normalized_header if name]
        match_count = len(required.intersection(detected))
        if match_count > best_match_count:
            best_detected, best_match_count = detected, match_count
        if required.issubset(normalized_header):
            header_at = idx
            break
    if header_at is None:
        missing = sorted(required.difference(best_detected))
        raise ValueError(
            "Medius header not found in first 100 rows. "
            f"Missing required columns: {missing}. Detected columns: {best_detected}"
        )
    names = normalized_header
    duplicates = sorted({x for x in names if x and names.count(x) > 1})
    if duplicates:
        raise ValueError(f"Duplicate normalized columns: {duplicates}")
    df = raw.iloc[header_at + 1 :].copy()
    df.columns = names
    df = df.loc[:, [bool(x) for x in names]]
    missing = [x for x in REQUIRED_MEDIUS_KEYS if x not in df.columns]
    if missing:
        raise ValueError(
            f"Missing required Medius columns: {missing}. Detected columns: {list(df.columns)}"
        )
    for optional in (M_DATE, M_ACTIVE):
        if optional not in df.columns:
            df[optional] = np.nan
    df = df.map(visible_text).dropna(how="all").reset_index(drop=True)
    if df.empty:
        raise ValueError("Medius input data is empty")
    df["_medius_id"] = np.arange(len(df), dtype=np.int64)
    for col in (M_PO, M_ITEM):
        df[col] = df[col].map(normalize_identifier)
        df[f"_{col}_key"] = df[col].map(
            lambda x: normalize_identifier(x, casefold=True)
        )
    df["_container"] = df[M_ADDITIONAL].map(extract_container_number)
    df["_container_key"] = df["_container"].map(
        lambda x: normalize_identifier(x, casefold=True)
    )
    original_amount = df[M_AMOUNT].copy()
    df[M_AMOUNT] = original_amount.map(parse_amount)
    df["_amount_invalid"] = original_amount.notna() & df[M_AMOUNT].isna()
    df[M_DATE] = pd.to_datetime(df[M_DATE], errors="coerce")
    return df


def read_medius(path: Path, sheet: str | int = 0) -> pd.DataFrame:
    if not path.is_file():
        raise FileNotFoundError(f"Medius file not found: {path}")
    if path.suffix.casefold() not in EXCEL_EXTENSIONS:
        raise ValueError(
            f"Unsupported file extension {path.suffix!r}; supported: .xls, .xlsx, .xlsm. Other formats require an appropriate installed engine."
        )
    return clean_medius_raw(_raw_excel(path, sheet))


def select_input_file(
    explicit: str | None, pattern: str, force_picker: bool, input_dir: Path
) -> Path:
    if explicit:
        selected = [Path(explicit)]
    else:
        candidates = (
            []
            if force_picker
            else sorted(
                p
                for p in input_dir.glob("*")
                if p.suffix.casefold() in EXCEL_EXTENSIONS
                and fnmatch.fnmatch(p.name.casefold(), pattern.casefold())
            )
        )
        selected = candidates if len(candidates) == 1 else list(pick_files())
    if len(selected) != 1:
        raise ValueError(
            f"Exactly one valid Medius Excel file must be selected; received {len(selected)}"
        )
    path = Path(selected[0])
    if not path.is_file():
        raise FileNotFoundError(f"Medius file not found: {path}")
    if path.suffix.casefold() not in EXCEL_EXTENSIONS:
        raise ValueError(f"Unsupported Medius extension: {path.suffix}")
    return path


def fetch_oracle(
    db_name: str, sql: str, lookback_days: int, max_rows: int, timeout_seconds: int
) -> pd.DataFrame:
    log.info("opening Oracle connection")
    rows: list[tuple[Any, ...]] = []
    try:
        with get_connection(db_name) as connection:
            connection.call_timeout = timeout_seconds * 1000
            log.info("connection established")
            with connection.cursor() as cursor:
                cursor.arraysize = FETCH_SIZE
                cursor.prefetchrows = FETCH_SIZE
                log.info("executing SQL")
                cursor.execute(sql, {"lookback_days": lookback_days})
                columns = [d[0] for d in cursor.description]
                log.info("fetching rows")
                while len(rows) <= max_rows:
                    batch = cursor.fetchmany(min(FETCH_SIZE, max_rows + 1 - len(rows)))
                    if not batch:
                        break
                    rows.extend(batch)
                    if len(rows) % 50_000 < len(batch):
                        log.info("fetched %d Oracle rows", len(rows))
    except Exception as exc:
        message = str(exc).casefold()
        if "timeout" in message or "dpi-1067" in message or "dpy-4024" in message:
            raise TimeoutError(
                f"Oracle SQL exceeded timeout of {timeout_seconds} seconds"
            ) from None
        raise RuntimeError(
            "Oracle connection or query failed; connection details were suppressed"
        ) from None
    if len(rows) > max_rows:
        raise ValueError(
            f"Oracle result exceeds --max-rows {max_rows}; no report was generated"
        )
    log.info("fetch complete: %d rows", len(rows))
    if not rows:
        raise ValueError("No Oracle rows returned")
    return pd.DataFrame.from_records(rows, columns=columns).rename(
        columns=ORACLE_RENAME
    )


def prepare_oracle(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    if df.empty:
        raise ValueError("No Oracle rows returned")
    out = df.rename(columns=ORACLE_RENAME).copy()
    needed = [
        "trepa_to_number",
        "chargeCode",
        "lineAmount",
        "selfBillingReference",
        "blNumber",
        "containerNumber",
        "created_date",
    ]
    missing = [c for c in needed if c not in out]
    if missing:
        raise ValueError(
            f"Oracle result missing columns {missing}; detected {list(out.columns)}"
        )
    out["_oracle_id"] = np.arange(len(out), dtype=np.int64)
    for col in ("trepa_to_number", "chargeCode", "containerNumber"):
        out[col] = out[col].map(normalize_identifier)
        out[f"_{col}_key"] = out[col].map(
            lambda x: normalize_identifier(x, casefold=True)
        )
    original_amount = out["lineAmount"].copy()
    out["lineAmount"] = original_amount.map(parse_amount)
    out["_amount_invalid"] = original_amount.notna() & out["lineAmount"].isna()
    out["created_date"] = pd.to_datetime(out["created_date"], errors="coerce")
    return out, int(out["_amount_invalid"].sum())


def _sum_or_nan(s: pd.Series) -> float:
    return s.sum(min_count=1)


def _set_text(s: Iterable[Any]) -> str:
    return ";".join(
        sorted(
            {str(x) for x in s if not pd.isna(x) and str(x).strip()}, key=str.casefold
        )
    )


def is_created_today(value: Any, as_of: datetime) -> bool:
    """Return True when an Oracle timestamp is on/after midnight of the run date."""
    if value is None or pd.isna(value):
        return False
    try:
        return pd.Timestamp(value).date() >= as_of.date()
    except (TypeError, ValueError, OverflowError):
        return False


def reconcile_headers(
    oracle: pd.DataFrame,
    medius: pd.DataFrame,
    tolerance: float,
    as_of: datetime | None = None,
) -> pd.DataFrame:
    as_of = as_of or datetime.now()
    og = (
        oracle.groupby("_trepa_to_number_key", dropna=False, sort=True)
        .agg(
            trepa_to_number=("trepa_to_number", "first"),
            chargeCode=("chargeCode", _set_text),
            **{
                "PO Amount": ("lineAmount", _sum_or_nan),
                "created_date": ("created_date", "max"),
            },
        )
        .reset_index()
    )
    eligible_medius = _eligible_medius(medius)
    mg = (
        eligible_medius.groupby(f"_{M_PO}_key", dropna=False, sort=True)
        .agg(
            **{
                MEDIUS_LABELS[0]: (M_PO, "first"),
                MEDIUS_LABELS[1]: (M_ITEM, _set_text),
                "Medius PO Amount": (M_AMOUNT, _sum_or_nan),
            }
        )
        .reset_index()
    )
    out = og.merge(
        mg,
        left_on="_trepa_to_number_key",
        right_on=f"_{M_PO}_key",
        how="left",
        indicator=True,
    )

    def reasons(r: pd.Series) -> tuple[str, ...]:
        if r["_merge"] == "left_only":
            if is_created_today(r["created_date"], as_of):
                return ("PO_AWAITING_MEDIUS",)
            return ("PO_MISSING",)
        result = []
        if pd.isna(r["PO Amount"]):
            result.append("ORACLE_AMOUNT_MISSING")
        if pd.isna(r["Medius PO Amount"]):
            result.append("MEDIUS_AMOUNT_MISSING")
        if (
            not pd.isna(r["PO Amount"])
            and not pd.isna(r["Medius PO Amount"])
            and not np.isclose(
                r["PO Amount"], r["Medius PO Amount"], atol=tolerance, rtol=0
            )
        ):
            result.append("AMOUNT_MISMATCH")
        if {x.casefold() for x in str(r["chargeCode"]).split(";") if x} != {
            x.casefold()
            for x in str(r[MEDIUS_LABELS[1]]).split(";")
            if x and x != "nan"
        }:
            result.append("CHARGE_MISMATCH")
        return tuple(result)

    out["_reasons"] = out.apply(reasons, axis=1)
    labels = {
        "PO_AWAITING_MEDIUS": "PO awaiting Medius import — created today",
        "PO_MISSING": "PO is missing in Medius",
        "AMOUNT_MISMATCH": "PO Amount is unmatched",
        "ORACLE_AMOUNT_MISSING": "Oracle PO amount is missing",
        "MEDIUS_AMOUNT_MISSING": "Medius PO amount is missing",
        "CHARGE_MISMATCH": "Charge Code is unmatched",
    }
    out["PO Match Status"] = out["_reasons"].map(
        lambda reasons: "PENDING" if "PO_AWAITING_MEDIUS" in reasons else not reasons
    )
    out["Comments"] = out["_reasons"].map(lambda x: "; ".join(labels[y] for y in x))
    return out


def _eligible_medius(m: pd.DataFrame) -> pd.DataFrame:
    """Exclude only Medius lines whose confirmed amount is zero and line Active is No."""
    active_no = (
        m[M_ACTIVE]
        .map(lambda value: normalize_identifier(value, casefold=True))
        .eq("no")
    )
    excluded = m[M_AMOUNT].notna() & m[M_AMOUNT].eq(0) & active_no
    return m.loc[~excluded].copy()


def reconcile_lines(
    oracle: pd.DataFrame,
    medius: pd.DataFrame,
    tolerance: float,
    as_of: datetime | None = None,
) -> pd.DataFrame:
    """Reconcile lines with indexed group lookups and one-to-one assignment."""
    as_of = as_of or datetime.now()
    started = time.monotonic()
    m = _eligible_medius(medius).reset_index(drop=True)
    po_groups: dict[Any, list[int]] = {}
    code_groups: dict[tuple[Any, Any], list[int]] = {}
    for mi, values in enumerate(
        m[[f"_{M_PO}_key", f"_{M_ITEM}_key"]].itertuples(index=False, name=None)
    ):
        po, code = values
        po_groups.setdefault(po, []).append(mi)
        code_groups.setdefault((po, code), []).append(mi)
    oracle_groups: dict[tuple[Any, Any], list[int]] = {}
    for oi, values in zip(
        oracle.index,
        oracle[["_trepa_to_number_key", "_chargeCode_key"]].itertuples(
            index=False, name=None
        ),
    ):
        oracle_groups.setdefault(values, []).append(oi)
    log.info(
        "line reconciliation: indexes built; Oracle groups=%d, eligible Medius lines=%d (%.2fs)",
        len(oracle_groups),
        len(m),
        time.monotonic() - started,
    )

    def container_rank(oracle_index: int, medius_index: int) -> int:
        oracle_container = oracle.at[oracle_index, "_containerNumber_key"]
        medius_container = m.at[medius_index, "_container_key"]
        if (
            pd.notna(oracle_container)
            and pd.notna(medius_container)
            and oracle_container == medius_container
        ):
            return 0
        if pd.isna(oracle_container) or pd.isna(medius_container):
            return 1
        return 2

    assigned: dict[int, int] = {}
    for group_number, (key, oracle_indexes) in enumerate(oracle_groups.items(), 1):
        available = set(code_groups.get(key, ()))
        ordered = sorted(
            oracle_indexes,
            key=lambda oi: (
                pd.isna(oracle.at[oi, "lineAmount"]),
                oracle.at[oi, "lineAmount"]
                if pd.notna(oracle.at[oi, "lineAmount"])
                else float("inf"),
                int(oracle.at[oi, "_oracle_id"]),
            ),
        )
        # Exact amount matches have priority, preventing their rows being consumed by mismatches.
        for oi in ordered:
            amount = oracle.at[oi, "lineAmount"]
            if pd.isna(amount) or not available:
                continue
            exact = [
                mi
                for mi in available
                if pd.notna(m.at[mi, M_AMOUNT])
                and np.isclose(amount, m.at[mi, M_AMOUNT], atol=tolerance, rtol=0)
            ]
            if exact:
                chosen = min(
                    exact,
                    key=lambda mi: (
                        container_rank(oi, mi),
                        abs(amount - m.at[mi, M_AMOUNT]),
                        int(m.at[mi, "_medius_id"]),
                    ),
                )
                assigned[oi] = chosen
                available.remove(chosen)
        # Assign remaining same-key rows by closest amount, still without Medius row reuse.
        for oi in ordered:
            if oi in assigned or not available:
                continue
            amount = oracle.at[oi, "lineAmount"]
            chosen = min(
                available,
                key=lambda mi: (
                    abs(amount - m.at[mi, M_AMOUNT])
                    if pd.notna(amount) and pd.notna(m.at[mi, M_AMOUNT])
                    else float("inf"),
                    container_rank(oi, mi),
                    int(m.at[mi, "_medius_id"]),
                ),
            )
            assigned[oi] = chosen
            available.remove(chosen)
        if group_number % 10_000 == 0:
            log.info(
                "line reconciliation: matched %d/%d key groups",
                group_number,
                len(oracle_groups),
            )
    log.info(
        "line reconciliation: assignment complete; assigned=%d of %d Oracle lines (%.2fs)",
        len(assigned),
        len(oracle),
        time.monotonic() - started,
    )

    comments = {
        "PO_AWAITING_MEDIUS": "PO awaiting Medius import — created today",
        "PO_MISSING": "PO is missing in Medius",
        "CHARGE_MISMATCH": "Charge Code is unmatched",
        "MEDIUS_LINE_MISSING": "Medius line item is missing",
        "ORACLE_AMOUNT_MISSING": "Oracle line amount is missing",
        "MEDIUS_AMOUNT_MISSING": "Medius line amount is missing",
        "AMOUNT_MISMATCH": "Line Amount is unmatched",
    }

    def closest(pool: list[int], amount: Any) -> int | None:
        if not pool:
            return None
        return min(
            pool,
            key=lambda mi: (
                abs(amount - m.at[mi, M_AMOUNT])
                if pd.notna(amount) and pd.notna(m.at[mi, M_AMOUNT])
                else float("inf"),
                int(m.at[mi, "_medius_id"]),
            ),
        )

    def container_warning(oracle_value: Any, medius_value: Any) -> tuple[str, str]:
        oracle_missing = pd.isna(oracle_value)
        medius_missing = pd.isna(medius_value)
        if oracle_missing and medius_missing:
            return (
                "CONTAINER_MISSING_BOTH",
                "Container number is missing in LARA and Medius; row is ambiguous",
            )
        if oracle_missing:
            return (
                "CONTAINER_MISSING_LARA",
                "Container number is missing in LARA; row is ambiguous",
            )
        if medius_missing:
            return (
                "CONTAINER_MISSING_MEDIUS",
                "Container number is missing in Medius; row is ambiguous",
            )
        if normalize_identifier(oracle_value, casefold=True) != normalize_identifier(
            medius_value, casefold=True
        ):
            return (
                "CONTAINER_UNMATCHED",
                "Container number is unmatched; row is ambiguous",
            )
        return "", ""

    rows: list[dict[str, Any]] = []
    ordered_oracle = oracle.sort_values("_oracle_id")
    for position, (oi, o) in enumerate(ordered_oracle.iterrows(), 1):
        po_key = o["_trepa_to_number_key"]
        code_key = o["_chargeCode_key"]
        po_pool = po_groups.get(po_key, [])
        code_pool = code_groups.get((po_key, code_key), [])
        assigned_mi = assigned.get(oi)
        display_mi = assigned_mi
        if display_mi is None:
            display_mi = closest(code_pool or po_pool, o["lineAmount"])
        mr = m.iloc[display_mi] if display_mi is not None else None
        if not po_pool:
            reason = (
                "PO_AWAITING_MEDIUS"
                if is_created_today(o["created_date"], as_of)
                else "PO_MISSING"
            )
        elif not code_pool:
            reason = "CHARGE_MISMATCH"
        elif assigned_mi is None:
            reason = "MEDIUS_LINE_MISSING"
        elif pd.isna(o["lineAmount"]):
            reason = "ORACLE_AMOUNT_MISSING"
        elif pd.isna(mr[M_AMOUNT]):
            reason = "MEDIUS_AMOUNT_MISSING"
        elif not np.isclose(o["lineAmount"], mr[M_AMOUNT], atol=tolerance, rtol=0):
            reason = "AMOUNT_MISMATCH"
        else:
            reason = ""
        container_issue, container_comment = ("", "")
        if not reason and mr is not None:
            container_issue, container_comment = container_warning(
                o["containerNumber"], mr["_container"]
            )
        row = {column: o.get(column, np.nan) for column in LINE_COLUMNS[:7]}
        row.update(
            {
                MEDIUS_LABELS[0]: mr[M_PO] if mr is not None else np.nan,
                MEDIUS_LABELS[1]: mr[M_ITEM] if mr is not None else np.nan,
                MEDIUS_LABELS[2]: mr[M_AMOUNT] if mr is not None else np.nan,
                "Medius Container Number": mr["_container"]
                if mr is not None
                else np.nan,
                "Item Line Status": "PENDING"
                if reason == "PO_AWAITING_MEDIUS"
                else not reason,
                "Comments": comments.get(reason, "") or container_comment,
                "_reason": reason,
                "_container_issue": container_issue,
                "_medius_id": int(mr["_medius_id"])
                if assigned_mi is not None
                else np.nan,
            }
        )
        rows.append(row)
        if position % 50_000 == 0:
            log.info(
                "line reconciliation: built %d/%d output rows",
                position,
                len(ordered_oracle),
            )
    result = pd.DataFrame.from_records(rows)
    log.info(
        "line reconciliation complete: %d rows in %.2fs",
        len(result),
        time.monotonic() - started,
    )
    return result


def build_summary(
    headers: pd.DataFrame,
    lines: pd.DataFrame,
    *,
    timestamp: datetime,
    filename: str,
    environment: str,
    lookback_days: int,
    tolerance: float,
    oracle_rows: int,
    medius_rows: int,
    duration: float,
    medius_only: int,
    invalid_medius: int,
    invalid_oracle: int,
    status: str = "SUCCESS",
) -> pd.DataFrame:
    hr = headers["_reasons"]
    lr = lines["_reason"]
    container_issues = lines["_container_issue"]
    values = [
        ("Execution", "execution timestamp", timestamp),
        ("Execution", "Medius input filename", filename),
        ("Execution", "Oracle environment", environment),
        ("Execution", "lookback days", lookback_days),
        ("Execution", "amount tolerance", tolerance),
        ("Execution", "Oracle row count", oracle_rows),
        ("Execution", "Medius row count", medius_rows),
        ("Execution", "execution duration seconds", round(duration, 2)),
        ("Execution", "final execution status", status),
        ("PO Header", "total Oracle POs", len(headers)),
        ("PO Header", "matched POs", int(headers["PO Match Status"].eq(True).sum())),
        ("PO Header", "unmatched POs", int(headers["PO Match Status"].eq(False).sum())),
        (
            "PO Header",
            "pending POs",
            int(headers["PO Match Status"].eq("PENDING").sum()),
        ),
        ("PO Header", "POs missing in Medius", sum("PO_MISSING" in x for x in hr)),
        ("PO Header", "PO amount mismatches", sum("AMOUNT_MISMATCH" in x for x in hr)),
        (
            "PO Header",
            "missing Oracle PO amounts",
            sum("ORACLE_AMOUNT_MISSING" in x for x in hr),
        ),
        (
            "PO Header",
            "missing Medius PO amounts",
            sum("MEDIUS_AMOUNT_MISSING" in x for x in hr),
        ),
        (
            "PO Header",
            "charge-code mismatches",
            sum("CHARGE_MISMATCH" in x for x in hr),
        ),
        (
            "PO Header",
            "PO match percentage",
            100 * headers["PO Match Status"].eq(True).sum() / len(headers)
            if len(headers)
            else 0,
        ),
        ("Line Item", "total Oracle lines", len(lines)),
        ("Line Item", "matched lines", int(lines["Item Line Status"].eq(True).sum())),
        (
            "Line Item",
            "unmatched lines",
            int(lines["Item Line Status"].eq(False).sum()),
        ),
        (
            "Line Item",
            "pending lines",
            int(lines["Item Line Status"].eq("PENDING").sum()),
        ),
        ("Line Item", "missing Medius POs", int(lr.eq("PO_MISSING").sum())),
        ("Line Item", "charge-code mismatches", int(lr.eq("CHARGE_MISMATCH").sum())),
        (
            "Line Item",
            "container mismatches or missing",
            int(container_issues.ne("").sum()),
        ),
        ("Line Item", "missing Medius lines", int(lr.eq("MEDIUS_LINE_MISSING").sum())),
        (
            "Line Item",
            "missing Oracle line amounts",
            int(lr.eq("ORACLE_AMOUNT_MISSING").sum()),
        ),
        (
            "Line Item",
            "missing Medius line amounts",
            int(lr.eq("MEDIUS_AMOUNT_MISSING").sum()),
        ),
        ("Line Item", "line-amount mismatches", int(lr.eq("AMOUNT_MISMATCH").sum())),
        (
            "Line Item",
            "line match percentage",
            100 * lines["Item Line Status"].eq(True).sum() / len(lines)
            if len(lines)
            else 0,
        ),
        ("Additional", "Medius-only PO count", medius_only),
        ("Additional", "invalid Medius amount count", invalid_medius),
        ("Additional", "invalid Oracle amount count", invalid_oracle),
    ]
    return pd.DataFrame(values, columns=["Section", "Metric", "Value"])


def _excel_value(v: Any) -> Any:
    if v is None or pd.isna(v):
        return "NA"
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if isinstance(v, np.generic):
        return v.item()
    return v


def _detail_sheet(
    wb: Workbook,
    name: str,
    frame: pd.DataFrame,
    columns: list[str],
    groups: list[tuple[str, int, int]],
    amount_cols: set[int],
    status_col: int,
    text_cols: set[int],
) -> None:
    if len(frame) + 2 > EXCEL_MAX_ROWS:
        raise ValueError(f"Worksheet {name!r} exceeds Excel row limit")
    ws = wb.create_sheet(name)
    colors = {"LARATO": "1F4E78", "Medius": "548235", "Reconciliation": "7030A0"}
    for label, start, end in groups:
        ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=end)
        cell = ws.cell(1, start, label)
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = PatternFill("solid", fgColor=colors[label])
        cell.alignment = Alignment(horizontal="center")
    for j, c in enumerate(columns, 1):
        cell = ws.cell(2, j, c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
    for i, row in enumerate(frame[columns].itertuples(index=False, name=None), 3):
        for j, v in enumerate(row, 1):
            cell = ws.cell(i, j, _excel_value(v))
            if j in amount_cols and isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.00"
            if j in text_cols:
                cell.number_format = "@"
        written = i - 2
        if written % 10_000 == 0:
            log.info("workbook: %s wrote %d/%d rows", name, written, len(frame))
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(columns))}{len(frame) + 2}"
    rng = f"{get_column_letter(status_col)}3:{get_column_letter(status_col)}{len(frame) + 2}"
    ws.conditional_formatting.add(
        rng,
        CellIsRule(
            operator="equal",
            formula=["TRUE"],
            fill=PatternFill("solid", fgColor="C6EFCE"),
        ),
    )
    ws.conditional_formatting.add(
        rng,
        CellIsRule(
            operator="equal",
            formula=["FALSE"],
            fill=PatternFill("solid", fgColor="FFC7CE"),
        ),
    )
    ws.conditional_formatting.add(
        rng,
        CellIsRule(
            operator="equal",
            formula=['"PENDING"'],
            fill=PatternFill("solid", fgColor="FFEB9C"),
        ),
    )
    for j, c in enumerate(columns, 1):
        width = (
            max(len(c), *(len(str(_excel_value(v))) for v in frame[c].head(1000)))
            if len(frame)
            else len(c)
        )
        ws.column_dimensions[get_column_letter(j)].width = min(width + 2, 50)


def write_workbook(
    path: Path, headers: pd.DataFrame, lines: pd.DataFrame, summary: pd.DataFrame
) -> Path:
    for name, frame, header_rows in (
        ("PO Header Reconciliation", headers, 2),
        ("PO Line Item Reconciliation", lines, 2),
        ("Summary", summary, 1),
    ):
        if len(frame) + header_rows > EXCEL_MAX_ROWS:
            raise ValueError(f"Worksheet {name!r} exceeds Excel row limit")
    try:
        started = time.monotonic()
        path.parent.mkdir(parents=True, exist_ok=True)
        log.info(
            "workbook: creating PO Header Reconciliation sheet (%d rows)", len(headers)
        )
        wb = Workbook()
        wb.remove(wb.active)
        _detail_sheet(
            wb,
            "PO Header Reconciliation",
            headers,
            HEADER_COLUMNS,
            [("LARATO", 1, 4), ("Medius", 5, 7), ("Reconciliation", 8, 9)],
            {3, 7},
            8,
            {1, 2, 5, 6},
        )
        log.info(
            "workbook: creating PO Line Item Reconciliation sheet (%d rows)", len(lines)
        )
        _detail_sheet(
            wb,
            "PO Line Item Reconciliation",
            lines,
            LINE_COLUMNS,
            [("LARATO", 1, 7), ("Medius", 8, 11), ("Reconciliation", 12, 13)],
            {3, 10},
            12,
            {1, 2, 4, 5, 6, 8, 9, 11},
        )
        log.info("workbook: creating Summary sheet")
        ws = wb.create_sheet("Summary")
        ws.append(list(summary.columns))
        for row in summary.itertuples(index=False, name=None):
            ws.append([_excel_value(x) for x in row])
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:C{len(summary) + 1}"
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4472C4")
        log.info("workbook: saving %s", path.name)
        wb.save(path)
        log.info("workbook complete in %.2fs", time.monotonic() - started)
        return path
    except Exception as exc:
        raise RuntimeError(f"Report write failed for {path}: {exc}") from exc


def run(args: Any) -> JobResult:
    started = time.monotonic()
    now = datetime.now()
    if args.tolerance < 0 or not np.isfinite(args.tolerance):
        raise ValueError("--tolerance must be a finite non-negative number")
    if args.lookback_days <= 0 or args.max_rows <= 0 or args.db_timeout_seconds <= 0:
        raise ValueError("lookback days, max rows, and timeout must be positive")
    settings = get_settings()
    log.info("input selection: resolving Medius report")
    source = select_input_file(
        args.medius_file, args.input_pattern, args.pick, settings.input_dir
    )
    log.info("input reading: %s", source.name)
    phase_started = time.monotonic()
    sheet = int(args.sheet) if str(args.sheet).isdigit() else args.sheet
    medius = read_medius(source, sheet)
    log.info(
        "input reading complete: %d Medius rows in %.2fs",
        len(medius),
        time.monotonic() - phase_started,
    )
    invalid_medius = int(medius["_amount_invalid"].sum())
    missing_medius = int(medius[M_AMOUNT].isna().sum()) - invalid_medius
    log.info(
        "Medius amounts: %d invalid, %d missing",
        invalid_medius,
        missing_medius,
    )
    original_medius_rows = len(medius)
    medius = _eligible_medius(medius).reset_index(drop=True)
    excluded_medius_rows = original_medius_rows - len(medius)
    log.info(
        "Medius filtering complete: excluded %d zero/inactive lines; %d rows remain",
        excluded_medius_rows,
        len(medius),
    )
    if medius.empty:
        raise ValueError(
            "Medius data is empty after excluding lines with confirmed amount 0 and line Active No"
        )
    sql_path = Path(__file__).parents[3] / "sql" / "lara_medius_reconciliation.sql"
    sql = sql_path.read_text(encoding="utf-8")
    raw = fetch_oracle(
        args.db, sql, args.lookback_days, args.max_rows, args.db_timeout_seconds
    )
    log.info("Oracle preparation started")
    phase_started = time.monotonic()
    oracle, invalid_oracle = prepare_oracle(raw)
    log.info("Oracle preparation complete in %.2fs", time.monotonic() - phase_started)
    missing_oracle = int(oracle["lineAmount"].isna().sum()) - invalid_oracle
    log.info(
        "Oracle amounts: %d invalid, %d missing",
        invalid_oracle,
        missing_oracle,
    )
    log.info("PO header reconciliation started")
    phase_started = time.monotonic()
    headers = reconcile_headers(oracle, medius, args.tolerance, as_of=now)
    log.info(
        "PO header reconciliation complete: %d POs in %.2fs",
        len(headers),
        time.monotonic() - phase_started,
    )
    log.info("PO line item reconciliation started: %d Oracle lines", len(oracle))
    lines = reconcile_lines(oracle, medius, args.tolerance, as_of=now)
    log.info("summary creation started")
    phase_started = time.monotonic()
    oracle_pos = set(oracle["_trepa_to_number_key"].dropna())
    medius_pos = set(medius[f"_{M_PO}_key"].dropna())
    db = get_database(args.db)
    summary = build_summary(
        headers,
        lines,
        timestamp=now,
        filename=source.name,
        environment=db.env,
        lookback_days=args.lookback_days,
        tolerance=args.tolerance,
        oracle_rows=len(oracle),
        medius_rows=len(medius),
        duration=time.monotonic() - started,
        medius_only=len(medius_pos - oracle_pos),
        invalid_medius=invalid_medius,
        invalid_oracle=invalid_oracle,
    )
    log.info("summary creation complete in %.2fs", time.monotonic() - phase_started)
    out = (
        settings.output_dir
        / f"LARA_MEDIUS_Reconciliation_Report_{now:%Y%m%d_%H%M%S}.xlsx"
    )
    log.info("report writing started: %s", out.name)
    write_workbook(out, headers, lines, summary)
    text = f"Oracle rows: {len(oracle)}; Medius rows: {len(medius)}; output: {out.name}\nSUCCESS"
    log.info(text)
    return JobResult(summary=text, files=[out])
