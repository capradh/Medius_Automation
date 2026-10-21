r"""CMONA vs Medius PO / GR reconciliation.

Asks for the five reports one by one (file picker window), cleans them, reconciles
PO and GR data and saves a 3-sheet Excel report to your Downloads folder.

    python scripts\run_job.py cmona_medius_reconciliation
    python scripts\run_job.py cmona_medius_reconciliation --auto
    python src\automation\jobs\cmona_medius_reconciliation.py          (VS Code "Run Python File")

Sheets: 1) PO Reconciliation  2) GR Reconciliation  3) PO Summary
The script ends with a SUCCESS or FAILED message.
"""

from __future__ import annotations

import argparse
import csv
import glob
import html
import io
import logging
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path
from typing import Sequence

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

try:  # normal use: part of the automation package
    from ..config import get_settings
    from ..results import JobResult
except ImportError:  # file executed directly from VS Code
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
    from automation.config import get_settings  # type: ignore[no-redef]
    from automation.results import JobResult  # type: ignore[no-redef]

log = logging.getLogger(__name__)

# --------------------------------------------------------------------------- #
# Constants
# --------------------------------------------------------------------------- #
NA = "NA"
REPORT_PREFIX = "CMONA_MEDIUS_Reconciliation_Report"
DEFAULT_TOLERANCE = 0.01
HEADER_SCAN_ROWS = 200
EXCEL_MAX_ROWS = 1_048_576
STYLE_CELL_LIMIT = 300_000  # above this many rows, skip per-cell number formatting
MAX_CELL_CHARS = 32_000
NULL_TOKENS = frozenset({"", "nan", "none", "null", "na", "n/a", "<na>"})

SHEET_PO = "PO Reconciliation"
SHEET_GR = "GR Reconciliation"
SHEET_SUMMARY = "PO Summary"
SHEET_TRANSFER = "TRANSFER DETAILS"
SHEET_INVOICE = "INVOICE DETAILS"

_TAG_PATTERN = r"<[^>]*>"
_ENTITY_PATTERN = r"&#?\w+;"
_TAG_RE = re.compile(_TAG_PATTERN)
_ENTITY_RE = re.compile(_ENTITY_PATTERN)
_WS_RE = re.compile(r"\s+")
_SCI_RE = re.compile(r"[+-]?\d+(\.\d+)?[eE][+-]?\d+")


def norm_col(name: object) -> str:
    """Standardise a column name: 'Order Identifier 2 ' -> 'order_identifier_2'."""
    return re.sub(r"[^0-9a-z]+", "_", str(name).strip().lower()).strip("_")


# Required source columns (as named in the reports) and their standardised names
COL_CM_PO = ("PO Id", "Business Unit Id", "sum(PO Spend in Original Currency)")
COL_CM_GR = (
    "[RCT] Receipt Id",
    "[PO-P&I] PO Id",
    "[PO-P&I]Purchasing Company (Purchase Organization Id)",
)
COL_ME_PO = (
    "Order Identifier 2",
    "Purchase Order: Company",
    "Ordered Amount",
    "Purchase Order: Imported Timestamp",
)
COL_ME_GR = (
    "Delivery Note",
    "Purchase Order: Company",
    "Purchase Order: Imported Timestamp",
)
COL_LIF2_TRANSFER = ("Source Transaction ID", "Message Text")  # TRANSFER DETAILS
COL_LIF2_INVOICE = ("Source ID", "Message Text")  # INVOICE DETAILS

N_CM_PO_ID, N_CM_BU, N_CM_AMT = (norm_col(c) for c in COL_CM_PO)
N_CM_RCT, N_CM_GR_PO, N_CM_GR_CO = (norm_col(c) for c in COL_CM_GR)

N_ME_ORDER, N_ME_CO, N_ME_AMT, N_ME_IMPORTED_TS = (norm_col(c) for c in COL_ME_PO)
N_ME_NOTE, N_ME_GR_CO, N_ME_GR_IMPORTED_TS = (norm_col(c) for c in COL_ME_GR)

N_SRC_TXN, N_MSG = (norm_col(c) for c in COL_LIF2_TRANSFER)
N_SRC_ID = norm_col(COL_LIF2_INVOICE[0])

PO_COLUMNS = [
    "Cmona PO ID",
    "Company Code",
    "CMONA Amount",
    "Medius Order Identifier",
    "Purchase Order Company",
    "Medius Amount",
    "Purchase Order: Imported Timestamp",
    "PO Match Status",
    "Amount Match Status",
    "Action Required",
    "Comments",
]

GR_COLUMNS = [
    "Cmona [RCT] Receipt Id",
    "Cmona [PO-P&I] PO Id",
    "Cmona [PO-P&I]Purchasing Company (Purchase Organization Id)",
    "Medius Delivery Note (Extracted Text)",
    "Medius Company Code",
    "Purchase Order: Imported Timestamp",
    "GR Match Status",
    "Action Required",
    "Comments",
]


# --------------------------------------------------------------------------- #
# Input file definitions + filename-pattern auto-detection
# --------------------------------------------------------------------------- #
_PO_WORD = r"(?<![a-z])po(?![a-z])|purchase[ _\-]?order"
_GR_WORD = r"(?<![a-z])gr(?![a-z])|goods[ _\-]?receipt|receipt"


@dataclass(frozen=True)
class FileSpec:
    key: str
    prompt: str
    include: tuple[str, ...]  # ALL of these regexes must match the lower-case file stem
    exclude: tuple[str, ...]  # NONE of these may match
    required: tuple[str, ...]


FILE_SPECS: tuple[FileSpec, ...] = (
    FileSpec(
        "cmona_po",
        "Source CMONA PO report",
        ("cmona", _PO_WORD),
        (_GR_WORD,),
        COL_CM_PO,
    ),
    FileSpec("cmona_gr", "Source CMONA GR report", ("cmona", _GR_WORD), (), COL_CM_GR),
    FileSpec(
        "medius_po", "Medius PO report", ("medius", _PO_WORD), (_GR_WORD,), COL_ME_PO
    ),
    FileSpec("medius_gr", "Medius GR report", ("medius", _GR_WORD), (), COL_ME_GR),
    FileSpec(
        "lif2", "Lif2 error report", (r"lif[ _\-]?2|error",), (), COL_LIF2_TRANSFER
    ),
)


def matches_spec(path: Path, spec: FileSpec) -> bool:
    """True when the file name (any extension) fits the spec's filename pattern."""
    stem = path.stem.lower()
    return all(re.search(p, stem) for p in spec.include) and not any(
        re.search(p, stem) for p in spec.exclude
    )


def find_candidates(spec: FileSpec, folder: Path) -> list[Path]:
    """Files in `folder` matching the pattern, newest first."""
    if not folder.is_dir():
        return []
    paths = [Path(p) for p in glob.glob(str(Path(glob.escape(str(folder))) / "*"))]
    hits = [
        p
        for p in paths
        if p.is_file() and not p.name.startswith("~$") and matches_spec(p, spec)
    ]
    return sorted(hits, key=lambda p: p.stat().st_mtime, reverse=True)


class _DialogUnavailable(Exception):
    """tkinter file dialog cannot be used (headless / not installed)."""


def _dialog_pick(title: str, initial_dir: Path) -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
    except Exception as exc:  # ImportError, TclError ...
        raise _DialogUnavailable(str(exc)) from exc
    try:
        return filedialog.askopenfilename(
            parent=root,
            title=title,
            initialdir=str(initial_dir if initial_dir.is_dir() else Path.cwd()),
            filetypes=[
                ("Excel / report files", "*.xls *.xlsx *.xlsm *.csv *.htm *.html"),
                ("All files", "*.*"),
            ],
        )
    finally:
        root.destroy()


def prompt_for_file(
    index: int, spec: FileSpec, folder: Path, use_dialog: bool
) -> Path | None:
    """Ask the user for ONE file. Returns None when nothing was chosen."""
    label = f"{index}. {spec.prompt}"
    print(f"\nPlease upload: {label}")
    chosen = ""
    if use_dialog:
        try:
            chosen = _dialog_pick(f"Upload {label}", folder)
        except _DialogUnavailable as exc:
            log.warning(
                "File dialog unavailable (%s) - falling back to console input", exc
            )
            use_dialog = False
    if not use_dialog:
        chosen = (
            input("  Full file path (Enter = auto-detect from input folder): ")
            .strip()
            .strip('"')
            .strip("'")
        )
    return Path(chosen) if chosen else None


def collect_input_files(
    input_dir: Path, auto: bool, use_dialog: bool
) -> dict[str, Path]:
    """Collect the 5 reports in the required order (prompted, or auto-detected with --auto)."""
    chosen: dict[str, Path] = {}
    for idx, spec in enumerate(FILE_SPECS, start=1):
        candidates = find_candidates(spec, input_dir)
        path: Path | None
        if auto:
            path = candidates[0] if candidates else None
        else:
            path = prompt_for_file(idx, spec, input_dir, use_dialog)
            if path is None and candidates:
                path = candidates[0]
                log.warning(
                    "Nothing selected for '%s' - using auto-detected file %s",
                    spec.prompt,
                    path.name,
                )
        if path is None:
            raise FileNotFoundError(
                f"{spec.prompt}: no file selected and no file matching the filename pattern found in '{input_dir}'."
            )
        if not path.is_file():
            raise FileNotFoundError(f"{spec.prompt}: file not found: {path}")
        if not matches_spec(path, spec):
            log.warning(
                "'%s' does not look like a '%s' by filename - content will be validated",
                path.name,
                spec.prompt,
            )
        for other_spec in FILE_SPECS:
            other = chosen.get(other_spec.key)
            if other is not None and other.resolve() == path.resolve():
                raise ValueError(
                    f"The same file '{path.name}' was selected for '{spec.prompt}' and for '{other_spec.prompt}'."
                )
        chosen[spec.key] = path
        log.info("[%d/5] %-24s -> %s", idx, spec.prompt, path)
    return chosen


# --------------------------------------------------------------------------- #
# Reading any report: xlsx / xls / csv / HTML-as-xls, header auto-detection, cleaning
# --------------------------------------------------------------------------- #
def sniff_kind(path: Path) -> str:
    """Detect real format from content (the extension can be anything)."""
    with path.open("rb") as fh:
        head = fh.read(4096)
    if head.startswith(b"PK\x03\x04"):
        return "xlsx"
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        return "xls"
    text = head.decode("utf-8-sig", errors="ignore").lstrip().lower()
    if "urn:schemas-microsoft-com:office:spreadsheet" in text:
        raise ValueError(
            f"'{path.name}' is an XML Spreadsheet 2003 file - please re-save it as .xlsx."
        )
    if text.startswith(
        ("<!doctype", "<html", "<table", "<head", "<body", "<meta", "<?xml")
    ):
        return "html"
    return "text"


def _read_text(path: Path) -> str:
    data = path.read_bytes()
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("cp1252", errors="replace")


def _rows_to_frame(rows: list[list[str]]) -> pd.DataFrame:
    width = max((len(r) for r in rows), default=0)
    return pd.DataFrame([r + [None] * (width - len(r)) for r in rows], dtype=object)


class _TableParser(HTMLParser):
    """Minimal HTML table extractor (stdlib only); text outside cells is ignored."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.tables: list[list[list[str]]] = []
        self._row: list[str] | None = None
        self._cell: list[str] | None = None
        self._skip = 0

    def handle_starttag(self, tag: str, attrs: list) -> None:
        if tag in ("script", "style"):
            self._skip += 1
        elif tag == "table":
            self.tables.append([])
        elif tag == "tr" and self.tables:
            self._row = []
        elif tag in ("td", "th") and self._row is not None:
            self._cell = []
        elif tag == "br" and self._cell is not None:
            self._cell.append(" ")

    def handle_endtag(self, tag: str) -> None:
        if tag in ("script", "style"):
            self._skip = max(0, self._skip - 1)
        elif tag in ("td", "th") and self._cell is not None and self._row is not None:
            self._row.append("".join(self._cell))
            self._cell = None
        elif tag == "tr" and self._row is not None and self.tables:
            self.tables[-1].append(self._row)
            self._row = None

    def handle_data(self, data: str) -> None:
        if not self._skip and self._cell is not None:
            self._cell.append(data)


def clean_cell(value: object) -> str:
    """Strip HTML tags/entities, non-breaking spaces and extra whitespace from one cell."""
    text = _TAG_RE.sub(" ", str(value))
    if _ENTITY_RE.search(text):
        text = html.unescape(text)
    return _WS_RE.sub(" ", text.replace("\xa0", " ")).strip()


def clean_text_series(s: pd.Series) -> pd.Series:
    """Trim, drop HTML, and replace every empty value with 'NA'."""
    out = s.astype("string").str.replace(_TAG_PATTERN, " ", regex=True)
    has_entity = out.str.contains(_ENTITY_PATTERN, regex=True, na=False)
    if has_entity.any():
        out.loc[has_entity] = [html.unescape(v) for v in out.loc[has_entity]]
    out = (
        out.str.replace("\xa0", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    is_null = out.isna() | out.str.lower().isin(NULL_TOKENS)
    return out.mask(is_null, NA).astype(object)


def detect_header_row(raw: pd.DataFrame, required: Sequence[str]) -> tuple[int, int]:
    """Row index of the header: first row (column A filled) containing the most required columns."""
    req = {norm_col(c) for c in required}
    best_idx, best_hits = -1, 0
    for i, row in enumerate(
        raw.iloc[:HEADER_SCAN_ROWS].itertuples(index=False, name=None)
    ):
        first = row[0] if row else None
        if not isinstance(first, str) or not clean_cell(first):
            continue  # headers always start in column A
        cells = {norm_col(clean_cell(v)) for v in row if isinstance(v, str)}
        hits = len(req & cells)
        if hits > best_hits:
            best_idx, best_hits = i, hits
            if hits == len(req):
                break
    return best_idx, best_hits


def frame_from_raw(
    raw: pd.DataFrame, required: Sequence[str], label: str, sheet: str
) -> pd.DataFrame:
    """Raw grid -> clean DataFrame with standardised column names (required columns only)."""
    if raw.empty:
        raise ValueError(f"{label} [{sheet}]: sheet is empty.")
    idx, _ = detect_header_row(raw, required)
    if idx < 0:
        raise ValueError(
            f"{label} [{sheet}]: none of the required columns {list(required)} were found in the first "
            f"{HEADER_SCAN_ROWS} rows (headers must start in column A)."
        )
    header = [
        clean_cell(v) if isinstance(v, str) else "" for v in raw.iloc[idx].tolist()
    ]
    names = [norm_col(h) for h in header]
    missing = [c for c in required if norm_col(c) not in names]
    if missing:
        raise ValueError(
            f"{label} [{sheet}]: missing required column(s) {missing}. "
            f"Columns found on header row {idx + 1}: {[h for h in header if h]}"
        )
    position = {
        n: i for i, n in reversed(list(enumerate(names))) if n
    }  # first duplicate wins
    body = raw.iloc[idx + 1 :]
    df = pd.DataFrame(
        {
            norm_col(c): clean_text_series(body.iloc[:, position[norm_col(c)]])
            for c in required
        }
    )
    df = df.reset_index(drop=True)
    before = len(df)
    df = df.loc[(df != NA).any(axis=1)].reset_index(drop=True)
    log.info(
        "%s [%s]: header on row %d, %d data rows (%d blank rows removed)",
        label,
        sheet,
        idx + 1,
        len(df),
        before - len(df),
    )
    return df


def _select_sheets(names: list[str], wanted: str | None, label: str) -> list[str]:
    if not wanted:
        return names
    squash = lambda t: _WS_RE.sub(" ", t).strip().lower()  # noqa: E731
    lookup = {squash(n): n for n in names}
    if squash(wanted) not in lookup:
        raise ValueError(
            f"{label}: sheet '{wanted}' not found. Sheets in file: {names}"
        )
    return [lookup[squash(wanted)]]


def load_table(
    path: Path, label: str, required: Sequence[str], sheet_name: str | None = None
) -> pd.DataFrame:
    """Read any supported report, detect the header row, clean it and validate required columns."""
    try:
        kind = sniff_kind(path)
        if kind in ("xlsx", "xls"):
            engine = "openpyxl" if kind == "xlsx" else "xlrd"
            problems: list[str] = []
            with pd.ExcelFile(path, engine=engine) as xl:
                for name in _select_sheets(
                    [str(n) for n in xl.sheet_names], sheet_name, label
                ):
                    raw = xl.parse(name, header=None, dtype=str)
                    try:
                        return frame_from_raw(raw, required, label, name)
                    except ValueError as exc:
                        problems.append(str(exc))
                    finally:
                        del raw
            raise ValueError(
                f"{label}: no sheet in '{path.name}' has the required columns.\n  "
                + "\n  ".join(problems)
            )
        if sheet_name:
            raise ValueError(
                f"{label}: '{path.name}' is not a workbook, cannot read sheet '{sheet_name}'."
            )
        if kind == "html":
            parser = _TableParser()
            parser.feed(_read_text(path))
            tables = [t for t in parser.tables if t]
            if not tables:
                raise ValueError(f"{label}: no table found in HTML file '{path.name}'.")
            raw = _rows_to_frame(max(tables, key=len))
        else:
            text = _read_text(path)
            try:
                delim = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|").delimiter
            except csv.Error:
                delim = ","
            raw = _rows_to_frame(list(csv.reader(io.StringIO(text), delimiter=delim)))
        return frame_from_raw(raw, required, label, "Sheet1")
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError(
            f"{label}: cannot read '{path.name}': {type(exc).__name__}: {exc}"
        ) from exc


# --------------------------------------------------------------------------- #
# Value helpers
# --------------------------------------------------------------------------- #
def parse_amount_value(text: str) -> float:
    """'1,234.56' / '1.234,56' / '(12.5)' / '12,5-' / '$ 1 234,50' -> float (NaN if unparsable)."""
    s = str(text).strip()
    if not s or s == NA:
        return float("nan")
    if _SCI_RE.fullmatch(s):
        return float(s)
    negative = "-" in s or (s.startswith("(") and s.endswith(")"))
    s = re.sub(r"[^0-9,.]", "", s)
    if not re.search(r"\d", s):
        return float("nan")
    last_comma, last_dot = s.rfind(","), s.rfind(".")
    if last_comma >= 0 and last_dot >= 0:
        if last_comma > last_dot:  # 1.234,56
            s = s.replace(".", "").replace(",", ".")
        else:  # 1,234.56
            s = s.replace(",", "")
    elif last_comma >= 0:
        if (
            s.count(",") == 1 and len(s) - last_comma - 1 != 3
        ):  # 12,5 or 12,50 -> decimal comma
            s = s.replace(",", ".")
        else:  # 1,234 or 1,234,567 -> thousands
            s = s.replace(",", "")
    elif s.count(".") > 1:  # 1.234.567
        s = s.replace(".", "")
    try:
        value = float(s)
    except ValueError:
        return float("nan")
    return -value if negative else value


def parse_amount(series: pd.Series, label: str = "amount") -> pd.Series:
    """Vectorised-by-unique-value conversion of an amount column to float64."""
    mapping = {u: parse_amount_value(u) for u in series.unique()}
    out = series.map(mapping).astype("float64")
    bad = int((out.isna() & (series != NA)).sum())
    if bad:
        log.warning(
            "%s: %d value(s) could not be converted to a number and were ignored",
            label,
            bad,
        )
    return out


def normalize_id(series: pd.Series) -> pd.Series:
    """Matching key: trimmed, upper-case, Excel '123.0' -> '123'."""
    return (
        series.astype(str).str.strip().str.replace(r"\.0+$", "", regex=True).str.upper()
    )


def po_without_suffix(series: pd.Series) -> pd.Series:
    """Remove the first '-' and everything after it: 'PO123-V2' -> 'PO123'."""
    out = series.astype(str).str.split("-", n=1).str[0].str.strip()
    return out.mask(out.str.lower().isin(NULL_TOKENS), NA)


def company_prefix(series: pd.Series) -> pd.Series:
    """Text before the first '-' ('1000-Company A' -> '1000')."""
    out = series.astype(str).str.split("-", n=1).str[0].str.strip()
    return out.mask(out.isin(["", NA]), NA)


def _truncate(text: str) -> str:
    return text if len(text) <= MAX_CELL_CHARS else text[: MAX_CELL_CHARS - 3] + "..."


# --------------------------------------------------------------------------- #
# PO reconciliation (pure functions)
# --------------------------------------------------------------------------- #
def aggregate_cmona_po(df: pd.DataFrame) -> pd.DataFrame:
    """One row per suffix-free PO Id: summed amount + first company code."""
    cmona_ids = po_without_suffix(df[N_CM_PO_ID])

    work = pd.DataFrame(
        {
            "key": normalize_id(cmona_ids),
            "cmona_id": cmona_ids,
            "cmona_company": df[N_CM_BU].mask(df[N_CM_BU] == NA),
            "cmona_amount": parse_amount(df[N_CM_AMT], "CMONA PO amount"),
        }
    )
    work = work.loc[work["key"] != NA]

    agg = (
        work.groupby("key", sort=False)
        .agg(
            cmona_id=("cmona_id", "first"),
            cmona_company=("cmona_company", "first"),
            cmona_amount=("cmona_amount", "sum"),
        )
        .reset_index()
    )
    agg["cmona_company"] = agg["cmona_company"].fillna(NA)
    return agg


def aggregate_medius_po(df: pd.DataFrame) -> pd.DataFrame:
    """One row per suffix-free Order Identifier 2."""
    medius_ids = po_without_suffix(df[N_ME_ORDER])

    work = pd.DataFrame(
        {
            "key": normalize_id(medius_ids),
            "medius_id": medius_ids,
            "medius_company": company_prefix(df[N_ME_CO]).mask(lambda s: s == NA),
            "medius_amount": parse_amount(df[N_ME_AMT], "Medius Ordered Amount"),
            "medius_imported_timestamp": df[N_ME_IMPORTED_TS],
        }
    )
    work = work.loc[work["key"] != NA]

    agg = (
        work.groupby("key", sort=False)
        .agg(
            medius_id=("medius_id", "first"),
            medius_company=("medius_company", "first"),
            medius_amount=("medius_amount", "sum"),
            medius_imported_timestamp=("medius_imported_timestamp", "first"),
        )
        .reset_index()
    )

    agg["medius_company"] = agg["medius_company"].fillna(NA)
    agg["medius_imported_timestamp"] = agg["medius_imported_timestamp"].fillna(NA)

    return agg


def build_error_map(df: pd.DataFrame, key_col: str = N_SRC_TXN) -> pd.Series:
    """Lif2 key column (Source Transaction ID / Source ID) -> joined Message Text."""
    work = pd.DataFrame({"key": normalize_id(df[key_col]), "msg": df[N_MSG]})
    work = work.loc[work["key"] != NA].drop_duplicates()
    if work.empty:
        return pd.Series(dtype=object)
    return work.groupby("key", sort=False)["msg"].agg(" | ".join)


def _po_comment(action: str, t_msg: object, i_msg: object) -> str:
    """Return all matching Lif2 messages, one sheet/message per line."""
    if action != "Checked":
        return ""

    messages: list[str] = []

    if isinstance(t_msg, str) and t_msg.strip():
        messages.append(f"{SHEET_TRANSFER}: {t_msg}")

    if isinstance(i_msg, str) and i_msg.strip():
        messages.append(f"{SHEET_INVOICE}: {i_msg}")

    return _truncate("\n".join(messages))


def reconcile_po(
    cm: pd.DataFrame,
    me: pd.DataFrame,
    transfer: pd.Series,
    invoice: pd.Series,
    tolerance: float = DEFAULT_TOLERANCE,
) -> pd.DataFrame:
    """CMONA is the master: every CMONA PO appears exactly once in the result."""
    m = cm.merge(me, on="key", how="left")
    found = m["medius_id"].notna().to_numpy()
    diff = (m["cmona_amount"] - m["medius_amount"]).round(6)
    amount_ok = (diff.abs() <= tolerance + 1e-9).to_numpy()
    t_msg = (
        m["key"].map(transfer)
        if len(transfer)
        else pd.Series(np.nan, index=m.index, dtype=object)
    )
    i_msg = (
        m["key"].map(invoice)
        if len(invoice)
        else pd.Series(np.nan, index=m.index, dtype=object)
    )
    # Check both Lif2 sheets independently for unmatched POs.
    in_i = ~found & i_msg.notna().to_numpy()
    in_t = ~found & t_msg.notna().to_numpy()

    po_status = np.where(found, "TRUE", NA)
    amt_status = np.select([~found, amount_ok], [NA, "TRUE"], default="FALSE")
    action = np.select(
        [found & amount_ok, found & ~amount_ok, in_t | in_i],
        ["Not Required", "Amount Mismatch", "Checked"],
        default="Need Investigation",
    )
    comments = [
        _po_comment(a, t, i)
        for a, t, i in zip(action, t_msg.to_numpy(), i_msg.to_numpy())
    ]
    return pd.DataFrame(
        {
            PO_COLUMNS[0]: m["cmona_id"],
            PO_COLUMNS[1]: m["cmona_company"],
            PO_COLUMNS[2]: m["cmona_amount"].round(2),
            PO_COLUMNS[3]: m["medius_id"].where(found, NA),
            PO_COLUMNS[4]: m["medius_company"].where(found, NA),
            PO_COLUMNS[5]: (
                m["medius_amount"].round(2).astype(object).where(found, NA)
            ),
            PO_COLUMNS[6]: m["medius_imported_timestamp"].where(found, NA),
            PO_COLUMNS[7]: po_status,
            PO_COLUMNS[8]: amt_status,
            PO_COLUMNS[9]: action,
            PO_COLUMNS[10]: comments,
        }
    )


# --------------------------------------------------------------------------- #
# GR reconciliation (pure functions)
# --------------------------------------------------------------------------- #
def _join_unique(values: pd.Series) -> str:
    items = [v for v in dict.fromkeys(values) if v != NA]
    return ", ".join(items) if items else NA


def prepare_cmona_gr(df: pd.DataFrame) -> pd.DataFrame:
    """One row per unique [RCT] Receipt Id (PO Ids / companies joined when several)."""
    work = pd.DataFrame(
        {
            "key": normalize_id(df[N_CM_RCT]),
            "receipt": df[N_CM_RCT],
            "po": po_without_suffix(df[N_CM_GR_PO]),
            "company": df[N_CM_GR_CO],
        }
    )
    work = work.loc[work["key"] != NA].drop_duplicates()
    agg = (
        work.groupby("key", sort=False)
        .agg(
            receipt=("receipt", "first"),
            po=("po", _join_unique),
            company=("company", _join_unique),
        )
        .reset_index()
    )
    agg["order"] = np.arange(len(agg))
    return agg


def _char_class(ch: str) -> str:
    return "d" if ch.isdigit() else ("a" if ch.isalpha() else "")


def _boundary_ok(text: str, start: int, end: int) -> bool:
    """Reject a hit that is glued to more digits (or letters) of the same kind: '123' inside '1234'."""
    for edge, neighbour_idx in ((text[start], start - 1), (text[end - 1], end)):
        if 0 <= neighbour_idx < len(text):
            cls = _char_class(edge)
            if cls and cls == _char_class(text[neighbour_idx]):
                return False
    return True


def find_receipts_in_note(
    note: str, keys: set[str], lengths: Sequence[int]
) -> list[tuple[str, str]]:
    """All CMONA receipt ids contained in a Delivery Note -> [(KEY, text as written in the note)]."""
    upper = note.upper()
    n = len(upper)
    same_len = len(note) == n
    hits: list[tuple[str, str]] = []
    for length in lengths:  # lengths sorted ascending
        if length > n:
            break
        for i in range(n - length + 1):
            cand = upper[i : i + length]
            if cand in keys and _boundary_ok(upper, i, i + length):
                hits.append((cand, note[i : i + length] if same_len else cand))
    return hits


def reconcile_gr(
    cm_gr: pd.DataFrame,
    me_gr: pd.DataFrame,
    transfer: pd.Series | None = None,
    invoice: pd.Series | None = None,
) -> pd.DataFrame:
    """CMONA GR is the master: every unique receipt appears once; match = id found inside a Delivery Note."""
    keys = set(cm_gr["key"])
    lengths = sorted({len(k) for k in keys})
    pairs = pd.DataFrame(
        {
            "note": me_gr[N_ME_NOTE],
            "company": company_prefix(me_gr[N_ME_GR_CO]),
            "imported_timestamp": me_gr[N_ME_GR_IMPORTED_TS],
        }
    ).drop_duplicates()
    pairs = pairs.loc[pairs["note"] != NA]
    rows: list[tuple[str, str, str, str, str]] = []
    for note, company, imported_timestamp in zip(
        pairs["note"].to_numpy(),
        pairs["company"].to_numpy(),
        pairs["imported_timestamp"].to_numpy(),
    ):
        for key, extracted in find_receipts_in_note(note, keys, lengths):
            rows.append(
                (
                    key,
                    extracted,
                    company,
                    note,
                    imported_timestamp,
                )
            )

    hits = pd.DataFrame(
        rows,
        columns=[
            "key",
            "extracted",
            "medius_company",
            "note",
            "medius_imported_timestamp",
        ],
    )
    m = cm_gr.merge(hits, on="key", how="left")
    cm_companies = [
        {c.strip().upper() for c in str(v).split(",")} for v in m["company"]
    ]
    m["company_ok"] = [
        isinstance(mc, str) and mc.upper() in s
        for mc, s in zip(m["medius_company"], cm_companies)
    ]
    m = (
        m.sort_values(
            "company_ok", ascending=False, kind="stable"
        )  # prefer the same-company hit
        .drop_duplicates("key")
        .sort_values("order")
    )
    found = m["extracted"].notna().to_numpy()

    # ---- Lif2 error report lookup for unmatched GRs (same logic as PO sheet)
    transfer = transfer if transfer is not None else pd.Series(dtype=object)
    invoice = invoice if invoice is not None else pd.Series(dtype=object)
    t_msg = (
        m["key"].map(transfer)
        if len(transfer)
        else pd.Series(np.nan, index=m.index, dtype=object)
    )
    i_msg = (
        m["key"].map(invoice)
        if len(invoice)
        else pd.Series(np.nan, index=m.index, dtype=object)
    )
    # Check both Lif2 sheets independently for unmatched GRs.
    in_i = ~found & i_msg.notna().to_numpy()
    in_t = ~found & t_msg.notna().to_numpy()

    action = np.select(
        [found, in_t | in_i],
        ["Not Required", "Checked"],
        default="Need Investigation",  # 3) not found in either sheet
    )
    comments = [  # 4) comment only for "Checked", blank otherwise
        _po_comment(a, t, i)
        for a, t, i in zip(action, t_msg.to_numpy(), i_msg.to_numpy())
    ]

    return pd.DataFrame(
        {
            GR_COLUMNS[0]: m["receipt"].to_numpy(),
            GR_COLUMNS[1]: m["po"].to_numpy(),
            GR_COLUMNS[2]: m["company"].to_numpy(),
            GR_COLUMNS[3]: m["extracted"].where(found, NA).to_numpy(),
            GR_COLUMNS[4]: (m["medius_company"].where(found, NA).fillna(NA).to_numpy()),
            GR_COLUMNS[5]: (
                m["medius_imported_timestamp"].where(found, NA).fillna(NA).to_numpy()
            ),
            GR_COLUMNS[6]: np.where(found, "TRUE", NA),
            GR_COLUMNS[7]: action,
            GR_COLUMNS[8]: comments,
        }
    )


# --------------------------------------------------------------------------- #
# Summary
# --------------------------------------------------------------------------- #
def build_po_summary(
    cm_df: pd.DataFrame, me_df: pd.DataFrame, po: pd.DataFrame
) -> pd.DataFrame:
    cm_keys = normalize_id(po_without_suffix(cm_df[N_CM_PO_ID]))
    me_keys = normalize_id(po_without_suffix(me_df[N_ME_ORDER]))

    metrics = [
        ("Total CMONA Rows", len(cm_df)),
        ("Total CMONA PO Count", int((cm_keys != NA).sum())),
        ("Distinct CMONA PO Count", int(cm_keys[cm_keys != NA].nunique())),
        ("Total Medius Rows", len(me_df)),
        ("Total Medius PO Count", int((me_keys != NA).sum())),
        ("Distinct Medius PO Count", int(me_keys[me_keys != NA].nunique())),
        ("Matched POs", int((po["PO Match Status"] == "TRUE").sum())),
        ("Unmatched POs", int((po["PO Match Status"] == NA).sum())),
        ("Amount Matches", int((po["Amount Match Status"] == "TRUE").sum())),
        ("Amount Mismatches", int((po["Amount Match Status"] == "FALSE").sum())),
        ("Error History Matches", int((po["Action Required"] == "Checked").sum())),
        (
            "Pending Investigations",
            int((po["Action Required"] == "Need Investigation").sum()),
        ),
    ]
    return pd.DataFrame(metrics, columns=["Metric", "Value"])


# --------------------------------------------------------------------------- #
# Excel output
# --------------------------------------------------------------------------- #
_HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
_CM_FILL = PatternFill("solid", fgColor="2E75B6")
_ME_FILL = PatternFill("solid", fgColor="548235")
_GREY_FILL = PatternFill("solid", fgColor="7F7F7F")
_WHITE_BOLD = Font(bold=True, color="FFFFFF")
_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _escape_formulas(df: pd.DataFrame, columns: Sequence[str]) -> pd.DataFrame:
    """Text starting with '=' would be written as an Excel formula - neutralise it."""
    out = df.copy()
    for col in columns:
        s = out[col].astype(str)
        out[col] = s.where(~s.str.startswith("="), "'" + s)
    return out


def _style_header(
    ws, row: int, first_col: int, last_col: int, fill: PatternFill = _HEAD_FILL
) -> None:
    for col in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font, cell.fill, cell.border = _WHITE_BOLD, fill, _BORDER
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )


def _autosize(ws, header_row: int, max_width: int = 60, sample: int = 2000) -> None:
    last_row = min(ws.max_row, header_row + sample)
    for cells in ws.iter_cols(min_row=header_row, max_row=last_row):
        width = max(
            (len(str(c.value)) for c in cells if c.value is not None), default=8
        )
        ws.column_dimensions[cells[0].column_letter].width = min(
            max(width + 2, 10), max_width
        )


def _check_row_limit(name: str, rows: int) -> None:
    if rows + 2 > EXCEL_MAX_ROWS:
        raise ValueError(
            f"Sheet '{name}' would have {rows:,} rows - more than Excel's limit of {EXCEL_MAX_ROWS:,}."
        )


def write_report(
    path: Path,
    po: pd.DataFrame,
    gr: pd.DataFrame,
    summary: pd.DataFrame,
    execution: pd.DataFrame,
) -> Path:
    _check_row_limit(SHEET_PO, len(po))
    _check_row_limit(SHEET_GR, len(gr))
    path.parent.mkdir(parents=True, exist_ok=True)
    po_out = _escape_formulas(po, [PO_COLUMNS[10]])
    gr_out = _escape_formulas(gr, [GR_COLUMNS[3], GR_COLUMNS[8]])

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # ---- Sheet 1: combined Cmona / Medius group header on row 1, column headers on row 2
        po_out.to_excel(writer, sheet_name=SHEET_PO, index=False, startrow=1)
        ws = writer.sheets[SHEET_PO]
        ws.merge_cells("A1:C1")
        ws.merge_cells("D1:G1")
        ws["A1"], ws["D1"] = "Cmona", "Medius"

        _style_header(ws, 1, 1, 3, _CM_FILL)
        _style_header(ws, 1, 4, 7, _ME_FILL)
        _style_header(ws, 1, 8, 11, _GREY_FILL)

        _style_header(ws, 2, 1, len(PO_COLUMNS))
        ws.auto_filter.ref = (
            f"A2:{ws.cell(row=2, column=len(PO_COLUMNS)).column_letter}{len(po) + 2}"
        )
        ws.freeze_panes = "A3"
        _autosize(ws, 2)
        ws.column_dimensions["K"].width = 70
        # Wrap multi-line comments in PO Reconciliation.
        for row in range(3, len(po) + 3):
            ws.cell(row=row, column=11).alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

        if len(po) <= STYLE_CELL_LIMIT:
            for r in range(3, len(po) + 3):
                for c in (3, 6):
                    ws.cell(row=r, column=c).number_format = "#,##0.00"

        # ---- Sheet 2: Cmona / Medius group header on row 1
        gr_out.to_excel(
            writer,
            sheet_name=SHEET_GR,
            index=False,
            startrow=1,
        )
        ws = writer.sheets[SHEET_GR]

        ws.merge_cells("A1:C1")
        ws.merge_cells("D1:F1")
        ws["A1"], ws["D1"] = "Cmona", "Medius"

        _style_header(ws, 1, 1, 3, _CM_FILL)
        _style_header(ws, 1, 4, 6, _ME_FILL)
        _style_header(ws, 1, 7, 9, _GREY_FILL)
        _style_header(ws, 2, 1, len(GR_COLUMNS))

        ws.auto_filter.ref = (
            f"A2:{ws.cell(row=2, column=len(GR_COLUMNS)).column_letter}{len(gr) + 2}"
        )
        ws.freeze_panes = "A3"
        _autosize(ws, 2)
        ws.column_dimensions["I"].width = 70
        # Wrap multi-line comments in GR Reconciliation.
        for row in range(3, len(gr) + 3):
            ws.cell(row=row, column=9).alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

        # ---- Sheet 3: metrics + execution summary
        summary.to_excel(writer, sheet_name=SHEET_SUMMARY, index=False)
        execution.to_excel(
            writer, sheet_name=SHEET_SUMMARY, index=False, startrow=len(summary) + 3
        )
        ws = writer.sheets[SHEET_SUMMARY]
        _style_header(ws, 1, 1, 2)
        _style_header(ws, len(summary) + 4, 1, 2)
        _autosize(ws, 1, max_width=90)
        for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
            row[0].alignment = Alignment(horizontal="left", wrap_text=True)
    return path


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #
def add_arguments(p: argparse.ArgumentParser) -> None:
    p.add_argument(
        "--auto",
        action="store_true",
        help="do not ask for files; use the newest file in the input folder matching each filename pattern",
    )
    p.add_argument(
        "--output-dir",
        default=None,
        help="where to save the report (default: your Downloads folder)",
    )
    p.add_argument(
        "--input-dir",
        default=None,
        help="folder opened by the file picker / searched by --auto (default: input\\)",
    )
    p.add_argument(
        "--tolerance",
        type=float,
        default=DEFAULT_TOLERANCE,
        help="amount tolerance (default 0.01)",
    )
    p.add_argument(
        "--no-dialog",
        action="store_true",
        help="type file paths in the console instead of the file picker",
    )


def _resolve_dirs(args) -> tuple[Path, Path, Path]:
    try:
        s = get_settings()
        input_dir, output_dir, log_dir = (
            Path(s.input_dir),
            Path(s.output_dir),
            Path(s.log_dir),
        )
    except (
        Exception
    ) as exc:  # settings unavailable -> work relative to the current folder
        log.warning(
            "Could not load settings (%s) - using folders relative to %s",
            exc,
            Path.cwd(),
        )
        input_dir, output_dir, log_dir = Path("input"), Path("output"), Path("logs")
    if getattr(args, "input_dir", None):
        input_dir = Path(args.input_dir)
    output_dir = (
        Path(args.output_dir)
        if getattr(args, "output_dir", None)
        else Path.home() / "Downloads"
    )
    return input_dir, output_dir, log_dir


def _execute(args, started: datetime) -> JobResult:
    t0 = time.perf_counter()
    tolerance = float(getattr(args, "tolerance", DEFAULT_TOLERANCE))
    input_dir, output_dir, _ = _resolve_dirs(args)

    # 1) files (sequence 1-5), then everything below runs unattended
    files = collect_input_files(
        input_dir,
        bool(getattr(args, "auto", False)),
        not getattr(args, "no_dialog", False),
    )

    # 2) read + clean + validate
    spec = {s.key: s for s in FILE_SPECS}
    cm_po = load_table(files["cmona_po"], spec["cmona_po"].prompt, COL_CM_PO)
    cm_gr = load_table(files["cmona_gr"], spec["cmona_gr"].prompt, COL_CM_GR)
    me_po = load_table(files["medius_po"], spec["medius_po"].prompt, COL_ME_PO)
    me_gr = load_table(files["medius_gr"], spec["medius_gr"].prompt, COL_ME_GR)
    transfer = build_error_map(
        load_table(
            files["lif2"], "Lif2 error report", COL_LIF2_TRANSFER, SHEET_TRANSFER
        ),
        N_SRC_TXN,
    )
    invoice = build_error_map(
        load_table(files["lif2"], "Lif2 error report", COL_LIF2_INVOICE, SHEET_INVOICE),
        N_SRC_ID,
    )

    # 3) reconcile
    po = reconcile_po(
        aggregate_cmona_po(cm_po),
        aggregate_medius_po(me_po),
        transfer,
        invoice,
        tolerance,
    )
    cm_gr_unique = prepare_cmona_gr(cm_gr)
    gr = reconcile_gr(cm_gr_unique, me_gr, transfer, invoice)
    summary = build_po_summary(cm_po, me_po, po)

    # 4) execution summary (logs + report)
    finished = datetime.now()
    gr_matched = int((gr["GR Match Status"] == "TRUE").sum())
    exec_rows = [
        ("Run started", f"{started:%Y-%m-%d %H:%M:%S}"),
        ("Run finished", f"{finished:%Y-%m-%d %H:%M:%S}"),
        ("Duration (seconds)", f"{time.perf_counter() - t0:.1f}"),
        ("Amount tolerance", f"{tolerance}"),
        *[(f"Input - {s.prompt}", files[s.key].name) for s in FILE_SPECS],
        ("GR - unique CMONA receipts", len(gr)),
        ("GR - matched", gr_matched),
        ("GR - unmatched", len(gr) - gr_matched),
        (
            "GR - unmatched, checked (Lif2)",
            int((gr["Action Required"] == "Checked").sum()),
        ),
        (
            "GR - unmatched, need investigation",
            int((gr["Action Required"] == "Need Investigation").sum()),
        ),
    ]
    execution = pd.DataFrame(exec_rows, columns=["Execution Summary", "Value"]).astype(
        {"Value": str}
    )
    for _, row in pd.concat(
        [
            summary.astype({"Value": str}),
            execution.set_axis(["Metric", "Value"], axis=1),
        ]
    ).iterrows():
        log.info("%-34s %s", row["Metric"], row["Value"])

    # 5) write report straight into the Downloads folder
    out_path = output_dir / f"{REPORT_PREFIX}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    write_report(out_path, po, gr, summary, execution)
    log.info("Report written: %s", out_path)

    text = (
        f"POs: {len(po)} | matched {s_val(summary, 'Matched POs')} | unmatched {s_val(summary, 'Unmatched POs')} "
        f"(checked {s_val(summary, 'Error History Matches')}, need investigation {s_val(summary, 'Pending Investigations')}) | "
        f"amount mismatches {s_val(summary, 'Amount Mismatches')} | GR receipts {len(gr)} (matched {gr_matched}) | "
        f"report: {out_path}"
    )
    return JobResult(summary=text, files=[out_path])


def s_val(summary: pd.DataFrame, metric: str) -> int:
    return int(summary.loc[summary["Metric"] == metric, "Value"].iloc[0])


def run(args) -> JobResult:
    """Job entry point (scripts\\run_job.py). Logs a final SUCCESS / FAILED line."""
    started = datetime.now()
    try:
        result = _execute(args, started)
    except Exception as exc:
        log.exception("FAILED: %s", exc)
        raise
    log.info("SUCCESS: %s", result.summary)
    return result


def _configure_logging(log_dir: Path | None) -> None:
    handlers: list[logging.Handler] = [logging.StreamHandler()]
    if log_dir is not None:
        try:
            log_dir.mkdir(parents=True, exist_ok=True)
            handlers.append(
                logging.FileHandler(
                    log_dir
                    / f"cmona_medius_reconciliation_{datetime.now():%Y%m%d_%H%M%S}.log",
                    encoding="utf-8",
                )
            )
        except OSError:
            pass
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)-8s | %(message)s",
        handlers=handlers,
        force=True,
    )


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="CMONA vs Medius PO/GR reconciliation")
    add_arguments(parser)
    args = parser.parse_args(argv)
    _configure_logging(_resolve_dirs(args)[2])
    try:
        result = run(args)
    except Exception as exc:
        print(f"\nFAILED: {exc}")
        return 1
    print(f"\nSUCCESS: {result.summary}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
