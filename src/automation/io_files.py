"""Read csv / tsv / xls / xlsx / json / jsonl into DataFrames, and write formatted .xlsx."""
from __future__ import annotations

import glob
import json
import logging
import re
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

log = logging.getLogger(__name__)

SUPPORTED = {".csv", ".tsv", ".txt", ".xlsx", ".xlsm", ".xls", ".json", ".jsonl", ".ndjson"}
_BAD_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")


# ------------------------------------------------------------------ discovery
def collect_input_files(paths: Iterable[str | Path]) -> list[Path]:
    """Expand files, folders and wildcards (cmd.exe does not expand * for you)."""
    found: list[Path] = []
    for raw in paths:
        s = str(raw)
        matches = [Path(m) for m in sorted(glob.glob(s))] if any(c in s for c in "*?[") else [Path(s)]
        if not matches:
            raise FileNotFoundError(f"No files match: {s}")
        for m in matches:
            if m.is_dir():
                found += sorted(
                    f for f in m.iterdir()
                    if f.is_file() and f.suffix.lower() in SUPPORTED and not f.name.startswith("~$")
                )
            elif m.is_file():
                found.append(m)
            else:
                raise FileNotFoundError(f"File or folder not found: {m}")
    unique: list[Path] = []
    for f in found:
        if f not in unique:
            unique.append(f)
    return unique


def pick_files() -> list[Path]:
    """Open a Windows file-picker window (tkinter ships with the python.org installer)."""
    try:
        import tkinter
        from tkinter import filedialog
    except ImportError as e:  # pragma: no cover
        raise RuntimeError("tkinter is not available; pass --input <path> instead.") from e
    root = tkinter.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    chosen = filedialog.askopenfilenames(
        title="Select data files",
        filetypes=[("Data files", "*.csv *.tsv *.xls *.xlsx *.json *.jsonl"), ("All files", "*.*")],
    )
    root.destroy()
    return [Path(f) for f in chosen]


# ------------------------------------------------------------------ reading
def dig(obj: Any, path: str | None) -> Any:
    """Follow a dotted path through dicts/lists, e.g. 'data.items' or 'results.0.rows'."""
    if not path:
        return obj
    for key in path.split("."):
        if isinstance(obj, list) and key.isdigit():
            obj = obj[int(key)]
        elif isinstance(obj, dict) and key in obj:
            obj = obj[key]
        else:
            raise KeyError(f"Path '{path}' not found (stopped at '{key}')")
    return obj


def json_to_frame(obj: Any, record_path: str | None = None) -> pd.DataFrame:
    """Turn parsed JSON into a flat table (nested objects become 'a.b' columns)."""
    obj = dig(obj, record_path)
    if isinstance(obj, dict):
        lists = [v for v in obj.values()
                 if isinstance(v, list) and v and all(isinstance(i, dict) for i in v)]
        obj = lists[0] if len(lists) == 1 else [obj]
    if isinstance(obj, list):
        if obj and all(isinstance(i, dict) for i in obj):
            return pd.json_normalize(obj)
        return pd.DataFrame({"value": obj})
    return pd.DataFrame({"value": [obj]})


def _read_csv(path: Path, sep: str | None) -> pd.DataFrame:
    last: Exception | None = None
    for enc in ("utf-8-sig", "cp1252"):  # cp1252 = typical Excel-on-Windows export
        try:
            return pd.read_csv(path, sep=sep, engine="python", encoding=enc)
        except UnicodeDecodeError as e:
            last = e
    raise ValueError(f"Cannot decode {path.name}: {last}")


def read_table(path: str | Path, sheet: str | int = 0, record_path: str | None = None) -> pd.DataFrame:
    """Read any supported file into a DataFrame."""
    path = Path(path)
    ext = path.suffix.lower()
    if ext not in SUPPORTED:
        raise ValueError(f"Unsupported file type '{ext}' ({path.name}). Supported: {sorted(SUPPORTED)}")
    if ext == ".tsv":
        df = _read_csv(path, "\t")
    elif ext in (".csv", ".txt"):
        df = _read_csv(path, None)
    elif ext in (".xlsx", ".xlsm", ".xls"):
        sheet_name: str | int = int(sheet) if str(sheet).isdigit() else sheet
        df = pd.read_excel(path, sheet_name=sheet_name)  # openpyxl for xlsx, xlrd for xls
    elif ext in (".jsonl", ".ndjson"):
        df = pd.read_json(path, lines=True)
    else:
        with path.open(encoding="utf-8-sig") as f:
            df = json_to_frame(json.load(f), record_path)
    log.info("Read %s: %d rows x %d columns", path.name, len(df), df.shape[1])
    return df


# ------------------------------------------------------------------ writing
def _safe_sheet_name(name: str, used: set[str]) -> str:
    base = _BAD_SHEET_CHARS.sub("_", str(name))[:31] or "Sheet"
    candidate, n = base, 2
    while candidate.lower() in used:
        suffix = f"_{n}"
        candidate, n = base[: 31 - len(suffix)] + suffix, n + 1
    used.add(candidate.lower())
    return candidate


def _format_sheet(ws) -> None:
    header_fill = PatternFill("solid", fgColor="1F3864")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        longest = max((len(str(c.value)) for c in list(col)[:1000] if c.value is not None), default=8)
        ws.column_dimensions[col[0].column_letter].width = min(max(longest + 2, 8), 60)


def write_excel(sheets: dict[str, pd.DataFrame], path: str | Path) -> Path:
    """Write several DataFrames to one formatted .xlsx (bold header, filters, frozen row, auto width)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    used: set[str] = set()
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for name, df in sheets.items():
                df = df.copy()
                for col in df.select_dtypes(include=["datetimetz"]).columns:  # Excel has no tz support
                    df[col] = df[col].dt.tz_localize(None)
                sheet = _safe_sheet_name(name, used)
                df.to_excel(writer, sheet_name=sheet, index=False)
                _format_sheet(writer.sheets[sheet])
    except PermissionError as e:
        raise PermissionError(f"Cannot write {path}. Is it open in Excel? Close it and retry.") from e
    log.info("Wrote %s (%d sheet(s))", path, len(sheets))
    return path
