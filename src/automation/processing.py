"""Cleaning and calculations with pandas + numpy. All functions are pure (DataFrame in, DataFrame out)."""
from __future__ import annotations

import re
import warnings

import numpy as np
import pandas as pd
from pandas.api import types as pt

_ASSIGN = re.compile(r"^\s*([A-Za-z_]\w*)\s*=(?!=)\s*(.+)$")


def clean_column_name(name) -> str:
    return re.sub(r"[^0-9a-zA-Z]+", "_", str(name).strip()).strip("_").lower() or "col"


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """snake_case, unique column names ('Order Date ' -> 'order_date')."""
    seen: dict[str, int] = {}
    new = []
    for c in df.columns:
        base = clean_column_name(c)
        seen[base] = seen.get(base, 0) + 1
        new.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
    out = df.copy()
    out.columns = new
    return out


def _looks_like_date_column(name: str) -> bool:
    return any(t in name for t in ("date", "time")) or name.endswith("_at") or name.endswith("_on")


def _is_text(s: pd.Series) -> bool:
    return pt.is_object_dtype(s) or pt.is_string_dtype(s)


def clean_dataframe(df: pd.DataFrame, dayfirst: bool = False) -> pd.DataFrame:
    """Tidy a raw table: names, blanks, whitespace, numbers-as-text, dates, duplicates."""
    df = clean_columns(df)
    # 1) trim text and turn blank strings into real missing values
    for col in df.columns:
        if _is_text(df[col]):
            s = df[col].map(lambda v: v.strip() if isinstance(v, str) else v)
            df[col] = s.where(s != "", np.nan)
    df = df.dropna(how="all").dropna(axis=1, how="all")
    # 2) convert text columns to numbers / dates when ALL values convert
    for col in df.columns:
        s = df[col]
        if not _is_text(s):
            continue
        non_null = s.dropna()
        if non_null.empty:
            continue
        has_leading_zero = non_null.astype(str).str.fullmatch(r"0\d+").any()  # keep IDs like "00123" as text
        if not has_leading_zero:
            num = pd.to_numeric(s, errors="coerce")
            if num.notna().sum() == s.notna().sum():
                df[col] = num
                continue
        if _looks_like_date_column(col):
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                dt = pd.to_datetime(s, errors="coerce", dayfirst=dayfirst, format="mixed")
            if dt.notna().sum() == s.notna().sum():
                df[col] = dt
    return df.drop_duplicates().reset_index(drop=True)


def apply_calculations(df: pd.DataFrame, exprs: list[str] | None) -> pd.DataFrame:
    """Add computed columns from expressions like 'total=price*qty' (uses cleaned column names)."""
    out = df.copy()
    for expr in exprs or []:
        m = _ASSIGN.match(expr)
        if not m:
            raise ValueError(f"Bad --calc '{expr}'. Use the form  new_column=expression  e.g. total=price*qty")
        name, rhs = m.groups()
        try:
            out[name] = out.eval(rhs)
        except Exception as e:  # noqa: BLE001
            raise ValueError(f"--calc '{expr}' failed: {e}. Columns: {list(out.columns)}") from e
    return out


def numeric_columns(df: pd.DataFrame) -> list[str]:
    return [c for c in df.columns if pt.is_numeric_dtype(df[c]) and not pt.is_bool_dtype(df[c])]


def numeric_summary(df: pd.DataFrame, cols: list[str] | None = None) -> pd.DataFrame:
    """Per-column statistics computed with numpy."""
    rows = []
    for col in cols or numeric_columns(df):
        nulls = int(df[col].isna().sum())
        a = df[col].dropna().to_numpy(dtype=float)
        if a.size == 0:
            rows.append({"column": col, "count": 0, "nulls": nulls})
            continue
        rows.append({
            "column": col, "count": int(a.size), "nulls": nulls,
            "sum": a.sum(), "mean": a.mean(), "median": float(np.median(a)),
            "std": a.std(ddof=1) if a.size > 1 else np.nan,
            "min": a.min(), "p25": np.percentile(a, 25), "p75": np.percentile(a, 75),
            "p95": np.percentile(a, 95), "max": a.max(),
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=["column"])


def column_profile(df: pd.DataFrame) -> pd.DataFrame:
    """Data-quality overview: type, filled, missing and distinct values per column."""
    return pd.DataFrame({
        "column": df.columns,
        "dtype": [str(t) for t in df.dtypes],
        "non_null": [int(df[c].notna().sum()) for c in df.columns],
        "nulls": [int(df[c].isna().sum()) for c in df.columns],
        "unique": [int(df[c].nunique(dropna=True)) for c in df.columns],
    })


def group_summary(df: pd.DataFrame, group_by: list[str], value_cols: list[str] | None = None) -> pd.DataFrame:
    missing = [c for c in group_by + (value_cols or []) if c not in df.columns]
    if missing:
        raise ValueError(f"Column(s) not found: {missing}. Available: {list(df.columns)}")
    values = value_cols or [c for c in numeric_columns(df) if c not in group_by]
    if not values:
        return df.groupby(group_by, dropna=False).size().reset_index(name="count")
    agg = df.groupby(group_by, dropna=False)[values].agg(["count", "sum", "mean", "min", "max"])
    agg.columns = [f"{c}_{stat}" for c, stat in agg.columns]
    return agg.reset_index()


def flag_outliers(df: pd.DataFrame, cols: list[str] | None = None, z_threshold: float = 3.0) -> pd.DataFrame:
    """Rows where any numeric column has |z-score| above the threshold (numpy)."""
    cols = cols or numeric_columns(df)
    flagged = np.zeros(len(df), dtype=bool)
    reasons = np.array([""] * len(df), dtype=object)
    for col in cols:
        a = df[col].to_numpy(dtype=float, na_value=np.nan)
        std = np.nanstd(a)
        if not np.isfinite(std) or std == 0:
            continue
        z = np.abs((a - np.nanmean(a)) / std)
        hit = np.nan_to_num(z, nan=0.0) > z_threshold
        flagged |= hit
        reasons[hit] = reasons[hit] + f"{col}; "
    out = df[flagged].copy()
    out["outlier_columns"] = [r.rstrip("; ") for r in reasons[flagged]]
    return out


def process_dataframe(
    df: pd.DataFrame,
    *,
    group_by: list[str] | None = None,
    value_cols: list[str] | None = None,
    calcs: list[str] | None = None,
    z_threshold: float = 3.0,
    dayfirst: bool = False,
) -> dict[str, pd.DataFrame]:
    """Full pipeline -> dict of sheet name -> DataFrame, ready for write_excel()."""
    group_by = [clean_column_name(c) for c in group_by or []]
    value_cols = [clean_column_name(c) for c in value_cols or []] or None

    data = apply_calculations(clean_dataframe(df, dayfirst=dayfirst), calcs)
    sheets: dict[str, pd.DataFrame] = {
        "Data": data,
        "Summary": numeric_summary(data, value_cols),
        "Columns": column_profile(data),
    }
    if group_by:
        sheets["By " + "+".join(group_by)] = group_summary(data, group_by, value_cols)
    outliers = flag_outliers(data, value_cols, z_threshold)
    if not outliers.empty:
        sheets["Outliers"] = outliers
    return sheets
