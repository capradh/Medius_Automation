"""Options and helpers shared by the file/API/database -> Excel jobs."""
from __future__ import annotations

import argparse
import logging
from datetime import datetime
from pathlib import Path

import pandas as pd

from ..config import get_settings
from ..io_files import write_excel
from ..processing import process_dataframe

log = logging.getLogger(__name__)


def add_processing_arguments(p: argparse.ArgumentParser) -> None:
    g = p.add_argument_group("processing options")
    g.add_argument("--group-by", nargs="+", metavar="COL", help="add a summary sheet grouped by these columns")
    g.add_argument("--value-cols", nargs="+", metavar="COL", help="numeric columns to summarise (default: all numeric)")
    g.add_argument("--calc", action="append", default=[], metavar="EXPR",
                   help='computed column, e.g. --calc "total=price*qty" (repeatable)')
    g.add_argument("--z-threshold", type=float, default=3.0, help="outlier z-score threshold (default 3.0)")
    g.add_argument("--dayfirst", action="store_true", help="read 01/02/2025 as 1 Feb (DD/MM) instead of 2 Jan")
    g.add_argument("--output-name", help="output file name without extension")


def process_to_excel(df: pd.DataFrame, args: argparse.Namespace, default_name: str) -> Path:
    sheets = process_dataframe(
        df,
        group_by=args.group_by, value_cols=args.value_cols, calcs=args.calc,
        z_threshold=args.z_threshold, dayfirst=args.dayfirst,
    )
    name = args.output_name or f"{default_name}_{datetime.now():%Y%m%d_%H%M%S}"
    return write_excel(sheets, get_settings().output_dir / f"{name}.xlsx")
