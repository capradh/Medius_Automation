"""Run a read-only SELECT on an Oracle database and export a processed Excel report.

    python scripts/run_job.py db_to_excel --db corsair --sql-file sql\\example_orders.sql --group-by region_id
    python scripts/run_job.py db_to_excel --db lara --sql "SELECT * FROM orders WHERE status = :s" --bind s=OPEN
"""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from ..db import fetch_all
from ..results import JobResult
from ._common import add_processing_arguments, process_to_excel
from .fetch_api import parse_key_values

REQUIRES_DB = True
_FORBIDDEN = re.compile(
    r"\b(insert|update|delete|merge|drop|alter|truncate|grant|revoke|create|commit|rollback|execute|call)\b", re.I)


def validate_read_only(sql: str) -> str:
    """Basic guard: single SELECT/WITH statement only. (Your real protection is a read-only DB account.)"""
    cleaned = re.sub(r"--.*?$|/\*.*?\*/", " ", sql, flags=re.S | re.M).strip().rstrip(";").strip()
    if not re.match(r"(select|with)\b", cleaned, re.I):
        raise ValueError("Only SELECT / WITH queries are allowed in this job.")
    if _FORBIDDEN.search(cleaned) or ";" in cleaned:
        raise ValueError("Query contains a forbidden keyword or multiple statements; only plain SELECTs are allowed.")
    return cleaned


def add_arguments(p) -> None:
    g = p.add_argument_group("query options")
    src = g.add_mutually_exclusive_group(required=True)
    src.add_argument("--sql", help="SELECT statement")
    src.add_argument("--sql-file", help="path to a .sql file containing one SELECT")
    g.add_argument("--bind", action="append", default=[], metavar="NAME=VALUE", help="bind variable (repeatable)")
    g.add_argument("--max-rows", type=int, default=100_000, help="safety limit (default 100000)")
    add_processing_arguments(p)


def run(args) -> JobResult:
    sql = args.sql or Path(args.sql_file).read_text(encoding="utf-8-sig")
    sql = validate_read_only(sql)
    rows = fetch_all(args.db, sql, parse_key_values(args.bind), max_rows=args.max_rows)
    df = pd.DataFrame(rows)
    note = f" (limited to {args.max_rows} rows)" if len(df) >= args.max_rows else ""
    out = process_to_excel(df, args, f"{args.db.lower()}_report")
    return JobResult(f"Fetched {len(df)} row(s) from {args.db}{note} -> {out.name}", [out])
