"""Minimal read-only example: query -> CSV in output/.  (See db_to_excel for the full version.)

    python scripts/run_job.py daily_report --db corsair
"""
import csv
import logging
from datetime import datetime

from ..config import get_settings
from ..db import fetch_all
from ..results import JobResult

log = logging.getLogger(__name__)
REQUIRES_DB = True

# TODO: replace with a real query for your schema.
SQL = "SELECT SYSDATE AS run_time, USER AS db_user FROM dual"


def run(args) -> JobResult:
    rows = fetch_all(args.db, SQL, max_rows=1000)
    if not rows:
        return JobResult("Query returned no rows.")
    out_dir = get_settings().output_dir
    out_dir.mkdir(exist_ok=True)
    path = out_dir / f"daily_report_{args.db.lower()}_{datetime.now():%Y%m%d_%H%M%S}.csv"
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    return JobResult(f"Fetched {len(rows)} row(s) from {args.db} -> {path.name}", [path])
