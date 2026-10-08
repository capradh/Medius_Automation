"""Example WRITE job: preview -> dry run (rollback) -> commit only with --execute.

Adapt the table/SQL before use.
    python scripts/run_job.py data_cleanup --db lara              # dry run (rolled back)
    python scripts/run_job.py data_cleanup --db lara --execute    # commits (prod asks to confirm)
"""
import logging

from ..db import execute_write, fetch_all
from ..results import JobResult

log = logging.getLogger(__name__)
REQUIRES_DB = True
WRITES = True

# TODO: replace with real table/conditions. Keep the SELECT and DELETE conditions identical.
COUNT_SQL = "SELECT COUNT(*) AS cnt FROM my_temp_table WHERE created_on < SYSDATE - :days"
DELETE_SQL = "DELETE FROM my_temp_table WHERE created_on < SYSDATE - :days"
PARAMS = {"days": 90}


def run(args) -> JobResult:
    dry_run = not args.execute
    (row,) = fetch_all(args.db, COUNT_SQL, PARAMS)
    affected = execute_write(args.db, DELETE_SQL, PARAMS, dry_run=dry_run)
    mode = "DRY RUN (rolled back)" if dry_run else "COMMITTED"
    return JobResult(f"{mode}: {affected} row(s) matched on {args.db} (count check: {row['CNT']}).")
