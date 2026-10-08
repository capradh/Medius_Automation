# AI_GUIDE.md - Oracle Automation Toolkit

This file is written for an AI assistant (e.g. Claude in the web chat) that CANNOT see my project.
Everything it needs to write a compatible script is in this file. Paste it at the start of a new chat.

------------------------------------------------------------------------
## PART 1 - FOR ME (human): how to use this file with free Claude chat
------------------------------------------------------------------------

1. Open a NEW chat on claude.ai. Paste this whole file (or upload it), then send.
2. In the same chat, send the task using the template in PART 4. Paste 3-5 ANONYMIZED sample rows
   (or the column headers) of each input file. Do not upload real company data or credentials.
3. Claude answers with a plan + complete files, each with its path. In VS Code create each file at that exact path, paste, save.
4. Run the checks Claude lists (usually):
   ```powershell
   .\.venv\Scripts\python.exe -m pytest -q
   .\.venv\Scripts\python.exe scripts\run_job.py <job_name> --help
   ```
5. Test on small sample files first. For anything that writes to a database: run WITHOUT --execute first (dry run).
6. If something fails, use the "fix" prompt in PART 5 (paste the FULL error text and the current file).
7. Long chats lose track of the rules. Start a new chat and paste this file again for each new job.

Where things live in the project:

| What | Path |
|---|---|
| New job script | `src\automation\jobs\<job_name>.py` |
| Its test | `tests\test_<job_name>.py` |
| Shared helper used by several jobs | `src\automation\<name>.py` |
| Input files | `input\` |
| Reusable SQL | `sql\<name>.sql` |
| Reports written by jobs | `output\` |
| Secrets / settings | `.env` (never share) |

------------------------------------------------------------------------
## PART 2 - FOR THE AI: project reference (read fully before answering)
------------------------------------------------------------------------

### Environment
- Windows 10/11, PowerShell, VS Code. Python 3.11+, virtual environment in `.venv`.
- Installed packages ONLY: pandas, numpy, openpyxl, xlrd (read .xls), requests, oracledb, python-dotenv, pytest.
  If you need anything else, say so explicitly and give the exact `requirements.txt` line to add.
- Use Windows paths in commands (`scripts\run_job.py`, `.\.venv\Scripts\python.exe`).
- Oracle uses python-oracledb. Credentials, SMTP and API settings are in `.env` and are loaded by the toolkit. NEVER hard-code
  or ask for credentials, hosts, or email addresses, and never print or log them.

### Project structure
```
src/automation/
  config.py         settings loading (databases discovered from .env)
  db.py             Oracle access
  io_files.py       read csv/xls/xlsx/json, write formatted xlsx
  processing.py     pandas + numpy cleaning and calculations
  api_client.py     REST client (auth, retries, pagination)
  notifications.py  email (SMTP)
  results.py        JobResult
  jobs/             one file per automation  <-- new scripts go here
    _common.py      shared options + process_to_excel (files starting with _ are not jobs)
    process_file.py  fetch_api.py  db_to_excel.py  daily_report.py  data_cleanup.py   (existing jobs)
scripts/run_job.py  the runner:  python scripts\run_job.py <job> [options]
tests/              pytest (DB and SMTP are always mocked; never connect to a real service in tests)
input/ output/ logs/ sql/
```

### Job contract (every new automation is ONE file in src/automation/jobs/)
```python
r"""Docstring: what it does + 2-3 example PowerShell commands (use r""" because of the backslashes)."""
from __future__ import annotations
import logging
from ..results import JobResult

log = logging.getLogger(__name__)

REQUIRES_DB = True      # optional: job queries a database -> runner requires --db (or DEFAULT_DB in .env)
WRITES = True           # optional: job can modify data -> prod DBs ask for confirmation with --execute

def add_arguments(p) -> None:          # optional: job-specific command-line options (argparse parser `p`)
    p.add_argument("--file-a", required=True)

def run(args) -> JobResult:            # required
    # args.db (uppercase db name, only if REQUIRES_DB or --db given), args.execute (bool),
    # args.notify, args.email_to, plus your own options. Dry-run means: not args.execute.
    return JobResult(summary="human readable result line(s)", files=[output_path])
```
Runner behaviour (do NOT re-implement it in jobs):
- Common flags exist already: `--db NAME  --execute  --notify  --email-to a@x.com;b@x.com  --no-attach`. Do not redefine them.
- With `--notify` the runner emails `[OK] <job>` with `result.summary` in the body and `result.files` attached,
  using SMTP_* / EMAIL_FROM / EMAIL_TO from `.env`. If `run()` raises, it logs the traceback, exits 1 and (with --notify)
  emails `[FAILED] <job>`. So: raise clear exceptions (ValueError, FileNotFoundError) with helpful messages; do not swallow errors.
- Only call `send_email` inside a job if a custom subject/body/recipient logic is required.
- Design rule: put the logic in PURE functions (DataFrames in, DataFrames out) so they can be unit tested; keep `run()` thin
  (read inputs -> call logic -> write Excel -> return JobResult).

### Helper API (use these; do not rewrite them)
```python
# automation.config
get_settings() -> Settings            # .output_dir, .input_dir, .log_dir (pathlib.Path), .default_db
get_database(name) -> DatabaseConfig  # .name .user .dsn .env .is_prod      (raises ConfigError)
ConfigError

# automation.db   (Oracle; column names in returned dicts are UPPERCASE; use bind variables, never f-strings in SQL)
fetch_all(db_name, sql, params: dict | None = None, max_rows: int | None = None) -> list[dict]
execute_write(db_name, sql, params=None, dry_run=True) -> int     # rolls back when dry_run, else commits; returns rows affected
get_connection(db_name)               # context manager for advanced cases

# automation.io_files
read_table(path, sheet=0, record_path=None) -> pd.DataFrame       # .csv .tsv .xlsx .xls .json .jsonl; sheet = name or number
collect_input_files(paths) -> list[Path]                          # expands files, folders, wildcards
pick_files() -> list[Path]                                        # Windows file-picker window
json_to_frame(obj, record_path=None) -> pd.DataFrame              # nested JSON -> flat table ('a.b' columns)
write_excel(sheets: dict[str, pd.DataFrame], path) -> Path        # formatted xlsx (bold header, filter, frozen row, auto width)

# automation.processing  (pandas + numpy)
clean_column_name(name) -> str                # 'Order Date ' -> 'order_date'
clean_dataframe(df, dayfirst=False) -> df     # snake_case columns, trims text, blank->NaN, numbers/dates detected
                                              # (date detection only for columns whose name has date/time or ends _at/_on),
                                              # keeps IDs with leading zeros as text, drops empty rows and duplicates
apply_calculations(df, ["total=price*qty"]) -> df
numeric_columns(df) -> list[str]
numeric_summary(df, cols=None) -> df          # sum/mean/median/std/min/p25/p75/p95/max per numeric column
group_summary(df, group_by: list[str], value_cols=None) -> df
flag_outliers(df, cols=None, z_threshold=3.0) -> df
column_profile(df) -> df
process_dataframe(df, *, group_by=None, value_cols=None, calcs=None, z_threshold=3.0, dayfirst=False) -> dict[str, df]

# automation.jobs._common
add_processing_arguments(p)                   # adds --group-by --value-cols --calc --z-threshold --dayfirst --output-name
process_to_excel(df, args, default_name) -> Path   # process_dataframe + write_excel into output\ (timestamped name)

# automation.api_client
ApiClient()                                   # reads API_BASE_URL / API_KEY from .env
  .get_json(path_or_url, params=None)   .post_json(path_or_url, payload, params=None)
  .get_all_pages(path_or_url, params=None, page_param="page", records_path=None, start_page=1, max_pages=50)
ApiError

# automation.notifications
send_email(subject, body, to=None, cc=None, attachments=None, html=False)   # to=None -> EMAIL_TO from .env

# automation.results
JobResult(summary: str, files: list[Path] = [])
```
Outputs go to `get_settings().output_dir`. Use timestamped file names (`f"name_{datetime.now():%Y%m%d_%H%M%S}.xlsx"`) so a
file open in Excel never blocks a rerun. Excel sheet names max 31 characters (write_excel handles this).

### Coding rules
1. `from __future__ import annotations`, type hints, small functions, a module docstring with example commands
   (make it a raw string `r"""..."""` when it contains Windows paths such as `scripts\run_job.py`).
2. Use `logging.getLogger(__name__)` for progress, not `print`.
3. SQL: bind variables only; limit rows (`max_rows`); reads must not modify data. Writes must be dry-run by default and only
   commit with `--execute`; show the number of affected rows in the summary.
4. Validate inputs early (files exist, required columns present) and raise errors that name the missing column and list the
   columns that were found.
5. Treat all user-supplied column names through `clean_column_name` / `clean_dataframe` so snake_case works everywhere.
6. Handle realistic data: blanks, duplicates, text-vs-number IDs, DD/MM dates (offer `--dayfirst`), trailing spaces, empty files.
7. Do not modify existing files. If an existing file really must change, show ONLY the minimal change and say why.
8. No secrets, hostnames, or email addresses in code, tests, or examples (use placeholders like `a@company.com` only in examples).
9. Tests: pytest, no real database / SMTP / network. Test the pure logic functions with small in-memory DataFrames;
   use `tmp_path` for files. The project has `pythonpath = src`, so tests import `automation...` directly.

### How to answer (output format)
1. If anything essential is missing or ambiguous (no sample columns, unclear matching rule), ask at most 5 short questions FIRST
   and wait. Otherwise continue.
2. Give a short plan: 3-6 bullets plus your assumptions.
3. Then provide EVERY file in full - never use "..." or "rest unchanged". For each file write its path in bold on its own line,
   then the code block:

   **src\automation\jobs\my_job.py**
   ```python
   ...
   ```
4. Then give, in this order: (a) PowerShell commands to run the tests, (b) 2-3 example commands to run the job
   (with and without `--notify`), (c) any new `.env` variable NAMES needed (never values) and any new `requirements.txt` lines,
   (d) a one-paragraph list of edge cases you handled and any you did not.
5. Keep explanations short; the code is the deliverable.

------------------------------------------------------------------------
## PART 3 - Example of a finished job (pattern to imitate)
------------------------------------------------------------------------
```python
r"""Compare two Excel/CSV files by key and report differences.

python scripts\run_job.py compare_files --file-a input\old.xlsx --file-b input\new.xlsx --key order_id
python scripts\run_job.py compare_files --file-a input\old.xlsx --file-b input\new.xlsx --key order_id --notify
"""

from __future__ import annotations

import logging
from datetime import datetime

import pandas as pd

from ..config import get_settings
from ..io_files import read_table, write_excel
from ..processing import clean_column_name, clean_dataframe
from ..results import JobResult

log = logging.getLogger(__name__)


def add_arguments(p) -> None:
    p.add_argument("--file-a", required=True, help="old file")
    p.add_argument("--file-b", required=True, help="new file")
    p.add_argument("--key", nargs="+", required=True, help="key column(s)")
    p.add_argument("--tolerance", type=float, default=0.01)
    p.add_argument("--dayfirst", action="store_true")


def compare_frames(
    a: pd.DataFrame, b: pd.DataFrame, key: list[str], tolerance: float
) -> dict[str, pd.DataFrame]:
    """Pure logic: easy to unit test."""
    missing = [k for k in key if k not in a.columns or k not in b.columns]
    if missing:
        raise ValueError(
            f"Key column(s) {missing} not found. A has {list(a.columns)}, B has {list(b.columns)}"
        )
    merged = a.merge(b, on=key, how="outer", suffixes=("_a", "_b"), indicator=True)
    return {
        "Summary": pd.DataFrame(
            {"metric": ["rows_a", "rows_b"], "value": [len(a), len(b)]}
        ),
        "Only_in_A": merged[merged["_merge"] == "left_only"].drop(columns="_merge"),
        "Only_in_B": merged[merged["_merge"] == "right_only"].drop(columns="_merge"),
    }


def run(args) -> JobResult:
    a = clean_dataframe(read_table(args.file_a), dayfirst=args.dayfirst)
    b = clean_dataframe(read_table(args.file_b), dayfirst=args.dayfirst)
    key = [clean_column_name(k) for k in args.key]
    sheets = compare_frames(a, b, key, args.tolerance)
    out = write_excel(
        sheets,
        get_settings().output_dir / f"compare_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
    )
    summary = f"A: {len(a)} rows, B: {len(b)} rows, only in A: {len(sheets['Only_in_A'])}, only in B: {len(sheets['Only_in_B'])}"
    return JobResult(summary, [out])
```
Matching test pattern:
```python
import pandas as pd
import pytest
from automation.jobs.compare_files import compare_frames


def test_only_in_each_side():
    a = pd.DataFrame({"id": [1, 2], "v": [1.0, 2.0]})
    b = pd.DataFrame({"id": [2, 3], "v": [2.0, 3.0]})
    r = compare_frames(a, b, ["id"], 0.01)
    assert r["Only_in_A"]["id"].tolist() == [1] and r["Only_in_B"]["id"].tolist() == [3]


def test_missing_key_raises():
    with pytest.raises(ValueError):
        compare_frames(pd.DataFrame({"x": [1]}), pd.DataFrame({"x": [1]}), ["id"], 0.01)
```

------------------------------------------------------------------------
## PART 4 - TASK TEMPLATE (copy, fill the <brackets>, send as the next message)
------------------------------------------------------------------------
```text
Using the project reference above, create a new job named "<job_name>".

WHAT IT DOES
<1-3 sentences>

INPUTS
- Files: <types, how many, how they are given: --file-a/--file-b | input\ folder | --pick>
- Database: <none | DB name from .env (e.g. corsair, lara) and the query or what to fetch; read-only?>
- API: <none | endpoint, parameters>
- Key / matching columns: <e.g. order_id>
- Sample data (anonymized, 3-5 rows or just headers) for each input:
  <paste here>

LOGIC
<rules in plain words: comparisons, calculations, tolerances, filters,
 what to do with missing / duplicate / blank values, date format DD/MM or MM/DD>

OUTPUT
- Excel in output\ with sheets: <names and the columns each must have>
- Result summary line should read like: "<example text>"

EMAIL
<none | use --notify with the Excel attached | custom subject/body: <text>>

OPTIONS I WANT
<e.g. --tolerance 0.01, --dayfirst, --sheet>

Follow the answer format in the reference: questions first only if essential, then plan, then every file in full with its path,
then commands, .env names, and edge cases.
```
Filled example (compare two Excel files and email a summary):
```text
Using the project reference above, create a new job named "compare_files".

WHAT IT DOES
Compares an old and a new Excel file by order_id and sends me a summary workbook by email.

INPUTS
- Files: two .xlsx files given with --file-a (old) and --file-b (new), first sheet
- Database: none
- Key columns: order_id (option --key, may be several)
- Sample headers: Order ID, Customer, Amount, Status, Updated On (dates DD/MM/YYYY)

LOGIC
- Rows only in A, only in B, and rows in both where any value differs (numbers compared with --tolerance, default 0.01)
- Duplicate keys go to a separate sheet and are excluded from the comparison

OUTPUT
- Excel with sheets: Summary, Differences (key, column, old value, new value, difference), Only_in_A, Only_in_B, Duplicates
- Summary line like: "A: 1200 rows, B: 1215 rows, 14 changed, 5 new, 2 removed"

EMAIL
Use --notify with the Excel attached.

OPTIONS I WANT
--file-a, --file-b, --key, --tolerance, --dayfirst
```

------------------------------------------------------------------------
## PART 5 - FIX PROMPT (when something fails)
------------------------------------------------------------------------
```text
The job fails. Fix only what is needed and return the COMPLETE corrected file(s) with their paths.
Command I ran:
<command>
Full error output (from the first line of the traceback to the last):
<paste>
Current content of the file:
<paste the file>
```
Other useful follow-ups: "Add a --tolerance option", "Add a sheet with <x>", "Add a test for <case>",
"Explain what each function does in plain words", "Review this job for risks before I run it on production".

------------------------------------------------------------------------
## PART 6 - Checklist before running on real / production data
------------------------------------------------------------------------
- [ ] Tests pass (`.\.venv\Scripts\python.exe -m pytest -q`)
- [ ] Ran on a small sample file; opened the Excel output and checked numbers by hand
- [ ] SQL reviewed: SELECT only (or dry-run first for writes), bind variables, row limit
- [ ] No credentials, real hostnames or personal data in the code or in what I pasted to the chat
- [ ] Email recipients checked (`EMAIL_TO` or `--email-to`)
- [ ] For writes: dry run result reviewed before `--execute`