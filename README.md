# Automation Toolkit

Python automation for multiple Oracle databases (CORSAIR, LARA, ...), file processing (csv / xls / xlsx / json -> Excel),
REST APIs, pandas / numpy calculations and email notifications.

**Windows users: follow [SETUP_WINDOWS.md](SETUP_WINDOWS.md)** for the complete step-by-step guide.

## Quick start (Windows)

```powershell
.\setup.bat                                   # creates .venv, installs packages, creates .env
# edit .env with your credentials
.\.venv\Scripts\Activate.ps1
python scripts\test_db_connectivity.py
python scripts\run_job.py process_file        # processes the sample files in input\ -> output\
```

## Jobs

| Job | What it does |
|---|---|
| `process_file` | Read csv / xls / xlsx / json, clean, calculate (pandas + numpy), write formatted Excel |
| `db_to_excel` | Read-only Oracle SELECT -> processed Excel |
| `fetch_api` | Call a REST API (auth, retries, pagination) -> processed Excel |
| `daily_report` | Minimal Oracle query -> CSV example |
| `data_cleanup` | Example write job: dry-run by default, `--execute` to commit |

Common options: `--db NAME`, `--execute`, `--notify` (email result), `--email-to`, `--no-attach`.
Job options: `python scripts\run_job.py <job> --help`.

## Layout

```
src/automation/   config | db | connectivity | io_files | processing | api_client | notifications | jobs/
scripts/          run_job.py | test_db_connectivity.py | test_email.py | test_api.py
input/ output/ logs/ sql/   working folders (input, output and logs are git-ignored)
tests/            49 unit tests, no Oracle needed
setup.bat run_job.bat       Windows helpers (setup, Task Scheduler wrapper)
```

## Safety rules

- Credentials live only in `.env` (git-ignored); each person uses their own account.
- Prefer read-only DB accounts for reporting. `db_to_excel` accepts only a single SELECT/WITH.
- Writes are dry-run by default; `--execute` is required; PROD requires typing the DB name.
- Always use bind variables (`:name`) in SQL.
