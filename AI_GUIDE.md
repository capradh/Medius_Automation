# Project guide for AI agents

Windows project, Python 3.11+, venv in .venv. Code lives in src/automation, run via scripts/run_job.py.

## Structure
- src/automation/jobs/<name>.py : one file per automation (a "job"). Files starting with _ are not jobs.
- src/automation/ : config.py, db.py, io_files.py, processing.py, api_client.py, notifications.py, results.py
- scripts/run_job.py : `python scripts\run_job.py <job> [options]`; common flags --db --execute --notify --email-to --no-attach
- tests/ (pytest, DB and SMTP are mocked), input/, output/, sql/, logs/

## Job contract
- Optional `add_arguments(p)` for job-specific CLI options (reuse `jobs/_common.add_processing_arguments` when relevant)
- Required `run(args) -> JobResult(summary: str, files: list[Path])`
- Set `REQUIRES_DB = True` if it queries a database; `WRITES = True` if it can modify data
- Do not send emails inside the job when --notify is enough: run_job.py emails the summary and attaches result.files

## Use the existing helpers, do not reinvent them
- Files: io_files.read_table(path, sheet, record_path), collect_input_files, pick_files, write_excel(dict_of_dataframes, path)
- Data: processing.clean_dataframe, numeric_summary, group_summary, flag_outliers, process_dataframe (pandas + numpy)
- Oracle: db.fetch_all(db_name, sql, binds, max_rows), db.execute_write(db_name, sql, binds, dry_run=True)
- API: api_client.ApiClient (auth, retries, pagination)
- Email (only for custom bodies): notifications.send_email(subject, body, to, attachments)
- Paths: config.get_settings().output_dir / input_dir

## Rules
- Never hard-code credentials, hosts or email addresses; they come from .env (names only: DB_<NAME>_USER/_PASSWORD/_DSN/_ENV, SMTP_*, EMAIL_FROM, EMAIL_TO, API_*)
- Never print or log secrets
- SQL: bind variables only, never f-strings; reads limited with max_rows
- Writes: dry-run by default, commit only with --execute
- Use logging.getLogger(__name__), not print, for progress; raise clear errors (ValueError) for bad input
- Excel output through write_excel only (formatted sheets); timestamped file names
- Windows paths, no Linux-only commands
- Do not modify existing files unless needed; list any file you changed