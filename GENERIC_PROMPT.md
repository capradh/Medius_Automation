Read AI_GUIDE.md and the existing code in src/automation first, and follow its structure.

TASK
Create a new job named "<job_name>" in src/automation/jobs/<job_name>.py that:
<one or two sentences: what it does>

INPUTS
- Files: <types, how many, where: input\ folder / --input paths / --pick>
- Database: <none | which DB name e.g. corsair, and what to query or a SQL file path>
- API: <none | endpoint and parameters>
- Key columns / how rows match: <e.g. order_id>
- Sheet/column names or an example row: <paste 3-5 sample rows or headers>

LOGIC
<rules in plain words: what to compare/calculate, tolerances, edge cases such as
 missing in one file, duplicate keys, blank values, date format DD/MM>

OUTPUT
- Excel report in output\ with sheets: <e.g. Summary, Differences, Only_in_A, Only_in_B>
- Columns required in each sheet: <...>

EMAIL
- <none | send via --notify with the Excel attached | custom subject/body text>
- Recipients come from EMAIL_TO in .env unless --email-to is given

CLI OPTIONS
<e.g. --file-a, --file-b, --key, --tolerance 0.01>

DELIVERABLES
1. The job file
2. A pytest test in tests/test_<job_name>.py using small in-memory or tmp_path data
3. The exact PowerShell command to run it, with and without --notify
4. A list of any other files you changed or new .env variables needed (ask me before adding any)

#---------------------------------------------------------------------------------------

#### Example

Read AI_GUIDE.md and the existing code in src/automation first, and follow its structure.

TASK
Create a job named "compare_files" in src/automation/jobs/compare_files.py that compares
two Excel files (old vs new) and produces a summary workbook of the differences.

INPUTS
- Files: two .xlsx or .xls files passed with --file-a and --file-b (first sheet by default, --sheet optional)
- Database: none
- Key column: order_id (option --key, may take several columns)
- Sample headers: order_id, customer, amount, status, updated_on (dates are DD/MM/YYYY)

LOGIC
- Match rows on the key after cleaning column names with the existing cleaning helper
- Report rows only in A, rows only in B, and rows in both whose values differ
- For numeric columns use a tolerance (--tolerance, default 0.01); show old value, new value and difference
- Duplicate keys: report them in a separate sheet and exclude them from the comparison

OUTPUT
- Excel in output\ with sheets: Summary (counts per category, percent changed), Differences, Only_in_A, Only_in_B, Duplicates

EMAIL
- Use --notify so run_job.py sends the Excel; the summary line should read
  "A: 1200 rows, B: 1215 rows, 14 changed, 5 new, 2 removed"

DELIVERABLES
1. Job file  2. pytest test  3. PowerShell commands  4. List of changed files