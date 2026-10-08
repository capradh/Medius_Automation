# Windows Setup Guide

Oracle Automation Toolkit: Oracle databases, csv/xls/json processing to Excel, REST APIs, email notifications.

## 1. Install the prerequisites (one time)

| Tool | Where | Notes |
|---|---|---|
| Python 3.11 or 3.12 (64-bit) | https://www.python.org/downloads/windows/ | On the first installer screen **tick "Add python.exe to PATH"** |
| Visual Studio Code | https://code.visualstudio.com/ | Install the **Python** extension (ms-python.python) |
| Git for Windows (optional, for sharing) | https://git-scm.com/download/win | Gives you Git Bash too |

You also need network access to the databases (VPN / allow-list as per company policy).

Open a **new** PowerShell window (Start menu -> "PowerShell") and verify:

```powershell
python --version
pip --version
git --version
```

If `python` is not recognised, try `py --version`. If that works, use `py` wherever this guide says `python` in step 3.

## 2. Get the project

1. Right-click the zip -> **Extract All** into a short path without spaces, for example `C:\projects\oracle-automation`.
2. Open the folder in VS Code:

```powershell
cd C:\projects\oracle-automation
code .
```

3. Open a terminal in VS Code: **Terminal -> New Terminal** (PowerShell).

## 3. Create the environment and install packages

### Option A - one command (recommended)

```powershell
.\setup.bat
```

It creates `.venv`, installs everything from `requirements.txt`, installs the project, and creates `.env` from the template.

### Option B - step by step

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install -e .
copy .env.example .env
```

If PowerShell says *"running scripts is disabled on this system"*, run this once and retry the activation:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

How to activate the environment in other shells (do this every time you open a new terminal, unless VS Code does it for you):

| Shell | Command |
|---|---|
| PowerShell | `.\.venv\Scripts\Activate.ps1` |
| Command Prompt | `.venv\Scripts\activate.bat` |
| Git Bash | `source .venv/Scripts/activate` |

Your prompt shows `(.venv)` when it is active.

### Tell VS Code to use it

Press `Ctrl+Shift+P` -> **Python: Select Interpreter** -> choose `.\.venv\Scripts\python.exe`. New terminals then activate it automatically.

## 4. Configure `.env`

Open `.env` in VS Code and fill in **your own** values. Never share or commit this file (it is git-ignored).

**Databases** - one block per database. The name between `DB_` and `_USER` is what you pass to `--db`.

```
DB_CORSAIR_USER=your_user
DB_CORSAIR_PASSWORD=your_password
DB_CORSAIR_DSN=host:1521/SERVICE_NAME
DB_CORSAIR_ENV=prod

DB_LARA_USER=...
DB_LARA_PASSWORD=...
DB_LARA_DSN=host:1521/SERVICE_NAME
DB_LARA_ENV=prod
```

To add another database, copy a block and rename it (e.g. `DB_NEWDB_*`). Passwords containing `#` or spaces must be in double quotes.

**Email (SMTP)**

```
# Office 365
SMTP_HOST=smtp.office365.com
SMTP_PORT=587
SMTP_SECURITY=starttls
SMTP_USER=you@company.com
SMTP_PASSWORD=...
EMAIL_FROM=you@company.com
EMAIL_TO=team@company.com
```

For Gmail use `smtp.gmail.com` with an App Password. For an internal relay use `SMTP_PORT=25`, `SMTP_SECURITY=none`, and leave user and password empty. Ask IT which one applies.

**API**

```
API_BASE_URL=https://api.example.com
API_KEY=your_token
API_KEY_HEADER=Authorization
API_KEY_PREFIX=Bearer
```

The key is sent as `Authorization: Bearer <key>`. For APIs that want `X-API-Key: <key>`, set `API_KEY_HEADER=X-API-Key` and `API_KEY_PREFIX=` (empty).

## 5. Verify everything works

```powershell
python scripts\test_db_connectivity.py          # all databases
python scripts\test_db_connectivity.py corsair  # one database
python scripts\test_email.py --to you@company.com
python scripts\test_api.py --url /posts/1
pytest                                          # unit tests, no Oracle needed
```

The database test checks network reachability, then login, and prints a hint for common errors (wrong password, wrong service name, VPN down).

## 6. Run the automations

Always run from the project folder with the environment active. Use `python scripts\run_job.py <job> --help` to see every option of a job.

### Process uploaded files -> Excel

Drop `.csv`, `.xls`, `.xlsx` or `.json` files into the `input\` folder (samples are included), then:

```powershell
python scripts\run_job.py process_file                                   # everything in input\
python scripts\run_job.py process_file --pick                            # file-picker window
python scripts\run_job.py process_file -i C:\data\sales.csv --group-by region --calc "total=quantity*unit_price" --dayfirst
python scripts\run_job.py process_file -i input\*.csv input\*.xlsx --combine --output-name all_sales
python scripts\run_job.py process_file -i input\sample_customers.json --record-path data
```

Each report is saved in `output\` as a formatted `.xlsx` with these sheets:

- **Data** - cleaned data (snake_case column names, trimmed text, numbers and dates detected, duplicates removed) plus your `--calc` columns
- **Summary** - numpy statistics per numeric column (sum, mean, median, std, percentiles)
- **Columns** - data-quality profile (type, missing, unique)
- **By <columns>** - grouped totals when you use `--group-by`
- **Outliers** - rows with values more than 3 standard deviations from the mean (change with `--z-threshold`)

Use `--dayfirst` for dates like 25/12/2026 (DD/MM). Close the output file in Excel before re-running with the same `--output-name`.

### Oracle query -> Excel

```powershell
python scripts\run_job.py db_to_excel --db corsair --sql-file sql\example_orders.sql --group-by region_id
python scripts\run_job.py db_to_excel --db lara --sql "SELECT * FROM orders WHERE status = :s" --bind s=OPEN
```

Only a single SELECT / WITH statement is accepted. For real protection also use a read-only database account.

### API -> Excel

```powershell
python scripts\run_job.py fetch_api --url /posts --group-by user_id
python scripts\run_job.py fetch_api --url https://api.example.com/v1/orders --param status=open --record-path data.items
python scripts\run_job.py fetch_api --url /orders --paginate --page-param page --max-pages 20
```

Failed calls (429 / 5xx) are retried automatically with back-off.

### Email notifications

Add `--notify` to any job. You get `[OK] job` with the summary and the Excel file attached, or `[FAILED] job` with the error.

```powershell
python scripts\run_job.py process_file --notify
python scripts\run_job.py db_to_excel --db corsair --sql-file sql\example_orders.sql --notify --email-to a@company.com;b@company.com
python scripts\run_job.py fetch_api --url /posts --notify --no-attach
```

### Database writes (safe pattern)

```powershell
python scripts\run_job.py data_cleanup --db lara              # dry run, rolled back
python scripts\run_job.py data_cleanup --db lara --execute    # commits; prod asks you to type the DB name
```

Edit the SQL in `src\automation\jobs\data_cleanup.py` first - it contains placeholder table names.

## 7. Run from VS Code

**Run and Debug** (`Ctrl+Shift+D`) has ready-made configurations: connectivity tests, process a file, db to Excel, API to Excel, and a dry-run job picker. Pick one and press F5.

## 8. Add your own job

Create `src\automation\jobs\my_job.py`:

```python
from automation.results import JobResult

def add_arguments(p):                 # optional: your own options
    p.add_argument("--threshold", type=float, default=10)

def run(args):
    # args.db, args.execute, args.threshold ...
    return JobResult("what happened", files=[])
```

Run it with `python scripts\run_job.py my_job --threshold 5`. Helpers you can reuse: `automation.db.fetch_all`, `automation.io_files.read_table / write_excel`, `automation.processing.process_dataframe`, `automation.api_client.ApiClient`, `automation.notifications.send_email`.

## 9. Share with colleagues (Git)

```powershell
git init
git add .
git commit -m "Initial commit"
git remote add origin <your-repo-url>
git push -u origin main
```

`.env`, `input\`, `output\` and `logs\` are git-ignored, so credentials and data are never uploaded. A colleague then does: install prerequisites (step 1), `git clone <repo-url>`, `.\setup.bat`, edit `.env`, run the tests (step 5).

## 10. Schedule on Windows (Task Scheduler)

`run_job.bat` activates the environment and runs a job, so it works from the scheduler. Example: every day at 08:00 with an email:

```powershell
schtasks /Create /SC DAILY /ST 08:00 /TN "OracleAutomation_DailyOrders" /TR "C:\projects\oracle-automation\run_job.bat db_to_excel --db corsair --sql-file sql\example_orders.sql --notify"
```

Run it once on demand: `schtasks /Run /TN "OracleAutomation_DailyOrders"`. Delete it: `schtasks /Delete /TN "OracleAutomation_DailyOrders" /F`. The task needs the PC to be on, you to be logged in (unless you configure "Run whether user is logged on or not" in the Task Scheduler app), and VPN connected if the databases need it. Keep the project path free of spaces.

## 11. Troubleshooting

| Problem | Fix |
|---|---|
| `python` is not recognized / opens Microsoft Store | Reinstall Python with "Add python.exe to PATH", or use `py`. Turn off the Store alias in Settings -> Apps -> Advanced app settings -> App execution aliases |
| Activation blocked in PowerShell | `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, or use `activate.bat` in Command Prompt |
| `.venv/bin/activate: No such file` in Git Bash | On Windows use `source .venv/Scripts/activate` |
| `pip install` fails with SSL / proxy errors | Ask IT for the proxy and CA bundle; then `pip install -r requirements.txt --proxy http://proxy:port` or set `PIP_CERT` |
| `ModuleNotFoundError` | The environment is not active (no `(.venv)` in the prompt), or run `pip install -r requirements.txt` again |
| `PermissionError ... .xlsx` | The file is open in Excel. Close it or change `--output-name` |
| Reading `.xls` fails | `pip install xlrd` (already in requirements); the file may really be html/csv renamed to `.xls` - save it as `.xlsx` from Excel |
| Dates look wrong (month and day swapped) | Add `--dayfirst` |
| ORA-01017 | Wrong username or password |
| ORA-12514 / DPY-6001 | Wrong service name in `DB_x_DSN` |
| Network UNREACHABLE | VPN not connected, firewall, wrong host or port |
| DPY-3010 | Old Oracle server: install Oracle Instant Client and set `ORACLE_CLIENT_LIB_DIR` in `.env` |
| Email `535` authentication error | Wrong password; Gmail needs an App Password; Office 365 needs SMTP AUTH enabled by IT |
| Email or API `certificate verify failed` | Corporate TLS inspection: ask IT for the CA bundle `.pem`, then set `API_VERIFY_SSL=C:\path\ca.pem` (API) or `REQUESTS_CA_BUNDLE` |
| Logs | See `logs\<job>_<date>.log` for full details of every run |
