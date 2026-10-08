import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import run_job  # noqa: E402
from automation.jobs.db_to_excel import validate_read_only  # noqa: E402


@pytest.fixture
def isolated(tmp_path, monkeypatch):
    """Send output/ and logs/ to a temp folder."""
    import automation.config as cfg
    real = cfg.get_settings()
    fake = type(real)(**{**real.__dict__, "output_dir": tmp_path / "out", "log_dir": tmp_path / "logs",
                         "input_dir": tmp_path / "in"})
    for mod in ("automation.config", "automation.jobs._common", "automation.jobs.process_file",
                "automation.logging_setup"):
        monkeypatch.setattr(f"{mod}.get_settings", lambda: fake, raising=False)
    (tmp_path / "in").mkdir()
    return tmp_path


def test_process_file_end_to_end(isolated):
    rc = run_job.main(["process_file", "-i", str(ROOT / "input" / "sample_sales.csv"),
                       "--group-by", "region", "--calc", "total=quantity*unit_price", "--dayfirst",
                       "--output-name", "report"])
    assert rc == 0
    sheets = pd.read_excel(isolated / "out" / "report.xlsx", sheet_name=None)
    assert {"Data", "Summary", "Columns", "By region", "Outliers"} <= set(sheets)
    assert "total" in sheets["Data"].columns


def test_process_file_json_and_xlsx_and_combine(isolated):
    rc = run_job.main(["process_file", "-i", str(ROOT / "input" / "sample_customers.json"),
                       str(ROOT / "input" / "sample_sales.xlsx"), "--combine", "--output-name", "c"])
    assert rc == 0
    data = pd.read_excel(isolated / "out" / "c.xlsx", sheet_name="Data")
    assert "source_file" in data.columns and len(data) == 55


def test_missing_file_fails_cleanly(isolated):
    assert run_job.main(["process_file", "-i", "nope.csv"]) == 1


def test_unknown_job_and_db_required(isolated, capsys):
    assert run_job.main(["does_not_exist"]) == 2
    assert run_job.main(["daily_report"]) == 2  # no --db and no DEFAULT_DB


def test_notify_sends_email_with_attachment(isolated, monkeypatch):
    sent = {}
    import automation.notifications as n
    monkeypatch.setattr(n, "send_email", lambda subject, body, to=None, attachments=None, **k: sent.update(
        subject=subject, to=to, attachments=attachments))
    rc = run_job.main(["process_file", "-i", str(ROOT / "input" / "sample_sales.csv"),
                       "--notify", "--email-to", "a@x.com;b@x.com", "--output-name", "m"])
    assert rc == 0
    assert sent["subject"] == "[OK] process_file" and sent["to"] == ["a@x.com", "b@x.com"]
    assert [p.name for p in sent["attachments"]] == ["m.xlsx"]


def test_failure_sends_failed_email(isolated, monkeypatch):
    sent = {}
    import automation.notifications as n
    monkeypatch.setattr(n, "send_email", lambda subject, body, **k: sent.update(subject=subject, body=body))
    rc = run_job.main(["process_file", "-i", str(ROOT / "input" / "sample_sales.csv"), "--calc", "bad", "--notify"])
    assert rc == 1 and sent["subject"] == "[FAILED] process_file" and "Bad --calc" in sent["body"]


@pytest.mark.parametrize("sql", ["SELECT 1 FROM dual", "with a as (select 1 x from dual) select * from a;",
                                 "-- note\nSELECT created FROM t"])
def test_sql_guard_allows(sql):
    validate_read_only(sql)


@pytest.mark.parametrize("sql", ["DELETE FROM t", "SELECT 1 FROM dual; DROP TABLE t", "SELECT * FROM t FOR UPDATE"[:0] + "update t set a=1",
                                 "SELECT 1 FROM dual WHERE 1=1 AND commit"])
def test_sql_guard_blocks(sql):
    with pytest.raises(ValueError):
        validate_read_only(sql)


def test_one_bad_file_does_not_stop_batch(isolated):
    # --calc only fits the sales file; the JSON file fails but the sales report is still produced
    rc = run_job.main(["process_file", "-i", str(ROOT / "input" / "sample_sales.csv"),
                       str(ROOT / "input" / "sample_customers.json"),
                       "--calc", "total=quantity*unit_price", "--output-name", "ok"])
    assert rc == 0
    assert (isolated / "out" / "ok.xlsx").exists()
