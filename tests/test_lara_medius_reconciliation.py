from __future__ import annotations

import argparse
from datetime import datetime
from types import SimpleNamespace

import numpy as np
import pandas as pd
import pytest
from openpyxl import load_workbook

from automation.jobs import lara_medius_reconciliation as job


def raw(rows, metadata=0):
    return pd.DataFrame([[None]*6]*metadata + [job.MEDIUS_LABELS] + rows)


def medius(rows=None, metadata=0):
    rows = rows or [["PO1","C1","10.00","CONT1;x","2024-01-01","Yes"]]
    return job.clean_medius_raw(raw(rows, metadata))


def oracle(rows=None):
    rows = rows or [[1,"PO1","1","C1","EUR",10,None,None,"CONT1",datetime(2024,1,1)]]
    return job.prepare_oracle(pd.DataFrame(rows, columns=["trepa_uid","trepa_to_number","lineNumber","chargeCode","mainCurrency","lineAmount","selfBillingReference","blNumber","containerNumber","created_date"]))[0]


def test_header_first_row(): assert len(medius()) == 1
def test_header_after_metadata(): assert len(medius(metadata=3)) == 1
def test_html_headers_and_values():
    labels=[f"<b>{x}</b>" for x in job.MEDIUS_LABELS]
    df=job.clean_medius_raw(pd.DataFrame([labels,["<i>PO1</i>"," C1 ","10","CONT1&nbsp;;x","2024-01-01","Yes"]]))
    assert df.iloc[0][job.M_PO] == "PO1" and df.iloc[0]["_container"] == "CONT1"
def test_required_columns_validation():
    with pytest.raises(ValueError,match="header not found"): job.clean_medius_raw(pd.DataFrame([["x"]]))
def test_duplicate_normalized_columns():
    labels=job.MEDIUS_LABELS+["Item Number"]
    # Duplicate a genuinely required normalized name.
    labels[-1]="Item: Number"
    with pytest.raises(ValueError,match="Duplicate"): job.clean_medius_raw(pd.DataFrame([labels,[1]*7]))
def test_blank_row_removal(): assert len(job.clean_medius_raw(raw([[None]*6,["PO1","C1",1,"X","2024-01-01","Yes"]]))) == 1
@pytest.mark.parametrize(("value","expected"),[("1,234.56",1234.56),("1.234,56",1234.56),("€ 1 234,56",1234.56),("-12.50",-12.5),("(12.50)",-12.5)])
def test_amount_formats(value,expected): assert job.parse_amount(value) == expected
def test_missing_amount_nan(): assert np.isnan(job.parse_amount(None))
def test_missing_does_not_match_zero(): assert not np.isclose(job.parse_amount(None),0,atol=.01,rtol=0,equal_nan=False)
def test_container_extraction(): assert job.extract_container_number(" CONT123 ; OTHER ") == "CONT123"
def test_po_missing(): assert "PO is missing" in job.reconcile_headers(oracle(),medius([["OTHER","C1",10,"CONT1","2024-01-01","Yes"]]),.01).iloc[0].Comments
def test_created_today_missing_po_is_pending_in_both_sheets():
    as_of = datetime(2026, 10, 22, 15, 10)
    o = oracle([[1,"PO1","1","C1","EUR",10,None,None,"CONT1",datetime(2026,10,22,0,0,0)]])
    m = medius([["OTHER","C1",10,"CONT1","2026-10-22","Yes"]])
    header = job.reconcile_headers(o, m, .01, as_of=as_of).iloc[0]
    line = job.reconcile_lines(o, m, .01, as_of=as_of).iloc[0]
    assert header["PO Match Status"] == "PENDING"
    assert line["Item Line Status"] == "PENDING"
    assert header["Comments"] == "PO awaiting Medius import — created today"
    assert line["Comments"] == "PO awaiting Medius import — created today"
def test_created_before_today_missing_po_is_false():
    as_of = datetime(2026, 10, 22, 0, 0)
    o = oracle([[1,"PO1","1","C1","EUR",10,None,None,"CONT1",datetime(2026,10,21,23,59,59)]])
    m = medius([["OTHER","C1",10,"CONT1","2026-10-22","Yes"]])
    header = job.reconcile_headers(o, m, .01, as_of=as_of).iloc[0]
    line = job.reconcile_lines(o, m, .01, as_of=as_of).iloc[0]
    assert header["PO Match Status"] is False or header["PO Match Status"] == False
    assert line["Item Line Status"] is False or line["Item Line Status"] == False
    assert header["Comments"] == "PO is missing in Medius"
    assert line["Comments"] == "PO is missing in Medius"
def test_po_amount_mismatch(): assert "PO Amount" in job.reconcile_headers(oracle(),medius([["PO1","C1",12,"CONT1","2024-01-01","Yes"]]),.01).iloc[0].Comments
def test_charge_set_mismatch(): assert "Charge Code" in job.reconcile_headers(oracle(),medius([["PO1","C2",10,"CONT1","2024-01-01","Yes"]]),.01).iloc[0].Comments
def test_additional_number_mismatch_fails_line_match():
    row = job.reconcile_lines(
        oracle(), medius([["PO1","C1",10,"OTHER","2024-01-01","Yes"]]), .01
    ).iloc[0]
    assert not bool(row["Item Line Status"])
    assert row["_reason"] == "REFERENCE_MISMATCH"
    assert bool(row["_reference_issue"])
    assert "additional numbers" in row["Comments"]
    assert pd.isna(row["Medius containerNumber"])


def test_additional_numbers_match_oracle_fields_unordered_and_independently():
    o = oracle(
        [[
            1,
            "PO1",
            "1",
            "C1",
            "EUR",
            10,
            "ANT2004587",
            "SANT2611466",
            "CMAU2405511",
            datetime(2024, 1, 1),
        ]]
    )
    m = medius(
        [["PO1", "C1", 10, " CMAU2405511 ; SANT2611466 ; ANT2004587 ", "2024-01-01", "Yes"]]
    )

    row = job.reconcile_lines(o, m, .01).iloc[0]

    assert bool(row["Item Line Status"])
    assert row["Medius selfBillingReference"] == "ANT2004587"
    assert row["Medius blNumber"] == "SANT2611466"
    assert row["Medius containerNumber"] == "CMAU2405511"


def test_missing_additional_number_match_is_na():
    o = oracle(
        [[
            1,
            "PO1",
            "1",
            "C1",
            "EUR",
            10,
            "ANT2004587",
            None,
            "CMAU2405511",
            datetime(2024, 1, 1),
        ]]
    )
    m = medius([["PO1", "C1", 10, "CMAU2405511", "2024-01-01", "Yes"]])

    row = job.reconcile_lines(o, m, .01).iloc[0]

    assert not bool(row["Item Line Status"])
    assert row["_reason"] == "REFERENCE_MISMATCH"
    assert pd.isna(row["Medius selfBillingReference"])
    assert pd.isna(row["Medius blNumber"])
    assert row["Medius containerNumber"] == "CMAU2405511"


def test_one_or_two_available_oracle_references_can_match():
    o = oracle(
        [[
            1,
            "PO1",
            "1",
            "C1",
            "EUR",
            10,
            "ANT2004587",
            None,
            "CMAU2405511",
            datetime(2024, 1, 1),
        ]]
    )
    m = medius([["PO1", "C1", 10, "CMAU2405511; ANT2004587", "2024-01-01", "Yes"]])

    row = job.reconcile_lines(o, m, .01).iloc[0]

    assert bool(row["Item Line Status"])
    assert row["Medius selfBillingReference"] == "ANT2004587"
    assert pd.isna(row["Medius blNumber"])
    assert row["Medius containerNumber"] == "CMAU2405511"
def test_line_amount_mismatch(): assert job.reconcile_lines(oracle(),medius([["PO1","C1",20,"CONT1","2024-01-01","Yes"]]),.01).iloc[0]._reason == "AMOUNT_MISMATCH"
def test_tolerance_match(): assert bool(job.reconcile_lines(oracle(),medius([["PO1","C1",10.009,"CONT1","2024-01-01","Yes"]]),.01).iloc[0]["Item Line Status"])
def test_duplicate_order_independent_and_no_reuse():
    o=oracle([[1,"PO1","1","C1","EUR",20,None,None,"CONT1",datetime.now()],[2,"PO1","2","C1","EUR",10,None,None,"CONT1",datetime.now()]])
    m=medius([["PO1","C1",10,"CONT1","2024-01-01","Yes"],["PO1","C1",20,"CONT1","2024-01-01","Yes"]])
    r=job.reconcile_lines(o,m,.01)
    assert r["Item Line Status"].all() and r["_medius_id"].nunique()==2
def test_actual_medius_value_output():
    r=job.reconcile_lines(oracle(),medius([["PO1","C1",12,"CONT1","2024-01-01","Yes"]]),.01)
    assert r.iloc[0][job.MEDIUS_LABELS[2]] == 12
def test_every_oracle_po_remains():
    o=oracle([[1,"PO1","1","C1","EUR",1,None,None,"X",datetime.now()],[2,"PO2","1","C1","EUR",1,None,None,"X",datetime.now()]])
    assert len(job.reconcile_headers(o,medius(),.01))==2
def test_every_oracle_line_remains():
    o=pd.concat([oracle(),oracle()],ignore_index=True); o["_oracle_id"]=[0,1]
    assert len(job.reconcile_lines(o,medius(),.01))==2
def test_empty_oracle():
    with pytest.raises(ValueError,match="No Oracle"): job.prepare_oracle(pd.DataFrame())
def test_empty_medius():
    with pytest.raises(ValueError,match="empty"): job.clean_medius_raw(raw([[None]*6]))

class Cursor:
    def __init__(self, rows, description=None):
        self.rows=rows
        self.description=description or [("TREPA_UID",)]
        self.executed=None
    def __enter__(self): return self
    def __exit__(self,*a): pass
    def execute(self,sql,params): self.executed=(sql,params)
    def fetchmany(self,n): out=self.rows[:n]; self.rows=self.rows[n:]; return out
class Conn:
    def __init__(self,rows,description=None): self.c=Cursor(rows,description)
    def __enter__(self): return self
    def __exit__(self,*a): pass
    def cursor(self): return self.c
def test_max_rows_overflow(monkeypatch):
    monkeypatch.setattr(job,"get_connection",lambda _:Conn([(1,),(2,)]))
    with pytest.raises(ValueError,match="exceeds"): job.fetch_oracle("LARA","select",1,1,5)
def test_oracle_uses_lookback_bind(monkeypatch):
    description=[
        ("TREPA_UID",),("TREPA_TO_NUMBER",),("CHARGECODE",),("LINEAMOUNT",),
        ("SELFBILLINGREFERENCE",),("BLNUMBER",),("CONTAINERNUMBER",),("CREATED_DATE",),
    ]
    connection=Conn([(1,"PO1","C1",10,None,None,"CONT1",datetime.now())],description)
    monkeypatch.setattr(job,"get_connection",lambda _:connection)
    result=job.fetch_oracle("LARA","select :lookback_days",7,10,5)
    assert connection.c.executed == ("select :lookback_days", {"lookback_days":7})
    assert len(result)==1
def test_exact_output_columns():
    assert job.HEADER_COLUMNS == [
        "trepa_to_number","chargeCode","PO Amount","created_date",
        "Purchase Order: Order Identifier 2","Item: Number","Medius PO Amount",
        "PO Match Status","Comments",
    ]
    assert job.LINE_COLUMNS == [
        "trepa_to_number","chargeCode","lineAmount","selfBillingReference",
        "blNumber","containerNumber","created_date",
        "Purchase Order: Order Identifier 2","Item: Number",
        "Purchase Order Line: Confirmed Amount","Medius selfBillingReference",
        "Medius blNumber","Medius containerNumber",
        "Item Line Status","Comments",
    ]
def test_workbook_features(tmp_path):
    h=job.reconcile_headers(oracle(),medius(),.01); l=job.reconcile_lines(oracle(),medius(),.01)
    s=job.build_summary(h,l,timestamp=datetime.now(),filename="x.xlsx",environment="T",lookback_days=15,tolerance=.01,oracle_rows=1,medius_rows=1,duration=1,medius_only=0,invalid_medius=0,invalid_oracle=0)
    p=job.write_workbook(tmp_path/"x.xlsx",h,l,s); wb=load_workbook(p)
    assert wb.sheetnames==["PO Header Reconciliation","PO Line Item Reconciliation","Summary"]
    ws=wb[wb.sheetnames[0]]; assert ws.freeze_panes=="A3" and ws.auto_filter.ref and "A1:D1" in {str(x) for x in ws.merged_cells.ranges}
    line_ws=wb["PO Line Item Reconciliation"]
    assert "H1:M1" in {str(x) for x in line_ws.merged_cells.ranges}
    assert "N1:O1" in {str(x) for x in line_ws.merged_cells.ranges}
    assert [line_ws.cell(2, col).value for col in range(1, 16)] == job.LINE_COLUMNS
def test_summary_metrics():
    h=job.reconcile_headers(oracle(),medius(),.01); l=job.reconcile_lines(oracle(),medius(),.01)
    s=job.build_summary(h,l,timestamp=datetime.now(),filename="x",environment="T",lookback_days=1,tolerance=.01,oracle_rows=1,medius_rows=1,duration=1,medius_only=0,invalid_medius=0,invalid_oracle=0)
    assert s.loc[s.Metric.eq("matched lines"),"Value"].iloc[0]==1 and s.iloc[-1].Value==0
def test_summary_counts_pending_separately():
    as_of = datetime(2026, 10, 22, 12, 0)
    o = oracle([[1,"PO1","1","C1","EUR",10,None,None,"CONT1",datetime(2026,10,22,1,0)]])
    m = medius([["OTHER","C1",10,"CONT1","2026-10-22","Yes"]])
    h = job.reconcile_headers(o, m, .01, as_of=as_of)
    l = job.reconcile_lines(o, m, .01, as_of=as_of)
    s = job.build_summary(h,l,timestamp=as_of,filename="x",environment="T",lookback_days=1,tolerance=.01,oracle_rows=1,medius_rows=1,duration=1,medius_only=1,invalid_medius=0,invalid_oracle=0)
    metrics = dict(zip(s["Metric"], s["Value"]))
    assert metrics["pending POs"] == 1
    assert metrics["pending lines"] == 1
    assert metrics["unmatched POs"] == 0
    assert metrics["unmatched lines"] == 0
def test_zero_inactive_excluded():
    r=job.reconcile_lines(oracle([[1,"PO1","1","C1","EUR",0,None,None,"CONT1",datetime.now()]]),medius([["PO1","C1",0,"CONT1","2024-01-01","No"]]),.01)
    assert not r.iloc[0]["Item Line Status"]
def test_mixed_po_ignores_only_zero_inactive_rows():
    m = medius([
        ["500355992","COS50","0.00 GB","X1","2026-10-22","No"],
        ["500355992","COS50","0.00 GB","X2","2026-10-22","No"],
        ["500355992","COS97","0.00 GB","X3","2026-10-22","No"],
        ["500355992","COS82","0.00 GB","X4","2026-10-22","No"],
        ["500355992","COS50","468.00 GB","CONT1","2026-10-22","Yes"],
        ["500355992","COS97","78.67 GB","CONT2","2026-10-22","Yes"],
        ["500355992","COS82","23.40 GB","CONT3","2026-10-22","Yes"],
    ])
    filtered = job._eligible_medius(m)
    assert len(filtered) == 3
    o = oracle([
        [1,"500355992","1","COS50","GBP",468,None,None,"CONT1",datetime(2026,10,21)],
        [2,"500355992","2","COS97","GBP",78.67,None,None,"CONT2",datetime(2026,10,21)],
        [3,"500355992","3","COS82","GBP",23.40,None,None,"CONT3",datetime(2026,10,21)],
    ])
    header = job.reconcile_headers(o, m, .01, as_of=datetime(2026,10,22)).iloc[0]
    lines = job.reconcile_lines(o, m, .01, as_of=datetime(2026,10,22))
    assert header["trepa_to_number"] == "500355992"
    assert header["Item: Number"] == "COS50;COS82;COS97"
    assert header["Medius PO Amount"] == pytest.approx(570.07)
    assert bool(header["PO Match Status"])
    assert lines["Item Line Status"].eq(True).all()
def test_select_exact_file(tmp_path):
    p=tmp_path/"Medius_LARATO.xlsx"; p.touch()
    assert job.select_input_file(None,"*medius*larato*",False,tmp_path)==p
def test_picker_on_multiple(monkeypatch,tmp_path):
    for x in ("Medius_LARATO_a.xlsx","Medius_LARATO_b.xlsx"): (tmp_path/x).touch()
    chosen=tmp_path/"Medius_LARATO_a.xlsx"; monkeypatch.setattr(job,"pick_files",lambda:[chosen])
    assert job.select_input_file(None,"*medius*larato*",False,tmp_path)==chosen
def test_add_arguments_defaults():
    p=argparse.ArgumentParser(); job.add_arguments(p); a=p.parse_args([])
    assert a.lookback_days==1 and a.max_rows==1_000_000 and a.tolerance==.01
    assert a.lif2_file is None


def test_read_lif2_errors_reads_both_sheets_and_prefixes_messages(tmp_path):
    lif2_path = tmp_path / "Lif2.xlsx"
    with pd.ExcelWriter(lif2_path) as writer:
        pd.DataFrame(
            {"Source ID": ["PO1"], "Message Text": ["Invoice issue"]}
        ).to_excel(writer, sheet_name="INVOICE DETAILS", index=False)
        pd.DataFrame(
            {"Source Transaction ID": ["PO1"], "Message Text": ["Transfer issue"]}
        ).to_excel(writer, sheet_name="TRANSFER DETAILS", index=False)

    messages = job.read_lif2_errors(lif2_path)

    assert messages["po1"] == [
        "INVOICE DETAILS: Invoice issue",
        "TRANSFER DETAILS: Transfer issue",
    ]


def test_lif2_comments_add_sheet_errors_and_medius_error_only_for_unmatched():
    headers = pd.DataFrame(
        [
            {"trepa_to_number": "PO1", "PO Match Status": False, "Comments": "PO Amount is unmatched"},
            {"trepa_to_number": "PO2", "PO Match Status": True, "Comments": ""},
        ]
    )
    lines = pd.DataFrame(
        [
            {"trepa_to_number": "PO1", "Item Line Status": False, "Comments": "Line Amount is unmatched"},
            {"trepa_to_number": "PO2", "Item Line Status": True, "Comments": ""},
        ]
    )
    errors = {
        "po1": [
            "INVOICE DETAILS: Invoice issue",
            "TRANSFER DETAILS: Transfer issue",
        ]
    }

    updated_headers, updated_lines = job.add_lif2_comments(
        headers, lines, errors, lif2_supplied=True
    )

    expected_prefix = (
        "INVOICE DETAILS: Invoice issue\n"
        "TRANSFER DETAILS: Transfer issue\n"
        "MEDIUS ERROR: "
    )
    assert updated_headers.iloc[0]["Comments"] == expected_prefix + "PO Amount is unmatched"
    assert updated_lines.iloc[0]["Comments"] == expected_prefix + "Line Amount is unmatched"
    assert updated_headers.iloc[1]["Comments"] == ""
    assert updated_lines.iloc[1]["Comments"] == ""


def test_lif2_comments_leave_existing_behavior_when_not_supplied():
    headers = pd.DataFrame(
        [{"trepa_to_number": "PO1", "PO Match Status": False, "Comments": "PO is missing"}]
    )
    lines = pd.DataFrame(
        [{"trepa_to_number": "PO1", "Item Line Status": False, "Comments": "Line is missing"}]
    )

    updated_headers, updated_lines = job.add_lif2_comments(
        headers, lines, {}, lif2_supplied=False
    )

    assert updated_headers.iloc[0]["Comments"] == "PO is missing"
    assert updated_lines.iloc[0]["Comments"] == "Line is missing"


def test_mocked_run(monkeypatch,tmp_path):
    source=tmp_path/"Medius_LARATO.xlsx"; source.touch(); captured={}
    monkeypatch.setattr(job,"get_settings",lambda:SimpleNamespace(input_dir=tmp_path,output_dir=tmp_path))
    monkeypatch.setattr(job,"read_medius",lambda *a:medius())
    def fetch(db,sql,lookback,max_rows,timeout): captured.update(db=db,sql=sql,lookback=lookback,max_rows=max_rows); return oracle().drop(columns=[c for c in oracle().columns if c.startswith("_")])
    monkeypatch.setattr(job,"fetch_oracle",fetch); monkeypatch.setattr(job,"get_database",lambda _:SimpleNamespace(env="TEST"))
    monkeypatch.setattr(job,"write_workbook",lambda p,*a:p)
    args=SimpleNamespace(tolerance=.01,lookback_days=1,max_rows=10,db_timeout_seconds=5,medius_file=str(source),input_pattern="*",pick=False,sheet=0)
    result=job.run(args)
    assert captured["db"] == job.DATABASE_NAME
    assert captured["lookback"]==1 and captured["max_rows"]==10 and result.files[0].name.startswith("LARA_MEDIUS_Reconciliation_Report_") and result.summary.endswith("SUCCESS")
