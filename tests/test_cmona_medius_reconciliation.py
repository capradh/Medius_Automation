import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

from automation.jobs.cmona_medius_reconciliation import (
    COL_CM_PO,
    COL_LIF2_INVOICE,
    COL_LIF2_TRANSFER,
    GR_COLUMNS,
    LINE_COLUMNS,
    N_CM_AMT,
    N_CM_BU,
    N_CM_ITEM,
    N_CM_PO_ID,
    N_CM_UNIT,
    N_ME_AMT,
    N_ME_CO,
    N_ME_ITEM,
    N_ME_ORDER,
    N_ME_UNIT,
    N_MSG,
    N_SRC_ID,
    N_SRC_TXN,
    PO_COLUMNS,
    SHEET_LINE,
    SHEET_PO,
    aggregate_cmona_po,
    aggregate_medius_po,
    build_error_map,
    find_receipts_in_note,
    load_table,
    matches_spec,
    FILE_SPECS,
    parse_amount_value,
    reconcile_po,
    reconcile_po_lines,
    write_report,
)


@pytest.mark.parametrize(
    "text,expected",
    [
        ("1,234.56", 1234.56),
        ("1.234,56", 1234.56),
        ("12,5", 12.5),
        ("1,234", 1234.0),
        ("(5.00)", -5.0),
        ("$ 1 234,50", 1234.5),
        ("1.5E+3", 1500.0),
        ("7-", -7.0),
    ],
)
def test_parse_amount(text, expected):
    assert parse_amount_value(text) == pytest.approx(expected)


def test_parse_amount_invalid():
    assert pd.isna(parse_amount_value("abc")) and pd.isna(parse_amount_value("NA"))


def test_header_detected_below_junk_rows_and_blank_rows_removed(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["Report title"])
    ws.append([])
    ws.append(list(COL_CM_PO))
    ws.append(["P1", "1000", "10"])
    ws.append([None, None, None])
    ws.append(["P2", None, "<b>5</b>"])
    path = tmp_path / "anything.dat"  # extension is irrelevant
    wb.save(path)
    df = load_table(path, "CMONA PO", COL_CM_PO)
    assert df[N_CM_PO_ID].tolist() == ["P1", "P2"]
    assert df[N_CM_BU].tolist() == ["1000", "NA"]
    assert df[N_CM_AMT].tolist() == ["10", "5"]


def test_html_disguised_as_xls(tmp_path):
    path = tmp_path / "x.xls"
    path.write_text(
        "<html><p>hi</p><table><tr><th>Source Transaction ID</th><th>Message Text</th></tr>"
        "<tr><td>P1</td><td>a&amp;b</td></tr></table></html>"
    )
    df = load_table(path, "Lif2", COL_LIF2_TRANSFER)
    assert df[N_MSG].tolist() == ["a&b"]


def test_missing_column_message_lists_found_columns(tmp_path):
    wb = Workbook()
    wb.active.append(["PO Id", "Other"])
    path = tmp_path / "f.xlsx"
    wb.save(path)
    with pytest.raises(ValueError, match="missing required column"):
        load_table(path, "CMONA PO", COL_CM_PO)


def _po_frames():
    cm = pd.DataFrame(
        {
            N_CM_PO_ID: ["P1", "P1", "P2", "P3", "P4"],
            N_CM_BU: ["1000"] * 5,
            N_CM_AMT: ["100", "50", "20", "5", "9"],
        }
    )
    me = pd.DataFrame(
        {
            N_ME_ORDER: ["P1", "P2", "p2"],
            N_ME_CO: ["1000-A"] * 3,
            N_ME_AMT: ["150", "10", "10.005"],
        }
    )
    return cm, me


def test_reconcile_po_statuses_and_actions():
    cm, me = _po_frames()
    transfer = build_error_map(pd.DataFrame({N_SRC_TXN: ["P3"], N_MSG: ["boom"]}))
    invoice = build_error_map(
        pd.DataFrame({N_SRC_ID: ["P3", "P9"], N_MSG: ["inv", "x"]}), N_SRC_ID
    )
    r = reconcile_po(
        aggregate_cmona_po(cm), aggregate_medius_po(me), transfer, invoice, 0.01
    )
    r = r.set_index("Cmona PO ID")
    assert len(r) == 4  # every CMONA PO present once
    assert r.loc["P1", "Action Required"] == "Not Required"
    assert (
        r.loc["P2", "Amount Match Status"] == "TRUE"
    )  # 10 + 10.005 vs 20 within tolerance
    assert r.loc[
        "P3", ["PO Match Status", "Amount Match Status", "Action Required"]
    ].tolist() == ["NA", "NA", "Checked"]
    assert (
        r.loc["P3", "Comments"] == "[TRANSFER DETAILS] boom"
    )  # transfer sheet wins over invoice
    assert r.loc["P4", "Action Required"] == "Need Investigation"


def test_amount_mismatch():
    cm = pd.DataFrame({N_CM_PO_ID: ["P1"], N_CM_BU: ["1"], N_CM_AMT: ["100"]})
    me = pd.DataFrame({N_ME_ORDER: ["P1"], N_ME_CO: ["1-X"], N_ME_AMT: ["99"]})
    r = reconcile_po(
        aggregate_cmona_po(cm),
        aggregate_medius_po(me),
        pd.Series(dtype=object),
        pd.Series(dtype=object),
    )
    assert r["Action Required"].tolist() == ["Amount Mismatch"] and r[
        "Amount Match Status"
    ].tolist() == ["FALSE"]


def test_duplicate_identical_lines_are_mapped_with_ambiguity_note():
    cmona = pd.DataFrame(
        {
            N_CM_PO_ID: ["AR118630", "AR118630"],
            N_CM_ITEM: ["O0801", "O0801"],
            N_CM_UNIT: ["1", "1"],
            N_CM_AMT: ["120", "120"],
        }
    )
    medius = pd.DataFrame(
        {
            N_ME_ORDER: ["AR118630", "AR118630"],
            N_ME_ITEM: ["O0801", "O0801"],
            N_ME_UNIT: ["1", "1"],
            N_ME_AMT: ["120", "120"],
        }
    )

    result = reconcile_po_lines(cmona, medius)

    assert len(result) == 2
    assert result["Line Match Status"].tolist() == ["TRUE", "TRUE"]
    assert result["Medius Item Number"].tolist() == ["O0801", "O0801"]
    assert result["Medius Item Quantity"].tolist() == [1, 1]
    assert result["Medius Ordered Amount"].tolist() == [120, 120]
    assert result["Review Reason"].tolist() == [
        "Ambiguous candidates",
        "Ambiguous candidates",
    ]


def test_find_receipts_respects_boundaries():
    keys, lengths = {"123", "RCT9"}, [3, 4]
    assert [h[0] for h in find_receipts_in_note("DN123 ok", keys, lengths)] == ["123"]
    assert find_receipts_in_note("x1234", keys, lengths) == []
    assert find_receipts_in_note("ref rct9/inv", keys, lengths) == [("RCT9", "rct9")]


def test_filename_patterns():
    specs = {s.key: s for s in FILE_SPECS}
    from pathlib import Path

    assert matches_spec(Path("CMONA_PO_report.xlsx"), specs["cmona_po"])
    assert not matches_spec(Path("CMONA_GR_report.xlsx"), specs["cmona_po"])
    assert matches_spec(Path("Medius GR 2026.xls"), specs["medius_gr"])
    assert matches_spec(Path("LIF2_errors.csv"), specs["lif2"])


def test_report_omits_requested_columns(tmp_path):
    report_path = tmp_path / "report.xlsx"
    write_report(
        report_path,
        pd.DataFrame(columns=PO_COLUMNS),
        pd.DataFrame(columns=LINE_COLUMNS),
        pd.DataFrame(columns=GR_COLUMNS),
        pd.DataFrame(columns=["Metric", "Value"]),
        pd.DataFrame(columns=["Execution Summary", "Value"]),
    )

    workbook = load_workbook(report_path, read_only=True)
    po_headers = next(
        workbook[SHEET_PO].iter_rows(min_row=2, max_row=2, values_only=True)
    )
    line_headers = next(
        workbook[SHEET_LINE].iter_rows(min_row=2, max_row=2, values_only=True)
    )
    removed_line_columns = {
        "CMONA PO Line Number",
        "Medius Line Number",
        "Amount Difference",
        "PO Match Status",
        "Match Method",
        "Confidence",
    }

    assert "Match score" not in po_headers
    assert not removed_line_columns.intersection(line_headers)
    assert "Match score" not in line_headers
    workbook.close()
