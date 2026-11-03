from datetime import datetime
from types import SimpleNamespace

import pandas as pd
from openpyxl import load_workbook
import pytest

from automation.jobs import vcr_medius_reconciliation as job


def oracle(rows):
    return job.normalize_oracle(
        pd.DataFrame(
            rows,
            columns=[
                "po_identifier_sq",
                "charge_code",
                "total_line_amount",
                "last_update_date",
            ],
        )
    )[0]


def medius(rows):
    return job.normalize_medius(
        pd.DataFrame(
            rows,
            columns=["external_system_id", "item_number", "ordered_amount"],
        )
    )[0]


def test_extract_po_id_uses_second_semicolon_segment():
    assert job.extract_po_id("254A;258;00010") == "258"


def test_oracle_failure_preserves_database_error_code(monkeypatch):
    class BrokenConnection:
        def __enter__(self):
            raise RuntimeError(
                "ORA-00932: inconsistent datatypes: expected CLOB got CHAR"
            )

        def __exit__(self, *_):
            return None

    monkeypatch.setattr(job, "get_connection", lambda _: BrokenConnection())

    with pytest.raises(job.ReconciliationError, match="ORA-00932: inconsistent datatypes"):
        job.fetch_oracle("CORSAIR", "select 1", 1, 10, 5)


def test_normalize_medius_excludes_dummy_external_system_ids():
    frame = pd.DataFrame(
        [
            ["088A;93568;DUMMY", "D1002", 999],
            ["088A;93568;00020", "D1002", 100],
        ],
        columns=["external_system_id", "item_number", "ordered_amount"],
    )

    normalized, invalid = job.normalize_medius(frame)

    assert len(normalized) == 1
    assert normalized.iloc[0]["external_po_id"] == "93568"
    assert normalized.iloc[0]["ordered_amount"] == 100
    assert invalid == 0


def test_reconcile_status_is_row_level_and_does_not_expand_oracle_rows():
    stamp = datetime(2026, 10, 26, 12, 0)
    oracle_data = oracle(
        [
            ["PO1", "C1", 40, stamp],
            ["PO1", "C2", 60, stamp],
            ["PO2", "C1", 10, stamp],
        ]
    )
    medius_data = medius(
        [
            ["254A;PO1;00010", "C1", 40],
            ["254A;PO1;00010", "C2", 70],
            ["254A;PO2;00010", "C1", 12],
            ["254A;PO1;00010", "EXTRA", 5],
        ]
    )

    result, _ = job.reconcile(oracle_data, medius_data, 0.01)

    assert len(result) == len(oracle_data)
    assert result["po_identifier_sq"].value_counts().to_dict() == {"PO1": 2, "PO2": 1}
    assert result["po_match"].tolist() == [True, True, True]
    assert len(result) == 3
    assert result["line_match"].tolist() == [True, True, True]
    assert result["line_amount_match"].tolist() == [True, False, False]
    assert "LINE_AMOUNT_MISMATCH" in result["reason_codes"].iloc[1]
    assert "LINE_AMOUNT_MISMATCH" in result["reason_codes"].iloc[2]


def test_oracle_charge_lines_are_aggregated_before_reconciliation():
    stamp = datetime(2026, 10, 26, 12, 0)
    oracle_data = oracle(
        [
            ["93568", "D1412", 92.2, stamp],
            ["93568", "D1412", 92.2, stamp],
            ["93568", "D1412", 92.2, stamp],
            ["93568", "D1412", 92.2, stamp],
            ["93568", "D1002", 13, stamp],
        ]
    )

    grouped = job.aggregate_oracle_charge_lines(oracle_data)

    assert len(grouped) == 2
    amounts = dict(zip(grouped["charge_code"], grouped["total_line_amount"]))
    assert amounts == {"D1412": 368.8, "D1002": 13}

    result, _ = job.reconcile(
        grouped,
        medius(
            [
                ["088A;93568;00020", "D1412", 368.8],
                ["088A;93568;00020", "D1002", 13],
            ]
        ),
        0.01,
    )
    assert len(result) == 2
    assert result["po_identifier_sq"].value_counts().to_dict() == {"93568": 2}
    assert result["line_amount_match"].all()


def test_repeated_oracle_rows_are_not_multiplied_by_duplicate_medius_rows():
    stamp = datetime(2026, 10, 26, 12, 0)
    oracle_data = oracle(
        [
            ["PO1", "C1", 10, stamp],
            ["PO1", "C1", 10, stamp],
        ]
    )
    medius_data = medius(
        [
            ["X;PO1;Z", "C1", 10],
            ["X;PO1;Z", "C1", 10],
            ["X;PO1;Z", "C1", 10],
        ]
    )

    result, _ = job.reconcile(oracle_data, medius_data, 0.01)

    assert len(result) == len(oracle_data) == 2
    assert result["po_match"].all()
    assert result["line_match"].all()


def test_charge_code_group_aggregates_oracle_and_medius_amounts():
    stamp = datetime(2026, 10, 26, 12, 0)
    oracle_data = oracle(
        [
            ["PO1", "C1", 10, stamp],
            ["PO1", "C1", 10, stamp],
            ["PO1", "C1", 10, stamp],
        ]
    )
    medius_data = medius([["X;PO1;Z", "C1", 10]])

    result, _ = job.reconcile(oracle_data, medius_data, 0.01)

    assert len(result) == len(oracle_data) == 3
    assert result["po_match"].all()
    assert result["line_match"].all()
    assert result["external_po_id"].notna().all()
    assert result["oracle_charge_amount"].tolist() == [30, 30, 30]
    assert result["medius_line_amount"].tolist() == [10, 10, 10]
    assert not result["line_amount_match"].any()
    assert result["reason_codes"].map(lambda reasons: "LINE_AMOUNT_MISMATCH" in reasons).all()


def test_duplicate_charge_lines_compare_grouped_amounts():
    stamp = datetime(2026, 10, 26, 12, 0)
    oracle_data = oracle(
        [
            ["PO1", "C1", 20, stamp],
            ["PO1", "C1", 10, stamp],
        ]
    )
    medius_data = medius(
        [
            ["X;PO1;Z", "C1", 10],
            ["X;PO1;Z", "C1", 20],
        ]
    )

    result, metrics = job.reconcile(oracle_data, medius_data, 0.01)

    assert result["oracle_charge_amount"].tolist() == [30, 30]
    assert result["medius_line_amount"].tolist() == [30, 30]
    assert result["line_amount_match"].all()
    assert metrics["matched_lines"] == 2


def test_four_equal_oracle_charge_lines_match_one_summed_medius_line():
    stamp = datetime(2026, 10, 26, 12, 0)
    oracle_data = oracle(
        [["93568", "D1412", 92.2, stamp] for _ in range(4)]
    )
    medius_data = medius([["088A;93568;00020", "D1412", 368.8]])

    result, _ = job.reconcile(oracle_data, medius_data, 0.01)
    report = job.presentation(job.add_comments(result, {}))

    assert len(result) == 4
    assert result["oracle_charge_amount"].tolist() == [368.8] * 4
    assert result["medius_line_amount"].tolist() == [368.8] * 4
    assert report["PO Line Amount"].tolist() == [368.8] * 4
    assert report["Medius PO Line Amount"].tolist() == [368.8] * 4
    assert result["po_match"].all()
    assert result["line_match"].all()
    assert result["line_amount_match"].all()


def test_po_amounts_are_grouped_and_statuses_are_boolean():
    stamp = datetime(2026, 10, 26, 12, 0)
    oracle_data = oracle(
        [
            ["PO1", "C1", 40, stamp],
            ["PO1", "C2", 60, stamp],
        ]
    )
    medius_data = medius(
        [
            ["X;PO1;Z", "C1", 40],
            ["X;PO1;Z", "C2", 70],
        ]
    )

    result, _ = job.reconcile(oracle_data, medius_data, 0.01)
    result = job.add_comments(result, {})
    report = job.presentation(result)

    assert result["oracle_po_amount"].tolist() == [100, 100]
    assert result["medius_po_amount"].tolist() == [110, 110]
    assert result["po_match"].all()
    assert not result["po_amount_match"].any()
    assert report.columns.tolist() == job.OUTPUT_COLUMNS
    assert report["PO Match Status"].map(type).eq(bool).all()
    assert "Error Aging" not in report.columns
    assert "MEDIUS ERROR: PO Amount is unmatched" in report.iloc[0]["Comments"]


def test_po_missing_comment_excludes_charge_code_error():
    stamp = datetime(2026, 10, 26, 12, 0)
    result, _ = job.reconcile(
        oracle([["PO1", "C1", 10, stamp]]),
        medius([["X;OTHER;Z", "C2", 10]]),
        0.01,
    )

    report = job.presentation(job.add_comments(result, {}))

    assert report.iloc[0]["Comments"] == "MEDIUS ERROR: PO is missing in Medius"


def test_decimal_line_difference_uses_oracle_amount_without_rounding():
    stamp = datetime(2026, 10, 26, 12, 0)
    result, _ = job.reconcile(
        oracle(
            [
                ["PO1", "C1", 15327796.5, stamp],
                ["PO2", "C1", 8441032.2, stamp],
            ]
        ),
        medius(
            [
                ["X;PO1;Z", "C1", 15327797],
                ["X;PO2;Z", "C1", 8441032],
            ]
        ),
        0.01,
    )

    assert result["po_amount_match"].all()
    assert result["line_amount_match"].all()
    report = job.presentation(result)
    assert report["PO Header Amount"].tolist() == [15327796.5, 8441032.2]
    assert report["Medius PO Amount"].tolist() == [15327796.5, 8441032.2]
    assert report["PO Line Amount"].tolist() == [15327796.5, 8441032.2]
    assert report["Medius PO Line Amount"].tolist() == [15327796.5, 8441032.2]


def test_one_unit_difference_remains_a_mismatch():
    stamp = datetime(2026, 10, 26, 12, 0)
    result, _ = job.reconcile(
        oracle([["PO1", "C1", 100, stamp]]),
        medius([["X;PO1;Z", "C1", 101]]),
        0.01,
    )

    assert not result["po_amount_match"].any()
    assert not result["line_amount_match"].any()


def test_po_headers_sum_adjusted_line_amounts_without_rounding():
    stamp = datetime(2026, 10, 26, 12, 0)
    result, _ = job.reconcile(
        oracle(
            [
                ["PO1", "C1", 123, stamp],
                ["PO1", "C2", 50.25, stamp],
            ]
        ),
        medius(
            [
                ["X;PO1;Z", "C1", 122.9],
                ["X;PO1;Z", "C2", 50.25],
            ]
        ),
        0.01,
    )

    assert result["oracle_po_amount"].tolist() == [173.25, 173.25]
    assert result["medius_po_amount"].tolist() == [173.25, 173.25]
    assert result["line_amount_match"].all()


def test_workbook_has_one_reconciliation_sheet_and_summary(tmp_path):
    row = {column: "" for column in job.OUTPUT_COLUMNS}
    row.update(
        {
            "PO Match Status": True,
            "PO Amount Match": True,
            "Line Match Status": True,
            "Line Amount Match": True,
        }
    )
    report = pd.DataFrame([row])
    path = tmp_path / "report.xlsx"

    job.write_workbook(path, report, {"Oracle row count": 1})
    workbook = load_workbook(path, data_only=True)

    assert workbook.sheetnames == ["PO Reconciliation", "Summary"]
    sheet = workbook["PO Reconciliation"]
    assert (sheet["A1"].value, sheet["F1"].value, sheet["J1"].value) == (
        "VCR",
        "Medius",
        "Reconciliation",
    )
    assert [sheet.cell(2, col).value for col in range(1, 15)] == job.OUTPUT_COLUMNS
