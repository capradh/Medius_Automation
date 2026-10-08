import json

import pandas as pd
import pytest

from automation.io_files import collect_input_files, json_to_frame, read_table, write_excel


def test_read_csv_and_tsv(tmp_path):
    (tmp_path / "a.csv").write_text("x,y\n1,2\n3,4\n", encoding="utf-8")
    (tmp_path / "b.tsv").write_text("x\ty\n1\t2\n", encoding="utf-8")
    assert read_table(tmp_path / "a.csv").shape == (2, 2)
    assert list(read_table(tmp_path / "b.tsv").columns) == ["x", "y"]


def test_read_csv_windows_encoding(tmp_path):
    (tmp_path / "w.csv").write_bytes("name,city\nRaj,Pune\nZoë,Mumbai\n".encode("cp1252"))
    assert read_table(tmp_path / "w.csv")["name"].tolist() == ["Raj", "Zoë"]


def test_json_variants(tmp_path):
    nested = {"status": "ok", "data": [{"id": 1, "a": {"b": 5}}, {"id": 2, "a": {"b": 6}}]}
    assert json_to_frame(nested)["a.b"].tolist() == [5, 6]
    assert json_to_frame(nested, "data").shape == (2, 2)
    assert json_to_frame([{"x": 1}]).shape == (1, 1)
    assert json_to_frame({"x": 1, "y": 2}).shape == (1, 2)
    with pytest.raises(KeyError):
        json_to_frame(nested, "nope")
    p = tmp_path / "f.json"
    p.write_text(json.dumps(nested), encoding="utf-8")
    assert read_table(p).shape == (2, 2)
    jl = tmp_path / "f.jsonl"
    jl.write_text('{"a":1}\n{"a":2}\n', encoding="utf-8")
    assert read_table(jl)["a"].tolist() == [1, 2]


def test_excel_roundtrip_and_formatting(tmp_path):
    out = write_excel({"Data": pd.DataFrame({"a": [1, 2], "b": ["x", "y"]}), "Bad/Name?": pd.DataFrame({"z": [1]})},
                      tmp_path / "o.xlsx")
    back = pd.read_excel(out, sheet_name=None)
    assert set(back) == {"Data", "Bad_Name_"}
    assert back["Data"]["a"].tolist() == [1, 2]
    from openpyxl import load_workbook
    ws = load_workbook(out)["Data"]
    assert ws.freeze_panes == "A2" and ws["A1"].font.bold


def test_read_back_via_read_table(tmp_path):
    out = write_excel({"Data": pd.DataFrame({"a": [1]})}, tmp_path / "o.xlsx")
    assert read_table(out).shape == (1, 1)


def test_collect_files_dir_glob_and_errors(tmp_path):
    for n in ("a.csv", "b.json", "~$tmp.xlsx", "notes.docx"):
        (tmp_path / n).write_text("x")
    assert [f.name for f in collect_input_files([tmp_path])] == ["a.csv", "b.json"]
    assert [f.name for f in collect_input_files([str(tmp_path / "*.csv")])] == ["a.csv"]
    with pytest.raises(FileNotFoundError):
        collect_input_files([tmp_path / "missing.csv"])
    with pytest.raises(ValueError):
        read_table(tmp_path / "notes.docx")
