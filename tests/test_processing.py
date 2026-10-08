import numpy as np
import pandas as pd
import pytest

from automation.processing import (apply_calculations, clean_columns, clean_dataframe, flag_outliers,
                                   group_summary, numeric_summary, process_dataframe)


def raw():
    return pd.DataFrame({
        " Order ID ": ["00123", "00124", "00124", None],
        "Order Date": ["01/02/2026", "02/02/2026", "02/02/2026", None],
        "Amount": ["10", " 20 ", "20", None],
        "Region": ["N", "S", "S", ""],
        "Empty": [None, None, None, None],
    })


def test_clean_columns_unique():
    assert list(clean_columns(pd.DataFrame(columns=["A b", "a-b", "C"])).columns) == ["a_b", "a_b_2", "c"]


def test_clean_dataframe():
    df = clean_dataframe(raw(), dayfirst=True)
    assert "empty" not in df.columns
    assert len(df) == 2                                   # empty row dropped, duplicate dropped
    assert df["order_id"].tolist() == ["00123", "00124"]  # leading zeros preserved
    assert df["amount"].tolist() == [10, 20]              # text -> numbers
    assert df["order_date"].iloc[0] == pd.Timestamp("2026-02-01")  # dayfirst


def test_calculations():
    df = pd.DataFrame({"price": [2.0, 3.0], "qty": [4, 5]})
    assert apply_calculations(df, ["total=price*qty", "double=total*2"])["double"].tolist() == [16.0, 30.0]
    with pytest.raises(ValueError):
        apply_calculations(df, ["nonsense"])
    with pytest.raises(ValueError):
        apply_calculations(df, ["x=missing_col+1"])


def test_numeric_summary_matches_numpy():
    df = pd.DataFrame({"v": [1.0, 2.0, 3.0, 4.0, np.nan]})
    r = numeric_summary(df).iloc[0]
    assert r["count"] == 4 and r["nulls"] == 1 and r["sum"] == 10 and r["median"] == 2.5
    assert r["std"] == pytest.approx(np.std([1, 2, 3, 4], ddof=1))


def test_group_summary():
    df = pd.DataFrame({"r": ["a", "a", "b"], "v": [1, 2, 10]})
    g = group_summary(df, ["r"])
    assert g.loc[g["r"] == "a", "v_sum"].iloc[0] == 3
    with pytest.raises(ValueError):
        group_summary(df, ["nope"])


def test_outliers():
    df = pd.DataFrame({"v": [10.0] * 30 + [1000.0], "c": [1.0] * 31})  # constant column must not crash
    out = flag_outliers(df)
    assert len(out) == 1 and out["outlier_columns"].iloc[0] == "v"


def test_pipeline_sheets():
    df = pd.DataFrame({"Region": ["N", "S"] * 20, "Qty": range(40), "Price": [5.0] * 39 + [900.0]})
    s = process_dataframe(df, group_by=["Region"], calcs=["total=qty*price"])
    assert {"Data", "Summary", "Columns", "By region", "Outliers"} <= set(s)
    assert "total" in s["Data"].columns
