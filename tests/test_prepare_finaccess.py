"""Unit and integration tests for the FinAccess 2024 streaming ETL pipeline."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pandas as pd

from prepare_finaccess import (
    PUBLISHED,
    USAGE_FLAGS,
    WANTED,
    build,
    col_letters,
    reconcile,
    stream_columns,
    write_dictionary,
)


def _sample_raw_df() -> pd.DataFrame:
    """Build synthetic raw rows covering all 36 WANTED survey variables."""
    base = {col: "" for col in WANTED}

    adult_1 = {
        **base,
        "county": "Nairobi City",
        "A08": "Urban",
        "A18": "29",
        "Sex": "Female",
        "Age": "26-35",
        "filt": "1",
        "indWeight": "848.4",
        "livelihoodcat": "Employed",
        "Quintiles": "Fourth",
        "Education": "Tertiary",
        "Marital": "Single",
        "ID_card": "Has ID (National)",
        "Overall_Access_fnl": "formal only",
        "Access_fnl": "Formal Prudential",
        "barriers_credit": "0",
        "mobile_money_access_primary": "Yes",
        "mobile_ownership_1": "Yes",
        "digital_acc_fnl": "Usage",
        "bank_usage_overall": "Usage",
        "Sacco_usage": "Usage",
        "Savings_usage": "Usage",
        "Savings_formal": "Usage",
        "Loan_usage_fnl": "Usage",
        "Loan_formal_2": "Usage",
        "Fuliza_loan": "Usage",
        "finhealth_fnl": "Yes",
    }
    adult_2 = {
        **base,
        "county": "Homabay",
        "A08": "Rural",
        "A18": "42",
        "Sex": "Male",
        "Age": "36-45",
        "filt": "1",
        "indWeight": "52.2",
        "livelihoodcat": "#NULL!",
        "Quintiles": "",
        "Education": "#NULL!",
        "Marital": "Married",
        "ID_card": "No",
        "Overall_Access_fnl": "informal only",
        "Access_fnl": "Informal",
        "barriers_credit": "Lack of collateral",
        "Savings_usage": "Usage",
        "Savings_informal": "Usage",
        "Loan_informal": "Usage",
    }
    adult_3 = {
        **base,
        "county": "Trans Nzoia",
        "A08": "Rural",
        "A18": "60",
        "Sex": "Female",
        "Age": "Above 55",
        "filt": "1",
        "indWeight": "99.4",
        "livelihoodcat": "Agriculture",
        "Quintiles": "Lowest",
        "Education": "Primary",
        "Marital": "Widowed",
        "ID_card": "",
        "Overall_Access_fnl": "excluded",
        "Access_fnl": "Excluded",
        "barriers_credit": "",
    }
    minor_row = {
        **adult_1,
        "county": "Nakuru",
        "filt": "0",
        "indWeight": "500.0",
    }
    return pd.DataFrame([adult_1, adult_2, adult_3, minor_row], columns=WANTED)


def test_col_letters_single_and_multi_letter() -> None:
    assert col_letters(1) == "A"
    assert col_letters(26) == "Z"
    assert col_letters(27) == "AA"
    assert col_letters(52) == "AZ"
    assert col_letters(702) == "ZZ"
    assert col_letters(703) == "AAA"


def test_build_filters_adults_normalizes_and_computes_weights() -> None:
    df = build(_sample_raw_df())

    assert len(df) == 3
    assert df["respondent_id"].tolist() == [1, 2, 3]
    assert df["map_location"].tolist() == [
        "Nairobi County, Kenya",
        "Homa Bay County, Kenya",
        "Trans-Nzoia County, Kenya",
    ]
    assert df["id_card"].tolist() == ["Has ID", "No ID", "No ID"]
    assert df["education"].tolist() == ["Tertiary", "Not stated", "Primary"]
    assert df["wealth_quintile"].tolist() == ["Fourth", "Not stated", "Lowest"]
    assert df["credit_barrier"].tolist() == [
        "None reported",
        "Lack of collateral",
        "None reported",
    ]
    assert df["age_group_sort"].tolist() == [2, 3, 5]
    assert df["wealth_quintile_sort"].tolist() == [4, 6, 1]
    assert df["education_sort"].tolist() == [4, 6, 2]

    assert df["formal_access"].tolist() == [1, 0, 0]
    assert df["informal_only"].tolist() == [0, 1, 0]
    assert df["excluded"].tolist() == [0, 0, 1]
    assert df["mobile_money"].tolist() == [1, 0, 0]
    assert df["fuliza"].tolist() == [1, 0, 0]

    for flag in ["formal_access", "informal_only", "excluded", *USAGE_FLAGS]:
        assert (df[f"w_{flag}"] == df["weight"] * df[flag]).all()


def test_reconcile_passes_and_catches_drift() -> None:
    df = build(_sample_raw_df())
    rec = reconcile(df)
    assert list(rec["indicator"]) == list(PUBLISHED)
    assert rec["matches"].all()

    drifted = df.copy()
    drifted.loc[0, "w_formal_access"] = 800.0
    bad = reconcile(drifted)
    assert not bad.loc[bad["indicator"] == "formal_access", "matches"].iloc[0]


def test_write_dictionary_covers_every_output_column(tmp_path: Path) -> None:
    df = build(_sample_raw_df())
    out_csv = tmp_path / "data_dictionary.csv"
    write_dictionary(df, out_csv)

    dic = pd.read_csv(out_csv)
    assert list(dic["column"]) == list(df.columns)
    assert dic["description"].notna().all()
    assert (dic["description"].str.len() > 0).all()


def test_stream_columns_decodes_shared_strings_and_multiletter_cols(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "mini_survey.xlsx"
    shared_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<si><t>Nairobi City</t></si>"
        "<si><t>Female</t></si>"
        "</sst>"
    )
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<sheetData>"
        '<row r="1"><c r="A1"><v>county</v></c><c r="B1"><v>indWeight</v></c><c r="AA1"><v>Sex</v></c></row>'
        '<row r="2"><c r="A2" t="s"><v>0</v></c><c r="B2"><v>123.5</v></c><c r="AA2" t="s"><v>1</v></c></row>'
        '<row r="3"><c r="A3" t="s"><v>0</v></c><c r="B3"/></row>'
        "</sheetData>"
        "</worksheet>"
    )
    with zipfile.ZipFile(xlsx_path, "w") as zf:
        zf.writestr("xl/sharedStrings.xml", shared_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)

    df = stream_columns(xlsx_path, {1: "county", 2: "indWeight", 27: "Sex"})
    assert list(df.columns) == ["county", "indWeight", "Sex"]
    assert df.iloc[0].to_dict() == {
        "county": "Nairobi City",
        "indWeight": "123.5",
        "Sex": "Female",
    }
    assert df.iloc[1].to_dict() == {
        "county": "Nairobi City",
        "indWeight": "",
        "Sex": "",
    }
