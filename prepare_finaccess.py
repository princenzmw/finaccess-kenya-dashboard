"""Prepare the FinAccess 2024 household survey for Power BI.

The source workbook is ~290 MB (3 GB of XML), so opening it in Excel or Power BI is painfully slow.
This script streams the big sheet, keeps ~40 columns, decodes the text labels, adds 0/1 indicators and
weighted columns, and writes small CSVs. It also checks the weighted results against the published
headline figures (84.8 % formal, 5.2 % informal only, 9.9 % excluded).

Usage:  python prepare_finaccess.py --xlsx "path/to/finaccess_2024.xlsx" --out output
Needs:  pandas, openpyxl (pip install pandas openpyxl)
"""
from __future__ import annotations

import argparse
import csv
import re
import time
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pandas as pd

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# Variables kept from the survey (names as in the workbook's 'Variables' sheet).
WANTED = [
    "county", "A08", "A18", "Sex", "Age", "filt", "indWeight", "livelihoodcat", "Quintiles", "Education",
    "Marital", "ID_card", "Sacco_usage", "Savings_usage", "Savings_formal", "Savings_informal",
    "Loan_informal", "Loan_formal_2", "formal_access_fnl", "Informal_usage_1", "Access_fnl", "finhealth_fnl",
    "Overall_Access_fnl", "Microfinance_usage_fnl", "mobile_money_access_primary", "bank_usage_overall",
    "barriers_credit", "Insurance_including_NHIF_fnl", "Loan_usage_fnl", "Mobile_banking_loan", "Fuliza_loan",
    "digital_app_loan", "hustler_fund_loan", "mobile_ownership_1", "Digital_credit_usage_fnl", "digital_acc_fnl",
]
# Deliberately NOT used: 'defaulted' (35.6 % of adults say yes, so it is not a default rate; definition unclear)
# and 'Access_Strand' (almost entirely zeros).

USAGE_FLAGS = {  # new_name: (source column, value meaning 1)
    "mobile_money": ("mobile_money_access_primary", "Yes"),
    "mobile_ownership": ("mobile_ownership_1", "Yes"),
    "digital_account": ("digital_acc_fnl", "Usage"),
    "bank_use": ("bank_usage_overall", "Usage"),
    "sacco_use": ("Sacco_usage", "Usage"),
    "microfinance_use": ("Microfinance_usage_fnl", "Usage"),
    "savings_any": ("Savings_usage", "Usage"),
    "savings_formal": ("Savings_formal", "Usage"),
    "savings_informal": ("Savings_informal", "Usage"),
    "loan_any": ("Loan_usage_fnl", "Usage"),
    "loan_formal": ("Loan_formal_2", "Usage"),
    "loan_informal": ("Loan_informal", "Usage"),
    "fuliza": ("Fuliza_loan", "Usage"),
    "hustler_fund": ("hustler_fund_loan", "Usage"),
    "mobile_banking_loan": ("Mobile_banking_loan", "Usage"),
    "digital_app_loan": ("digital_app_loan", "Usage"),
    "digital_credit_any": ("Digital_credit_usage_fnl", "Usage"),
    "insurance_any": ("Insurance_including_NHIF_fnl", "Usage"),
    "financially_healthy": ("finhealth_fnl", "Yes"),
}
AGE_ORDER = {"18-25": 1, "26-35": 2, "36-45": 3, "46-55": 4, "Above 55": 5}
WEALTH_ORDER = {"Lowest": 1, "Second": 2, "Middle": 3, "Fourth": 4, "Highest": 5, "Not stated": 6}
EDU_ORDER = {"None": 1, "Primary": 2, "Secondary": 3, "Tertiary": 4, "Other": 5, "Not stated": 6}
MAP_NAME_FIXES = {"Nairobi City": "Nairobi", "Homabay": "Homa Bay", "Trans Nzoia": "Trans-Nzoia"}
PUBLISHED = {"formal_access": 84.8, "informal_only": 5.2, "excluded": 9.9}  # FinAccess 2024 topline report


def col_letters(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def shared_strings(z: zipfile.ZipFile) -> list[str]:
    out = []
    for _, el in ET.iterparse(z.open("xl/sharedStrings.xml")):
        if el.tag == NS + "si":
            out.append("".join(t.text or "" for t in el.iter(NS + "t")))
            el.clear()
    return out


def stream_columns(xlsx: Path, positions: dict[int, str]) -> pd.DataFrame:
    z = zipfile.ZipFile(xlsx)
    ss = shared_strings(z)
    pos = sorted(positions)
    l2pos = {col_letters(p): p for p in pos}
    alt = "|".join(sorted(l2pos, key=len, reverse=True)).encode()
    cell_re = re.compile(rb'<c r="(' + alt + rb')\d+"([^>]*?)(?:/>|>(?:<v>([^<]*)</v>)?(?:<is>.*?</is>)?</c>)')
    row_re = re.compile(rb'<row r="(\d+)"')
    rows, t0, buf = [], time.time(), b""
    with z.open("xl/worksheets/sheet1.xml") as f:
        while True:
            chunk = f.read(16 * 1024 * 1024)
            buf += chunk
            parts = buf.split(b"</row>")
            buf = parts.pop() if chunk else b""
            for seg in parts:
                m = row_re.search(seg)
                if not m or int(m.group(1)) == 1:
                    continue
                rec = {}
                for cm in cell_re.finditer(seg):
                    raw = (cm.group(3) or b"").decode()
                    if b't="s"' in cm.group(2) and raw != "":
                        raw = ss[int(raw)]
                    rec[l2pos[cm.group(1).decode()]] = raw
                rows.append([rec.get(p, "") for p in pos])
            if not chunk:
                break
    print(f"streamed {len(rows):,} rows in {time.time() - t0:.0f}s")
    return pd.DataFrame(rows, columns=[positions[p] for p in pos])


def build(raw: pd.DataFrame) -> pd.DataFrame:
    raw = raw.replace({"": pd.NA, "#NULL!": pd.NA})
    raw["indWeight"] = pd.to_numeric(raw["indWeight"])
    raw["A18"] = pd.to_numeric(raw["A18"])
    a = raw[raw["filt"].astype(str) == "1"].reset_index(drop=True)  # adults 18+ only
    out = pd.DataFrame({"respondent_id": range(1, len(a) + 1)})
    out["county"] = a["county"]
    out["map_location"] = a["county"].replace(MAP_NAME_FIXES) + " County, Kenya"
    out["area_type"] = a["A08"]
    out["sex"] = a["Sex"]
    out["age_group"] = a["Age"]
    out["age"] = a["A18"]
    out["education"] = a["Education"].fillna("Not stated")
    out["wealth_quintile"] = a["Quintiles"].fillna("Not stated")
    out["livelihood"] = a["livelihoodcat"].fillna("Not stated")
    out["marital_status"] = a["Marital"].fillna("Not stated")
    out["id_card"] = a["ID_card"].map(lambda v: "Has ID" if str(v).startswith("Has ID") else "No ID")
    out["age_group_sort"] = out["age_group"].map(AGE_ORDER)
    out["wealth_quintile_sort"] = out["wealth_quintile"].map(WEALTH_ORDER)
    out["education_sort"] = out["education"].map(EDU_ORDER)
    out["weight"] = a["indWeight"]
    out["access_level"] = a["Access_fnl"]
    out["access_group"] = a["Overall_Access_fnl"].map(
        {"formal only": "Formal", "informal only": "Informal only", "excluded": "Excluded"})
    out["credit_barrier"] = a["barriers_credit"].replace({"0": "None reported"}).fillna("None reported")
    out["formal_access"] = (out["access_group"] == "Formal").astype(int)
    out["informal_only"] = (out["access_group"] == "Informal only").astype(int)
    out["excluded"] = (out["access_group"] == "Excluded").astype(int)
    for new, (src, yes) in USAGE_FLAGS.items():
        out[new] = (a[src] == yes).astype(int)
    for flag in ["formal_access", "informal_only", "excluded", *USAGE_FLAGS]:
        out[f"w_{flag}"] = out["weight"] * out[flag]
    return out


def reconcile(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for k, published in PUBLISHED.items():
        mine = df[f"w_{k}"].sum() / df["weight"].sum() * 100
        rows.append({"indicator": k, "this_dataset_pct": round(mine, 1), "published_topline_pct": published,
                     "matches": abs(mine - published) < 0.05})
    return pd.DataFrame(rows)


DESCRIPTIONS = {
    "respondent_id": "Sequential row number (not the survey interview key)",
    "county": "County of residence (47 counties)",
    "map_location": "County name formatted for Power BI map geocoding",
    "area_type": "Rural or Urban cluster",
    "sex": "Sex of respondent",
    "age_group": "Age group of respondent (adults 18+ only)",
    "age": "Age in completed years",
    "education": "Highest education level; 'Not stated' when missing",
    "wealth_quintile": "Household wealth quintile from the survey; 'Not stated' when missing",
    "livelihood": "Main livelihood category",
    "marital_status": "Marital status",
    "id_card": "Owns a national ID, passport or alien ID",
    "weight": "Individual survey weight (sum = adult population represented). ALWAYS use for percentages",
    "access_level": "Detailed access strand: Formal Prudential / Non-Prudential / Registered, Informal, Excluded",
    "access_group": "Formal / Informal only / Excluded (matches the published topline split)",
    "credit_barrier": "Main barrier to credit; 'None reported' for 0 (may include people not asked)",
    "formal_access": "1 if uses formal financial services",
    "informal_only": "1 if uses informal services only",
    "excluded": "1 if uses no financial services",
    "mobile_money": "1 if has accessed mobile money on own phone",
    "mobile_ownership": "1 if owns a mobile phone (primary ownership)",
    "digital_account": "1 if uses a digital account",
    "bank_use": "1 if uses a bank (incl. mobile banking providers regulated by CBK)",
    "sacco_use": "1 if uses a SACCO",
    "microfinance_use": "1 if uses microfinance",
    "savings_any": "1 if saves anywhere", "savings_formal": "1 if saves with a formal provider",
    "savings_informal": "1 if saves informally",
    "loan_any": "1 if has a loan (current or last 12 months)", "loan_formal": "1 if has a formal loan",
    "loan_informal": "1 if has an informal loan",
    "fuliza": "1 if used a Fuliza overdraft loan in the last 12 months",
    "hustler_fund": "1 if took a Hustler Fund loan in the last 12 months",
    "mobile_banking_loan": "1 if took a mobile banking loan in the last 12 months",
    "digital_app_loan": "1 if took a digital app loan in the last 12 months",
    "digital_credit_any": "1 if used any digital credit (Fuliza, apps, Hustler Fund, mobile banking, digital MFIs)",
    "insurance_any": "1 if has insurance including NHIF",
    "financially_healthy": "1 if classed as financially healthy by the survey index",
}


def write_dictionary(df: pd.DataFrame, path: Path) -> None:
    rows = []
    for c in df.columns:
        if c.startswith("w_"):
            desc = f"weight x {c[2:]} (sum / sum of weight = weighted share)"
        elif c.endswith("_sort"):
            desc = "Sort-order helper: use 'Sort by column' in Power BI"
        else:
            desc = DESCRIPTIONS.get(c, "")
        rows.append({"column": c, "description": desc, "type": str(df[c].dtype)})
    pd.DataFrame(rows).to_csv(path, index=False)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True, type=Path)
    ap.add_argument("--out", default=Path("output"), type=Path)
    args = ap.parse_args()
    args.out.mkdir(parents=True, exist_ok=True)

    var = pd.read_excel(args.xlsx, sheet_name="Variables")
    found = var[var["name"].isin(WANTED)]
    missing = set(WANTED) - set(found["name"])
    if missing:
        raise SystemExit(f"Variables not found in workbook: {sorted(missing)}")
    positions = {int(r.position): r["name"] for _, r in found.iterrows()}
    df = build(stream_columns(args.xlsx, positions))

    df.to_csv(args.out / "finaccess_adults.csv", index=False)
    counties = (df.groupby(["county", "map_location"]).size().rename("sample_size").reset_index())
    counties.to_csv(args.out / "dim_county.csv", index=False)
    write_dictionary(df, args.out / "data_dictionary.csv")
    rec = reconcile(df)
    rec.to_csv(args.out / "reconciliation.csv", index=False)
    print(rec.to_string(index=False))
    print(f"wrote {len(df):,} adult respondents; weighted adults = {df['weight'].sum() / 1e6:.2f} million")
    if not rec["matches"].all():
        raise SystemExit("Reconciliation FAILED: do not use these outputs.")


if __name__ == "__main__":
    main()
