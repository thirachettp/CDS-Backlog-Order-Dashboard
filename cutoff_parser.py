"""Parses the 'CDS' sheet of the Stock Cut-off schedule Excel file into a
tidy dataframe: one row per Store x Sub Dept x blackout period.
"""
import re
from datetime import date

import pandas as pd


def _parse_blackout_range(text):
    """Parses strings like '4-10/08/2026' into (date(2026,8,4), date(2026,8,10))."""
    if pd.isna(text):
        return (None, None)
    m = re.match(r"\s*(\d+)\s*-\s*(\d+)\s*/\s*(\d+)\s*/\s*(\d+)\s*", str(text))
    if not m:
        return (None, None)
    d1, d2, mo, yr = map(int, m.groups())
    try:
        end_date = date(yr, mo, d2)
        if d1 <= d2:
            start_date = date(yr, mo, d1)
        else:
            prev_month = mo - 1 if mo > 1 else 12
            prev_year = yr if mo > 1 else yr - 1
            start_date = date(prev_year, prev_month, d1)
    except ValueError:
        return (None, None)
    return (start_date, end_date)


def parse_cutoff_excel(file) -> pd.DataFrame:
    """Reads only the 'CDS' sheet and returns a tidy dataframe with columns:
    StoreCode, StoreName, SubDept, SubDeptNote, CutFrom, CutTo, ResumeDate, Remark
    """
    df = pd.read_excel(file, sheet_name="CDS", header=1)
    df = df.iloc[:, :10]
    df.columns = [
        "RowNo", "StoreCodeRaw", "StoreName", "SubDept", "SubDeptNote",
        "_blank", "EventDate", "BlackoutText", "ResumeDate", "Remark",
    ]
    df = df.drop(columns=["_blank", "EventDate"])
    df[["RowNo", "StoreCodeRaw", "StoreName"]] = df[["RowNo", "StoreCodeRaw", "StoreName"]].ffill()
    df = df.dropna(subset=["SubDept"]).copy()

    df["StoreCode"] = df["StoreCodeRaw"].astype(str).str.split(":").str[-1].str.strip()
    df["SubDept"] = df["SubDept"].astype(float).astype(int).astype(str)
    df["ResumeDate"] = pd.to_datetime(df["ResumeDate"], errors="coerce")

    ranges = df["BlackoutText"].apply(_parse_blackout_range)
    df["CutFrom"] = ranges.apply(lambda t: t[0])
    df["CutTo"] = ranges.apply(lambda t: t[1])
    df["CutFrom"] = pd.to_datetime(df["CutFrom"], errors="coerce")
    df["CutTo"] = pd.to_datetime(df["CutTo"], errors="coerce")

    out = df[["StoreCode", "StoreName", "SubDept", "SubDeptNote", "CutFrom", "CutTo",
              "ResumeDate", "Remark"]].reset_index(drop=True)
    return out
