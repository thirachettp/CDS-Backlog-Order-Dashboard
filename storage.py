"""Local-disk persistence helpers for the CDS Backlog Dashboard.

All data is stored under ./data relative to this file so it survives
across Streamlit reruns and app restarts (as long as the host filesystem
is persistent — see README for notes on ephemeral hosts).
"""
import json
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd

DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)

ORDER_FILE = DATA_DIR / "latest_order.xlsx"
ORDER_META = DATA_DIR / "latest_order_meta.json"

OWNER_MAP_FILE = DATA_DIR / "owner_mapping.json"

CUTOFF_FILE = DATA_DIR / "latest_stock_cutoff.csv"
CUTOFF_META = DATA_DIR / "latest_stock_cutoff_meta.json"


def _now() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


# ---------------------------------------------------------------------------
# Order status file
# ---------------------------------------------------------------------------
def save_order_file(uploaded_file) -> None:
    with open(ORDER_FILE, "wb") as f:
        f.write(uploaded_file.getbuffer())
    ORDER_META.write_text(json.dumps({
        "filename": uploaded_file.name,
        "uploaded_at": _now(),
    }, ensure_ascii=False))


def has_saved_order_file() -> bool:
    return ORDER_FILE.exists()


def load_saved_order_meta() -> dict | None:
    if ORDER_META.exists():
        return json.loads(ORDER_META.read_text())
    return None


# ---------------------------------------------------------------------------
# Owner mapping (Sub Dept -> Owner)
# ---------------------------------------------------------------------------
def load_owner_map() -> dict:
    if OWNER_MAP_FILE.exists():
        return json.loads(OWNER_MAP_FILE.read_text())
    return {}


def save_owner_map(mapping: dict) -> None:
    OWNER_MAP_FILE.write_text(json.dumps(mapping, ensure_ascii=False, indent=2))


# ---------------------------------------------------------------------------
# Stock cut-off schedule
# ---------------------------------------------------------------------------
def save_cutoff_data(df: pd.DataFrame, filename: str) -> None:
    df.to_csv(CUTOFF_FILE, index=False)
    CUTOFF_META.write_text(json.dumps({
        "filename": filename,
        "uploaded_at": _now(),
        "rows": len(df),
    }, ensure_ascii=False))


def load_cutoff_data() -> pd.DataFrame:
    if CUTOFF_FILE.exists():
        df = pd.read_csv(CUTOFF_FILE)
        for c in ["CutFrom", "CutTo", "ResumeDate"]:
            if c in df.columns:
                df[c] = pd.to_datetime(df[c], errors="coerce")
        df["SubDept"] = df["SubDept"].astype(str)
        df["StoreCode"] = df["StoreCode"].astype(str)
        return df
    return pd.DataFrame()


def load_cutoff_meta() -> dict | None:
    if CUTOFF_META.exists():
        return json.loads(CUTOFF_META.read_text())
    return None


def clear_all() -> None:
    for p in [ORDER_FILE, ORDER_META, OWNER_MAP_FILE, CUTOFF_FILE, CUTOFF_META]:
        if p.exists():
            p.unlink()
