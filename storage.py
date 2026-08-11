"""Local-disk persistence helpers for the CDS Backlog Dashboard.

All data is stored under ./data relative to this file so it survives
across Streamlit reruns and app restarts (as long as the host filesystem
is persistent — see README for notes on ephemeral hosts).
"""
import json
import shutil
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)

BANGKOK_TZ = ZoneInfo("Asia/Bangkok")

ORDER_FILE = DATA_DIR / "latest_order.xlsx"
ORDER_META = DATA_DIR / "latest_order_meta.json"

OWNER_MAP_FILE = DATA_DIR / "owner_mapping.json"

CUTOFF_FILE = DATA_DIR / "latest_stock_cutoff.csv"
CUTOFF_META = DATA_DIR / "latest_stock_cutoff_meta.json"


def _now() -> str:
    """Current time formatted in Thailand local time (Asia/Bangkok), regardless
    of what timezone the host server (e.g. Streamlit Cloud, which runs UTC)
    is set to."""
    return datetime.now(BANGKOK_TZ).strftime("%d/%m/%Y %H:%M:%S")


def _safe_read_json(path: Path):
    """Reads a JSON file, returning None if it's missing, empty, or corrupted
    (e.g. left in a partial state by a sync client like OneDrive) instead of
    raising, so the app can fall back gracefully rather than crash."""
    if not path.exists():
        return None
    try:
        text = path.read_text(encoding="utf-8").strip()
        if not text:
            return None
        return json.loads(text)
    except (json.JSONDecodeError, OSError, UnicodeDecodeError):
        return None


def _safe_write_json(path: Path, data) -> None:
    """Writes JSON atomically (write to temp file then replace) to avoid
    leaving a half-written/empty file behind if interrupted mid-write."""
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    tmp_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp_path.replace(path)


# ---------------------------------------------------------------------------
# Order status file
# ---------------------------------------------------------------------------
def save_order_file(uploaded_file) -> None:
    tmp_path = ORDER_FILE.with_suffix(ORDER_FILE.suffix + ".tmp")
    with open(tmp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    tmp_path.replace(ORDER_FILE)
    _safe_write_json(ORDER_META, {
        "filename": uploaded_file.name,
        "uploaded_at": _now(),
    })


def has_saved_order_file() -> bool:
    # Guard against a 0-byte file left behind by an interrupted write/sync.
    return ORDER_FILE.exists() and ORDER_FILE.stat().st_size > 0


def load_saved_order_meta() -> dict | None:
    return _safe_read_json(ORDER_META)


# ---------------------------------------------------------------------------
# Owner mapping (Sub Dept -> Owner)
# ---------------------------------------------------------------------------
def load_owner_map() -> dict:
    return _safe_read_json(OWNER_MAP_FILE) or {}


def save_owner_map(mapping: dict) -> None:
    _safe_write_json(OWNER_MAP_FILE, mapping)


# ---------------------------------------------------------------------------
# Stock cut-off schedule (append-only, active entries only)
# ---------------------------------------------------------------------------
def _today_ts() -> pd.Timestamp:
    return pd.Timestamp(datetime.now(BANGKOK_TZ).date())


def append_cutoff_data(new_df: pd.DataFrame, filename: str) -> dict:
    """Appends newly-uploaded, still-active cut-off rows onto the existing
    accumulated dataset, then prunes any row (old or new) whose end date
    (CutTo) has already passed. Returns a small summary dict for UI feedback.
    """
    today = _today_ts()

    new_df = new_df.copy()
    new_df["CutTo"] = pd.to_datetime(new_df["CutTo"], errors="coerce")
    new_active = new_df[new_df["CutTo"].notna() & (new_df["CutTo"] >= today)]
    skipped_expired = len(new_df) - len(new_active)

    existing = load_cutoff_data()
    combined = pd.concat([existing, new_active], ignore_index=True) if not existing.empty else new_active.copy()

    dedupe_cols = [c for c in ["StoreCode", "SubDept", "CutFrom", "CutTo"] if c in combined.columns]
    before_dedupe = len(combined)
    if dedupe_cols:
        combined = combined.drop_duplicates(subset=dedupe_cols, keep="last")
    duplicates_removed = before_dedupe - len(combined)

    combined["CutTo"] = pd.to_datetime(combined["CutTo"], errors="coerce")
    before_prune = len(combined)
    combined = combined[combined["CutTo"].notna() & (combined["CutTo"] >= today)]
    expired_pruned = before_prune - len(combined)

    combined = combined.sort_values(["CutTo", "StoreCode", "SubDept"]).reset_index(drop=True)

    tmp_path = CUTOFF_FILE.with_suffix(CUTOFF_FILE.suffix + ".tmp")
    combined.to_csv(tmp_path, index=False)
    tmp_path.replace(CUTOFF_FILE)

    meta = _safe_read_json(CUTOFF_META) or {}
    history = meta.get("upload_history", [])
    history.append({
        "filename": filename,
        "uploaded_at": _now(),
        "rows_added": len(new_active),
        "rows_skipped_expired": int(skipped_expired),
    })
    history = history[-20:]  # keep last 20 uploads only
    _safe_write_json(CUTOFF_META, {
        "last_filename": filename,
        "uploaded_at": _now(),
        "rows": len(combined),
        "upload_history": history,
    })

    return {
        "rows_added": len(new_active),
        "rows_skipped_expired": int(skipped_expired),
        "duplicates_removed": int(duplicates_removed),
        "expired_pruned": int(expired_pruned),
        "total_active_rows": len(combined),
    }


def load_cutoff_data() -> pd.DataFrame:
    if not CUTOFF_FILE.exists() or CUTOFF_FILE.stat().st_size == 0:
        return pd.DataFrame()
    try:
        df = pd.read_csv(CUTOFF_FILE)
    except (pd.errors.EmptyDataError, OSError):
        return pd.DataFrame()
    if df.empty:
        return pd.DataFrame()
    for c in ["CutFrom", "CutTo", "ResumeDate"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    if "SubDept" in df.columns:
        df["SubDept"] = df["SubDept"].astype(str)
    if "StoreCode" in df.columns:
        df["StoreCode"] = df["StoreCode"].astype(str)
    return df


def load_cutoff_meta() -> dict | None:
    return _safe_read_json(CUTOFF_META)


def clear_all() -> None:
    for p in [ORDER_FILE, ORDER_META, OWNER_MAP_FILE, CUTOFF_FILE, CUTOFF_META]:
        if p.exists():
            p.unlink()
