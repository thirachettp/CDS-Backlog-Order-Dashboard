"""Local-disk persistence helpers for the CDS Backlog Dashboard.

All data is stored under ./data relative to this file so it survives
across Streamlit reruns and app restarts (as long as the host filesystem
is persistent — see README for notes on ephemeral hosts).
"""
import json
import re
import shutil
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)

BANGKOK_TZ = ZoneInfo("Asia/Bangkok")

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
# Order status data — stored per BU so uploading one BU's file never wipes
# out another BU's previously saved data. Each BU gets its own CSV (parsed
# data, not the raw xlsx) plus a small meta file.
# ---------------------------------------------------------------------------
ORDER_DIR = DATA_DIR / "orders"
ORDER_DIR.mkdir(exist_ok=True)

_DATE_COLS = ["Order Create Date", "Start Ship Date", "Allocated Date", "Picked Date",
              "Shipped Date", "Scan Load Date"]
_NUMERIC_COLS = ["Required QTY", "Allocated QTY", "Pick QTY"]
_STR_CODE_COLS = ["Sub Dept", "To Store Str", "BU"]


def _safe_bu_key(bu: str) -> str:
    """Turns a BU code into a filesystem-safe file name stem."""
    key = re.sub(r"[^A-Za-z0-9_\-]+", "_", str(bu).strip())
    return key or "UNKNOWN"


def _order_csv_path(bu: str) -> Path:
    return ORDER_DIR / f"{_safe_bu_key(bu)}.csv"


def _order_meta_path(bu: str) -> Path:
    return ORDER_DIR / f"{_safe_bu_key(bu)}_meta.json"


def restore_order_dtypes(df: pd.DataFrame) -> pd.DataFrame:
    """Re-applies proper dtypes after a round-trip through CSV, where
    everything comes back as plain strings/objects."""
    df = df.copy()
    for c in _DATE_COLS:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    for c in _NUMERIC_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    for c in _STR_CODE_COLS:
        if c in df.columns:
            df[c] = df[c].astype(str).replace({"nan": pd.NA, "None": pd.NA, "<NA>": pd.NA})
    return df


def save_order_data(df: pd.DataFrame, filename: str) -> list[str]:
    """Splits the parsed order dataframe by its 'BU' column and saves each
    BU's slice to its own file, overwriting only that BU's previous data.
    Returns the list of BU codes that were saved."""
    if "BU" not in df.columns:
        df = df.copy()
        df["BU"] = "UNKNOWN"

    saved_bus = []
    for bu, sub_df in df.groupby(df["BU"].fillna("UNKNOWN")):
        bu = str(bu)
        csv_path = _order_csv_path(bu)
        tmp_path = csv_path.with_suffix(csv_path.suffix + ".tmp")
        sub_df.to_csv(tmp_path, index=False)
        tmp_path.replace(csv_path)
        _safe_write_json(_order_meta_path(bu), {
            "bu": bu,
            "filename": filename,
            "uploaded_at": _now(),
            "rows": len(sub_df),
        })
        saved_bus.append(bu)
    return saved_bus


def list_saved_bus() -> list[str]:
    return sorted({p.stem for p in ORDER_DIR.glob("*.csv")})


def load_order_data(bu: str) -> pd.DataFrame:
    path = _order_csv_path(bu)
    if not path.exists() or path.stat().st_size == 0:
        return pd.DataFrame()
    try:
        df = pd.read_csv(path)
    except (pd.errors.EmptyDataError, OSError):
        return pd.DataFrame()
    if df.empty:
        return df
    return restore_order_dtypes(df)


def load_order_meta(bu: str) -> dict | None:
    return _safe_read_json(_order_meta_path(bu))


def load_all_order_data() -> tuple[pd.DataFrame, dict]:
    """Loads and concatenates every saved BU's order data. Returns
    (combined_df, {bu: meta_dict}) — combined_df is empty if nothing saved."""
    bus = list_saved_bus()
    frames = []
    metas = {}
    for bu in bus:
        d = load_order_data(bu)
        if not d.empty:
            frames.append(d)
        m = load_order_meta(bu)
        if m:
            metas[bu] = m
    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return combined, metas


def has_saved_order_data() -> bool:
    return len(list_saved_bus()) > 0


def clear_order_bu(bu: str) -> None:
    for p in [_order_csv_path(bu), _order_meta_path(bu)]:
        if p.exists():
            p.unlink()


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
    # Free-text columns can come back as all-NaN float64 after a CSV
    # round-trip if every value was empty — coerce to a proper string dtype
    # so editors (e.g. st.column_config.TextColumn) don't choke on them.
    for c in ["StoreName", "SubDeptNote", "Remark"]:
        if c in df.columns:
            df[c] = df[c].apply(lambda v: "" if pd.isna(v) else str(v))
    return df


def load_cutoff_meta() -> dict | None:
    return _safe_read_json(CUTOFF_META)


def overwrite_cutoff_data(df: pd.DataFrame) -> dict:
    """Replaces the whole active Stock Cut-off dataset with an edited
    version (used by the editable table in Settings), pruning any row
    whose end date has already passed. Returns a small summary dict."""
    today = _today_ts()
    df = df.copy()
    df["CutFrom"] = pd.to_datetime(df["CutFrom"], errors="coerce")
    df["CutTo"] = pd.to_datetime(df["CutTo"], errors="coerce")

    before = len(df)
    df = df[df["CutTo"].notna() & (df["CutTo"] >= today)]
    expired_pruned = before - len(df)

    dedupe_cols = [c for c in ["StoreCode", "SubDept", "CutFrom", "CutTo"] if c in df.columns]
    if dedupe_cols:
        df = df.drop_duplicates(subset=dedupe_cols, keep="last")

    df = df.sort_values(["CutTo", "StoreCode", "SubDept"]).reset_index(drop=True)

    tmp_path = CUTOFF_FILE.with_suffix(CUTOFF_FILE.suffix + ".tmp")
    df.to_csv(tmp_path, index=False)
    tmp_path.replace(CUTOFF_FILE)

    meta = _safe_read_json(CUTOFF_META) or {}
    history = meta.get("upload_history", [])
    history.append({
        "filename": "(แก้ไขด้วยตนเองในหน้า Settings)",
        "uploaded_at": _now(),
        "rows_added": 0,
        "rows_skipped_expired": 0,
    })
    history = history[-20:]
    _safe_write_json(CUTOFF_META, {
        "last_filename": meta.get("last_filename", "-"),
        "uploaded_at": _now(),
        "rows": len(df),
        "upload_history": history,
    })

    return {"total_active_rows": len(df), "expired_pruned": int(expired_pruned)}


def clear_all() -> None:
    for p in [OWNER_MAP_FILE, CUTOFF_FILE, CUTOFF_META]:
        if p.exists():
            p.unlink()
    for bu in list_saved_bus():
        clear_order_bu(bu)
