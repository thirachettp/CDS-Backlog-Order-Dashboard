import html
import io
import re
from datetime import date

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import storage
from cutoff_parser import parse_cutoff_excel

APP_TITLE = "Sprint WMS FC - Backlog Order Dashboard"

st.set_page_config(page_title=APP_TITLE, layout="wide", initial_sidebar_state="expanded")

# ---------------------------------------------------------------------------
# STYLING
# ---------------------------------------------------------------------------
st.markdown("""
<style>
.kpi-card {
    background: #ffffff;
    border-radius: 10px;
    padding: 14px 16px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    border-top: 4px solid #6C5CE7;
    height: 100%;
}
.kpi-label { font-size: 11px; font-weight: 700; color: #6b7280; letter-spacing: .04em; text-transform: uppercase; }
.kpi-value { font-size: 24px; font-weight: 800; color: #111827; margin: 4px 0 2px 0; }
.kpi-sub   { font-size: 11px; color: #6b7280; }

.kpi2-card {
    background: #ffffff;
    border-radius: 10px;
    padding: 12px 14px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    border-top: 4px solid #6C5CE7;
    height: 100%;
}
.kpi2-card.excluded { background: #fafafa; }
.kpi2-card.headline { border-top-width: 6px; }
.kpi2-label { font-size: 12px; font-weight: 800; color: #111827; margin-bottom: 6px; }
.kpi2-block { margin-bottom: 4px; }
.kpi2-block:last-child { margin-bottom: 0; }
.kpi2-tag { font-size: 9px; font-weight: 700; color: #9CA3AF; letter-spacing: .05em; text-transform: uppercase; }
.kpi2-val { font-size: 19px; font-weight: 800; color: #111827; line-height: 1.25; }
.kpi2-pct { font-size: 11px; color: #6b7280; }
.kpi2-divider { border-top: 1px dashed #e5e7eb; margin: 6px 0; }

.section-title { font-size: 22px; font-weight: 800; color: #111827; margin: 6px 0 2px 0; }
.subsection-title { font-size: 15px; font-weight: 800; color: #374151; margin: 14px 0 4px 0; }
.group-title { font-size: 18px; font-weight: 800; color: #DC2626; margin: 18px 0 2px 0; }
.group-sub { font-size: 14px; font-weight: 700; color: #111827; margin-left: 6px; }

.pivot-wrap { overflow-x: auto; border: 1px solid #e5e7eb; border-radius: 8px; margin: 4px 0 14px 0; }
.pivot-table { border-collapse: collapse; width: max-content; min-width: 100%; font-size: 13px; }
.pivot-table th, .pivot-table td {
    padding: 6px 12px; text-align: right; white-space: nowrap; border-bottom: 1px solid #f1f5f9;
}
.pivot-table thead th {
    position: sticky; top: 0; background: #f9fafb; color: #374151; font-weight: 700;
    border-bottom: 1px solid #e5e7eb; z-index: 1;
}
.pivot-table th:first-child, .pivot-table td:first-child { text-align: left; }
.pivot-table td:last-child, .pivot-table th:last-child { font-weight: 700; }
.pivot-table tr.pivot-group-header td {
    background: #eef2ff; padding: 10px 12px; border-top: 2px solid #c7d2fe; border-bottom: none;
}
.pivot-table tr.pivot-total td { background: #fde68a; font-weight: 800; }
.pivot-table tr.pivot-empty td { color: #6b7280; font-style: italic; padding: 10px 12px; }
.flow-box { background: #EEF2FF; border: 1px solid #C7D2FE; border-radius: 8px; padding: 8px 12px;
            font-weight: 700; font-size: 12px; white-space: nowrap; }
div.block-container { padding-top: 1.2rem; }
</style>
""", unsafe_allow_html=True)

CANCELLED_LABEL = "ยกเลิก"
WAITING_LABEL = "รอเบิก"
ALLOCATED_LABELS = ["Allocated", "Allocated Partial", "Allocated ShortAll"]

FINAL_STATUS_ORDER = ["รอเบิก", "รอยืนยัน", "Allocated", "Allocated Partial", "Allocated ShortAll",
                       "ปิดเอกสาร", "Picked", "Loaded", "Shipped", "Cancelled"]
CARD_COLORS = ["#6C5CE7", "#DC2626", "#2563EB", "#059669", "#F59E0B", "#0EA5E9", "#16A34A"]
THAI_WD = ["จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส.", "อา."]

# Display-only relabeling: the underlying Final Status values (and the source
# Excel file) keep their original names — this map only changes what's shown
# on screen. "Allocated / Allocated Partial / Allocated ShortAll" are shown
# merged as one "Released" group; "ยกเลิก" (already normalized to internal
# "Cancelled") and "ปิดเอกสาร" are both shown merged as "Cancelled".
STATUS_DISPLAY_MAP = {
    "รอเบิก": "Not Printed",
    "Allocated": "Released",
    "Allocated Partial": "Released",
    "Allocated ShortAll": "Released",
    "Cancelled": "Cancelled",
    "ปิดเอกสาร": "Cancelled",
}


def display_status(final_status: str) -> str:
    return STATUS_DISPLAY_MAP.get(final_status, final_status)


STATUS_DISPLAY_ORDER = ["Not Printed", "รอยืนยัน", "Released", "Picked", "Loaded", "Shipped", "Cancelled"]

DONUT_COLORS = {
    "Not Printed": "#2563EB", "รอยืนยัน": "#7C3AED", "Released": "#059669",
    "Picked": "#F59E0B", "Loaded": "#0EA5E9", "Shipped": "#16A34A", "Cancelled": "#9CA3AF",
}

FLOW_EXPLAINER_HTML = """
<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
  <div class="flow-box">Status</div><div>→</div>
  <div class="flow-box">Is Picked?</div><div>→</div>
  <div class="flow-box">Is Scan Load?</div><div>→</div>
  <div class="flow-box">Is Shipped?</div><div>→</div>
  <div class="flow-box" style="background:#FEF2F2;border-color:#FECACA;">Final Logic Status</div>
</div>
"""


def kpi_card(label, value, sub, color):
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color:{color};">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)


def dual_kpi_card(label, orders_val, orders_pct, pcs_val, pcs_pct, color, muted=False, headline=False):
    classes = "kpi2-card"
    if muted:
        classes += " excluded"
    if headline:
        classes += " headline"
    st.markdown(f"""
    <div class="{classes}" style="border-top-color:{color};">
        <div class="kpi2-label">{label}</div>
        <div class="kpi2-block">
            <div class="kpi2-tag">Orders</div>
            <div class="kpi2-val">{orders_val:,}</div>
            <div class="kpi2-pct">{orders_pct:.1f}% ของ Total Orders</div>
        </div>
        <div class="kpi2-divider"></div>
        <div class="kpi2-block">
            <div class="kpi2-tag">Pcs.</div>
            <div class="kpi2-val">{pcs_val:,.0f}</div>
            <div class="kpi2-pct">{pcs_pct:.1f}% ของ Total Pcs.</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def date_col_label(d) -> str:
    return f"{THAI_WD[d.weekday()]} {d.month}/{d.day}"


def _to_code_str(v):
    """Normalizes a store/dept code that may arrive as int, float (e.g. 303.0),
    or already-clean string into a plain string code (e.g. '303')."""
    if pd.isna(v):
        return pd.NA
    s = str(v).strip()
    if s in ("", "nan", "None"):
        return pd.NA
    try:
        return str(int(float(s)))
    except ValueError:
        return s


def pct(part, whole) -> float:
    return (part / whole * 100) if whole else 0.0


# ---------------------------------------------------------------------------
# DATA LOADING
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def parse_order_file(file_bytes: bytes) -> pd.DataFrame:
    """Reads the Order Status by SKU export. Auto-detects the file layout:
    - New format (Aug 2026+): header on the first row, columns include
      'Owner' (BU code) and 'Warehouse'.
    - Old format: title row + blank column, real header on the 3rd row.
    """
    probe = pd.read_excel(io.BytesIO(file_bytes), header=0, nrows=3)
    if "Order No." in probe.columns and "Status" in probe.columns:
        header_row = 0
    else:
        header_row = 2

    df = pd.read_excel(io.BytesIO(file_bytes), header=header_row)
    drop_cols = [c for c in df.columns if str(c).startswith("Unnamed: 0")]
    df = df.drop(columns=drop_cols, errors="ignore")

    # Normalize column names: some exports use a literal newline inside the
    # header cell (e.g. "Pick\nQTY") — collapse any whitespace run to a
    # single space so downstream column lookups are consistent either way.
    df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]

    str_cols = df.select_dtypes(include=["object", "string"]).columns
    for c in str_cols:
        df[c] = df[c].astype(str).str.strip().replace({"nan": pd.NA, "None": pd.NA})

    for c in ["Required QTY", "Allocated QTY", "Pick QTY"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    for c in ["Order Create Date", "Allocated Date", "Picked Date", "Shipped Date", "Scan Load Date"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    if "Start Ship Date" in df.columns:
        df["Start Ship Date"] = pd.to_datetime(df["Start Ship Date"], format="%d/%m/%Y", errors="coerce")

    if "Sub Dept" in df.columns:
        df["Sub Dept"] = df["Sub Dept"].apply(_to_code_str)
    if "To Store" in df.columns:
        df["To Store Str"] = df["To Store"].apply(_to_code_str)

    # -- Business Unit (BU) --------------------------------------------------
    # New format carries the BU directly in the "Owner" column, e.g.
    # "CDS-CDS" -> BU "CDS". Old format only has "Company/Division", e.g.
    # "CDS - SCDC WH" -> BU "CDS". Falls back to a fixed "CDS" if neither
    # column is present.
    if "Owner" in df.columns:
        df["BU"] = df["Owner"].apply(lambda v: str(v).split("-")[0].strip() if pd.notna(v) else pd.NA)
    elif "Company/Division" in df.columns:
        df["BU"] = df["Company/Division"].apply(lambda v: str(v).split("-")[0].strip() if pd.notna(v) else pd.NA)
    else:
        df["BU"] = "CDS"

    # -- Final Logic Status ---------------------------------------------------
    # Status → Is Picked → Is Scan Load → Is Shipped → Final Logic Status.
    # Cancelled and รอเบิก are terminal shortcuts; everything else walks the
    # picked/scan-load/shipped chain. Old-format files have no scan-load
    # tracking, so that step is skipped for them (Picked jumps straight to
    # Shipped once Is Shipped = Y).
    has_scan_load = "Is Scan Load" in df.columns

    def _final_status(row):
        status = row.get("Status")
        if status == CANCELLED_LABEL:
            return "Cancelled"
        if status == WAITING_LABEL:
            return WAITING_LABEL
        if row.get("Is Picked") != "Y":
            return status
        if has_scan_load:
            if row.get("Is Scan Load") != "Y":
                return "Picked"
            if row.get("Is Shipped") != "Y":
                return "Loaded"
            return "Shipped"
        else:
            return "Shipped" if row.get("Is Shipped") == "Y" else "Picked"

    df["Final Status"] = df.apply(_final_status, axis=1)

    return df


def get_bu_order_df(bu: str):
    """Returns (df, source_note, meta) for a single BU's saved data only, so
    the Dashboard always shows exactly one BU's data at a time."""
    if not bu:
        return None, None, None
    d = storage.load_order_data(bu)
    if d.empty:
        return None, None, None
    meta = storage.load_order_meta(bu) or {}
    note = f"BU: {bu} | ไฟล์ล่าสุด: {meta.get('filename', '-')} (อัปโหลดเมื่อ {meta.get('uploaded_at', '-')}, {meta.get('rows', 0):,} แถว)"
    return d, note, meta


def pic_for_subdept(sub_dept, owner_map):
    """Looks up the assigned Person In Charge (PIC) for a Sub Dept code."""
    if pd.isna(sub_dept):
        return "ไม่ระบุคนดูแล"
    return owner_map.get(str(sub_dept), "ไม่ระบุคนดูแล")


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Backlog Items")
    return buf.getvalue()


def build_cutoff_preview(cdf: pd.DataFrame) -> pd.DataFrame:
    """Builds an operations-friendly preview table for a Stock Cut-off
    dataframe: Store, SDEPT, SDEPT Name, From-To, Duration, and status."""
    if cdf.empty:
        return cdf
    today = pd.Timestamp(date.today())
    out = cdf.copy()
    out["Store"] = out["StoreCode"].astype(str) + " - " + out["StoreName"].astype(str)
    out["SDEPT"] = out["SubDept"].astype(str)
    out["SDEPT Name"] = out["SubDeptNote"].fillna("")
    out["From - To"] = out["CutFrom"].dt.strftime("%d/%m/%Y") + " - " + out["CutTo"].dt.strftime("%d/%m/%Y")
    out["Duration (วัน)"] = (out["CutTo"] - out["CutFrom"]).dt.days + 1

    def _status(row):
        if row["CutFrom"] <= today <= row["CutTo"]:
            days_left = (row["CutTo"] - today).days
            return f"🔴 กำลัง Cutoff (เหลืออีก {days_left} วัน)"
        elif today < row["CutFrom"]:
            days_to_start = (row["CutFrom"] - today).days
            return f"🕒 อีก {days_to_start} วันจะเริ่ม"
        else:
            return "หมดอายุแล้ว"

    out["สถานะ"] = out.apply(_status, axis=1)
    return out[["Store", "SDEPT", "SDEPT Name", "From - To", "Duration (วัน)", "สถานะ"]].sort_values(
        ["Store", "SDEPT"]
    )


def build_subdept_name_lookup(order_df, cutoff_df_local) -> tuple[dict, str]:
    """Sub Dept code -> descriptive name. Prefers a name column carried in
    the order file itself (e.g. 'Sub Dept Name'), since that's the true
    master data; falls back to names harvested from the Stock Cut-off file
    if the order file doesn't have one. Returns (lookup, source_label)."""
    # Match candidate names loosely (case/space/underscore-insensitive) since
    # different Order Status exports have used slightly different spellings
    # for this column (e.g. "SDEPT NAME", "Sub_Dept_Name", "SubDeptName") —
    # an exact-match-only lookup silently returns nothing for those exports,
    # which is why the Settings table would show a Sub Dept code but no name.
    name_col_candidates = ["Sub Dept Name", "SDEPT Name", "Sub_Dept_Name", "Dept Name", "SubDept Name",
                            "SubDeptName", "Sub Department Name", "Sub Dept Desc", "SDEPT Desc"]

    def _norm(s: str) -> str:
        return re.sub(r"[\s_\-]+", "", str(s)).strip().lower()

    if order_df is not None:
        norm_to_actual = {_norm(c): c for c in order_df.columns}
        for cand in name_col_candidates:
            actual_col = norm_to_actual.get(_norm(cand))
            if actual_col:
                lookup = (
                    order_df.dropna(subset=["Sub Dept"])
                    .drop_duplicates("Sub Dept")
                    .set_index("Sub Dept")[actual_col].to_dict()
                )
                return lookup, f"จากไฟล์ Order Status (Master, คอลัมน์ '{actual_col}')"
    if not cutoff_df_local.empty:
        lookup = (
            cutoff_df_local.dropna(subset=["SubDept"])
            .drop_duplicates("SubDept")
            .set_index("SubDept")["SubDeptNote"].to_dict()
        )
        return lookup, "จากไฟล์ Stock Cut-off (สำรอง)"
    return {}, "-"


# ---------------------------------------------------------------------------
# SIDEBAR — BU selection gate + navigation
# ---------------------------------------------------------------------------
st.sidebar.title(f"📦 {APP_TITLE}")

saved_bus = storage.list_saved_bus()
if saved_bus:
    if "selected_bu" not in st.session_state or st.session_state["selected_bu"] not in saved_bus:
        st.session_state["selected_bu"] = saved_bus[0]
    selected_bu = st.sidebar.selectbox("🏢 เลือก BU", saved_bus, key="selected_bu")
    st.sidebar.caption("ข้อมูลทั้งหมดในหน้า Dashboard และ Settings จะอ้างอิงเฉพาะ BU ที่เลือกนี้เท่านั้น")
else:
    selected_bu = None
    st.sidebar.info("ยังไม่มี BU ในระบบ — อัปโหลดไฟล์ Order Status ในหน้า Dashboard ก่อน")

page = st.sidebar.radio("เมนู", ["📊 Dashboard", "⚙️ Settings"])

owner_map = storage.load_owner_map()
cutoff_df = storage.load_cutoff_data()
cutoff_meta = storage.load_cutoff_meta()

bu_df_for_lookup, _, _ = get_bu_order_df(selected_bu) if selected_bu else (None, None, None)
subdept_name_lookup, subdept_name_source = build_subdept_name_lookup(bu_df_for_lookup, cutoff_df)

# ===========================================================================
# SETTINGS PAGE
# ===========================================================================
if page == "⚙️ Settings":
    st.title("⚙️ Settings")
    if selected_bu:
        st.caption(f"กำลังตั้งค่าสำหรับ BU: **{selected_bu}**")

    st.markdown("### 1. จัดการคนดูแลแต่ละ Sub Dept")
    st.caption(f"กำหนดผู้ดูแล (PIC) สำหรับแต่ละ Sub Dept — แสดง Sub Dept **ครบทุกตัว** ที่พบในไฟล์ Order Status ของ BU นี้ "
               f"ไม่ใช่เฉพาะตัวที่เคย Mapping ไว้แล้ว (ชื่อ Sub Dept: {subdept_name_source})")

    if bu_df_for_lookup is not None and "Sub Dept" in bu_df_for_lookup.columns:
        sub_depts = sorted(bu_df_for_lookup["Sub Dept"].dropna().unique().tolist())
    else:
        sub_depts = sorted(owner_map.keys())

    if sub_depts:
        editor_df = pd.DataFrame({
            "Sub Dept": sub_depts,
            "ชื่อ Sub Dept": [subdept_name_lookup.get(sd, "") for sd in sub_depts],
            "คนดูแล": [owner_map.get(sd, "") for sd in sub_depts],
        })
        edited = st.data_editor(
            editor_df,
            column_config={
                "Sub Dept": st.column_config.TextColumn(disabled=True),
                "ชื่อ Sub Dept": st.column_config.TextColumn(disabled=True),
                "คนดูแล": st.column_config.TextColumn(help="ใส่ชื่อผู้ดูแล Sub Dept นี้"),
            },
            hide_index=True,
            width='stretch',
            key="owner_editor",
        )
        if st.button("💾 บันทึกคนดูแล", type="primary"):
            new_map = dict(owner_map)  # keep mappings for Sub Depts from other BUs too
            for _, row in edited.iterrows():
                if str(row["คนดูแล"]).strip():
                    new_map[row["Sub Dept"]] = row["คนดูแล"]
                elif row["Sub Dept"] in new_map:
                    del new_map[row["Sub Dept"]]
            storage.save_owner_map(new_map)
            st.success(f"บันทึกคนดูแลสำหรับ {len(edited)} Sub Dept ของ BU {selected_bu} เรียบร้อยแล้ว")
            st.rerun()
    else:
        st.info("ยังไม่มีข้อมูล Sub Dept — กรุณาอัปโหลดไฟล์ Order Status ในหน้า Dashboard ก่อน")

    st.markdown("---")
    st.markdown("### 2. ไฟล์ Stock Cut-off (POS Stock Cut off)")
    st.caption(
        "อัปโหลดไฟล์ Excel ตารางแจ้งช่วงเวลางดรับสินค้าระหว่างนับสต๊อก — ระบบจะอ่านเฉพาะชีทชื่อ 'CDS' เท่านั้น "
        "ข้อมูลที่อัปโหลดจะถูก **เพิ่มเข้าไปเรื่อยๆ (Append)** ไม่ทับของเดิม และระบบจะเก็บไว้เฉพาะรายการที่ **ยัง Active อยู่** "
        "(วันที่สิ้นสุด ≥ วันนี้) เท่านั้น รายการที่หมดอายุแล้วจะถูกลบออกจากระบบอัตโนมัติทุกครั้งที่มีการอัปโหลดใหม่"
    )

    if cutoff_meta:
        st.success(f"ข้อมูล Active ในระบบตอนนี้: {cutoff_meta.get('rows', 0)} รายการ "
                    f"(อัปเดตล่าสุดเมื่อ {cutoff_meta.get('uploaded_at', '-')} จากไฟล์ {cutoff_meta.get('last_filename', '-')})")

    cutoff_upload = st.file_uploader("อัปโหลดไฟล์ Stock Cut-off (.xlsx)", type=["xlsx"], key="cutoff_uploader")
    if cutoff_upload is not None:
        try:
            parsed = parse_cutoff_excel(cutoff_upload)
            today_ts = pd.Timestamp(date.today())
            n_active = (parsed["CutTo"] >= today_ts).sum()
            n_expired = len(parsed) - n_active
            st.write(f"พบข้อมูล {len(parsed)} แถว จากชีท 'CDS' — Active {n_active} แถว, หมดอายุแล้ว {n_expired} แถว (จะไม่ถูกเพิ่มเข้าระบบ)")
            st.dataframe(build_cutoff_preview(parsed), width='stretch', height=300, hide_index=True)
            if st.button("➕ เพิ่มเข้าระบบ (Append เฉพาะรายการ Active)", type="primary"):
                summary = storage.append_cutoff_data(parsed, cutoff_upload.name)
                st.success(
                    f"เพิ่มข้อมูลแล้ว: เพิ่มใหม่ {summary['rows_added']} แถว, "
                    f"ข้ามเพราะหมดอายุ {summary['rows_skipped_expired']} แถว, "
                    f"ตัดรายการซ้ำออก {summary['duplicates_removed']} แถว, "
                    f"ลบรายการที่หมดอายุออกจากระบบ {summary['expired_pruned']} แถว — "
                    f"รวม Active ทั้งหมดตอนนี้ {summary['total_active_rows']} แถว"
                )
                st.rerun()
        except Exception as e:
            st.error(f"อ่านไฟล์ไม่สำเร็จ กรุณาตรวจสอบว่ามีชีทชื่อ 'CDS' และรูปแบบตรงตามเทมเพลต ({e})")

    if not cutoff_df.empty:
        st.markdown("#### ✏️ แก้ไขรายการ Stock Cut-off ที่ Active อยู่")
        st.caption("แก้ไขค่าในตารางได้โดยตรง (ดับเบิลคลิกที่ช่อง), ลบแถวด้วยการเลือกแถวแล้วกดไอคอนถังขยะ, "
                   "หรือเพิ่มแถวใหม่ด้วยปุ่ม + ที่มุมล่างซ้ายของตาราง แล้วกด 'บันทึกการแก้ไข' เพื่อยืนยัน")

        edit_source = cutoff_df.copy().sort_values(["StoreCode", "SubDept"]).reset_index(drop=True)
        cutoff_edited = st.data_editor(
            edit_source,
            column_config={
                "StoreCode": st.column_config.TextColumn("Store Code"),
                "StoreName": st.column_config.TextColumn("Store Name"),
                "SubDept": st.column_config.TextColumn("SDEPT"),
                "SubDeptNote": st.column_config.TextColumn("SDEPT Name"),
                "CutFrom": st.column_config.DateColumn("Cut From", format="DD/MM/YYYY"),
                "CutTo": st.column_config.DateColumn("Cut To", format="DD/MM/YYYY"),
                "ResumeDate": st.column_config.DateColumn("Resume Date", format="DD/MM/YYYY"),
                "Remark": st.column_config.TextColumn("หมายเหตุ"),
            },
            hide_index=True,
            width='stretch',
            height=350,
            num_rows="dynamic",
            key="cutoff_editor",
        )
        if st.button("💾 บันทึกการแก้ไข Stock Cut-off", type="primary"):
            required_cols = ["StoreCode", "StoreName", "SubDept", "CutFrom", "CutTo"]
            bad_rows = cutoff_edited[required_cols].isna().any(axis=1).sum()
            if bad_rows:
                st.error(f"มี {bad_rows} แถวที่ยังกรอกไม่ครบ (ต้องมี Store Code, Store Name, SDEPT, Cut From, Cut To) กรุณาแก้ไขก่อนบันทึก")
            else:
                result = storage.overwrite_cutoff_data(cutoff_edited)
                msg = f"บันทึกแล้ว — เหลือรายการ Active {result['total_active_rows']} รายการ"
                if result["expired_pruned"]:
                    msg += f" (ตัดรายการที่หมดอายุระหว่างแก้ไขออก {result['expired_pruned']} รายการ)"
                st.success(msg)
                st.rerun()

        with st.expander("👁️ ดูตัวอย่างแบบอ่านง่าย (Store / SDEPT / From-To / สถานะ)"):
            st.dataframe(build_cutoff_preview(cutoff_df), width='stretch', height=350, hide_index=True)

    st.markdown("---")
    st.markdown("### 3. ข้อมูลไฟล์ที่บันทึกไว้ในระบบ")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**📄 Order Status (แยกตาม BU)**")
        _, order_metas = storage.load_all_order_data()
        if order_metas:
            for bu, m in sorted(order_metas.items()):
                marker = " ✅ (กำลังดูอยู่)" if bu == selected_bu else ""
                st.info(f"**{bu}**{marker}: {m.get('filename', '-')}\n\nอัปโหลดเมื่อ {m.get('uploaded_at', '-')} "
                        f"({m.get('rows', 0):,} แถว)")
        else:
            st.warning("ยังไม่มีไฟล์ Order Status ที่บันทึกไว้")
    with c2:
        st.markdown("**📄 Stock Cut-off**")
        if cutoff_meta:
            st.info(f"{cutoff_meta.get('rows', 0)} รายการ Active\n\n"
                    f"อัปเดตล่าสุดเมื่อ {cutoff_meta.get('uploaded_at', '-')}")
        else:
            st.warning("ยังไม่มีไฟล์ Stock Cut-off ที่บันทึกไว้")

    if st.button("🗑️ ล้างข้อมูลที่บันทึกไว้ทั้งหมด"):
        storage.clear_all()
        st.success("ล้างข้อมูลทั้งหมดแล้ว")
        st.rerun()

    st.stop()

# ===========================================================================
# DASHBOARD PAGE
# ===========================================================================
st.title(f"📦 {APP_TITLE}")

# -- File upload (also persists to disk, split per BU) -----------------------
uploaded = st.file_uploader(
    "อัปโหลดไฟล์ Excel Order Status by SKU ใหม่ — รองรับหลาย BU: อัปโหลดไฟล์ของ BU ไหนก็ได้ "
    "ระบบจะบันทึกแยกตาม BU ให้อัตโนมัติโดยไม่ทับข้อมูลของ BU อื่นที่เคยอัปโหลดไว้ แล้วเลือก BU ที่ต้องการดูได้จากแถบด้านซ้าย",
    type=["xlsx"],
    help="รองรับเฉพาะไฟล์รายงาน **031. Order Status by SKU** (.xlsx) ที่ export จากระบบ WMS เท่านั้น "
         "ไฟล์รูปแบบอื่นอาจอ่านคอลัมน์ไม่ครบหรือคำนวณ Final Status ผิดพลาด",
)
if uploaded is not None:
    file_bytes = uploaded.getvalue()
    # Guard against an infinite rerun loop: the file_uploader widget keeps
    # returning the same UploadedFile on every rerun until the user removes
    # it, so only (re)parse + save + rerun once per distinct upload — this
    # is also what makes the Sub Dept Name lookup below use fresh data on
    # this same run instead of lagging one page-load behind.
    file_sig = (uploaded.name, len(file_bytes))
    if st.session_state.get("_last_uploaded_order_sig") != file_sig:
        new_df = parse_order_file(file_bytes)
        just_saved_bus = storage.save_order_data(new_df, uploaded.name)
        st.session_state["_last_uploaded_order_sig"] = file_sig
        st.success(f"บันทึกข้อมูลแล้วสำหรับ BU: {', '.join(just_saved_bus)} ({len(new_df):,} แถว)")
        if just_saved_bus and st.session_state.get("selected_bu") not in just_saved_bus:
            st.session_state["selected_bu"] = just_saved_bus[0]
        st.rerun()

if not selected_bu:
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มดู Dashboard")
    st.stop()

df_raw, source_note, order_meta = get_bu_order_df(selected_bu)
if df_raw is None:
    st.info("ไม่พบข้อมูลสำหรับ BU นี้ กรุณาอัปโหลดไฟล์")
    st.stop()

snapshot_upload_label = f"ข้อมูล ณ ไฟล์ที่อัปโหลดเมื่อ {order_meta.get('uploaded_at', '-')}" if order_meta else ""

st.caption(f"{source_note} | แถวข้อมูลทั้งหมด: {len(df_raw):,}")

with st.expander("ℹ️ วิธีคำนวณ Final Logic Status"):
    st.markdown(FLOW_EXPLAINER_HTML, unsafe_allow_html=True)
    st.markdown("""
- **ยกเลิก** → `Cancelled` (ไม่นับเป็น Backlog)
- **รอเบิก** → คงสถานะ `รอเบิก`
- สถานะอื่นๆ (Allocated, Allocated Partial, Allocated ShortAll, รอยืนยัน, ปิดเอกสาร):
  ตรวจสอบต่อ — `Is Picked = N` → คงสถานะเดิม, `Is Picked = Y` → ตรวจ `Is Scan Load`
  → `N` = `Picked`, `Y` → ตรวจ `Is Shipped` → `N` = `Loaded`, `Y` = `Shipped`
- **ติด Stock Cutoff ได้เฉพาะงานที่สถานะเป็น `Loaded` เท่านั้น** — งานที่ยังไม่ถึง Loaded (เช่น รอเบิก/Picked)
  จะไม่ถูกนับเป็น Stock Cutoff แม้ Store/Sub Dept/วันที่จะตรงกับตาราง Cut-off ก็ตาม
  และระบบเช็คว่า **วันนี้** ยังอยู่ในช่วง Cutoff หรือไม่ (ไม่ใช่แค่ดูวันที่ตั้งใจส่ง) — พ้นช่วงแล้วจะกลับมานับเป็น
  `Loaded` ปกติทันทีโดยไม่ต้องรออัปเดตข้อมูลใน Settings
- **Backlog = งานที่ไม่ใช่ Cancelled / Stock Cutoff / Shipped และ Start Ship Date < วันนี้ (เกินกำหนดส่งแล้ว)**
- **Ontime = งานกลุ่มเดียวกัน แต่ Start Ship Date ≥ วันนี้ หรือไม่มีวันที่ระบุ (ยังไม่เกินกำหนด)**
- ตัวเลขในทุกส่วนของ Dashboard (Part 1–4) อ้างอิงจาก **Backlog (เกินกำหนด)** เท่านั้น ไม่รวม Ontime

**ชื่อสถานะที่แสดงในหน้าจอ** (ข้อมูลต้นฉบับ/ไฟล์ Excel ยังใช้ชื่อเดิมเหมือนเดิม เปลี่ยนแค่การแสดงผล):
| สถานะเดิม (Final Logic Status) | แสดงเป็น |
|---|---|
| รอเบิก | Not Printed |
| Allocated, Allocated Partial, Allocated ShortAll | Released |
| Cancelled (จาก ยกเลิก), ปิดเอกสาร | Cancelled |
| Picked / Loaded / Shipped / รอยืนยัน | เหมือนเดิม |

⚠️ **หมายเหตุ:** เนื่องจาก `ปิดเอกสาร` ตอนนี้แสดงชื่อรวมเป็น "Cancelled" เหมือนกับ ยกเลิก
ระบบเลย**ปรับ Backlog logic ให้ตัด ปิดเอกสาร ออกจาก Backlog ด้วย** (เหมือน ยกเลิก) เพื่อไม่ให้ตัวเลขขัดกับป้ายชื่อที่เห็น
ถ้าไม่ต้องการให้ ปิดเอกสาร ถูกตัดออกจาก Backlog (แค่อยากเปลี่ยนชื่อโชว์แต่นับเหมือนเดิม) แจ้งได้เลยครับ จะแก้กลับให้
    """)

# -- Global "Download Snapshot" button — captures Part 1 + 2 + 3 + 4 as one
# single image (everything inside dashboard_container below). -------------
dashboard_snapshot_btn = st.columns([1, 5])[0]
with dashboard_snapshot_btn:
    st.iframe(height=50, src="""
    <button id="dash-snapshot-btn" style="
        background:#6C5CE7;color:#fff;border:none;border-radius:8px;
        padding:8px 14px;font-weight:700;font-size:13px;cursor:pointer;width:100%;">
        📷 Download Snapshot
    </button>
    <script>
    function loadHtml2Canvas() {
      const doc = window.parent.document;
      if (doc.defaultView.html2canvas) return Promise.resolve(doc.defaultView.html2canvas);
      return new Promise((resolve, reject) => {
        const s = doc.createElement('script');
        s.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js';
        s.onload = () => resolve(doc.defaultView.html2canvas);
        s.onerror = reject;
        doc.head.appendChild(s);
      });
    }
    const btn = document.getElementById('dash-snapshot-btn');
    btn.addEventListener('click', async () => {
      const originalText = btn.innerText;
      btn.innerText = '⏳ กำลังสร้างรูปภาพ...';
      btn.disabled = true;
      try {
        const html2canvas = await loadHtml2Canvas();
        const doc = window.parent.document;
        const marker = doc.getElementById('dashboard-snapshot-marker');
        if (!marker) { alert('ไม่พบเนื้อหา Dashboard ในหน้า กรุณาลองรีเฟรชหน้าเว็บ'); return; }
        const container = marker.closest('[data-testid="stVerticalBlock"]');
        const target = container || marker.parentElement;
        const canvas = await html2canvas(target, {backgroundColor: '#ffffff', scale: 2, useCORS: true});
        const link = doc.createElement('a');
        link.download = 'backlog_dashboard_snapshot.png';
        link.href = canvas.toDataURL('image/png');
        doc.body.appendChild(link);
        link.click();
        link.remove();
      } catch (e) {
        alert('เกิดข้อผิดพลาดในการสร้างรูปภาพ: ' + e);
      } finally {
        btn.innerText = originalText;
        btn.disabled = false;
      }
    });
    </script>
    """)

dashboard_container = st.container()

with dashboard_container:
    # Invisible marker — lets the "Download Snapshot" button (above) find the
    # single DOM node that wraps Part 1 through Part 4, so it can screenshot
    # exactly that region (and nothing from Settings/sidebar) as one image.
    st.markdown('<div id="dashboard-snapshot-marker"></div>', unsafe_allow_html=True)
    if snapshot_upload_label:
        st.caption(f"📅 {snapshot_upload_label} — ข้อความนี้จะติดไปกับรูป Snapshot ด้วย")

    # -----------------------------------------------------------------------
    # PART 1: FILTER & KPI CARDS
    # -----------------------------------------------------------------------
    st.markdown('<div class="section-title">Part 1 · Filter & KPI Cards</div>', unsafe_allow_html=True)

    f1, f2, f3, f4, f5 = st.columns(5)
    final_status_options = sorted(df_raw["Final Status"].apply(display_status).dropna().unique().tolist())
    brand_options = sorted(df_raw["Brand"].dropna().unique().tolist())
    store_options = sorted(df_raw["Store Name"].dropna().unique().tolist())
    type_options = sorted(df_raw["Order Type"].dropna().unique().tolist())
    pic_options = sorted(set(owner_map.values())) if owner_map else []

    sel_status = f1.multiselect("Status", final_status_options, default=[])
    sel_brand = f2.multiselect("Brand", brand_options, default=[])
    sel_store = f3.multiselect("To Store", store_options, default=[])
    sel_type = f4.multiselect("Order Type", type_options, default=[])
    sel_pic = f5.multiselect("คนดูแล (PIC)", pic_options, default=[]) if pic_options else []

    f6, f7 = st.columns(2)
    min_create, max_create = df_raw["Order Create Date"].min(), df_raw["Order Create Date"].max()
    min_ship, max_ship = df_raw["Start Ship Date"].min(), df_raw["Start Ship Date"].max()

    create_range = f6.date_input("Order Create Date", value=(),
                                  min_value=min_create.date() if pd.notna(min_create) else None,
                                  max_value=max_create.date() if pd.notna(max_create) else None)
    ship_range = f7.date_input("Start Ship Date", value=(),
                                min_value=min_ship.date() if pd.notna(min_ship) else None,
                                max_value=max_ship.date() if pd.notna(max_ship) else None)

    df = df_raw.copy()
    df["PIC"] = df["Sub Dept"].apply(lambda sd: pic_for_subdept(sd, owner_map))
    df["SDEPT Name"] = df["Sub Dept"].map(subdept_name_lookup).fillna("")
    df["Status Display"] = df["Final Status"].apply(display_status)

    if sel_status:
        df = df[df["Status Display"].isin(sel_status)]
    if sel_brand:
        df = df[df["Brand"].isin(sel_brand)]
    if sel_store:
        df = df[df["Store Name"].isin(sel_store)]
    if sel_type:
        df = df[df["Order Type"].isin(sel_type)]
    if sel_pic:
        df = df[df["PIC"].isin(sel_pic)]
    if isinstance(create_range, tuple) and len(create_range) == 2:
        df = df[(df["Order Create Date"].dt.date >= create_range[0]) & (df["Order Create Date"].dt.date <= create_range[1])]
    if isinstance(ship_range, tuple) and len(ship_range) == 2:
        df = df[(df["Start Ship Date"].dt.date >= ship_range[0]) & (df["Start Ship Date"].dt.date <= ship_range[1])]

    # -- Stock cut-off flag ------------------------------------------------------
    # Only a job that has reached "Loaded" can be stuck on a stock cutoff (it's
    # ready to ship but the store is mid stock-count and can't receive it).
    # Jobs at earlier stages (รอเบิก / Allocated / Picked) are never tagged as
    # Stock Cutoff even if their Store+Sub Dept+date matches a cutoff window —
    # whatever else is blocking them is a separate issue.
    #
    # The match also requires the cutoff window to still be active *today*
    # (CutFrom <= today <= CutTo), not just that the Start Ship Date falls in
    # the window. That way, once the cutoff period ends, the row automatically
    # reverts to being counted as a normal "Loaded" backlog item on the very
    # next page load — no manual cleanup in Settings required.
    today_ts = pd.Timestamp(date.today())
    cutoff_df_active_today = (
        cutoff_df[(cutoff_df["CutFrom"] <= today_ts) & (today_ts <= cutoff_df["CutTo"])]
        if not cutoff_df.empty else cutoff_df
    )


    def compute_cutoff_flag(row, cdf):
        if row.get("Final Status") != "Loaded":
            return False
        if cdf.empty or pd.isna(row.get("To Store Str")) or pd.isna(row.get("Sub Dept")):
            return False
        matches = cdf[(cdf["StoreCode"] == str(row["To Store Str"])) & (cdf["SubDept"] == str(row["Sub Dept"]))]
        return not matches.empty


    if not cutoff_df_active_today.empty:
        df["ติด Stock Cutoff"] = df.apply(lambda r: compute_cutoff_flag(r, cutoff_df_active_today), axis=1)
    else:
        df["ติด Stock Cutoff"] = False


    # -- Bucket: Cancelled / Shipped / Stock Cutoff (excluded) / Backlog / Ontime -
    # "ปิดเอกสาร" (closed document, never picked) is treated as functionally the
    # same as Cancelled — it's now also shown merged under the "Cancelled" label,
    # so it must be excluded from Backlog the same way, or the numbers wouldn't
    # line up with what's displayed.
    #
    # "Backlog" vs "Ontime": any job that isn't Cancelled/Shipped/Stock-Cutoff is
    # split by whether its Start Ship Date has already passed "today":
    #   - Start Ship Date < today  -> Backlog (เกินกำหนดส่งแล้วแต่ยังไม่ Shipped)
    #   - Start Ship Date >= today, or no date at all -> Ontime (ยังอยู่ในเวลา)
    today_date = date.today()


    def compute_bucket(row):
        if row["Final Status"] in ("Cancelled", "ปิดเอกสาร"):
            return "Cancelled"
        if row["Final Status"] == "Shipped":
            return "Shipped"
        if row["ติด Stock Cutoff"]:
            return "Stock Cutoff"
        ship_date = row.get("Start Ship Date")
        if pd.notna(ship_date) and ship_date.date() < today_date:
            return "Backlog"
        return "Ontime"


    df["Bucket"] = df.apply(compute_bucket, axis=1)

    df_backlog = df[df["Bucket"] == "Backlog"]
    df_ontime = df[df["Bucket"] == "Ontime"]
    df_cancelled = df[df["Bucket"] == "Cancelled"]
    df_cutoff_excluded = df[df["Bucket"] == "Stock Cutoff"]
    df_shipped = df[df["Bucket"] == "Shipped"]

    total_orders = df["Order No."].nunique()
    total_pcs = df["Required QTY"].sum()


    def orders_pcs(sub_df):
        return sub_df["Order No."].nunique(), sub_df["Required QTY"].sum()


    backlog_orders, backlog_pcs = orders_pcs(df_backlog)
    ontime_orders, ontime_pcs = orders_pcs(df_ontime)
    cancelled_orders, cancelled_pcs = orders_pcs(df_cancelled)
    cutoff_orders, cutoff_pcs = orders_pcs(df_cutoff_excluded)
    shipped_orders, shipped_pcs = orders_pcs(df_shipped)

    waiting_df = df_backlog[df_backlog["Final Status"] == WAITING_LABEL]
    allocated_df = df_backlog[df_backlog["Final Status"].isin(ALLOCATED_LABELS)]
    picked_df = df_backlog[df_backlog["Final Status"] == "Picked"]
    loaded_df = df_backlog[df_backlog["Final Status"] == "Loaded"]

    waiting_o, waiting_p = orders_pcs(waiting_df)
    allocated_o, allocated_p = orders_pcs(allocated_df)
    picked_o, picked_p = orders_pcs(picked_df)
    loaded_o, loaded_p = orders_pcs(loaded_df)

    # -- Headline row: Total Orders / Backlog / Ontime / รอเบิก ------------------
    r1c1, r1c2, r1c3, r1c4 = st.columns(4)
    with r1c1:
        dual_kpi_card("Total Orders / Jobs", total_orders, 100.0, total_pcs, 100.0, CARD_COLORS[0], headline=True)
    with r1c2:
        dual_kpi_card("Backlog (เกินกำหนด)", backlog_orders, pct(backlog_orders, total_orders), backlog_pcs,
                      pct(backlog_pcs, total_pcs), CARD_COLORS[1], headline=True)
    with r1c3:
        dual_kpi_card("Ontime (ยังไม่เกินกำหนด)", ontime_orders, pct(ontime_orders, total_orders), ontime_pcs,
                      pct(ontime_pcs, total_pcs), "#0EA5E9", headline=True)
    with r1c4:
        dual_kpi_card("Not Printed", waiting_o, pct(waiting_o, total_orders), waiting_p, pct(waiting_p, total_pcs), CARD_COLORS[2])

    # -- Stage row: Released / Picked / Loaded / Shipped ------------------------
    r2c1, r2c2, r2c3, r2c4 = st.columns(4)
    with r2c1:
        dual_kpi_card("Released", allocated_o, pct(allocated_o, total_orders), allocated_p,
                      pct(allocated_p, total_pcs), CARD_COLORS[3])
    with r2c2:
        dual_kpi_card("Picked", picked_o, pct(picked_o, total_orders), picked_p, pct(picked_p, total_pcs), CARD_COLORS[4])
    with r2c3:
        dual_kpi_card("Loaded", loaded_o, pct(loaded_o, total_orders), loaded_p, pct(loaded_p, total_pcs), CARD_COLORS[5])
    with r2c4:
        dual_kpi_card("Shipped", shipped_orders, pct(shipped_orders, total_orders), shipped_pcs,
                      pct(shipped_pcs, total_pcs), CARD_COLORS[6])

    # -- Excluded row: Cancelled / Stock Cutoff ----------------------------------
    st.markdown('<div class="subsection-title">ไม่รวมใน Backlog (Excluded)</div>', unsafe_allow_html=True)
    r3c1, r3c2 = st.columns(2)
    with r3c1:
        dual_kpi_card("Cancelled", cancelled_orders, pct(cancelled_orders, total_orders), cancelled_pcs,
                      pct(cancelled_pcs, total_pcs), "#9CA3AF", muted=True)
    with r3c2:
        dual_kpi_card("Stock Cutoff", cutoff_orders, pct(cutoff_orders, total_orders), cutoff_pcs,
                      pct(cutoff_pcs, total_pcs), "#9CA3AF", muted=True)

    exp_c1, exp_c2, exp_c3 = st.columns(3)
    with exp_c1:
        with st.expander(f"👁️ ดูรายการ Cancelled ({cancelled_orders:,} orders)"):
            cols_show = [c for c in ["Order No.", "Store Name", "Sub Dept", "SDEPT Name", "Status", "Required QTY"]
                         if c in df_cancelled.columns]
            st.dataframe(df_cancelled[cols_show], width='stretch', height=250, hide_index=True)
    with exp_c2:
        with st.expander(f"👁️ ดูรายการ Stock Cutoff ({cutoff_orders:,} orders)"):
            cols_show = [c for c in ["Order No.", "Store Name", "Sub Dept", "SDEPT Name", "Status", "Start Ship Date",
                                      "Required QTY"] if c in df_cutoff_excluded.columns]
            st.dataframe(df_cutoff_excluded[cols_show], width='stretch', height=250, hide_index=True)
    with exp_c3:
        with st.expander(f"👁️ ดูรายการ Ontime ({ontime_orders:,} orders)"):
            cols_show = [c for c in ["Order No.", "Store Name", "Sub Dept", "SDEPT Name", "Status", "Start Ship Date",
                                      "Required QTY"] if c in df_ontime.columns]
            st.dataframe(df_ontime[cols_show], width='stretch', height=250, hide_index=True)

    st.markdown("---")

    # ---------------------------------------------------------------------------
    # PART 2: BACKLOG ตามคนดูแล (PIC) — by Start Ship Date
    # ---------------------------------------------------------------------------
    st.markdown('<div class="section-title">Part 2 · Backlog ตามคนดูแล</div>', unsafe_allow_html=True)
    st.caption("นับเฉพาะรายการที่อยู่ใน Backlog **ที่เกินกำหนดส่งแล้ว** (ไม่รวม Cancelled / Stock Cutoff / Shipped / Ontime) "
               "จัดกลุ่มตามวันที่ **Start Ship Date**")

    part2_container = st.container()

    with part2_container:
        metric_choice = st.radio("แสดงค่าด้วย", ["Required QTY (ผลรวม)", "จำนวน Order (ไม่ซ้ำ)"],
                                  horizontal=True, key="part2_metric")
        value_col = "Required QTY" if metric_choice.startswith("Required") else "Order No."
        agg_func = "sum" if value_col == "Required QTY" else "nunique"

        # All groups (Over All + every PIC) share this SAME set of date
        # columns — derived once from the whole Backlog, not per-group — so
        # every table lines up column-for-column and can live inside one
        # combined <table> with a single, synced horizontal scrollbar.
        backlog_dated = df_backlog.dropna(subset=["Start Ship Date"])
        all_date_cols = sorted(backlog_dated["Start Ship Date"].dt.date.unique()) if not backlog_dated.empty else []

        def fmt(v):
            return "-" if v == 0 else f"{v:,.0f}"

        def build_group_pivot(sub_df: pd.DataFrame):
            """Returns a DataFrame indexed by Status Display + a trailing
            'Total' row, columns = all_date_cols + 'Total', reindexed onto
            the shared date_cols so it aligns with every other group. None
            if this group has no dated rows at all."""
            dated = sub_df.dropna(subset=["Start Ship Date"]).copy()
            if dated.empty:
                return None
            dated["Date"] = dated["Start Ship Date"].dt.date
            pivot = pd.pivot_table(dated, index="Status Display", columns="Date", values=value_col,
                                    aggfunc=agg_func, fill_value=0)
            pivot = pivot.reindex(columns=all_date_cols, fill_value=0)
            status_order = [s for s in STATUS_DISPLAY_ORDER if s in pivot.index] + \
                           [s for s in pivot.index if s not in STATUS_DISPLAY_ORDER]
            pivot = pivot.reindex(status_order, fill_value=0)
            pivot["Total"] = pivot.sum(axis=1)
            total_row = pivot.sum(axis=0)
            total_row.name = "Total"
            return pd.concat([pivot, total_row.to_frame().T])

        def group_rows_html(sub_df: pd.DataFrame, title: str, subtitle: str = "") -> str:
            n_cols = len(all_date_cols) + 2  # Status + date cols + Total
            header = f'<tr class="pivot-group-header"><td colspan="{n_cols}">'
            header += f'<span style="color:#DC2626;font-weight:800;font-size:14px;">{title}</span>'
            if subtitle:
                header += f'<span style="color:#111827;font-weight:700;font-size:13px;margin-left:8px;">{subtitle}</span>'
            header += '</td></tr>'

            if sub_df.empty:
                return header + f'<tr class="pivot-empty"><td colspan="{n_cols}">ไม่มีข้อมูลในกลุ่มนี้</td></tr>'

            pivot = build_group_pivot(sub_df)
            if pivot is None:
                return header + f'<tr class="pivot-empty"><td colspan="{n_cols}">ไม่มีข้อมูลวันที่ในกลุ่มนี้</td></tr>'

            body_rows = []
            for status_name, row in pivot.iterrows():
                cls = "pivot-total" if status_name == "Total" else ""
                cells = f'<td>{status_name}</td>'
                for d in all_date_cols:
                    cells += f'<td>{fmt(row[d])}</td>'
                cells += f'<td>{fmt(row["Total"])}</td>'
                body_rows.append(f'<tr class="{cls}">{cells}</tr>')

            return header + "".join(body_rows)

        def render_pivot_table(groups: list) -> None:
            """groups: list of (sub_df, title, subtitle) tuples, all sharing
            all_date_cols, rendered as ONE <table> so scrolling any part of
            it scrolls every group at once."""
            date_headers = "".join(f'<th>{date_col_label(d)}</th>' for d in all_date_cols)
            html = ['<div class="pivot-wrap"><table class="pivot-table"><thead><tr>',
                    '<th>Status</th>', date_headers, '<th>Total</th>',
                    '</tr></thead><tbody>']
            for sub_df, title, subtitle in groups:
                html.append(group_rows_html(sub_df, title, subtitle))
            html.append('</tbody></table></div>')
            st.markdown("".join(html), unsafe_allow_html=True)

        if not all_date_cols:
            st.caption("ไม่มีข้อมูลวันที่ Start Ship Date ในข้อมูล Backlog นี้เลย")
        else:
            groups = [(df_backlog, f"{selected_bu} Over All", "")]

            if owner_map:
                # Build PIC -> sorted list of Sub Dept codes assigned to them,
                # so each PIC's section title can show which Sub Depts they're
                # responsible for — including each Sub Dept's descriptive
                # name, not just its raw code, so the header is readable.
                pic_to_subdepts = {}
                for sd, ow in owner_map.items():
                    pic_to_subdepts.setdefault(ow, []).append(sd)
                for ow in pic_to_subdepts:
                    pic_to_subdepts[ow] = sorted(pic_to_subdepts[ow], key=lambda x: (len(x), x))

                def _sd_label(sd: str) -> str:
                    name = subdept_name_lookup.get(sd, "")
                    return f"{sd} - {name}" if name else sd

                pics_present = [o for o in sorted(df_backlog["PIC"].unique()) if o != "ไม่ระบุคนดูแล"]
                for pic in pics_present:
                    sub_depts_str = ", ".join(_sd_label(sd) for sd in pic_to_subdepts.get(pic, []))
                    subtitle = f"(SDEPT: {sub_depts_str})" if sub_depts_str else ""
                    groups.append((df_backlog[df_backlog["PIC"] == pic], f"{selected_bu} PICK — {pic}", subtitle))

                render_pivot_table(groups)

                if "ไม่ระบุคนดูแล" in df_backlog["PIC"].unique():
                    with st.expander("Sub Dept ที่ยังไม่ได้กำหนดคนดูแล"):
                        render_pivot_table([(df_backlog[df_backlog["PIC"] == "ไม่ระบุคนดูแล"],
                                              "ไม่ระบุคนดูแล", "")])
            else:
                render_pivot_table(groups)
                st.info("ยังไม่ได้ตั้งค่าคนดูแล Sub Dept — ไปที่เมนู ⚙️ Settings เพื่อกำหนดคนดูแลแต่ละ Sub Dept")

    st.markdown("---")

    # ---------------------------------------------------------------------------
    # PART 3: BACKLOG ITEMS DETAILS
    # ---------------------------------------------------------------------------
    st.markdown('<div class="section-title">Part 3 · Backlog Items Details</div>', unsafe_allow_html=True)
    st.caption("แสดงเฉพาะรายการที่อยู่ใน Backlog **ที่เกินกำหนดส่งแล้ว** (ไม่รวม Cancelled / Stock Cutoff / Shipped / Ontime)")

    p3_status_opts = sorted(df_backlog["Status Display"].dropna().unique().tolist())
    p3_store_opts = sorted(df_backlog["Store Name"].dropna().unique().tolist())
    p3_brand_opts = sorted(df_backlog["Brand"].dropna().unique().tolist())
    p3_sdept_opts = sorted(df_backlog["Sub Dept"].dropna().unique().tolist())
    p3_pic_opts = sorted(df_backlog["PIC"].dropna().unique().tolist())

    p3f1, p3f2, p3f3, p3f4, p3f5 = st.columns(5)
    p3_sel_status = p3f1.multiselect("Status", p3_status_opts, default=[], key="p3_status")
    p3_sel_store = p3f2.multiselect("Store", p3_store_opts, default=[], key="p3_store")
    p3_sel_brand = p3f3.multiselect("Brand", p3_brand_opts, default=[], key="p3_brand")
    p3_sel_sdept = p3f4.multiselect("SDEPT", p3_sdept_opts, default=[], key="p3_sdept")
    p3_sel_pic = p3f5.multiselect("คนดูแล (PIC)", p3_pic_opts, default=[], key="p3_pic")

    search_term = st.text_input("🔍 Search Order No, Store, IBC...", value="",
                                 placeholder="Search Order No, Store, IBC...")

    detail_cols = ["Order No.", "Order Create Date", "Start Ship Date", "Brand", "BU", "To Store", "Store Name",
                    "Status Display", "Sub Dept", "SDEPT Name", "IBC", "PIC",
                    "Required QTY", "Allocated QTY", "Pick QTY"]
    detail_cols = [c for c in detail_cols if c in df_backlog.columns]

    p3_df = df_backlog.copy()
    if p3_sel_status:
        p3_df = p3_df[p3_df["Status Display"].isin(p3_sel_status)]
    if p3_sel_store:
        p3_df = p3_df[p3_df["Store Name"].isin(p3_sel_store)]
    if p3_sel_brand:
        p3_df = p3_df[p3_df["Brand"].isin(p3_sel_brand)]
    if p3_sel_sdept:
        p3_df = p3_df[p3_df["Sub Dept"].isin(p3_sel_sdept)]
    if p3_sel_pic:
        p3_df = p3_df[p3_df["PIC"].isin(p3_sel_pic)]

    details = p3_df[detail_cols].copy()
    details = details.rename(columns={"Sub Dept": "SDEPT", "PIC": "คนดูแล (PIC)", "Status Display": "Status"})

    if search_term:
        mask = pd.Series(False, index=details.index)
        for c in ["Order No.", "Store Name", "IBC"]:
            if c in details.columns:
                mask |= details[c].astype(str).str.contains(search_term, case=False, na=False)
        details = details[mask]

    details_display = details.copy()
    if "Order Create Date" in details_display.columns:
        details_display["Order Create Date"] = details_display["Order Create Date"].dt.strftime("%d/%m/%Y")
    if "Start Ship Date" in details_display.columns:
        details_display["Start Ship Date"] = details_display["Start Ship Date"].dt.strftime("%d/%m/%Y")

    st.caption(f"แสดง {len(details_display):,} แถว")
    st.dataframe(details_display, width='stretch', height=420, hide_index=True)

    excel_bytes = to_excel_bytes(details_display)
    st.download_button("⬇️ Export Table Details (Excel)", data=excel_bytes,
                        file_name="backlog_items_details.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown("---")

    # ---------------------------------------------------------------------------
    # PART 4: DASHBOARD (CHARTS)
    # ---------------------------------------------------------------------------
    st.markdown('<div class="section-title">Part 4 · Dashboard</div>', unsafe_allow_html=True)
    st.caption("ข้อมูลในทุกกราฟด้านล่างนับเฉพาะ Backlog **ที่เกินกำหนดส่งแล้ว** เช่นกัน")

    col1, col2 = st.columns(2)

    status_order_counts = df_backlog.groupby("Status Display")["Order No."].nunique().reset_index(name="Unique Orders")
    status_qty_counts = df_backlog.groupby("Status Display")["Required QTY"].sum().reset_index(name="Required QTY")

    with col1:
        st.markdown("#### Backlog Orders by Status")
        fig = go.Figure(data=[go.Pie(
            labels=status_order_counts["Status Display"], values=status_order_counts["Unique Orders"], hole=0.55,
            marker=dict(colors=[DONUT_COLORS.get(s, "#9CA3AF") for s in status_order_counts["Status Display"]]))])
        fig.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=340)
        st.plotly_chart(fig, width='stretch')

    with col2:
        st.markdown("#### Backlog QTY by Status")
        fig2 = go.Figure(data=[go.Pie(
            labels=status_qty_counts["Status Display"], values=status_qty_counts["Required QTY"], hole=0.55,
            marker=dict(colors=[DONUT_COLORS.get(s, "#9CA3AF") for s in status_qty_counts["Status Display"]]))])
        fig2.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=340)
        st.plotly_chart(fig2, width='stretch')

    st.markdown("#### Backlog by Start Ship Date (Pcs. แยกตามสถานะ)")
    ship_daily_status = (
        df_backlog.dropna(subset=["Start Ship Date"])
        .groupby([df_backlog["Start Ship Date"].dt.date, "Status Display"])["Required QTY"]
        .sum()
        .reset_index()
        .rename(columns={"Start Ship Date": "Date"})
        .sort_values("Date")
    )
    fig4 = go.Figure()
    for status in STATUS_DISPLAY_ORDER:
        s_df = ship_daily_status[ship_daily_status["Status Display"] == status]
        if s_df.empty:
            continue
        fig4.add_trace(go.Bar(x=s_df["Date"], y=s_df["Required QTY"], name=status,
                               marker_color=DONUT_COLORS.get(status, "#9CA3AF")))
    fig4.update_layout(height=420, margin=dict(t=10, b=10, l=10, r=10), barmode="stack",
                        legend=dict(orientation="h", y=1.1), yaxis=dict(title="Pcs."), xaxis=dict(title="Start Ship Date"))
    st.plotly_chart(fig4, width='stretch')
