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
.section-title { font-size: 22px; font-weight: 800; color: #111827; margin: 6px 0 2px 0; }
.group-title { font-size: 18px; font-weight: 800; color: #DC2626; margin: 18px 0 2px 0; }
.group-sub { font-size: 14px; font-weight: 700; color: #111827; margin-left: 6px; }
div.block-container { padding-top: 1.2rem; }
</style>
""", unsafe_allow_html=True)

BACKLOG_STATUSES = ["รอเบิก", "รอยืนยัน"]
ALL_STATUS_ORDER = ["รอเบิก", "รอยืนยัน", "Allocated", "Allocated Partial", "Allocated ShortAll",
                     "ปิดเอกสาร", "ยกเลิก"]
CARD_COLORS = ["#6C5CE7", "#2563EB", "#059669", "#D97706", "#DB2777", "#7C3AED", "#0EA5E9", "#DC2626"]
THAI_WD = ["จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส.", "อา."]

DONUT_COLORS = {
    "รอเบิก": "#2563EB", "Allocated": "#059669", "รอยืนยัน": "#7C3AED",
    "Allocated Partial": "#F59E0B", "Allocated ShortAll": "#D97706",
    "ปิดเอกสาร": "#9CA3AF", "ยกเลิก": "#DC2626",
}


def kpi_card(label, value, sub, color):
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color:{color};">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
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

    return df


def get_active_order_df():
    """Returns (df, source_note) using session upload if present this run,
    otherwise falling back to the persisted latest file on disk."""
    if "order_file_bytes" in st.session_state:
        df = parse_order_file(st.session_state["order_file_bytes"])
        meta = st.session_state.get("order_file_meta", {})
        note = f"ไฟล์ที่อัปโหลด: {meta.get('filename', '-')} (อัปโหลดเมื่อ {meta.get('uploaded_at', '-')})"
        return df, note
    if storage.has_saved_order_file():
        with open(storage.ORDER_FILE, "rb") as f:
            file_bytes = f.read()
        df = parse_order_file(file_bytes)
        meta = storage.load_saved_order_meta() or {}
        note = f"ใช้ไฟล์ล่าสุดที่เคยอัปโหลดไว้: {meta.get('filename', '-')} (อัปโหลดเมื่อ {meta.get('uploaded_at', '-')})"
        return df, note
    return None, None


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


# ---------------------------------------------------------------------------
# SIDEBAR NAVIGATION
# ---------------------------------------------------------------------------
st.sidebar.title(f"📦 {APP_TITLE}")
page = st.sidebar.radio("เมนู", ["📊 Dashboard", "⚙️ Settings"])

owner_map = storage.load_owner_map()
cutoff_df = storage.load_cutoff_data()
cutoff_meta = storage.load_cutoff_meta()

# ===========================================================================
# SETTINGS PAGE
# ===========================================================================
if page == "⚙️ Settings":
    st.title("⚙️ Settings")

    st.markdown("### 1. จัดการคนดูแลแต่ละ Sub Dept")
    st.caption("กำหนดผู้ดูแล (PIC) สำหรับแต่ละ Sub Dept เพื่อใช้แบ่งกลุ่มข้อมูลในหน้า Dashboard (Part 2)")

    df_for_settings, note = get_active_order_df()
    subdept_names = {}
    if not cutoff_df.empty:
        subdept_names = (
            cutoff_df.dropna(subset=["SubDept"])
            .drop_duplicates("SubDept")
            .set_index("SubDept")["SubDeptNote"].to_dict()
        )

    if df_for_settings is not None and "Sub Dept" in df_for_settings.columns:
        sub_depts = sorted(df_for_settings["Sub Dept"].dropna().unique().tolist())
    else:
        sub_depts = sorted(owner_map.keys())

    if sub_depts:
        editor_df = pd.DataFrame({
            "Sub Dept": sub_depts,
            "ชื่อ Sub Dept (จากไฟล์ Stock Cut-off ถ้ามี)": [subdept_names.get(sd, "") for sd in sub_depts],
            "คนดูแล": [owner_map.get(sd, "") for sd in sub_depts],
        })
        edited = st.data_editor(
            editor_df,
            column_config={
                "Sub Dept": st.column_config.TextColumn(disabled=True),
                "ชื่อ Sub Dept (จากไฟล์ Stock Cut-off ถ้ามี)": st.column_config.TextColumn(disabled=True),
                "คนดูแล": st.column_config.TextColumn(help="ใส่ชื่อผู้ดูแล Sub Dept นี้"),
            },
            hide_index=True,
            width='stretch',
            key="owner_editor",
        )
        if st.button("💾 บันทึกคนดูแล", type="primary"):
            new_map = {row["Sub Dept"]: row["คนดูแล"] for _, row in edited.iterrows() if str(row["คนดูแล"]).strip()}
            storage.save_owner_map(new_map)
            st.success(f"บันทึกคนดูแลสำหรับ {len(new_map)} Sub Dept เรียบร้อยแล้ว")
            st.rerun()
    else:
        st.info("ยังไม่มีข้อมูล Sub Dept — กรุณาอัปโหลดไฟล์ Order Status ในหน้า Dashboard ก่อน หรือมีการบันทึกคนดูแลไว้ก่อนหน้านี้แล้ว")

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
        with st.expander(f"📋 ดูรายการ Stock Cut-off ที่ Active อยู่ในระบบทั้งหมด ({len(cutoff_df)} รายการ)"):
            st.dataframe(build_cutoff_preview(cutoff_df), width='stretch', height=400, hide_index=True)

    st.markdown("---")
    st.markdown("### 3. ข้อมูลไฟล์ที่บันทึกไว้ในระบบ")
    order_meta = storage.load_saved_order_meta()
    c1, c2 = st.columns(2)
    with c1:
        if order_meta:
            st.info(f"📄 Order Status ล่าสุด: {order_meta['filename']}\n\nอัปโหลดเมื่อ {order_meta['uploaded_at']}")
        else:
            st.warning("ยังไม่มีไฟล์ Order Status ที่บันทึกไว้")
    with c2:
        if cutoff_meta:
            st.info(f"📄 Stock Cut-off: {cutoff_meta.get('rows', 0)} รายการ Active\n\n"
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

# -- File upload (also persists to disk) ------------------------------------
uploaded = st.file_uploader("อัปโหลดไฟล์ Excel Order Status by SKU ใหม่ (ถ้าไม่อัปโหลด ระบบจะใช้ไฟล์ล่าสุดที่เคยอัปโหลดไว้)",
                             type=["xlsx"])
if uploaded is not None:
    file_bytes = uploaded.getvalue()
    st.session_state["order_file_bytes"] = file_bytes
    st.session_state["order_file_meta"] = {
        "filename": uploaded.name,
        "uploaded_at": storage._now(),
    }
    storage.save_order_file(uploaded)

df_raw, source_note = get_active_order_df()

if df_raw is None:
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มดู Dashboard")
    st.stop()

bu_list = sorted(df_raw["BU"].dropna().unique().tolist()) if "BU" in df_raw.columns else []
bu_note = f" | BU ในไฟล์: {', '.join(bu_list)}" if bu_list else ""
st.caption(f"{source_note} | แถวข้อมูลทั้งหมด: {len(df_raw):,}{bu_note}")

# ---------------------------------------------------------------------------
# PART 1: FILTER & KPI CARDS
# ---------------------------------------------------------------------------
st.markdown('<div class="section-title">Part 1 · Filter & KPI Cards</div>', unsafe_allow_html=True)

f1, f2, f3, f4, f5, f6 = st.columns(6)
status_options = sorted(df_raw["Status"].dropna().unique().tolist())
brand_options = sorted(df_raw["Brand"].dropna().unique().tolist())
store_options = sorted(df_raw["Store Name"].dropna().unique().tolist())
type_options = sorted(df_raw["Order Type"].dropna().unique().tolist())
pic_options = sorted(set(owner_map.values())) if owner_map else []

sel_bu = f1.multiselect("BU", bu_list, default=[]) if bu_list else []
sel_status = f2.multiselect("Status", status_options, default=[])
sel_brand = f3.multiselect("Brand", brand_options, default=[])
sel_store = f4.multiselect("To Store", store_options, default=[])
sel_type = f5.multiselect("Order Type", type_options, default=[])
sel_pic = f6.multiselect("คนดูแล (PIC)", pic_options, default=[]) if pic_options else []

f7, f8 = st.columns(2)
min_create, max_create = df_raw["Order Create Date"].min(), df_raw["Order Create Date"].max()
min_ship, max_ship = df_raw["Start Ship Date"].min(), df_raw["Start Ship Date"].max()

create_range = f7.date_input("Order Create Date", value=(),
                              min_value=min_create.date() if pd.notna(min_create) else None,
                              max_value=max_create.date() if pd.notna(max_create) else None)
ship_range = f8.date_input("Start Ship Date", value=(),
                            min_value=min_ship.date() if pd.notna(min_ship) else None,
                            max_value=max_ship.date() if pd.notna(max_ship) else None)

df = df_raw.copy()
df["PIC"] = df["Sub Dept"].apply(lambda sd: pic_for_subdept(sd, owner_map))

if sel_bu:
    df = df[df["BU"].isin(sel_bu)]
if sel_status:
    df = df[df["Status"].isin(sel_status)]
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

# -- Stock cut-off flag -------------------------------------------------------
def compute_cutoff_flag(row, cdf):
    if cdf.empty or pd.isna(row.get("To Store Str")) or pd.isna(row.get("Sub Dept")):
        return False
    matches = cdf[(cdf["StoreCode"] == str(row["To Store Str"])) & (cdf["SubDept"] == str(row["Sub Dept"]))]
    if matches.empty or pd.isna(row.get("Start Ship Date")):
        return False
    ship_date = row["Start Ship Date"]
    return bool(((matches["CutFrom"] <= ship_date) & (ship_date <= matches["CutTo"])).any())


if not cutoff_df.empty:
    df["ติด Stock Cutoff"] = df.apply(lambda r: compute_cutoff_flag(r, cutoff_df), axis=1)
else:
    df["ติด Stock Cutoff"] = False

df_backlog = df[df["Status"].isin(BACKLOG_STATUSES)]
df_shipped = df_backlog[df_backlog["Is Shipped"] == "Y"]
df_picked_rows = df_backlog[df_backlog["Pick QTY"] > 0]

total_backlog_orders = df_backlog["Order No."].nunique()
backlog_required_qty = df_backlog["Required QTY"].sum()
backlog_items_ibc = df_backlog["IBC"].nunique() if "IBC" in df_backlog.columns else 0
allocated_qty = df_backlog["Allocated QTY"].sum()
picked_qty = df_backlog["Pick QTY"].sum()
shipped_qty = df_shipped["Pick QTY"].sum()
manpower_picked = df_picked_rows["User Pick"].nunique()
productivity = (picked_qty / manpower_picked) if manpower_picked else 0
cutoff_affected = df_backlog["ติด Stock Cutoff"].sum()

pct_allocated = (allocated_qty / backlog_required_qty * 100) if backlog_required_qty else 0
pct_picked = (picked_qty / backlog_required_qty * 100) if backlog_required_qty else 0
pct_shipped = (shipped_qty / backlog_required_qty * 100) if backlog_required_qty else 0

cols = st.columns(8)
with cols[0]:
    kpi_card("Total Backlog Orders", f"{total_backlog_orders:,}", "Unique pending orders (รอเบิก+รอยืนยัน)", CARD_COLORS[0])
with cols[1]:
    kpi_card("Backlog Required QTY", f"{backlog_required_qty:,.0f}", f"{backlog_items_ibc:,} items (IBCs)", CARD_COLORS[1])
with cols[2]:
    kpi_card("Allocated QTY", f"{allocated_qty:,.0f}", f"{pct_allocated:.1f}% of required QTY", CARD_COLORS[2])
with cols[3]:
    kpi_card("Picked QTY", f"{picked_qty:,.0f}", f"{pct_picked:.1f}% of required QTY", CARD_COLORS[3])
with cols[4]:
    kpi_card("Shipped QTY", f"{shipped_qty:,.0f}", f"{pct_shipped:.1f}% of required QTY", CARD_COLORS[4])
with cols[5]:
    kpi_card("Manpower Picked", f"{manpower_picked:,}", "Unique pick users", CARD_COLORS[5])
with cols[6]:
    kpi_card("Productivity", f"{productivity:,.1f}", "Pick QTY / Manpower", CARD_COLORS[6])
with cols[7]:
    kpi_card("ติด Stock Cutoff", f"{cutoff_affected:,}", "แถว backlog ที่ตรงช่วงงดรับสินค้า", CARD_COLORS[7])

st.markdown("---")

# ---------------------------------------------------------------------------
# PART 2: BACKLOG ตามคนดูแล (PIC)
# ---------------------------------------------------------------------------
st.markdown('<div class="section-title">Part 2 · Backlog ตามคนดูแล</div>', unsafe_allow_html=True)

metric_choice = st.radio("แสดงค่าด้วย", ["Required QTY (ผลรวม)", "จำนวน Order (ไม่ซ้ำ)"],
                          horizontal=True, key="part2_metric")
value_col = "Required QTY" if metric_choice.startswith("Required") else "Order No."
agg_func = "sum" if value_col == "Required QTY" else "nunique"

st.caption("หมายเหตุ: ไฟล์ Order Status ที่อัปโหลดมีสถานะจริงคือ "
           "รอเบิก / รอยืนยัน / Allocated / Allocated Partial / Allocated ShortAll / ปิดเอกสาร / ยกเลิก "
           "ตารางด้านล่างจึงใช้สถานะเหล่านี้แทนสถานะแพ็ก/โหลดขึ้นรถในตัวอย่าง")


def render_owner_pivot(sub_df: pd.DataFrame, title: str, subtitle: str = ""):
    header_html = f'<div class="group-title">{title}'
    if subtitle:
        header_html += f'<span class="group-sub">{subtitle}</span>'
    header_html += '</div>'
    st.markdown(header_html, unsafe_allow_html=True)

    if sub_df.empty:
        st.caption("ไม่มีข้อมูลในกลุ่มนี้")
        return

    dated = sub_df.dropna(subset=["Order Create Date"]).copy()
    if dated.empty:
        st.caption("ไม่มีข้อมูลวันที่ในกลุ่มนี้")
        return
    dated["Date"] = dated["Order Create Date"].dt.date

    pivot = pd.pivot_table(dated, index="Status", columns="Date", values=value_col,
                            aggfunc=agg_func, fill_value=0)

    status_order = [s for s in ALL_STATUS_ORDER if s in pivot.index] + \
                   [s for s in pivot.index if s not in ALL_STATUS_ORDER]
    pivot = pivot.reindex(status_order)

    date_cols = sorted(pivot.columns)
    pivot = pivot[date_cols]
    pivot["Total"] = pivot.sum(axis=1)

    total_row = pivot.sum(axis=0)
    total_row.name = "Total"
    pivot = pd.concat([pivot, total_row.to_frame().T])

    display_pivot = pivot.copy()
    display_pivot.columns = [date_col_label(c) if c != "Total" else "Total" for c in display_pivot.columns]

    def fmt(v):
        return "-" if v == 0 else f"{v:,.0f}"

    styled = display_pivot.style.format(fmt)
    styled = styled.apply(
        lambda row: ["background-color:#fde68a; font-weight:800;" if row.name == "Total" else "" for _ in row],
        axis=1,
    )
    styled = styled.set_properties(subset=display_pivot.columns[-1:], **{"font-weight": "700"})
    st.dataframe(styled, width='stretch')


render_owner_pivot(df, "CDS Over All")

if owner_map:
    # Build PIC -> sorted list of Sub Dept codes assigned to them, so each
    # PIC's section title can show which Sub Depts they're responsible for.
    pic_to_subdepts = {}
    for sd, ow in owner_map.items():
        pic_to_subdepts.setdefault(ow, []).append(sd)
    for ow in pic_to_subdepts:
        pic_to_subdepts[ow] = sorted(pic_to_subdepts[ow], key=lambda x: (len(x), x))

    pics_present = [o for o in sorted(df["PIC"].unique()) if o != "ไม่ระบุคนดูแล"]
    for pic in pics_present:
        sub_depts_str = ", ".join(pic_to_subdepts.get(pic, []))
        subtitle = f"{pic} (SDEPT: {sub_depts_str})" if sub_depts_str else pic
        render_owner_pivot(df[df["PIC"] == pic], "CDS PICK", subtitle)
    if "ไม่ระบุคนดูแล" in df["PIC"].unique():
        with st.expander("Sub Dept ที่ยังไม่ได้กำหนดคนดูแล"):
            render_owner_pivot(df[df["PIC"] == "ไม่ระบุคนดูแล"], "ไม่ระบุคนดูแล")
else:
    st.info("ยังไม่ได้ตั้งค่าคนดูแล Sub Dept — ไปที่เมนู ⚙️ Settings เพื่อกำหนดคนดูแลแต่ละ Sub Dept")

st.markdown("---")

# ---------------------------------------------------------------------------
# PART 3: BACKLOG ITEMS DETAILS
# ---------------------------------------------------------------------------
st.markdown('<div class="section-title">Part 3 · Backlog Items Details</div>', unsafe_allow_html=True)

top_row1, top_row2 = st.columns([3, 2])
search_term = top_row2.text_input("🔍 Search Order No, Store, IBC...", value="",
                                    placeholder="Search Order No, Store, IBC...")

detail_cols = ["Order No.", "Order Create Date", "Start Ship Date", "Brand", "BU", "To Store", "Store Name",
                "Status", "Sub Dept", "IBC", "PIC", "ติด Stock Cutoff", "Required QTY", "Allocated QTY", "Pick QTY"]
detail_cols = [c for c in detail_cols if c in df_backlog.columns]
details = df_backlog[detail_cols].copy()
details = details.rename(columns={"Sub Dept": "SDEPT", "PIC": "คนดูแล (PIC)"})

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

col1, col2 = st.columns(2)

status_order_counts = df.groupby("Status")["Order No."].nunique().reset_index(name="Unique Orders")
status_qty_counts = df.groupby("Status")["Required QTY"].sum().reset_index(name="Required QTY")

with col1:
    st.markdown("#### Backlog Orders by Status")
    fig = go.Figure(data=[go.Pie(
        labels=status_order_counts["Status"], values=status_order_counts["Unique Orders"], hole=0.55,
        marker=dict(colors=[DONUT_COLORS.get(s, "#9CA3AF") for s in status_order_counts["Status"]]))])
    fig.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=340)
    st.plotly_chart(fig, width='stretch')

with col2:
    st.markdown("#### Backlog QTY by Status")
    fig2 = go.Figure(data=[go.Pie(
        labels=status_qty_counts["Status"], values=status_qty_counts["Required QTY"], hole=0.55,
        marker=dict(colors=[DONUT_COLORS.get(s, "#9CA3AF") for s in status_qty_counts["Status"]]))])
    fig2.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=340)
    st.plotly_chart(fig2, width='stretch')

col3, col4 = st.columns(2)

with col3:
    st.markdown("#### Top 10 Stores by Backlog QTY")
    store_agg = (
        df_backlog.groupby(["To Store", "Store Name"])
        .agg(**{"Backlog QTY": ("Required QTY", "sum"), "Unique Orders": ("Order No.", "nunique")})
        .reset_index()
    )
    store_agg["Store Label"] = store_agg["To Store"].astype(str) + " - " + store_agg["Store Name"].astype(str)
    store_agg = store_agg.sort_values("Backlog QTY", ascending=False).head(10).sort_values("Backlog QTY")
    fig3 = go.Figure()
    fig3.add_trace(go.Bar(y=store_agg["Store Label"], x=store_agg["Backlog QTY"], name="Backlog QTY",
                           orientation="h", marker_color="#2563EB"))
    fig3.update_layout(height=380, margin=dict(t=10, b=10, l=10, r=10))
    st.plotly_chart(fig3, width='stretch')

with col4:
    st.markdown("#### Order Backlog (Daily)")
    created = df.groupby(df["Order Create Date"].dt.date)["Order No."].nunique().reset_index(name="Created Orders")
    created.columns = ["Date", "Created Orders"]
    picked_mask = df["Is Picked"] == "Y"
    picked = (
        df[picked_mask].groupby(df.loc[picked_mask, "Picked Date"].dt.date)["Order No."]
        .nunique().reset_index(name="Picked Orders")
    )
    picked.columns = ["Date", "Picked Orders"]
    daily = pd.merge(created, picked, on="Date", how="outer").fillna(0).sort_values("Date")

    fig4 = go.Figure()
    fig4.add_trace(go.Scatter(x=daily["Date"], y=daily["Created Orders"], name="Created Orders",
                               mode="lines+markers", line=dict(color="#2563EB"), fill="tozeroy",
                               fillcolor="rgba(37,99,235,0.1)"))
    fig4.add_trace(go.Scatter(x=daily["Date"], y=daily["Picked Orders"], name="Picked Orders",
                               mode="lines+markers", line=dict(color="#059669")))
    fig4.update_layout(height=380, margin=dict(t=10, b=10, l=10, r=10), legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig4, width='stretch')
