import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime

st.set_page_config(page_title="CDS Backlog Dashboard", layout="wide", initial_sidebar_state="expanded")

# ---------------------------------------------------------------------------
# STYLING
# ---------------------------------------------------------------------------
st.markdown("""
<style>
.kpi-card {
    background: #ffffff;
    border-radius: 10px;
    padding: 16px 18px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    border-top: 4px solid #6C5CE7;
    height: 100%;
}
.kpi-label { font-size: 12px; font-weight: 700; color: #6b7280; letter-spacing: .04em; text-transform: uppercase; }
.kpi-value { font-size: 28px; font-weight: 800; color: #111827; margin: 4px 0 2px 0; }
.kpi-sub   { font-size: 12px; color: #6b7280; }
div.block-container { padding-top: 1.2rem; }
</style>
""", unsafe_allow_html=True)

BACKLOG_STATUSES = ["รอเบิก", "รอยืนยัน"]  # canonical (post-strip) backlog statuses
CARD_COLORS = ["#6C5CE7", "#2563EB", "#059669", "#D97706", "#DB2777", "#7C3AED", "#DC2626"]


def kpi_card(label, value, sub, color):
    st.markdown(f"""
    <div class="kpi-card" style="border-top-color:{color};">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# DATA LOADING
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_data(file) -> pd.DataFrame:
    df = pd.read_excel(file, header=2)
    # drop the leading blank/unnamed helper column if present
    drop_cols = [c for c in df.columns if str(c).startswith("Unnamed: 0")]
    df = df.drop(columns=drop_cols, errors="ignore")

    # strip whitespace from all string/object columns
    str_cols = df.select_dtypes(include=["object", "string"]).columns
    for c in str_cols:
        df[c] = df[c].astype(str).str.strip().replace({"nan": pd.NA, "None": pd.NA})

    # numeric coercion
    for c in ["Required QTY", "Allocated QTY", "Pick QTY"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # dates
    for c in ["Order Create Date", "Allocated Date", "Picked Date", "Shipped Date"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    if "Start Ship Date" in df.columns:
        df["Start Ship Date"] = pd.to_datetime(df["Start Ship Date"], format="%d/%m/%Y", errors="coerce")

    df["Order Create Date Only"] = df["Order Create Date"].dt.date
    return df


st.title("📦 CDS Backlog Order Dashboard")

uploaded = st.file_uploader("อัปโหลดไฟล์ Excel (031_Order_Status_by_SKU...xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มดู Dashboard")
    st.stop()

df_raw = load_data(uploaded)

st.caption(f"CDS: Updated {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}  |  แถวข้อมูลทั้งหมด: {len(df_raw):,}")

# ---------------------------------------------------------------------------
# FILTERS
# ---------------------------------------------------------------------------
st.markdown("### ตัวกรอง")
f1, f2, f3, f4, f5, f6 = st.columns(6)

status_options = sorted(df_raw["Status"].dropna().unique().tolist())
brand_options = sorted(df_raw["Brand"].dropna().unique().tolist())
store_options = sorted(df_raw["Store Name"].dropna().unique().tolist())
type_options = sorted(df_raw["Order Type"].dropna().unique().tolist())

sel_status = f1.multiselect("Status", status_options, default=[])
sel_brand = f2.multiselect("Brand", brand_options, default=[])
sel_store = f3.multiselect("To Store", store_options, default=[])
sel_type = f4.multiselect("Order Type", type_options, default=[])

min_create = df_raw["Order Create Date"].min()
max_create = df_raw["Order Create Date"].max()
min_ship = df_raw["Start Ship Date"].min()
max_ship = df_raw["Start Ship Date"].max()

create_range = f5.date_input(
    "Order Create Date",
    value=(),
    min_value=min_create.date() if pd.notna(min_create) else None,
    max_value=max_create.date() if pd.notna(max_create) else None,
)
ship_range = f6.date_input(
    "Start Ship Date",
    value=(),
    min_value=min_ship.date() if pd.notna(min_ship) else None,
    max_value=max_ship.date() if pd.notna(max_ship) else None,
)

df = df_raw.copy()
if sel_status:
    df = df[df["Status"].isin(sel_status)]
if sel_brand:
    df = df[df["Brand"].isin(sel_brand)]
if sel_store:
    df = df[df["Store Name"].isin(sel_store)]
if sel_type:
    df = df[df["Order Type"].isin(sel_type)]
if isinstance(create_range, tuple) and len(create_range) == 2:
    df = df[(df["Order Create Date"].dt.date >= create_range[0]) & (df["Order Create Date"].dt.date <= create_range[1])]
if isinstance(ship_range, tuple) and len(ship_range) == 2:
    df = df[(df["Start Ship Date"].dt.date >= ship_range[0]) & (df["Start Ship Date"].dt.date <= ship_range[1])]

# ---------------------------------------------------------------------------
# KPI CALCULATIONS
# ---------------------------------------------------------------------------
df_backlog = df[df["Status"].isin(BACKLOG_STATUSES)]
df_shipped = df_backlog[df_backlog["Is Shipped"] == "Y"]
df_picked_rows = df_backlog[df_backlog["Pick QTY"] > 0]

total_backlog_orders = df_backlog["Order No."].nunique()
backlog_required_qty = df_backlog["Required QTY"].sum()
backlog_items_sbc = df_backlog["SBC"].nunique()
allocated_qty = df_backlog["Allocated QTY"].sum()
picked_qty = df_backlog["Pick QTY"].sum()
shipped_qty = df_shipped["Pick QTY"].sum()
manpower_picked = df_picked_rows["User Pick"].nunique()
productivity = (picked_qty / manpower_picked) if manpower_picked else 0

pct_allocated = (allocated_qty / backlog_required_qty * 100) if backlog_required_qty else 0
pct_picked = (picked_qty / backlog_required_qty * 100) if backlog_required_qty else 0
pct_shipped = (shipped_qty / backlog_required_qty * 100) if backlog_required_qty else 0

# ---------------------------------------------------------------------------
# KPI CARDS
# ---------------------------------------------------------------------------
c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
with c1:
    kpi_card("Total Backlog Orders", f"{total_backlog_orders:,}", "Unique pending orders (รอเบิก+รอยืนยัน)", CARD_COLORS[0])
with c2:
    kpi_card("Backlog Required QTY", f"{backlog_required_qty:,.0f}", f"{backlog_items_sbc:,} items (SBCs)", CARD_COLORS[1])
with c3:
    kpi_card("Allocated QTY", f"{allocated_qty:,.0f}", f"{pct_allocated:.1f}% of required QTY", CARD_COLORS[2])
with c4:
    kpi_card("Picked QTY", f"{picked_qty:,.0f}", f"{pct_picked:.1f}% of required QTY", CARD_COLORS[3])
with c5:
    kpi_card("Shipped QTY", f"{shipped_qty:,.0f}", f"{pct_shipped:.1f}% of required QTY", CARD_COLORS[4])
with c6:
    kpi_card("Manpower Picked", f"{manpower_picked:,}", "Unique pick users", CARD_COLORS[5])
with c7:
    kpi_card("Productivity", f"{productivity:,.1f}", "Pick QTY / Manpower", CARD_COLORS[6])

st.markdown("---")

# ---------------------------------------------------------------------------
# CHARTS ROW 1: donuts
# ---------------------------------------------------------------------------
col1, col2 = st.columns(2)

status_order_counts = df.groupby("Status")["Order No."].nunique().reset_index(name="Unique Orders")
status_qty_counts = df.groupby("Status")["Required QTY"].sum().reset_index(name="Required QTY")

donut_colors = {
    "รอเบิก": "#2563EB",
    "Allocated": "#059669",
    "รอยืนยัน": "#7C3AED",
    "Allocated ShortAll": "#D97706",
    "ปิดเอกสาร": "#9CA3AF",
    "ยกเลิก": "#DC2626",
}

with col1:
    st.markdown("#### Backlog Orders by Status")
    fig = go.Figure(data=[go.Pie(
        labels=status_order_counts["Status"],
        values=status_order_counts["Unique Orders"],
        hole=0.55,
        marker=dict(colors=[donut_colors.get(s, "#9CA3AF") for s in status_order_counts["Status"]]),
    )])
    fig.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=340, legend=dict(orientation="v"))
    st.plotly_chart(fig, use_container_width=True)

with col2:
    st.markdown("#### Backlog QTY by Status")
    fig2 = go.Figure(data=[go.Pie(
        labels=status_qty_counts["Status"],
        values=status_qty_counts["Required QTY"],
        hole=0.55,
        marker=dict(colors=[donut_colors.get(s, "#9CA3AF") for s in status_qty_counts["Status"]]),
    )])
    fig2.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=340, legend=dict(orientation="v"))
    st.plotly_chart(fig2, use_container_width=True)

# ---------------------------------------------------------------------------
# CHARTS ROW 2: top 10 stores + daily backlog
# ---------------------------------------------------------------------------
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
    fig3.update_layout(height=380, margin=dict(t=10, b=10, l=10, r=10),
                        xaxis_title="", yaxis_title="", legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig3, use_container_width=True)

with col4:
    st.markdown("#### Order Backlog (Daily)")
    created = df.groupby(df["Order Create Date"].dt.date)["Order No."].nunique().reset_index(name="Created Orders")
    created.columns = ["Date", "Created Orders"]
    picked = (
        df[df["Is Picked"] == "Y"]
        .groupby(df.loc[df["Is Picked"] == "Y", "Picked Date"].dt.date)["Order No."]
        .nunique()
        .reset_index(name="Picked Orders")
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
    st.plotly_chart(fig4, use_container_width=True)

st.markdown("---")

# ---------------------------------------------------------------------------
# BACKLOG ITEMS DETAILS TABLE
# ---------------------------------------------------------------------------
top_row1, top_row2 = st.columns([3, 2])
top_row1.markdown("#### Backlog Items Details")
search_term = top_row2.text_input("🔍 Search Order No, Store, SBC...", value="", label_visibility="collapsed",
                                    placeholder="Search Order No, Store, SBC...")

detail_cols = ["Order No.", "Order Create Date", "Start Ship Date", "Brand", "To Store", "Store Name",
                "Status", "SBC", "Required QTY", "Allocated QTY", "Pick QTY"]
detail_cols = [c for c in detail_cols if c in df_backlog.columns]
details = df_backlog[detail_cols].copy()

if search_term:
    mask = pd.Series(False, index=details.index)
    for c in ["Order No.", "Store Name", "SBC"]:
        if c in details.columns:
            mask |= details[c].astype(str).str.contains(search_term, case=False, na=False)
    details = details[mask]

details_display = details.copy()
if "Order Create Date" in details_display.columns:
    details_display["Order Create Date"] = details_display["Order Create Date"].dt.strftime("%d/%m/%Y")
if "Start Ship Date" in details_display.columns:
    details_display["Start Ship Date"] = details_display["Start Ship Date"].dt.strftime("%d/%m/%Y")

st.dataframe(details_display, use_container_width=True, height=420, hide_index=True)

csv_bytes = details.to_csv(index=False).encode("utf-8-sig")
st.download_button("⬇️ Export Table Details (CSV)", data=csv_bytes, file_name="backlog_items_details.csv",
                    mime="text/csv")
