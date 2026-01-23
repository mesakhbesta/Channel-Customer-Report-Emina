import streamlit as st
import pandas as pd
from io import BytesIO
import datetime

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(layout="wide", page_title="Channel & Customer Group Report")

st.title("📊 Channel & Customer Group Report")
st.subheader("Performance Overview")
st.divider()

# =====================================================
# SIDEBAR – DATE & FILE UPLOAD
# =====================================================
st.sidebar.header("🗓️ Reporting Settings")

cutoff_date = st.sidebar.date_input("Cut-off Date", datetime.date.today())
cutoff_str = cutoff_date.strftime("%d %B %Y")
st.sidebar.info(f"📌 Cut-off Date: **{cutoff_str}**")
st.sidebar.divider()

with st.sidebar.expander("📁 Upload Excel Files", expanded=True):
    master_file = st.file_uploader("Master Data", type=["xlsx"])
    channel_file = st.file_uploader("Channel Metrics", type=["xlsx"])
    customer_file = st.file_uploader("Customer Metrics", type=["xlsx"])

if not all([master_file, channel_file, customer_file]):
    st.warning("⚠️ Upload all 3 files to continue.")
    st.stop()

# =====================================================
# HELPERS
# =====================================================
def parse_percent(val):
    if pd.isna(val):
        return None
    if isinstance(val, str):
        return round(float(val.replace("%", "").replace(",", ".")), 1)
    return round(float(val) * 100, 1)

def parse_number(val):
    if pd.isna(val):
        return None
    return round(float(val), 0)

# ⚠️ FIX UTAMA ADA DI SINI (parser ➜ _parser)
@st.cache_data(show_spinner=False)
def load_sheet(file, sheet, key_col, val_col, _parser=None, skip=0):
    df = pd.read_excel(file, sheet_name=sheet, skiprows=skip)
    out = {}
    for _, r in df.iterrows():
        key = r[key_col]
        val = r[val_col]
        if _parser:
            val = _parser(val)
        out[key] = val
    return out

# =====================================================
# LOAD METRICS
# =====================================================
with st.spinner("🔄 Loading Channel Metrics..."):
    channel_metrics = {
        "cont": load_sheet(
            channel_file, "Sheet 18", "Customer P",
            "% of Total Current DO TP2 along Customer P, Customer P Hidden",
            _parser=parse_percent
        ),
        "mtd": load_sheet(channel_file, "Sheet 1", "Customer P", "Current DO", _parser=parse_number),
        "ytd": load_sheet(channel_file, "Sheet 1", "Customer P", "Current DO TP2", _parser=parse_number),
        "g_mtd": load_sheet(channel_file, "Sheet 4", "Customer P", "vs LY", _parser=parse_percent, skip=1),
        "g_l3m": load_sheet(channel_file, "Sheet 3", "Customer P", "vs L3M", _parser=parse_percent, skip=1),
        "g_ytd": load_sheet(channel_file, "Sheet 5", "Customer P", "vs LY", _parser=parse_percent, skip=1),
        "a_mtd": load_sheet(channel_file, "Sheet 13", "Customer P", "Current Achievement", _parser=parse_percent),
        "a_ytd": load_sheet(channel_file, "Sheet 14", "Customer P", "Current Achievement TP2", _parser=parse_percent),
    }

with st.spinner("🔄 Loading Customer Metrics..."):
    customer_metrics = {
        "cont": load_sheet(
            customer_file, "Sheet 18", "Customer P",
            "% of Total Current DO TP2 along Customer P, Customer P Hidden",
            _parser=parse_percent
        ),
        "mtd": load_sheet(customer_file, "Sheet 1", "Customer P", "Current DO", _parser=parse_number),
        "ytd": load_sheet(customer_file, "Sheet 1", "Customer P", "Current DO TP2", _parser=parse_number),
        "g_mtd": load_sheet(customer_file, "Sheet 4", "Customer P", "vs LY", _parser=parse_percent, skip=1),
        "g_l3m": load_sheet(customer_file, "Sheet 3", "Customer P", "vs L3M", _parser=parse_percent, skip=1),
        "g_ytd": load_sheet(customer_file, "Sheet 5", "Customer P", "vs LY", _parser=parse_percent, skip=1),
        "a_mtd": load_sheet(customer_file, "Sheet 13", "Customer P", "Current Achievement", _parser=parse_percent),
        "a_ytd": load_sheet(customer_file, "Sheet 14", "Customer P", "Current Achievement TP2", _parser=parse_percent),
    }

# =====================================================
# LOAD MASTER
# =====================================================
master_df = pd.read_excel(master_file)

# =====================================================
# FLEXIBLE COLUMN MAPPING
# =====================================================
CHANNEL_COL_CANDIDATES = [
    "CHANNEL_REPORT_NAME", "CHANNEL_NAME", "CHANNEL", "SALES_CHANNEL"
]

CUSTOMER_COL_CANDIDATES = [
    "CUSTOMER_GROUP", "CUSTOMER_NAME", "CUSTOMER", "CUST_GROUP"
]

def find_column(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

st.sidebar.header("🧩 Master Column Mapping")

channel_col = find_column(master_df, CHANNEL_COL_CANDIDATES)
customer_col = find_column(master_df, CUSTOMER_COL_CANDIDATES)

all_cols = master_df.columns.tolist()

if not channel_col:
    channel_col = st.sidebar.selectbox("Pilih kolom CHANNEL", [""] + all_cols)

if not customer_col:
    customer_col = st.sidebar.selectbox("Pilih kolom CUSTOMER", [""] + all_cols)

if not channel_col or not customer_col:
    st.error("❌ Kolom CHANNEL / CUSTOMER belum valid")
    st.stop()

# =====================================================
# BUILD CHANNEL → CUSTOMER MAP
# =====================================================
@st.cache_data
def build_channel_to_customer(df, ch_col, cust_col):
    return {
        ch: sorted(g[cust_col].dropna().astype(str).unique().tolist())
        for ch, g in df.groupby(ch_col)
    }

channel_to_customer = build_channel_to_customer(master_df, channel_col, customer_col)

# =====================================================
# PRECOMPUTE ROWS
# =====================================================
@st.cache_data
def build_rows(metrics):
    rows = {}
    for k in metrics["mtd"].keys():
        rows[k] = [
            metrics["cont"].get(k, 0),
            metrics["mtd"].get(k, 0),
            metrics["ytd"].get(k, 0),
            metrics["g_mtd"].get(k, 0),
            metrics["g_l3m"].get(k, 0),
            metrics["g_ytd"].get(k, 0),
            metrics["a_mtd"].get(k, 0),
            metrics["a_ytd"].get(k, 0),
        ]
    return rows

channel_rows = build_rows(channel_metrics)
customer_rows = build_rows(customer_metrics)

# =====================================================
# FILTERS
# =====================================================
st.sidebar.header("🎯 Filters")

channels = list(channel_to_customer.keys())

if "channel" not in st.session_state:
    st.session_state.channel = channels

st.session_state.channel = st.sidebar.multiselect(
    "Channel",
    channels,
    default=st.session_state.channel
)

customers = sorted({
    c for ch in st.session_state.channel
    for c in channel_to_customer.get(ch, [])
})

if "customer" not in st.session_state:
    st.session_state.customer = customers

st.session_state.customer = st.sidebar.multiselect(
    "Customer Group",
    customers,
    default=[c for c in st.session_state.customer if c in customers]
)

# =====================================================
# BUILD DISPLAY ROWS
# =====================================================
rows = [["GRAND TOTAL"] + channel_rows.get("GRAND TOTAL", [0]*8)]

for ch in st.session_state.channel:
    rows.append([ch] + channel_rows.get(ch, [0]*8))
    for cu in st.session_state.customer:
        if cu in channel_to_customer.get(ch, []):
            rows.append([f"    {cu}"] + customer_rows.get(cu, [0]*8))

# =====================================================
# DISPLAY
# =====================================================
st.subheader("📈 Performance Table")

df = pd.DataFrame(rows, columns=[
    "Channel / Customer",
    "Cont YTD","Value MTD","Value YTD",
    "Growth MTD","%Gr L3M","Growth YTD","Ach MTD","Ach YTD"
])

for c in ["Cont YTD","Growth MTD","%Gr L3M","Growth YTD","Ach MTD","Ach YTD"]:
    df[c] = df[c].apply(lambda x: f"{x:.1f}%")

st.dataframe(df, use_container_width=True)

# =====================================================
# DOWNLOAD
# =====================================================
output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    df.to_excel(writer, index=False, sheet_name="Report")

output.seek(0)

st.download_button(
    "📥 Download Excel",
    output,
    "Channel_Customer_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
