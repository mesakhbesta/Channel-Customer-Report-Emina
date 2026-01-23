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
# INIT SESSION STATE (ANTI RESET)
# =====================================================
for k, v in {
    "channel": [],
    "customer": [],
    "lock_selection": False
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

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
    if pd.isna(val): return None
    if isinstance(val, str):
        return round(float(val.replace("%","").replace(",", ".")), 1)
    return round(float(val) * 100, 1)

def parse_number(val):
    if pd.isna(val): return None
    return round(float(val), 0)

@st.cache_data(show_spinner=False)
def load_sheet(file, sheet, key_col, val_col, parser=None, skip=0):
    df = pd.read_excel(file, sheet_name=sheet, skiprows=skip)
    out = {}
    for _, r in df.iterrows():
        v = parser(r[val_col]) if parser else r[val_col]
        out[r[key_col]] = v
    return out

def sanitize_selection(old, options, lock):
    if lock:
        return old
    return [x for x in old if x in options]

# =====================================================
# LOAD METRICS
# =====================================================
channel_metrics = {
    "cont": load_sheet(channel_file,"Sheet 18","Customer P",
        "% of Total Current DO TP2 along Customer P, Customer P Hidden",parse_percent),
    "mtd": load_sheet(channel_file,"Sheet 1","Customer P","Current DO",parse_number),
    "ytd": load_sheet(channel_file,"Sheet 1","Customer P","Current DO TP2",parse_number),
    "g_mtd": load_sheet(channel_file,"Sheet 4","Customer P","vs LY",parse_percent,1),
    "g_l3m": load_sheet(channel_file,"Sheet 3","Customer P","vs L3M",parse_percent,1),
    "g_ytd": load_sheet(channel_file,"Sheet 5","Customer P","vs LY",parse_percent,1),
    "a_mtd": load_sheet(channel_file,"Sheet 13","Customer P","Current Achievement",parse_percent),
    "a_ytd": load_sheet(channel_file,"Sheet 14","Customer P","Current Achievement TP2",parse_percent),
}

customer_metrics = {
    "cont": load_sheet(customer_file,"Sheet 18","Customer P",
        "% of Total Current DO TP2 along Customer P, Customer P Hidden",parse_percent),
    "mtd": load_sheet(customer_file,"Sheet 1","Customer P","Current DO",parse_number),
    "ytd": load_sheet(customer_file,"Sheet 1","Customer P","Current DO TP2",parse_number),
    "g_mtd": load_sheet(customer_file,"Sheet 4","Customer P","vs LY",parse_percent,1),
    "g_l3m": load_sheet(customer_file,"Sheet 3","Customer P","vs L3M",parse_percent,1),
    "g_ytd": load_sheet(customer_file,"Sheet 5","Customer P","vs LY",parse_percent,1),
    "a_mtd": load_sheet(customer_file,"Sheet 13","Customer P","Current Achievement",parse_percent),
    "a_ytd": load_sheet(customer_file,"Sheet 14","Customer P","Current Achievement TP2",parse_percent),
}

# =====================================================
# LOAD MASTER
# =====================================================
master_df = pd.read_excel(master_file)

CHANNEL_COLS = ["CHANNEL_REPORT_NAME","CHANNEL_NAME","CHANNEL","SALES_CHANNEL"]
CUSTOMER_COLS = ["CUSTOMER_GROUP","CUSTOMER_NAME","CUSTOMER","CUST_GROUP"]

def find_col(df, cands):
    for c in cands:
        if c in df.columns:
            return c
    return None

channel_col = find_col(master_df, CHANNEL_COLS)
customer_col = find_col(master_df, CUSTOMER_COLS)

if not channel_col or not customer_col:
    st.error("❌ Channel / Customer column not found in Master")
    st.stop()

# =====================================================
# BUILD CHANNEL → CUSTOMER MAP
# =====================================================
@st.cache_data
def build_map(df, ch, cust):
    out = {}
    for c, g in df.groupby(ch):
        out[c] = sorted(g[cust].astype(str).unique())
    return out

channel_to_customers = build_map(master_df, channel_col, customer_col)

# =====================================================
# 🔒 FILTER SECTION
# =====================================================
st.sidebar.header("🎯 Filter Data")

lock = st.sidebar.toggle(
    "🔒 Lock Selection",
    value=st.session_state.lock_selection,
    key="lock_selection",
    help="Jika ON, pilihan Channel & Customer tidak berubah walau upload ulang"
)

channels = sorted(channel_to_customers.keys())
st.session_state.channel = sanitize_selection(
    st.session_state.channel, channels, lock
)

st.session_state.channel = st.sidebar.multiselect(
    "Channel",
    channels,
    default=st.session_state.channel,
    disabled=lock
)

customers = []
for ch in st.session_state.channel:
    customers.extend(channel_to_customers.get(ch, []))
customers = sorted(set(customers))

st.session_state.customer = sanitize_selection(
    st.session_state.customer, customers, lock
)

st.session_state.customer = st.sidebar.multiselect(
    "Customer Group",
    customers,
    default=st.session_state.customer,
    disabled=lock
)

if lock:
    st.sidebar.caption("🔒 Selection terkunci")

# =====================================================
# PREPARE ROWS
# =====================================================
def build_row(label, metrics, indent=False):
    name = f"    {label}" if indent else label
    vals = [metrics[k].get(label,0) for k in
            ["cont","mtd","ytd","g_mtd","g_l3m","g_ytd","a_mtd","a_ytd"]]
    return [name] + vals

rows = [build_row("GRAND TOTAL", channel_metrics)]

for ch in st.session_state.channel:
    rows.append(build_row(ch, channel_metrics))
    for cust in st.session_state.customer:
        if cust in channel_to_customers.get(ch, []):
            rows.append(build_row(cust, customer_metrics, True))

# =====================================================
# DISPLAY
# =====================================================
df_disp = pd.DataFrame(rows, columns=[
    "Channel / Customer","Cont YTD","Value MTD","Value YTD",
    "Growth MTD","%Gr L3M","Growth YTD","Ach MTD","Ach YTD"
])

def pct(x): return f"{x:.1f}%" if x is not None else "0%"

for c in ["Cont YTD","Growth MTD","%Gr L3M","Growth YTD","Ach MTD","Ach YTD"]:
    df_disp[c] = df_disp[c].apply(pct)

st.subheader("📈 Performance Table")
st.dataframe(df_disp, use_container_width=True)

# =====================================================
# DOWNLOAD
# =====================================================
output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    df_disp.to_excel(writer, index=False, sheet_name="Report")

st.download_button(
    "📥 Download Excel Report",
    output.getvalue(),
    "Channel_Customer_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
