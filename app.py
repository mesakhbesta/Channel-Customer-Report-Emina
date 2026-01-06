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

@st.cache_data(show_spinner=False)
def load_sheet(file, sheet, key_col, val_col, _parser=None, skip=0):
    df = pd.read_excel(file, sheet_name=sheet, skiprows=skip)
    d = {}
    for _, row in df.iterrows():
        key = row[key_col]
        val = row[val_col]
        if _parser:
            val = _parser(val)
        d[key] = val
    return d

# =====================================================
# LOAD METRICS
# =====================================================
with st.spinner("🔄 Loading Channel Metrics..."):
    channel_metrics = {
        "cont": load_sheet(channel_file, "Sheet 18", "Customer P",
                           "% of Total Current DO TP2 along Customer P, Customer P Hidden", _parser=parse_percent),
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
        "cont": load_sheet(customer_file, "Sheet 18", "Customer P",
                           "% of Total Current DO TP2 along Customer P, Customer P Hidden", _parser=parse_percent),
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
# FLEXIBLE COLUMN MAPPING (CHANNEL & CUSTOMER)
# =====================================================
CHANNEL_COL_CANDIDATES = [
    "CHANNEL_REPORT_NAME",
    "CHANNEL_NAME",
    "CHANNEL",
    "SALES_CHANNEL"
]

CUSTOMER_COL_CANDIDATES = [
    "CUSTOMER_GROUP",
    "CUSTOMER_NAME",
    "CUSTOMER",
    "CUST_GROUP"
]

def find_column(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

st.sidebar.divider()
st.sidebar.header("🧩 Master Column Mapping")

channel_col = find_column(master_df, CHANNEL_COL_CANDIDATES)
customer_col = find_column(master_df, CUSTOMER_COL_CANDIDATES)

all_cols = master_df.columns.tolist()

if not channel_col:
    channel_col = st.sidebar.selectbox("Pilih kolom CHANNEL di Master", options=[""] + all_cols)

if not customer_col:
    customer_col = st.sidebar.selectbox("Pilih kolom CUSTOMER di Master", options=[""] + all_cols)

missing = []
if not channel_col or channel_col not in master_df.columns:
    missing.append("CHANNEL")
if not customer_col or customer_col not in master_df.columns:
    missing.append("CUSTOMER")

if missing:
    st.error(f"❌ Kolom berikut belum valid di Master: {', '.join(missing)}")
    st.stop()

# ===== RAPIH & COLLAPSIBLE DISPLAY =====
with st.sidebar.expander("✅ Master Column Mapping (click to expand)", expanded=False):
    col1, col2 = st.columns([1, 2])

    with col1:
        st.markdown("**Channel**")
        st.markdown("**Customer Group**")

    with col2:
        st.markdown(f":blue[{channel_col}]")
        st.markdown(f":blue[{customer_col}]")

# =====================================================
# BUILD CHANNEL → CUSTOMER MAPPING
# =====================================================
with st.spinner("🔄 Building Channel → Customer mapping..."):
    @st.cache_data
    def build_channel_to_customers(df, ch_col, cust_col):
        mapping = {}
        for ch, g in df.groupby(ch_col):
            customers = [str(c) if pd.notna(c) else "Data Kosong" for c in g[cust_col].unique()]
            mapping[ch] = customers
        return mapping

    channel_to_customers = build_channel_to_customers(master_df, channel_col, customer_col)

# =====================================================
# PRECOMPUTE METRICS ROWS
# =====================================================
with st.spinner("🔄 Precomputing Metrics rows..."):
    @st.cache_data
    def build_rows_dict(metrics_dict):
        rows = {}
        for key in metrics_dict["mtd"].keys():
            row = [
                metrics_dict["cont"].get(key),
                metrics_dict["mtd"].get(key),
                metrics_dict["ytd"].get(key),
                metrics_dict["g_mtd"].get(key),
                metrics_dict["g_l3m"].get(key),
                metrics_dict["g_ytd"].get(key),
                metrics_dict["a_mtd"].get(key),
                metrics_dict["a_ytd"].get(key),
            ]
            row = [v if v is not None else 0 for v in row]
            rows[key] = row
        return rows

    customer_rows_dict = build_rows_dict(customer_metrics)
    channel_rows_dict = build_rows_dict(channel_metrics)

# =====================================================
# FILTER SECTION
# =====================================================
st.sidebar.header("🎯 Filter Data")

channels = list(channel_to_customers.keys())

def select_all():
    st.session_state.channel_selector = channels

def unselect_all():
    st.session_state.channel_selector = []

select_option = st.sidebar.radio(
    "Channel Selection Mode",
    ["Custom", "Select All", "Unselect All"],
    index=0
)

if select_option == "Select All":
    select_all()
elif select_option == "Unselect All":
    unselect_all()

selected_channels = st.sidebar.multiselect(
    "Channel",
    options=channels,
    default=st.session_state.get("channel_selector", []),
    key="channel_selector"
)
st.session_state["channel"] = selected_channels

customers_filtered = []
for ch in st.session_state["channel"]:
    customers_filtered.extend(channel_to_customers.get(ch, []))

normalized_customers = sorted(list(set([str(c) if pd.notna(c) else "Data Kosong" for c in customers_filtered])))

old_customer_selection = st.session_state.get("customer", [])
default_selection = [c for c in old_customer_selection if c in normalized_customers]

selected_customers_ui = st.sidebar.multiselect(
    "Customer Group",
    options=normalized_customers,
    default=default_selection,
    key="customer_selector"
)
st.session_state["customer"] = selected_customers_ui

# =====================================================
# BUILD ROWS
# =====================================================
def build_row(label, metrics_dict, indent=False):
    lbl = f"    {label}" if indent else label
    return [lbl] + metrics_dict.get(label, [0]*8)

rows = []
rows.append(build_row("GRAND TOTAL", channel_rows_dict))

for ch in st.session_state["channel"]:
    rows.append(build_row(ch, channel_rows_dict))
    for cust in st.session_state["customer"]:
        if cust in channel_to_customers.get(ch, []):
            rows.append(build_row(cust, customer_rows_dict, indent=True))

# =====================================================
# DISPLAY TABLE
# =====================================================
st.subheader("📈 Performance Table (Filtered Preview)")
st.caption(f"Data as of **{cutoff_str}**. Preview hanya baris yang dipilih filter.")

display_df = pd.DataFrame(rows, columns=[
    "Channel/Customer","Cont YTD","Value MTD","Value YTD",
    "Growth MTD","Growth %Gr L3M","Growth YTD","Ach MTD","Ach YTD"
])

def fmt_pct(x):
    return f"{x:.1f}%" if pd.notna(x) else "0%"

for c in ["Cont YTD","Growth MTD","Growth %Gr L3M","Growth YTD","Ach MTD","Ach YTD"]:
    display_df[c] = display_df[c].apply(fmt_pct)

st.dataframe(display_df, use_container_width=True)

# =====================================================
# DOWNLOAD SECTION
# =====================================================
st.divider()
st.subheader("⬇️ Export Full Report")

output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    wb = writer.book
    ws = wb.add_worksheet("Report")
    writer.sheets["Report"] = ws

    header = wb.add_format({"bold": True,"align":"center","border":1})
    bold = wb.add_format({"bold": True,"border":1})
    ind = wb.add_format({"border":1,"indent":2,"font_color":"blue"})
    num = wb.add_format({"border":1,"num_format":"#,##0"})
    pct_g = wb.add_format({"border":1,"num_format":"0.0%","font_color":"green"})
    pct_r = wb.add_format({"border":1,"num_format":"0.0%","font_color":"red"})

    ws.write(0,0,f"Cut-off: {cutoff_str}",header)
    ws.write_row(1,0,display_df.columns.tolist(),header)

    for i, r in enumerate(rows,start=2):
        name_fmt = ind if r[0].startswith("    ") else bold
        ws.write(i,0,r[0].strip(),name_fmt)
        for c in range(1,9):
            v = r[c]
            if c==1 or c>=4:
                ws.write_number(i,c,v/100,pct_g if v>=0 else pct_r)
            else:
                ws.write_number(i,c,v or 0,num)

    ws.set_column("A:A",50)
    ws.set_column("B:I",18)

output.seek(0)

st.download_button(
    "📥 Download Excel Report",
    output,
    "Channel & Customer Group Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
