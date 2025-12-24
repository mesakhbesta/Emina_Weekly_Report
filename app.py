import streamlit as st
import pandas as pd
from io import BytesIO
import datetime

# ================== PAGE ==================
st.set_page_config(layout="wide")
st.title("Dynamic Emina Metrics Report")

cutoff_date = st.sidebar.date_input("Select Cut-off Date", datetime.date.today())
cutoff_str = cutoff_date.strftime("%d %B %Y")

# ================== UPLOAD ==================
with st.sidebar.expander("Upload Excel Files (click to expand)", expanded=False):
    master_file = st.file_uploader("Master Product", type=["xlsx"])
    format_file = st.file_uploader("Format Metrics", type=["xlsx"])
    variant_file = st.file_uploader("Variant Metrics", type=["xlsx"])
    product_file = st.file_uploader("Product Metrics", type=["xlsx"])

if not all([master_file, format_file, variant_file, product_file]):
    st.warning("Please upload all 4 Excel files.")
    st.stop()

# ================== CACHE ==================
@st.cache_data(show_spinner=False)
def load_excel(file, sheet, skip=0):
    return pd.read_excel(file, sheet_name=sheet, skiprows=skip)

# ================== PARSER ==================
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

def load_map(sheet, key_col, val_col, file, skip=0, parser=None):
    tmp = load_excel(file, sheet, skip)
    result = {}
    for _, r in tmp.iterrows():
        v = r[val_col]
        if parser:
            v = parser(v)
        result[r[key_col]] = v
    return result

# ================== DATA ==================
df = pd.read_excel(master_file)

def load_all_maps(file):
    return {
        "cont": load_map("Sheet 18", "Product P",
            "% of Total Current DO TP2 along Product P, Product P Hidden",
            file, parser=parse_percent),
        "mtd": load_map("Sheet 1", "Product P", "Current DO", file, parser=parse_number),
        "ytd": load_map("Sheet 1", "Product P", "Current DO TP2", file, parser=parse_number),
        "gr_mtd": load_map("Sheet 4", "Product P", "vs LY", file, 1, parse_percent),
        "gr_l3m": load_map("Sheet 3", "Product P", "vs L3M", file, 1, parse_percent),
        "gr_ytd": load_map("Sheet 5", "Product P", "vs LY", file, 1, parse_percent),
        "ach_mtd": load_map("Sheet 13", "Product P", "Current Achievement", file, parser=parse_percent),
        "ach_ytd": load_map("Sheet 14", "Product P", "Current Achievement TP2", file, parser=parse_percent),
    }

fmt = load_all_maps(format_file)
var = load_all_maps(variant_file)
prd = load_all_maps(product_file)

# ================== FILTER HELPER ==================
def make_option_map(values):
    values = list(values)
    return {i: v for i, v in enumerate(values)}

for k in ["format", "variant", "product"]:
    st.session_state.setdefault(k, [])

st.sidebar.title("Filter Products")

# ===== FORMAT FILTER (1 CLICK FIX) =====
all_formats = sorted(df["PRODUCT_FORMAT"].dropna().unique())
format_map = make_option_map(all_formats)

fmt_ids = st.sidebar.multiselect(
    "Format",
    options=list(format_map.keys()),
    format_func=lambda x: format_map[x],
    default=[k for k, v in format_map.items() if v in st.session_state["format"]],
    key="fmt_filter"
)
st.session_state["format"] = [format_map[i] for i in fmt_ids]

# ===== VARIANT FILTER =====
variant_pool = df[df["PRODUCT_FORMAT"].isin(st.session_state["format"])]["PRODUCT_VARIANT_NAME"].dropna().unique()
variant_pool = sorted(variant_pool)
variant_map = make_option_map(variant_pool)

var_ids = st.sidebar.multiselect(
    "Variant",
    options=list(variant_map.keys()),
    format_func=lambda x: variant_map[x],
    default=[k for k, v in variant_map.items() if v in st.session_state["variant"]],
    key="var_filter"
)
st.session_state["variant"] = [variant_map[i] for i in var_ids]

# ===== PRODUCT FILTER =====
product_pool = df[df["PRODUCT_VARIANT_NAME"].isin(st.session_state["variant"])]["PRODUCT_NAME"].dropna().unique()
product_pool = sorted(product_pool)
product_map = make_option_map(product_pool)

prd_ids = st.sidebar.multiselect(
    "Product",
    options=list(product_map.keys()),
    format_func=lambda x: product_map[x],
    default=[k for k, v in product_map.items() if v in st.session_state["product"]],
    key="prd_filter"
)
st.session_state["product"] = [product_map[i] for i in prd_ids]

# ================== ROW BUILD ==================
rows = []

rows.append([
    "GRAND TOTAL",
    fmt["cont"].get("GRAND TOTAL"),
    fmt["mtd"].get("GRAND TOTAL"),
    fmt["ytd"].get("GRAND TOTAL"),
    fmt["gr_mtd"].get("GRAND TOTAL"),
    fmt["gr_l3m"].get("GRAND TOTAL"),
    fmt["gr_ytd"].get("GRAND TOTAL"),
    fmt["ach_mtd"].get("GRAND TOTAL"),
    fmt["ach_ytd"].get("GRAND TOTAL"),
])

for f in st.session_state["format"]:
    rows.append([
        f,
        fmt["cont"].get(f),
        fmt["mtd"].get(f),
        fmt["ytd"].get(f),
        fmt["gr_mtd"].get(f),
        fmt["gr_l3m"].get(f),
        fmt["gr_ytd"].get(f),
        fmt["ach_mtd"].get(f),
        fmt["ach_ytd"].get(f),
    ])
    for v in st.session_state["variant"]:
        if v in df[df["PRODUCT_FORMAT"] == f]["PRODUCT_VARIANT_NAME"].values:
            rows.append([
                f"        {v}",
                var["cont"].get(v),
                var["mtd"].get(v),
                var["ytd"].get(v),
                var["gr_mtd"].get(v),
                var["gr_l3m"].get(v),
                var["gr_ytd"].get(v),
                var["ach_mtd"].get(v),
                var["ach_ytd"].get(v),
            ])
            for p in st.session_state["product"]:
                if p in df[df["PRODUCT_VARIANT_NAME"] == v]["PRODUCT_NAME"].values:
                    rows.append([
                        f"            {p}",
                        prd["cont"].get(p),
                        prd["mtd"].get(p),
                        prd["ytd"].get(p),
                        prd["gr_mtd"].get(p),
                        prd["gr_l3m"].get(p),
                        prd["gr_ytd"].get(p),
                        prd["ach_mtd"].get(p),
                        prd["ach_ytd"].get(p),
                    ])

# ================== DISPLAY ==================
display_df = pd.DataFrame(rows, columns=[
    "Produk", "Cont YTD", "Value MTD", "Value YTD",
    "Growth MTD", "Growth %Gr L3M", "Growth YTD",
    "Ach MTD", "Ach YTD"
])

def pct(x): return f"{x:.1f}%" if pd.notna(x) else ""

for c in ["Cont YTD","Growth MTD","Growth %Gr L3M","Growth YTD","Ach MTD","Ach YTD"]:
    display_df[c] = display_df[c].apply(pct)

display_df.columns = pd.MultiIndex.from_tuples([
    ("Cut-off: " + cutoff_str, ""),
    ("", "Cont YTD"),
    ("Value", "MTD"),
    ("Value", "YTD"),
    ("Growth", "MTD"),
    ("Growth", "%Gr L3M"),
    ("Growth", "YTD"),
    ("Ach", "MTD"),
    ("Ach", "YTD"),
])

st.dataframe(display_df, use_container_width=True)

# ================== DOWNLOAD ==================
output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    display_df.to_excel(writer, sheet_name="Report", startrow=2, index=False)
output.seek(0)

st.download_button(
    "Download Excel",
    output,
    "Report_Full_Level.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
