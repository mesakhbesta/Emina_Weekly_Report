import streamlit as st
import pandas as pd
from io import BytesIO
import datetime

# =============================
# PAGE CONFIG
# =============================
st.set_page_config(layout="wide")
st.title("Dynamic Emina Metrics Report")

# =============================
# CUTOFF DATE
# =============================
cutoff_date = st.sidebar.date_input("Select Cut-off Date", datetime.date.today())
cutoff_str = cutoff_date.strftime("%d %B %Y")

# =============================
# FILE UPLOAD
# =============================
with st.sidebar.expander("Upload Excel Files (click to expand)", expanded=False):
    master_file = st.file_uploader("Master Product", type=["xlsx"])
    format_file = st.file_uploader("Format Metrics", type=["xlsx"])
    variant_file = st.file_uploader("Variant Metrics", type=["xlsx"])
    product_file = st.file_uploader("Product Metrics", type=["xlsx"])

if not all([master_file, format_file, variant_file, product_file]):
    st.warning("Please upload all 4 Excel files.")
    st.stop()

# =============================
# CACHE EXCEL LOADING (PERFORMANCE FIX)
# =============================
@st.cache_data(show_spinner=False)
def load_excel(file, sheet, skip=0):
    return pd.read_excel(file, sheet_name=sheet, skiprows=skip)

# =============================
# PARSERS
# =============================
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

# =============================
# LOAD MASTER
# =============================
df = pd.read_excel(master_file)

# =============================
# LOAD FORMAT METRICS
# =============================
cont_map_fmt = load_map("Sheet 18", "Product P", "% of Total Current DO TP2 along Product P, Product P Hidden", format_file, parser=parse_percent)
value_mtd_fmt = load_map("Sheet 1", "Product P", "Current DO", format_file, parser=parse_number)
value_ytd_fmt = load_map("Sheet 1", "Product P", "Current DO TP2", format_file, parser=parse_number)
growth_mtd_fmt = load_map("Sheet 4", "Product P", "vs LY", format_file, skip=1, parser=parse_percent)
growth_l3m_fmt = load_map("Sheet 3", "Product P", "vs L3M", format_file, skip=1, parser=parse_percent)
growth_ytd_fmt = load_map("Sheet 5", "Product P", "vs LY", format_file, skip=1, parser=parse_percent)
ach_mtd_fmt = load_map("Sheet 13", "Product P", "Current Achievement", format_file, parser=parse_percent)
ach_ytd_fmt = load_map("Sheet 14", "Product P", "Current Achievement TP2", format_file, parser=parse_percent)

# =============================
# LOAD VARIANT METRICS
# =============================
cont_map_var = load_map("Sheet 18", "Product P", "% of Total Current DO TP2 along Product P, Product P Hidden", variant_file, parser=parse_percent)
value_mtd_var = load_map("Sheet 1", "Product P", "Current DO", variant_file, parser=parse_number)
value_ytd_var = load_map("Sheet 1", "Product P", "Current DO TP2", variant_file, parser=parse_number)
growth_mtd_var = load_map("Sheet 4", "Product P", "vs LY", variant_file, skip=1, parser=parse_percent)
growth_l3m_var = load_map("Sheet 3", "Product P", "vs L3M", variant_file, skip=1, parser=parse_percent)
growth_ytd_var = load_map("Sheet 5", "Product P", "vs LY", variant_file, skip=1, parser=parse_percent)
ach_mtd_var = load_map("Sheet 13", "Product P", "Current Achievement", variant_file, parser=parse_percent)
ach_ytd_var = load_map("Sheet 14", "Product P", "Current Achievement TP2", variant_file, parser=parse_percent)

# =============================
# LOAD PRODUCT METRICS
# =============================
cont_map_prod = load_map("Sheet 18", "Product P", "% of Total Current DO TP2 along Product P, Product P Hidden", product_file, parser=parse_percent)
value_mtd_prod = load_map("Sheet 1", "Product P", "Current DO", product_file, parser=parse_number)
value_ytd_prod = load_map("Sheet 1", "Product P", "Current DO TP2", product_file, parser=parse_number)
growth_mtd_prod = load_map("Sheet 4", "Product P", "vs LY", product_file, skip=1, parser=parse_percent)
growth_l3m_prod = load_map("Sheet 3", "Product P", "vs L3M", product_file, skip=1, parser=parse_percent)
growth_ytd_prod = load_map("Sheet 5", "Product P", "vs LY", product_file, skip=1, parser=parse_percent)
ach_mtd_prod = load_map("Sheet 13", "Product P", "Current Achievement", product_file, parser=parse_percent)
ach_ytd_prod = load_map("Sheet 14", "Product P", "Current Achievement TP2", product_file, parser=parse_percent)

# =============================
# HELPER: OPTION MAP (FILTER FIX)
# =============================
def make_option_map(values):
    values = list(values)
    return {i: v for i, v in enumerate(values)}

for k in ["format", "variant", "product"]:
    if k not in st.session_state:
        st.session_state[k] = []

# =============================
# SIDEBAR FILTERS (1 CLICK SELECT FIX)
# =============================
st.sidebar.title("Filter Products")

# FORMAT
all_formats = sorted(df["PRODUCT_FORMAT"].dropna().unique())
format_map = make_option_map(all_formats)

selected_format_ids = st.sidebar.multiselect(
    "Format",
    options=list(format_map.keys()),
    format_func=lambda x: format_map[x],
    default=[k for k, v in format_map.items() if v in st.session_state["format"]],
)

st.session_state["format"] = [format_map[i] for i in selected_format_ids]

# VARIANT
variant_pool = df[df["PRODUCT_FORMAT"].isin(st.session_state["format"])]["PRODUCT_VARIANT_NAME"].dropna().unique()
variant_pool = sorted(variant_pool)
variant_map = make_option_map(variant_pool)

selected_variant_ids = st.sidebar.multiselect(
    "Variant",
    options=list(variant_map.keys()),
    format_func=lambda x: variant_map[x],
    default=[k for k, v in variant_map.items() if v in st.session_state["variant"]],
)

st.session_state["variant"] = [variant_map[i] for i in selected_variant_ids]

# PRODUCT
product_pool = df[df["PRODUCT_VARIANT_NAME"].isin(st.session_state["variant"])]["PRODUCT_NAME"].dropna().unique()
product_pool = sorted(product_pool)
product_map = make_option_map(product_pool)

selected_product_ids = st.sidebar.multiselect(
    "Product",
    options=list(product_map.keys()),
    format_func=lambda x: product_map[x],
    default=[k for k, v in product_map.items() if v in st.session_state["product"]],
)

st.session_state["product"] = [product_map[i] for i in selected_product_ids]

# =============================
# BUILD ROWS (TIDAK DIUBAH)
# =============================
rows = []

rows.append([
    "GRAND TOTAL",
    cont_map_fmt.get("GRAND TOTAL"),
    value_mtd_fmt.get("GRAND TOTAL"),
    value_ytd_fmt.get("GRAND TOTAL"),
    growth_mtd_fmt.get("GRAND TOTAL"),
    growth_l3m_fmt.get("GRAND TOTAL"),
    growth_ytd_fmt.get("GRAND TOTAL"),
    ach_mtd_fmt.get("GRAND TOTAL"),
    ach_ytd_fmt.get("GRAND TOTAL")
])

for fmt in st.session_state["format"]:
    rows.append([
        fmt,
        cont_map_fmt.get(fmt),
        value_mtd_fmt.get(fmt),
        value_ytd_fmt.get(fmt),
        growth_mtd_fmt.get(fmt),
        growth_l3m_fmt.get(fmt),
        growth_ytd_fmt.get(fmt),
        ach_mtd_fmt.get(fmt),
        ach_ytd_fmt.get(fmt)
    ])
    fmt_df = df[df["PRODUCT_FORMAT"] == fmt]
    for var in st.session_state["variant"]:
        if var in fmt_df["PRODUCT_VARIANT_NAME"].values:
            rows.append([
                f"        {var}",
                cont_map_var.get(var),
                value_mtd_var.get(var),
                value_ytd_var.get(var),
                growth_mtd_var.get(var),
                growth_l3m_var.get(var),
                growth_ytd_var.get(var),
                ach_mtd_var.get(var),
                ach_ytd_var.get(var)
            ])
            var_df = fmt_df[fmt_df["PRODUCT_VARIANT_NAME"] == var]
            for prod in st.session_state["product"]:
                if prod in var_df["PRODUCT_NAME"].values:
                    rows.append([
                        f"            {prod}",
                        cont_map_prod.get(prod),
                        value_mtd_prod.get(prod),
                        value_ytd_prod.get(prod),
                        growth_mtd_prod.get(prod),
                        growth_l3m_prod.get(prod),
                        growth_ytd_prod.get(prod),
                        ach_mtd_prod.get(prod),
                        ach_ytd_prod.get(prod)
                    ])

# =============================
# DISPLAY & EXPORT (TIDAK DIUBAH)
# =============================
display_df = pd.DataFrame(rows, columns=[
    "Produk","Cont YTD","Value MTD","Value YTD",
    "Growth MTD","Growth %Gr L3M","Growth YTD","Ach MTD","Ach YTD"
])

def fmt_pct(x):
    return f"{x:.1f}%" if pd.notna(x) else ""

for c in ["Cont YTD","Growth MTD","Growth %Gr L3M","Growth YTD","Ach MTD","Ach YTD"]:
    display_df[c] = display_df[c].apply(fmt_pct)

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
