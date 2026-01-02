import streamlit as st
import pandas as pd
from io import BytesIO
import datetime

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(
    layout="wide",
    page_title="Weekly Performance Report"
)

# =====================================================
# HEADER
# =====================================================
st.title("📊 Weekly Performance Report")
st.subheader("Format, Variant & Product Performance Overview")

st.caption("Dynamic performance monitoring across product hierarchy")
st.divider()

# =====================================================
# SIDEBAR – DATE & FILE UPLOAD
# =====================================================
st.sidebar.header("🗓️ Reporting Settings")

cutoff_date = st.sidebar.date_input(
    "Cut-off Date",
    datetime.date.today()
)
cutoff_str = cutoff_date.strftime("%d %B %Y")

st.sidebar.info(f"📌 Cut-off Date: **{cutoff_str}**")

st.sidebar.divider()

with st.sidebar.expander("📁 Upload Excel Files", expanded=False):
    master_file = st.file_uploader("Master Product", type=["xlsx"])
    format_file = st.file_uploader("Format Metrics", type=["xlsx"])
    variant_file = st.file_uploader("Variant Metrics", type=["xlsx"])
    product_file = st.file_uploader("Product Metrics", type=["xlsx"])

if not all([master_file, format_file, variant_file, product_file]):
    st.warning("⚠️ Please upload **all 4 required Excel files** to proceed.")
    st.stop()

# =====================================================
# CACHE & HELPER
# =====================================================
@st.cache_data(show_spinner=False)
def load_excel(file, sheet, skip=0):
    return pd.read_excel(file, sheet_name=sheet, skiprows=skip)

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

# =====================================================
# LOAD MASTER
# =====================================================
df = pd.read_excel(master_file)

# =====================================================
# LOAD METRICS
# =====================================================
def load_all(file):
    return dict(
        cont=load_map("Sheet 18", "Product P",
            "% of Total Current DO TP2 along Product P, Product P Hidden",
            file, parser=parse_percent),
        mtd=load_map("Sheet 1", "Product P", "Current DO", file, parser=parse_number),
        ytd=load_map("Sheet 1", "Product P", "Current DO TP2", file, parser=parse_number),
        g_mtd=load_map("Sheet 4", "Product P", "vs LY", file, skip=1, parser=parse_percent),
        g_l3m=load_map("Sheet 3", "Product P", "vs L3M", file, skip=1, parser=parse_percent),
        g_ytd=load_map("Sheet 5", "Product P", "vs LY", file, skip=1, parser=parse_percent),
        a_mtd=load_map("Sheet 13", "Product P", "Current Achievement", file, parser=parse_percent),
        a_ytd=load_map("Sheet 14", "Product P", "Current Achievement TP2", file, parser=parse_percent),
    )

fmt = load_all(format_file)
var = load_all(variant_file)
prd = load_all(product_file)

# =====================================================
# FILTER SECTION
# =====================================================
st.sidebar.header("🎯 Product Filters")

for k in ["format", "variant", "product"]:
    if k not in st.session_state:
        st.session_state[k] = []

# ---------- FORMAT ----------
formats = list(dict.fromkeys(df["PRODUCT_FORMAT"].dropna()))
st.session_state["format"] = st.sidebar.multiselect(
    "Format",
    formats,
    default=st.session_state["format"]
)

# ---------- VARIANT ----------
variants = list(dict.fromkeys(
    df[df["PRODUCT_FORMAT"].isin(st.session_state["format"])]
    ["PRODUCT_VARIANT_NAME"].dropna()
))
st.session_state["variant"] = st.sidebar.multiselect(
    "Variant",
    variants,
    default=[v for v in st.session_state["variant"] if v in variants]
)

# ---------- PRODUCT ----------
products = list(dict.fromkeys(
    df[df["PRODUCT_VARIANT_NAME"].isin(st.session_state["variant"])]
    ["PRODUCT_NAME"].dropna()
))
st.session_state["product"] = st.sidebar.multiselect(
    "Product",
    products,
    default=[p for p in st.session_state["product"] if p in products]
)

# =====================================================
# BUILD ROWS (LOGIC AS IS)
# =====================================================
rows = []

rows.append([
    "GRAND TOTAL",
    fmt["cont"].get("GRAND TOTAL"),
    fmt["mtd"].get("GRAND TOTAL"),
    fmt["ytd"].get("GRAND TOTAL"),
    fmt["g_mtd"].get("GRAND TOTAL"),
    fmt["g_l3m"].get("GRAND TOTAL"),
    fmt["g_ytd"].get("GRAND TOTAL"),
    fmt["a_mtd"].get("GRAND TOTAL"),
    fmt["a_ytd"].get("GRAND TOTAL"),
])

for f in st.session_state["format"]:
    rows.append([
        f,
        fmt["cont"].get(f),
        fmt["mtd"].get(f),
        fmt["ytd"].get(f),
        fmt["g_mtd"].get(f),
        fmt["g_l3m"].get(f),
        fmt["g_ytd"].get(f),
        fmt["a_mtd"].get(f),
        fmt["a_ytd"].get(f),
    ])

    for v in st.session_state["variant"]:
        if v in df[df["PRODUCT_FORMAT"] == f]["PRODUCT_VARIANT_NAME"].values:
            rows.append([
                f"        {v}",
                var["cont"].get(v),
                var["mtd"].get(v),
                var["ytd"].get(v),
                var["g_mtd"].get(v),
                var["g_l3m"].get(v),
                var["g_ytd"].get(v),
                var["a_mtd"].get(v),
                var["a_ytd"].get(v),
            ])

            for p in st.session_state["product"]:
                if p in df[df["PRODUCT_VARIANT_NAME"] == v]["PRODUCT_NAME"].values:
                    rows.append([
                        f"            {p}",
                        prd["cont"].get(p),
                        prd["mtd"].get(p),
                        prd["ytd"].get(p),
                        prd["g_mtd"].get(p),
                        prd["g_l3m"].get(p),
                        prd["g_ytd"].get(p),
                        prd["a_mtd"].get(p),
                        prd["a_ytd"].get(p),
                    ])

# =====================================================
# DISPLAY TABLE
# =====================================================
st.subheader("📈 Performance Detail Table")
st.caption(f"Data as of **{cutoff_str}**")

display_df = pd.DataFrame(rows, columns=[
    "Produk","Cont YTD","Value MTD","Value YTD",
    "Growth MTD","Growth %Gr L3M","Growth YTD","Ach MTD","Ach YTD"
])

def fmt_pct(x):
    return f"{x:.1f}%" if pd.notna(x) else ""

for c in ["Cont YTD","Growth MTD","Growth %Gr L3M","Growth YTD","Ach MTD","Ach YTD"]:
    display_df[c] = display_df[c].apply(fmt_pct)

display_df.columns = pd.MultiIndex.from_tuples([
    (f"Cut-off: {cutoff_str}", ""),
    ("","Cont YTD"),
    ("Value","MTD"),
    ("Value","YTD"),
    ("Growth","MTD"),
    ("Growth","%Gr L3M"),
    ("Growth","YTD"),
    ("Achievement","MTD"),
    ("Achievement","YTD"),
])

def highlight_product(row):
    styles = [""] * len(row)
    if row.iloc[0].startswith("            "):
        styles[0] = "color: #1f77b4"
    return styles

st.dataframe(
    display_df.style.apply(highlight_product, axis=1),
    use_container_width=True
)

# =====================================================
# DOWNLOAD SECTION
# =====================================================
st.divider()
st.subheader("⬇️ Export Report")

output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    wb = writer.book
    ws = wb.add_worksheet("Report")
    writer.sheets["Report"] = ws

    header = wb.add_format({"bold": True, "align": "center", "border": 1})
    bold = wb.add_format({"bold": True, "border": 1})
    ind1 = wb.add_format({"border": 1, "indent": 2})
    ind2 = wb.add_format({"border": 1, "indent": 4, "font_color": "blue"})
    num = wb.add_format({"border": 1, "num_format": "#,##0"})
    pct_g = wb.add_format({"border": 1, "num_format": "0.0%", "font_color": "green"})
    pct_r = wb.add_format({"border": 1, "num_format": "0.0%", "font_color": "red"})

    ws.write(0, 0, f"Cut-off: {cutoff_str}", header)
    ws.write_row(
        1, 0,
        ["Produk","Cont YTD","Value MTD","Value YTD","Growth MTD","%Gr L3M","Growth YTD","Ach MTD","Ach YTD"],
        header
    )

    for i, r in enumerate(rows, start=2):
        if r[0].startswith("            "):
            name_fmt = ind2
        elif r[0].startswith("        "):
            name_fmt = ind1
        else:
            name_fmt = bold

        ws.write(i, 0, r[0].strip(), name_fmt)

        for c in range(1, 9):
            v = r[c]
            if c == 1 or c >= 4:
                if v is not None:
                    ws.write_number(i, c, v / 100, pct_g if v >= 0 else pct_r)
            else:
                ws.write_number(i, c, v or 0, num)

    ws.set_column("A:A", 50)
    ws.set_column("B:I", 18)

output.seek(0)

st.download_button(
    "📥 Download Excel Report",
    output,
    "Weekly_Performance_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
