import streamlit as st
import pandas as pd
from io import BytesIO
import datetime

# =============================
# CONFIG
# =============================
st.set_page_config(layout="wide")
st.title("Dynamic Emina Metrics Report")

cutoff_date = st.sidebar.date_input(
    "Select Cut-off Date", datetime.date.today()
)
cutoff_str = cutoff_date.strftime("%d %B %Y")

# =============================
# UPLOAD FILES
# =============================
with st.sidebar.expander("Upload Excel Files", expanded=False):
    master_file = st.file_uploader("Master Product", type=["xlsx"])
    format_file = st.file_uploader("Format Metrics", type=["xlsx"])
    variant_file = st.file_uploader("Variant Metrics", type=["xlsx"])
    product_file = st.file_uploader("Product Metrics", type=["xlsx"])

if not all([master_file, format_file, variant_file, product_file]):
    st.warning("Please upload all 4 Excel files.")
    st.stop()

# =============================
# HELPERS
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
    df = pd.read_excel(file, sheet_name=sheet, skiprows=skip)
    out = {}
    for _, r in df.iterrows():
        v = r[val_col]
        if parser:
            v = parser(v)
        out[r[key_col]] = v
    return out

# =============================
# LOAD MASTER
# =============================
df = pd.read_excel(master_file)

# =============================
# LOAD METRICS
# =============================
cont_fmt = load_map("Sheet 18", "Product P",
    "% of Total Current DO TP2 along Product P, Product P Hidden",
    format_file, parser=parse_percent)

val_mtd_fmt = load_map("Sheet 1", "Product P", "Current DO", format_file, parser=parse_number)
val_ytd_fmt = load_map("Sheet 1", "Product P", "Current DO TP2", format_file, parser=parse_number)
gr_mtd_fmt = load_map("Sheet 4", "Product P", "vs LY", format_file, skip=1, parser=parse_percent)
gr_l3m_fmt = load_map("Sheet 3", "Product P", "vs L3M", format_file, skip=1, parser=parse_percent)
gr_ytd_fmt = load_map("Sheet 5", "Product P", "vs LY", format_file, skip=1, parser=parse_percent)
ach_mtd_fmt = load_map("Sheet 13", "Product P", "Current Achievement", format_file, parser=parse_percent)
ach_ytd_fmt = load_map("Sheet 14", "Product P", "Current Achievement TP2", format_file, parser=parse_percent)

cont_var = load_map("Sheet 18", "Product P",
    "% of Total Current DO TP2 along Product P, Product P Hidden",
    variant_file, parser=parse_percent)

val_mtd_var = load_map("Sheet 1", "Product P", "Current DO", variant_file, parser=parse_number)
val_ytd_var = load_map("Sheet 1", "Product P", "Current DO TP2", variant_file, parser=parse_number)
gr_mtd_var = load_map("Sheet 4", "Product P", "vs LY", variant_file, skip=1, parser=parse_percent)
gr_l3m_var = load_map("Sheet 3", "Product P", "vs L3M", variant_file, skip=1, parser=parse_percent)
gr_ytd_var = load_map("Sheet 5", "Product P", "vs LY", variant_file, skip=1, parser=parse_percent)
ach_mtd_var = load_map("Sheet 13", "Product P", "Current Achievement", variant_file, parser=parse_percent)
ach_ytd_var = load_map("Sheet 14", "Product P", "Current Achievement TP2", variant_file, parser=parse_percent)

cont_prod = load_map("Sheet 18", "Product P",
    "% of Total Current DO TP2 along Product P, Product P Hidden",
    product_file, parser=parse_percent)

val_mtd_prod = load_map("Sheet 1", "Product P", "Current DO", product_file, parser=parse_number)
val_ytd_prod = load_map("Sheet 1", "Product P", "Current DO TP2", product_file, parser=parse_number)
gr_mtd_prod = load_map("Sheet 4", "Product P", "vs LY", product_file, skip=1, parser=parse_percent)
gr_l3m_prod = load_map("Sheet 3", "Product P", "vs L3M", product_file, skip=1, parser=parse_percent)
gr_ytd_prod = load_map("Sheet 5", "Product P", "vs LY", product_file, skip=1, parser=parse_percent)
ach_mtd_prod = load_map("Sheet 13", "Product P", "Current Achievement", product_file, parser=parse_percent)
ach_ytd_prod = load_map("Sheet 14", "Product P", "Current Achievement TP2", product_file, parser=parse_percent)

# =============================
# FILTER (ANTI DOUBLE CLICK)
# =============================
st.sidebar.title("Filter Products")

formats = sorted(df["PRODUCT_FORMAT"].dropna().unique())
fmt_ids = {i: v for i, v in enumerate(formats)}

fmt_sel = st.sidebar.multiselect(
    "Format",
    list(fmt_ids.keys()),
    format_func=lambda x: fmt_ids[x],
    key="fmt_sel"
)

selected_formats = [fmt_ids[i] for i in fmt_sel]

variants = sorted(
    df[df["PRODUCT_FORMAT"].isin(selected_formats)]
    ["PRODUCT_VARIANT_NAME"].dropna().unique()
)
var_ids = {i: v for i, v in enumerate(variants)}

var_sel = st.sidebar.multiselect(
    "Variant",
    list(var_ids.keys()),
    format_func=lambda x: var_ids[x],
    key="var_sel"
)

selected_variants = [var_ids[i] for i in var_sel]

products = sorted(
    df[df["PRODUCT_VARIANT_NAME"].isin(selected_variants)]
    ["PRODUCT_NAME"].dropna().unique()
)
prd_ids = {i: v for i, v in enumerate(products)}

prd_sel = st.sidebar.multiselect(
    "Product",
    list(prd_ids.keys()),
    format_func=lambda x: prd_ids[x],
    key="prd_sel"
)

selected_products = [prd_ids[i] for i in prd_sel]

# =============================
# BUILD TABLE
# =============================
rows = []

rows.append([
    "GRAND TOTAL",
    cont_fmt.get("GRAND TOTAL"),
    val_mtd_fmt.get("GRAND TOTAL"),
    val_ytd_fmt.get("GRAND TOTAL"),
    gr_mtd_fmt.get("GRAND TOTAL"),
    gr_l3m_fmt.get("GRAND TOTAL"),
    gr_ytd_fmt.get("GRAND TOTAL"),
    ach_mtd_fmt.get("GRAND TOTAL"),
    ach_ytd_fmt.get("GRAND TOTAL"),
])

for f in selected_formats:
    rows.append([
        f,
        cont_fmt.get(f),
        val_mtd_fmt.get(f),
        val_ytd_fmt.get(f),
        gr_mtd_fmt.get(f),
        gr_l3m_fmt.get(f),
        gr_ytd_fmt.get(f),
        ach_mtd_fmt.get(f),
        ach_ytd_fmt.get(f),
    ])

    for v in selected_variants:
        if v in df[df["PRODUCT_FORMAT"] == f]["PRODUCT_VARIANT_NAME"].values:
            rows.append([
                f"        {v}",
                cont_var.get(v),
                val_mtd_var.get(v),
                val_ytd_var.get(v),
                gr_mtd_var.get(v),
                gr_l3m_var.get(v),
                gr_ytd_var.get(v),
                ach_mtd_var.get(v),
                ach_ytd_var.get(v),
            ])

            for p in selected_products:
                if p in df[df["PRODUCT_VARIANT_NAME"] == v]["PRODUCT_NAME"].values:
                    rows.append([
                        f"            {p}",
                        cont_prod.get(p),
                        val_mtd_prod.get(p),
                        val_ytd_prod.get(p),
                        gr_mtd_prod.get(p),
                        gr_l3m_prod.get(p),
                        gr_ytd_prod.get(p),
                        ach_mtd_prod.get(p),
                        ach_ytd_prod.get(p),
                    ])

display_df = pd.DataFrame(rows, columns=[
    "Produk","Cont YTD","Value MTD","Value YTD",
    "Growth MTD","Growth %Gr L3M","Growth YTD",
    "Ach MTD","Ach YTD"
])

def fmt_pct(x):
    return f"{x:.1f}%" if pd.notna(x) else ""

for c in ["Cont YTD","Growth MTD","Growth %Gr L3M","Growth YTD","Ach MTD","Ach YTD"]:
    display_df[c] = display_df[c].apply(fmt_pct)

# =============================
# STREAMLIT STYLING (ONLY NAME BLUE)
# =============================
def highlight_product(row):
    styles = [""] * len(row)
    if row.iloc[0].startswith("            "):
        styles[0] = "color: blue"
    return styles

st.dataframe(
    display_df.style.apply(highlight_product, axis=1),
    use_container_width=True
)

# =============================
# EXPORT EXCEL
# =============================
output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    wb = writer.book
    ws = wb.add_worksheet("Report")
    writer.sheets["Report"] = ws

    header = wb.add_format({"bold": True, "border": 1, "align": "center"})
    bold = wb.add_format({"bold": True, "border": 1})
    ind1 = wb.add_format({"border": 1, "indent": 2})
    ind2 = wb.add_format({"border": 1, "indent": 4, "font_color": "blue"})
    num = wb.add_format({"border": 1, "num_format": "#,##0"})
    pct_g = wb.add_format({"border": 1, "num_format": "0.0%", "font_color": "green"})
    pct_r = wb.add_format({"border": 1, "num_format": "0.0%", "font_color": "red"})

    ws.write(0, 0, "Cut-off: " + cutoff_str, header)

    ws.write_row(1, 0, [
        "Produk","Cont YTD","Value MTD","Value YTD",
        "Growth MTD","%Gr L3M","Growth YTD","Ach MTD","Ach YTD"
    ], header)

    for i, r in enumerate(rows, start=2):
        if r[0].startswith("            "):
            name_fmt = ind2
        elif r[0].startswith("        "):
            name_fmt = ind1
        else:
            name_fmt = bold

        ws.write(i, 0, r[0].strip(), name_fmt)

        ws.write_number(i, 1, (r[1] or 0)/100, pct_g if (r[1] or 0) >= 0 else pct_r)
        ws.write_number(i, 2, r[2] or 0, num)
        ws.write_number(i, 3, r[3] or 0, num)

        for j, v in enumerate(r[4:7], start=4):
            if v is not None:
                ws.write_number(i, j, v/100, pct_g if v >= 0 else pct_r)

        for j, v in enumerate(r[7:9], start=7):
            if v is not None:
                ws.write_number(i, j, v/100, pct_g if v >= 0 else pct_r)

    ws.set_column("A:A", 55)
    ws.set_column("B:I", 18)

output.seek(0)
st.download_button(
    "Download Excel",
    output,
    "Report_Full_Level.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
