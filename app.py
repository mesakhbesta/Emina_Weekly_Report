import streamlit as st
import pandas as pd
from io import BytesIO
import datetime

st.set_page_config(layout="wide")
st.title("Dynamic Emina Metrics Report")


cutoff_date = st.sidebar.date_input("Select Cut-off Date", datetime.date.today())
cutoff_str = cutoff_date.strftime("%d %B %Y")  # ex: 18 Desember 2025

with st.sidebar.expander("Upload Excel Files (click to expand)", expanded=False):
    master_file = st.file_uploader("Master Product", type=["xlsx"])
    format_file = st.file_uploader("Format Metrics", type=["xlsx"])
    variant_file = st.file_uploader("Variant Metrics", type=["xlsx"])
    product_file = st.file_uploader("Product Metrics", type=["xlsx"])

if not all([master_file, format_file, variant_file, product_file]):
    st.warning("Please upload all 4 Excel files.")
    st.stop()

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
    tmp = pd.read_excel(file, sheet_name=sheet, skiprows=skip)
    result = {}
    for _, r in tmp.iterrows():
        v = r[val_col]
        if parser:
            v = parser(v)
        result[r[key_col]] = v
    return result

df = pd.read_excel(master_file)

cont_map_fmt = load_map("Sheet 18", "Product P", "% of Total Current DO TP2 along Product P, Product P Hidden", file=format_file, parser=parse_percent)
value_mtd_fmt = load_map("Sheet 1", "Product P", "Current DO", file=format_file, parser=parse_number)
value_ytd_fmt = load_map("Sheet 1", "Product P", "Current DO TP2", file=format_file, parser=parse_number)
growth_mtd_fmt = load_map("Sheet 4", "Product P", "vs LY", skip=1, file=format_file, parser=parse_percent)
growth_l3m_fmt = load_map("Sheet 3", "Product P", "vs L3M", skip=1, file=format_file, parser=parse_percent)
growth_ytd_fmt = load_map("Sheet 5", "Product P", "vs LY", skip=1, file=format_file, parser=parse_percent)
ach_mtd_fmt = load_map("Sheet 13", "Product P", "Current Achievement", file=format_file, parser=parse_percent)
ach_ytd_fmt = load_map("Sheet 14", "Product P", "Current Achievement TP2", file=format_file, parser=parse_percent)


cont_map_var = load_map("Sheet 18", "Product P", "% of Total Current DO TP2 along Product P, Product P Hidden", file=variant_file, parser=parse_percent)
value_mtd_var = load_map("Sheet 1", "Product P", "Current DO", file=variant_file, parser=parse_number)
value_ytd_var = load_map("Sheet 1", "Product P", "Current DO TP2", file=variant_file, parser=parse_number)
growth_mtd_var = load_map("Sheet 4", "Product P", "vs LY", skip=1, file=variant_file, parser=parse_percent)
growth_l3m_var = load_map("Sheet 3", "Product P", "vs L3M", skip=1, file=variant_file, parser=parse_percent)
growth_ytd_var = load_map("Sheet 5", "Product P", "vs LY", skip=1, file=variant_file, parser=parse_percent)
ach_mtd_var = load_map("Sheet 13", "Product P", "Current Achievement", file=variant_file, parser=parse_percent)
ach_ytd_var = load_map("Sheet 14", "Product P", "Current Achievement TP2", file=variant_file, parser=parse_percent)


cont_map_prod = load_map("Sheet 18", "Product P", "% of Total Current DO TP2 along Product P, Product P Hidden", file=product_file, parser=parse_percent)
value_mtd_prod = load_map("Sheet 1", "Product P", "Current DO", file=product_file, parser=parse_number)
value_ytd_prod = load_map("Sheet 1", "Product P", "Current DO TP2", file=product_file, parser=parse_number)
growth_mtd_prod = load_map("Sheet 4", "Product P", "vs LY", skip=1, file=product_file, parser=parse_percent)
growth_l3m_prod = load_map("Sheet 3", "Product P", "vs L3M", skip=1, file=product_file, parser=parse_percent)
growth_ytd_prod = load_map("Sheet 5", "Product P", "vs LY", skip=1, file=product_file, parser=parse_percent)
ach_mtd_prod = load_map("Sheet 13", "Product P", "Current Achievement", file=product_file, parser=parse_percent)
ach_ytd_prod = load_map("Sheet 14", "Product P", "Current Achievement TP2", file=product_file, parser=parse_percent)


for k in ["format", "variant", "product"]:
    if k not in st.session_state:
        st.session_state[k] = []

st.sidebar.title("Filter Products")

all_formats = sorted(df["PRODUCT_FORMAT"].dropna().unique())
st.session_state["format"] = st.sidebar.multiselect("Format", all_formats, default=st.session_state["format"])

variant_pool = df[df["PRODUCT_FORMAT"].isin(st.session_state["format"])]["PRODUCT_VARIANT_NAME"].dropna().unique()
variant_pool = sorted(variant_pool)
st.session_state["variant"] = [v for v in st.session_state["variant"] if v in variant_pool]
st.session_state["variant"] = st.sidebar.multiselect("Variant", variant_pool, default=st.session_state["variant"])

product_pool = df[df["PRODUCT_VARIANT_NAME"].isin(st.session_state["variant"])]["PRODUCT_NAME"].dropna().unique()
product_pool = sorted(product_pool)
st.session_state["product"] = [p for p in st.session_state["product"] if p in product_pool]
st.session_state["product"] = st.sidebar.multiselect("Product", product_pool, default=st.session_state["product"])

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

display_df = pd.DataFrame(rows, columns=[
    "Produk",
    "Cont YTD",
    "Value MTD",
    "Value YTD",
    "Growth MTD",
    "Growth %Gr L3M",
    "Growth YTD",
    "Ach MTD",
    "Ach YTD"
])

def fmt_pct(x):
    return f"{x:.1f}%" if pd.notna(x) else ""

display_df["Cont YTD"] = display_df["Cont YTD"].apply(fmt_pct)
display_df["Growth MTD"] = display_df["Growth MTD"].apply(fmt_pct)
display_df["Growth %Gr L3M"] = display_df["Growth %Gr L3M"].apply(fmt_pct)
display_df["Growth YTD"] = display_df["Growth YTD"].apply(fmt_pct)
display_df["Ach MTD"] = display_df["Ach MTD"].apply(fmt_pct)
display_df["Ach YTD"] = display_df["Ach YTD"].apply(fmt_pct)

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

output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    wb = writer.book
    ws = wb.add_worksheet("Report")
    writer.sheets["Report"] = ws

    header = wb.add_format({"bold": True, "align": "center", "border": 1})
    fmt_f = wb.add_format({"bold": True, "border": 1})
    fmt_v = wb.add_format({"align": "right", "indent": 2, "border": 1})
    fmt_p = wb.add_format({"align": "right", "indent": 4, "font_color": "blue", "border": 1})
    pct_green = wb.add_format({"num_format": "0.0%", "font_color": "green", "border": 1})
    pct_red = wb.add_format({"num_format": "0.0%", "font_color": "red", "border": 1})
    num = wb.add_format({"num_format": "#,##0", "border": 1})

    ws.write(0, 0, "Cut-off: " + cutoff_str, header)

    ws.write(1, 0, "Produk", header)
    ws.write(1, 1, "Cont YTD", header)
    ws.merge_range(1, 2, 1, 3, "Value", header)
    ws.merge_range(1, 4, 1, 6, "Growth", header)
    ws.merge_range(1, 7, 1, 8, "Ach", header)
    ws.write_row(2, 2, ["MTD", "YTD", "MTD", "%Gr L3M", "YTD", "MTD", "YTD"], header)

    r0 = 3
    for i, row in enumerate(rows):
        r = r0 + i
        if row[0].startswith("            "):
            ws.write(r, 0, row[0].strip(), fmt_p)
        elif row[0].startswith("        "):
            ws.write(r, 0, row[0].strip(), fmt_v)
        else:
            ws.write(r, 0, row[0], fmt_f)

        val = row[1] or 0
        if val >= 0:
            ws.write_number(r, 1, val/100, pct_green)
        else:
            ws.write_number(r, 1, val/100, pct_red)

        ws.write_number(r, 2, row[2] or 0, num)
        ws.write_number(r, 3, row[3] or 0, num)

        for col_idx, val in enumerate(row[4:7], start=4):
            if val is None:
                ws.write_blank(r, col_idx, None, num)
            elif val >= 0:
                ws.write_number(r, col_idx, val/100, pct_green)
            else:
                ws.write_number(r, col_idx, val/100, pct_red)

        for col_idx, val in enumerate(row[7:9], start=7):
            if val is None:
                ws.write_blank(r, col_idx, None, num)
            elif val >= 0:
                ws.write_number(r, col_idx, val/100, pct_green)
            else:
                ws.write_number(r, col_idx, val/100, pct_red)

    ws.set_column("A:A", 55)
    ws.set_column("B:I", 18)

output.seek(0)
st.download_button(
    "Download Excel",
    output,
    "Report_Full_Level.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)