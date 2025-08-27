import io
import re
import sys
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="PO Generator", page_icon="📄", layout="wide")
st.title("📄 Purchase Order Generator")
st.caption("Upload your input file and template to generate purchase orders. Each Control No + Item No becomes a new sheet.")

# --- Sidebar inputs ---
with st.sidebar:
    st.header("Inputs")
    xlsm_file = st.file_uploader("Upload input .xlsm file", type=["xlsm", "xlsx"])  # allow xlsx too
    template_file = st.file_uploader("Upload template .xlsx (purchase order format)", type=["xlsx"])  # openpyxl best with .xlsx
    input_sheetname = st.text_input("Input sheet name", value="250826")
    template_sheetname = st.text_input("Template sheet name (blank = active)", value="")
    start_row = st.number_input("Table start row in template", min_value=1, value=9, step=1)
    remove_template_sheet = st.checkbox("Remove the original template sheet from output", value=True)
    generate_btn = st.button("Generate Purchase Orders")

# --- Helpers ---

def sanitize_sheet_title(title: str) -> str:
    invalid = r"[:\\\\/?*\[\]]"
    title = re.sub(invalid, "-", str(title))
    title = title.strip() or "Sheet"
    return title[:31]


def auto_width(ws: Worksheet, cols: List[int]):
    for col_idx in cols:
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for cell in ws[col_letter]:
            v = cell.value
            if v is None:
                continue
            v = str(v)
            if len(v) > max_len:
                max_len = len(v)
        ws.column_dimensions[col_letter].width = max(10, min(max_len + 2, 60))


if generate_btn:
    if not xlsm_file or not template_file:
        st.error("Please upload both the input .xlsm and the template .xlsx.")
        st.stop()

    # --- Load input sheet ---
    try:
        df = pd.read_excel(xlsm_file, sheet_name=input_sheetname, engine="openpyxl")
    except Exception as e:
        st.error(f"Failed to read input sheet '{input_sheetname}': {e}")
        st.stop()

    if df.empty:
        st.error("The input sheet appears to be empty.")
        st.stop()

    st.subheader("Step 1: Map your columns")
    st.write("Select which columns in your input correspond to the required fields.")

    mapping = {}
    for field in ["control_no", "item_no", "description", "qty", "unit", "price", "amount"]:
        mapping[field] = st.selectbox(
            f"Select column for {field.replace('_',' ').title()}",
            options=[None] + list(df.columns),
            index=0
        )

    # Required fields check
    if not mapping["control_no"] or not mapping["item_no"] or not mapping["qty"]:
        st.error("Please select at least Control No, Item No, and Qty.")
        st.stop()

    # --- Load template workbook ---
    try:
        wb = load_workbook(template_file, data_only=False, keep_vba=False)
    except Exception as e:
        st.error(f"Failed to load template workbook: {e}")
        st.stop()

    if template_sheetname and template_sheetname in wb.sheetnames:
        tpl_ws = wb[template_sheetname]
    else:
        tpl_ws = wb.active

    created_sheets = []

    try:
        grouped = df.groupby([mapping["control_no"], mapping["item_no"]], dropna=False)
    except Exception as e:
        st.error(f"Failed to group by Control No and Item No: {e}")
        st.stop()

    for (control_no, item_no), g in grouped:
        ws = wb.copy_worksheet(tpl_ws)
        ws.title = sanitize_sheet_title(f"{control_no}_{item_no}")
        created_sheets.append(ws.title)

        row_ptr = int(start_row)
        for _, r in g.iterrows():
            ws.cell(row=row_ptr, column=1, value=r.get(mapping["item_no"], ""))
            ws.cell(row=row_ptr, column=2, value=r.get(mapping["description"], ""))
            ws.cell(row=row_ptr, column=3, value=r.get(mapping["qty"], None))
            ws.cell(row=row_ptr, column=4, value=r.get(mapping["unit"], ""))
            ws.cell(row=row_ptr, column=5, value=r.get(mapping["price"], None))
            ws.cell(row=row_ptr, column=6, value=r.get(mapping["amount"], None))
            row_ptr += 1

        auto_width(ws, [1, 2, 3, 4, 5, 6])

    if remove_template_sheet:
        try:
            wb.remove(tpl_ws)
        except Exception:
            pass

    output_buf = io.BytesIO()
    wb.save(output_buf)
    output_buf.seek(0)

    st.success(f"Created {len(created_sheets)} purchase order sheet(s).")
    st.download_button(
        label="⬇️ Download PurchaseOrders.xlsx",
        data=output_buf,
        file_name="PurchaseOrders.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("Sheets created"):
        st.write(created_sheets)
