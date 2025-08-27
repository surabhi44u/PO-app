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
st.caption("Create a single .xlsx file with one sheet per Control No + Item No using your template, preserving headers/footers/logos.")

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
    # Excel sheet title constraints: max 31 chars, no : \\ / ? * [ ]
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
        # add padding
        ws.column_dimensions[col_letter].width = max(10, min(max_len + 2, 60))


def infer_columns(df: pd.DataFrame) -> Dict[str, str]:
    """Infer mapping from input columns to canonical keys used by the template table.
    Returns a dict mapping canonical -> input column name.
    Canonical keys: control_no, item_no, description, qty, unit, price, amount
    """
    cols = {c: str(c).strip().lower() for c in df.columns}

    def find(*keywords):
        for col, low in cols.items():
            if all(k in low for k in keywords):
                return col
        # try startswith any of keywords (looser)
        for col, low in cols.items():
            if any(low.startswith(k) for k in keywords):
                return col
        return None

    mapping = {}
    mapping["control_no"] = find("control", "no") or find("control") or find("ctrl")
    mapping["item_no"] = find("item", "no") or find("item code") or find("item")
    mapping["description"] = find("description") or find("desc") or find("spec")
    mapping["qty"] = find("qty") or find("quantity") or find("order qty")
    mapping["unit"] = find("unit") or find("uom") or find("unit of measure")
    mapping["price"] = find("price") or find("rate") or find("unit price")
    mapping["amount"] = find("amount") or find("total") or find("line total")

    return mapping


def validate_mapping(mapping: Dict[str, str]):
    required = ["control_no", "item_no", "description", "qty"]
    missing = [k for k in required if not mapping.get(k)]
    if missing:
        raise ValueError(f"Could not infer columns for: {', '.join(missing)}. Please rename your input headers or pre-clean your file.")


def group_key(row: pd.Series, mapping: Dict[str, str]) -> Tuple[str, str]:
    c = str(row[mapping["control_no"]]).strip()
    i = str(row[mapping["item_no"]]).strip()
    return c, i


def ensure_numeric(s):
    return pd.to_numeric(s, errors="coerce")


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

    # Infer & validate mapping
    mapping = infer_columns(df)
    try:
        validate_mapping(mapping)
    except ValueError as e:
        st.error(str(e))
        st.stop()

    # Clean & compute fields
    df[mapping["qty"]] = ensure_numeric(df[mapping["qty"]])
    if mapping.get("price"):
        df[mapping["price"]] = ensure_numeric(df[mapping["price"]])
    if mapping.get("amount") is None and mapping.get("price"):
        df["__amount__"] = df[mapping["qty"]] * df[mapping["price"]]
        amount_col = "__amount__"
    else:
        amount_col = mapping.get("amount")

    # --- Load template workbook (base for output) ---
    try:
        wb = load_workbook(template_file, data_only=False, keep_vba=False)
    except Exception as e:
        st.error(f"Failed to load template workbook: {e}")
        st.stop()

    # Choose template sheet
    if template_sheetname and template_sheetname in wb.sheetnames:
        tpl_ws = wb[template_sheetname]
    else:
        tpl_ws = wb.active

    # We'll create all copies within the same workbook to preserve styles/images
    created_sheets = []

    # Group by control + item
    try:
        grouped = df.groupby([mapping["control_no"], mapping["item_no"]], dropna=False)
    except Exception as e:
        st.error(f"Failed to group by Control No and Item No: {e}")
        st.stop()

    # For each group, duplicate template, rename and fill rows starting at start_row
    for (control_no, item_no), g in grouped:
        # Duplicate the template sheet within the same workbook
        ws = wb.copy_worksheet(tpl_ws)
        # Rename per rule
        ws.title = sanitize_sheet_title(f"{control_no}_{item_no}")
        created_sheets.append(ws.title)

        # Fill table rows
        row_ptr = int(start_row)
        for _, r in g.iterrows():
            # Safe get values
            val_item = r.get(mapping.get("item_no"), "")
            val_desc = r.get(mapping.get("description"), "")
            val_qty = r.get(mapping.get("qty"), None)
            val_unit = r.get(mapping.get("unit"), "") if mapping.get("unit") else ""
            val_price = r.get(mapping.get("price"), None) if mapping.get("price") else None
            val_amount = r.get(amount_col, None) if amount_col else None

            # Write values into first 6 columns (A..F). Adjust if your template differs.
            ws.cell(row=row_ptr, column=1, value=val_item)
            ws.cell(row=row_ptr, column=2, value=val_desc)
            ws.cell(row=row_ptr, column=3, value=val_qty)
            ws.cell(row=row_ptr, column=4, value=val_unit)
            ws.cell(row=row_ptr, column=5, value=val_price)
            ws.cell(row=row_ptr, column=6, value=val_amount)
            row_ptr += 1

        # Auto-width the first 6 columns by default
        auto_width(ws, [1, 2, 3, 4, 5, 6])

    # Optionally remove the original template sheet
    if remove_template_sheet:
        try:
            wb.remove(tpl_ws)
        except Exception:
            pass

    # Save workbook to bytes and offer download
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

    with st.expander("Debug: Column Mapping Inferred"):
        st.json(mapping)

    with st.expander("Sheets created"):
        st.write(created_sheets)

st.markdown("---")
st.markdown(
    """
**How it works**
1. Upload your `.xlsm` input and the `.xlsx` purchase order template.
2. The app infers columns (Control No, Item No, Description, Qty, Unit, Price, Amount).
3. It duplicates the template sheet once per (Control No, Item No) and fills rows starting at the chosen start row.
4. Logos, headers, and footers from your template are preserved because sheets are duplicated within the same workbook.

> Tip: If your headers don't match common names, rename them to something recognizable like `Control No`, `Item No`, `Description`, `Qty`, `Unit`, `Price`, `Amount`.
"""
)
