import io
import re
from typing import Dict, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ----------------------------------
# Streamlit UI
# ----------------------------------
st.set_page_config(page_title="PO Generator (Fixed Mapping)", page_icon="📄", layout="wide")
st.title("📄 Purchase Order Generator — Fixed Mapping")
st.caption("Creates one .xlsx with one sheet per (Control NO, Item NO) using your template. Green areas stay static; blue fields are filled from your input.")

with st.sidebar:
    st.header("Inputs")
    input_file = st.file_uploader("Upload INPUT Excel (.xlsm/.xlsx)", type=["xlsm", "xlsx"])
    sheet_name = st.text_input("Input sheet name", value="250826")  # you can change to 250729
    template_file = st.file_uploader("Upload TEMPLATE .xlsx", type=["xlsx"], help="Use the colored template you provided")
    remove_template_sheet = st.checkbox("Remove the original template sheet from output", value=True)
    btn = st.button("Generate Purchase Orders")

# ----------------------------------
# Helpers
# ----------------------------------
INVALID_SHEET_CHARS = r"[:\\\\/?*\[\]]"

def sanitize_sheet_title(title: str) -> str:
    title = re.sub(INVALID_SHEET_CHARS, "-", str(title)).strip() or "Sheet"
    return title[:31]

# Try to resolve column names flexibly but prefer your exact headers
PREFERRED = {
    "control_no": ["Control NO", "CONTROL NO", "Control No", "control no", "ControlNO", "Ctrl No"],
    "item_no": ["Item NO", "ITEM NO", "Item No", "item no", "Item code", "Item code ", "品番", "品番 / Item no"],
    "barcode": ["Barcode", "JAN", "JAN code", "JAN Code", "JANコード"],
    "qty": ["Qty", "QTY", "Quantity", "数量"],
    "price": ["Price", "単価", "Unit Price", "Unit price"],
    "delivery": ["Delivery time", "Delivery", "Delivery date", "納期"],
}

def find_col(df: pd.DataFrame, candidates) -> str:
    cols = list(df.columns)
    # exact first
    for cand in candidates:
        for c in cols:
            if c.strip() == cand:
                return c
    # case-insensitive exact
    for cand in candidates:
        for c in cols:
            if c.strip().lower() == cand.strip().lower():
                return c
    # contains
    for cand in candidates:
        low = cand.strip().lower()
        for c in cols:
            if low in c.strip().lower():
                return c
    return None

@st.cache_data(show_spinner=False)
def load_input(_file, _sheet) -> pd.DataFrame:
    return pd.read_excel(_file, sheet_name=_sheet, engine="openpyxl")

# ----------------------------------
# Main action
# ----------------------------------
if btn:
    if not input_file or not template_file:
        st.error("Please upload both the INPUT workbook and the TEMPLATE workbook.")
        st.stop()

    # Load input
    try:
        df = load_input(input_file, sheet_name)
    except Exception as e:
        st.error(f"Could not read sheet '{sheet_name}': {e}")
        st.stop()

    if df.empty:
        st.error("The input sheet appears to be empty.")
        st.stop()

    # Column resolution
    cols = {}
    for key in PREFERRED:
        cols[key] = find_col(df, PREFERRED[key])
    required = ["control_no", "item_no", "barcode", "qty", "price", "delivery"]
    missing = [k for k in required if not cols.get(k)]
    if missing:
        st.error("Missing required columns in input: " + ", ".join(missing))
        st.write("Detected columns:", cols)
        st.stop()

    # Deduplicate: first row per (Control NO, Item NO)
    df_sorted = df.copy()
    df_sorted["__group_key__"] = (
        df_sorted[cols["control_no"]].astype(str).str.strip() + "\u0001" +
        df_sorted[cols["item_no"]].astype(str).str.strip()
    )
    first_rows = df_sorted.drop_duplicates("__group_key__", keep="first").reset_index(drop=True)

    # Parse numerics safely (strip commas, currency)
    def to_float(x):
        if pd.isna(x):
            return None
        s = str(x)
        s = s.replace(",", "")
        s = s.replace("￥", "").replace("¥", "")
        s = s.strip()
        try:
            return float(s)
        except:
            # handle 60.600 etc.
            try:
                return float(s.replace(" ", ""))
            except:
                return None

    def to_int(x):
        f = to_float(x)
        return int(round(f)) if f is not None else None

    # Load template workbook
    try:
        wb = load_workbook(template_file, data_only=False, keep_vba=False)
    except Exception as e:
        st.error(f"Failed to open template: {e}")
        st.stop()

    tpl_ws = wb.active
    created = []

    for _, r in first_rows.iterrows():
        control_no = str(r[cols["control_no"]]).strip()
        item_no    = str(r[cols["item_no"]]).strip()
        barcode    = str(r[cols["barcode"]]).strip()
        qty        = to_int(r[cols["qty"]])
        price      = to_float(r[cols["price"]])
        delivery   = str(r[cols["delivery"]]).strip()

        ws = wb.copy_worksheet(tpl_ws)
        ws.title = sanitize_sheet_title(f"{control_no}_{item_no}")

        # ---------------------------
        # Fill BLUE dynamic cells
        # ---------------------------
        try: ws["AD9"] = control_no
        except: pass
        try: ws["E16"] = item_no
        except: pass
        try: ws["S16"] = barcode
        except: pass
        try: ws["B28"] = delivery
        except: pass
        try: ws["AA24"] = qty
        except: pass
        try: ws["N30"] = price
        except: pass
        try: ws["N32"] = price
        except: pass
        # Amount = Qty * Price → F37
        try:
            amount = (qty or 0) * (price or 0)
            ws["F37"] = amount
        except:
            pass

        # ---------------------------
        # Clear requested cells
        # ---------------------------
        for cell in ["E18", "E20", "E24", "N24", "B26", "A35", "R37", "F39"]:
            try:
                ws[cell] = None
            except:
                pass

        # ---------------------------
        # Delete rows 60–64
        # ---------------------------
        try:
            ws.delete_rows(60, 5)
        except:
            pass

        created.append(ws.title)

    # Remove the original template sheet if requested
    if remove_template_sheet:
        try: wb.remove(tpl_ws)
        except: pass

    # Save to buffer and offer download
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    st.success(f"Created {len(created)} sheet(s)")
    st.download_button(
        "⬇️ Download PurchaseOrders.xlsx",
        buf,
        file_name="PurchaseOrders.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("Sheets created"):
        st.write(created)

st.markdown("---")
st.markdown(
    """
**Fixed cell mapping (blue fields):**
- Control NO → `AD9`
- Item No → `E16`
- JAN / Barcode → `S16`
- Delivery time/date → `B28`
- Qty → `AA24`
- Price → `N30` and `N32`
- Amount → `F37` (calculated as Qty × Price)

**Cleanup applied to each sheet:**
- Clears cells: `E18, E20, E24, N24, B26, A35, R37, F39`
- Deletes rows `60–64`

**Notes:**
- One sheet per *(Control NO, Item NO)*. If duplicates exist, only the **first row** for each pair is used.
- Green boxes in the template remain untouched.
- Number formats (e.g., ¥, thousands separators, decimals) are preserved from your template.
"""
)

