import io
import re
from typing import List

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# ------------------------------
# App config
# ------------------------------
st.set_page_config(page_title="PO Generator – Manual Input", page_icon="📝", layout="wide")
st.title("📝 Purchase Order Generator — Manual Input")
st.caption("Enter Control NO, Item NO, JAN/Barcode, Qty, Price, Delivery (no spreadsheet needed). The app fills your template — green areas stay static; blue fields are filled.")

# Fixed cell mapping in the template (blue fields)
CELL = {
    "control_no": "AD9",
    "item_no": "E16",
    "barcode": "S16",
    "delivery": "B28",
    "qty": "AA24",
    "price1": "N30",
    "price2": "N32",
    "amount": "F37",
}

# Cells to clear and rows to delete as per your instructions
CELLS_TO_CLEAR = ["E18", "E20", "E24", "N24", "B26", "A35", "R37", "F39"]
DELETE_ROWS = (60, 5)  # start row, count

INVALID_SHEET_CHARS = r"[:\\\\/?*\[\]]"

def sanitize_sheet_title(title: str) -> str:
    title = re.sub(INVALID_SHEET_CHARS, "-", str(title)).strip() or "Sheet"
    return title[:31]

# ------------------------------
# Sidebar: template
# ------------------------------
with st.sidebar:
    st.header("Template")
    tpl_file = st.file_uploader("Upload TEMPLATE .xlsx", type=["xlsx"], help="Use your green/blue colored template")
    remove_template_sheet = st.checkbox("Remove original template sheet in output", value=True)

# ------------------------------
# Manual input table
# ------------------------------
st.subheader("Enter PO lines")

if "rows" not in st.session_state:
    st.session_state.rows = [
        {"Control NO": "", "Item NO": "", "Barcode": "", "Qty": "", "Price": "", "Delivery": ""}
    ]

cols = st.columns([1,1,1,1,1,1,0.6])
with cols[0]: ctrl = st.text_input("Control NO", value=st.session_state.rows[-1]["Control NO"]) 
with cols[1]: item = st.text_input("Item NO", value=st.session_state.rows[-1]["Item NO"]) 
with cols[2]: jan  = st.text_input("JAN / Barcode", value=st.session_state.rows[-1]["Barcode"]) 
with cols[3]: qty  = st.text_input("Qty", value=str(st.session_state.rows[-1]["Qty"]))
with cols[4]: price= st.text_input("Price (e.g. 60.600)", value=str(st.session_state.rows[-1]["Price"]))
with cols[5]: delv = st.text_input("Delivery", value=st.session_state.rows[-1]["Delivery"]) 
add = cols[6].button("➕ Add line")

if add:
    st.session_state.rows.append({
        "Control NO": ctrl, "Item NO": item, "Barcode": jan,
        "Qty": qty, "Price": price, "Delivery": delv
    })

# Show current lines
if st.session_state.rows:
    df = pd.DataFrame(st.session_state.rows)
    st.dataframe(df, use_container_width=True)

# Utility to parse numbers with commas/yen and 3 decimals

def to_float(x):
    if x is None: return None
    s = str(x).strip()
    if not s: return None
    s = s.replace(",", "").replace("￥", "").replace("¥", "")
    try:
        return float(s)
    except Exception:
        try:
            return float(s.replace(" ", ""))
        except Exception:
            return None

def to_int(x):
    f = to_float(x)
    return int(round(f)) if f is not None else None

st.markdown("---")
make_btn = st.button("📦 Generate PurchaseOrders.xlsx")

if make_btn:
    # Validations
    if not tpl_file:
        st.error("Please upload your TEMPLATE .xlsx in the sidebar.")
        st.stop()
    # Filter out empty rows (need Control NO and Item NO at minimum)
    entries: List[dict] = []
    for r in st.session_state.rows:
        if str(r.get("Control NO", "")).strip() and str(r.get("Item NO", "")).strip():
            entries.append(r)
    if not entries:
        st.error("Please add at least one line with Control NO and Item NO.")
        st.stop()

    # Load template workbook
    try:
        wb = load_workbook(tpl_file, data_only=False, keep_vba=False)
    except Exception as e:
        st.error(f"Failed to load template: {e}")
        st.stop()

    tpl_ws = wb.active
    created = []

    for r in entries:
        control_no = str(r.get("Control NO", "")).strip()
        item_no    = str(r.get("Item NO", "")).strip()
        barcode    = str(r.get("Barcode", "")).strip()
        qty        = to_int(r.get("Qty"))
        price      = to_float(r.get("Price"))
        delivery   = str(r.get("Delivery", "")).strip()

        ws = wb.copy_worksheet(tpl_ws)
        ws.title = sanitize_sheet_title(f"{control_no}_{item_no}")

        # Fill dynamic cells (BLUE)
        try: ws[CELL["control_no"]] = control_no
        except: pass
        try: ws[CELL["item_no"]] = item_no
        except: pass
        try: ws[CELL["barcode"]] = barcode
        except: pass
        try: ws[CELL["delivery"]] = delivery
        except: pass
        try: ws[CELL["qty"]] = qty
        except: pass
        try: ws[CELL["price1"]] = price
        except: pass
        try: ws[CELL["price2"]] = price
        except: pass
        try:
            amount = (qty or 0) * (price or 0)
            ws[CELL["amount"]] = amount
        except: pass

        # Clear non-blue cells requested
        for c in CELLS_TO_CLEAR:
            try: ws[c] = None
            except: pass

        # Delete extra rows
        try: ws.delete_rows(*DELETE_ROWS)
        except: pass

        created.append(ws.title)

    # Remove template sheet if asked
    if remove_template_sheet:
        try: wb.remove(tpl_ws)
        except: pass

    # Return the workbook
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

st.markdown("---")
st.markdown(
    """
**Template mapping (blue fields):**  
- Control NO → `AD9`  
- Item No → `E16`  
- JAN / Barcode → `S16`  
- Delivery time/date → `B28`  
- Qty → `AA24`  
- Price → `N30` and `N32`  
- Amount → `F37` (calculated)

**Other actions:** Clears `E18, E20, E24, N24, B26, A35, R37, F39` and deletes rows `60–64`.  
**Sheet name:** `ControlNo_ItemNo`.  
**Tip:** Enter numbers with or without commas/yen (e.g. `6,000`, `¥60.600`). The app will parse them correctly.
"""
)
