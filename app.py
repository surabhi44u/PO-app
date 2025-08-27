import io
import re
from typing import List

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill, NamedStyle, numbers
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="PO Generator – Manual Input (Per-Item)", page_icon="📝", layout="wide")
st.title("📝 Purchase Order Generator — Manual Input (Per Item, No Template)")
st.caption("No uploads. Enter lines manually. The app creates one sheet per (Control NO, Item NO) with a clean PO layout.")

# ------------------------------
# Helpers
# ------------------------------
INVALID_SHEET_CHARS = r"[:\\\\/?*\[\]]"

def sanitize_sheet_title(title: str) -> str:
    title = re.sub(INVALID_SHEET_CHARS, "-", str(title)).strip() or "Sheet"
    return title[:31]

thin = Side(style="thin")
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
header_fill = PatternFill(start_color="FFEDF2FF", end_color="FFEDF2FF", fill_type="solid")
colhdr_fill = PatternFill(start_color="FFECECEC", end_color="FFECECEC", fill_type="solid")

num_currency = numbers.FORMAT_CURRENCY_USD_SIMPLE.replace("$", "¥")  # displays like ¥#,##0.00
num_qty = "#,##0"
num_price3 = "#,##0.000"  # for cases like 60.600

# ------------------------------
# Manual input UI
# ------------------------------
if "rows" not in st.session_state:
    st.session_state.rows = []

st.subheader("Add PO Line")
cols = st.columns([1.2, 1.2, 1.6, 0.8, 1.0, 1.2, 0.4])
with cols[0]: ctrl = st.text_input("Control NO", key="ctrl_in")
with cols[1]: item = st.text_input("Item NO", key="item_in")
with cols[2]: jan = st.text_input("JAN / Barcode", key="jan_in")
with cols[3]: qty = st.text_input("Qty", key="qty_in", placeholder="e.g. 6,000")
with cols[4]: price = st.text_input("Price", key="price_in", placeholder="e.g. 60.600")
with cols[5]: delivery = st.text_input("Delivery", key="delv_in", placeholder="e.g. 8/ETD")
add = cols[6].button("➕")

if add:
    st.session_state.rows.append({
        "Control NO": ctrl.strip(),
        "Item NO": item.strip(),
        "Barcode": jan.strip(),
        "Qty": qty.strip(),
        "Price": price.strip(),
        "Delivery": delivery.strip(),
    })
    st.session_state.ctrl_in = ""
    st.session_state.item_in = ""
    st.session_state.jan_in = ""
    st.session_state.qty_in = ""
    st.session_state.price_in = ""
    st.session_state.delv_in = ""

st.subheader("Current Lines")
if st.session_state.rows:
    df = pd.DataFrame(st.session_state.rows)
    st.dataframe(df, use_container_width=True, height=220)
else:
    st.info("Add at least one line above.")

# ------------------------------
# Utils
# ------------------------------

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


def auto_width(ws, start_col=1, end_col=10):
    for col in range(start_col, end_col + 1):
        letter = get_column_letter(col)
        max_len = 0
        for cell in ws[letter]:
            v = cell.value
            if v is None:
                continue
            l = len(str(v))
            if l > max_len:
                max_len = l
        ws.column_dimensions[letter].width = min(max(10, max_len + 2), 60)

# ------------------------------
# Build workbook per item
# ------------------------------
make_btn = st.button("📦 Generate PurchaseOrders.xlsx", type="primary")

if make_btn:
    # filter valid lines (must have control + item)
    entries: List[dict] = [r for r in st.session_state.rows if str(r.get("Control NO", "")).strip() and str(r.get("Item NO", "")).strip()]
    if not entries:
        st.error("Please add at least one line with Control NO and Item NO.")
        st.stop()

    # Deduplicate per (Control NO, Item NO) keep first
    seen = set()
    uniq = []
    for r in entries:
        key = (r.get("Control NO",""), r.get("Item NO",""))
        if key not in seen:
            seen.add(key)
            uniq.append(r)

    wb = Workbook()
    # We'll remove the default sheet after creating ours
    default_ws = wb.active

    for r in uniq:
        control_no = str(r.get("Control NO", "")).strip()
        item_no = str(r.get("Item NO", "")).strip()
        barcode = str(r.get("Barcode", "")).strip()
        qty = to_int(r.get("Qty"))
        price = to_float(r.get("Price"))
        amount = (qty or 0) * (price or 0)
        delivery = str(r.get("Delivery", "")).strip()

        title = sanitize_sheet_title(f"{control_no}_{item_no}")
        ws = wb.create_sheet(title)

        # Header
        ws.merge_cells("A1:F1")
        ws["A1"] = "Purchase Order"
        ws["A1"].font = Font(size=16, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center")

        ws["A3"] = "Control NO"; ws["B3"] = control_no
        ws["D3"] = "Delivery";   ws["E3"] = delivery
        for c in ["A3","B3","D3","E3"]:
            ws[c].font = Font(bold=(c in ["A3","D3"]))

        # Table header at row 6-7 for spacing similar to template
        start_row = 8
        headers = ["Item No", "JAN Code", "Qty", "Price", "Amount"]
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(row=start_row, column=ci, value=h)
            cell.font = Font(bold=True)
            cell.fill = colhdr_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = border_all

        # Single line for this item
        row = start_row + 1
        values = [item_no, barcode, qty, price, amount]
        for ci, val in enumerate(values, 1):
            cell = ws.cell(row=row, column=ci, value=val)
            cell.border = border_all
            if ci == 3:  # Qty
                cell.number_format = num_qty
                cell.alignment = Alignment(horizontal="right")
            elif ci in (4, 5):  # Price, Amount
                # Price with 3 decimals and Amount with 2 decimals-like currency
                if ci == 4:
                    cell.number_format = num_price3
                else:
                    cell.number_format = num_currency
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="left")

        auto_width(ws, 1, len(headers))

    # Remove default
    wb.remove(default_ws)

    # Deliver file
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    st.success(f"Created {len(uniq)} sheet(s)")
    st.download_button(
        "⬇️ Download PurchaseOrders.xlsx",
        buf,
        file_name="PurchaseOrders.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown("---")
st.markdown(
    """
**Behavior:** One sheet per *(Control NO, Item NO)* using a fixed layout built in-code (no template).

**Columns:** Item No, JAN Code, Qty, Price (3 decimals), Amount (Qty×Price).  
**Formats:** Qty `#,##0`; Price `#,##0.000`; Amount currency styled as `¥`.
"""
)
