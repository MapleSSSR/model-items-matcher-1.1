
import re
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Model Number → Items Matcher (Preserve Format)", layout="wide")
st.title("🔗 Model Number → Items Matcher · Preserve Original Formatting (v3.2)")

st.markdown("""
**How it works**  
- Edits your uploaded **A** workbook *in place* (using `openpyxl`) to preserve cell styles, column widths, filters, freeze panes, etc.  
- Inserts a new **Items** column **right after “Model Number”** on each sheet that has that header (case-insensitive).  
- **Matching rule**: For each token in A, find the **longest** `Model Number` from B that is **contained** in that token (case-insensitive). If none found → `N/A`.  
- Ignores leading quantities like `49 x ...`, `289*...`. Supports multiple models separated by **`,`** (English) or **`，`** (Chinese).
- If any token in a row is `N/A`, the **entire row is highlighted in red**.
""")

with st.sidebar:
    st.header("Upload files (Drag & Drop supported)")
    file_a = st.file_uploader("Upload A (Excel to be matched)", type=["xlsx"])
    file_b = st.file_uploader("Upload B (Mapping: 2 columns)", type=["xlsx", "xls"])
    run = st.button("🚀 Run Matching (Preserve Format)", use_container_width=True)

def _clean_leading_qty(txt: str) -> str:
    return re.sub(r"^\s*\d+\s*[x\*]\s*", "", txt, flags=re.IGNORECASE)

def _split_models(cell: str):
    parts = re.split(r"[,\，]", cell)
    return [p.strip() for p in parts if p.strip()]

def _norm_cols(cols):
    return [str(c).strip() for c in cols]

def build_mapping(b_df: pd.DataFrame):
    b_df = b_df.copy()
    b_df.columns = _norm_cols(b_df.columns)
    # auto-detect columns (fallback to first two)
    col1 = None; col2 = None
    for c in b_df.columns:
        if c.lower() == "model number" and col1 is None:
            col1 = c
        if c.lower() == "items" and col2 is None:
            col2 = c
    if col1 is None or col2 is None:
        col1, col2 = b_df.columns[:2]
    tmp = b_df[[col1, col2]].copy()
    tmp[col1] = tmp[col1].astype(str).str.strip()
    tmp[col2] = tmp[col2].astype(str).str.strip()
    tmp = tmp[tmp[col1] != ""]
    mapping = dict(zip(tmp[col1], tmp[col2]))
    keys_sorted = sorted(mapping.keys(), key=lambda x: len(str(x)), reverse=True)
    keys_sorted_lower = [str(k).lower() for k in keys_sorted]
    return mapping, keys_sorted, keys_sorted_lower

def longest_substring_match(token: str, keys_sorted, keys_sorted_lower):
    t = token.lower()
    for k_lower, k in zip(keys_sorted_lower, keys_sorted):
        if k_lower in t:
            return k
    return None

def extend_autofilter_to_last_col(ws):
    try:
        if ws.auto_filter and ws.auto_filter.ref:
            ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    except Exception:
        pass

def process_workbook(a_bytes: bytes, b_df: pd.DataFrame) -> bytes:
    mapping, keys_sorted, keys_sorted_lower = build_mapping(b_df)
    red_fill = PatternFill(start_color="FFFF9999", end_color="FFFF9999", fill_type="solid")
    wb = load_workbook(BytesIO(a_bytes))
    for ws in wb.worksheets:
        header_row = 1
        model_col_idx = None
        model_header_style = None
        for cell in ws[header_row]:
            if cell.value and str(cell.value).strip().lower() == "model number":
                model_col_idx = cell.column
                model_header_style = cell._style
                break
        if model_col_idx is None:
            continue
        mn_letter = get_column_letter(model_col_idx)
        mn_width = ws.column_dimensions[mn_letter].width
        insert_at = model_col_idx + 1
        ws.insert_cols(insert_at, 1)
        header_cell = ws.cell(row=header_row, column=insert_at)
        header_cell.value = "Items"
        try:
            header_cell._style = model_header_style
        except Exception:
            pass
        ws.column_dimensions[get_column_letter(insert_at)].width = mn_width if mn_width else 15
        for r in range(2, ws.max_row + 1):
            raw = ws.cell(row=r, column=model_col_idx).value
            text = (str(raw).strip()) if raw is not None else ""
            if text == "":
                ws.cell(row=r, column=insert_at).value = ""
                continue
            cleaned = _clean_leading_qty(text)
            parts = _split_models(cleaned)
            out_items = []
            has_na = False
            for p in parts:
                mk = longest_substring_match(p, keys_sorted, keys_sorted_lower)
                if mk is None:
                    out_items.append("N/A"); has_na = True
                else:
                    val = mapping.get(mk, "N/A")
                    out_items.append(val)
                    if val == "N/A":
                        has_na = True
            ws.cell(row=r, column=insert_at).value = ",".join(out_items)
            if has_na:
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=c).fill = red_fill
        extend_autofilter_to_last_col(ws)
    out = BytesIO()
    wb.save(out)
    return out.getvalue()

if run:
    if not file_a or not file_b:
        st.error("Please upload both A and B files.")
        st.stop()
    try:
        b_df = pd.read_excel(file_b)
        processed = process_workbook(file_a.read(), b_df)

        # Dynamic output name: <A filename>_update.xlsx
        base_name = file_a.name.rsplit(".", 1)[0] if file_a.name else "A_matched"
        out_name = f"{base_name}_update.xlsx"

        st.success("Done! Formatting preserved as much as possible.")
        st.download_button("⬇️ Download result", data=processed,
                           file_name=out_name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.exception(e)
else:
    st.info("Drag & drop your A and B files on the left, then click the button.")
