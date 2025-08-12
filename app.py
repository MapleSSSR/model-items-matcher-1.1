
import re
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.table import TableColumn

st.set_page_config(page_title="Model Number → Items Matcher (Preserve Format)", layout="wide")
st.title("🔗 Model Number → Items Matcher · Preserve Original Formatting (v3.4)")

st.markdown("""
**What's new in v3.4**  
- **No more blank gridlines**: we only process up to the **last non-empty Model Number row**, so we don't touch empty rows at the bottom.  
- **Keep totals working**: after inserting the column, we adjust formulas that reference columns to the right (e.g. `SUM(R:R)` becomes `SUM(S:S)`).  
- **Consistent look**: the new **Items** column copies the body style from **Model Number** so borders align with neighboring columns.  
- Still includes: longest **substring** match, whole-row red highlight on `N/A`, ignore leading quantities, multi-models with commas, preserve filters/tables.
""")

with st.sidebar:
    st.header("Upload files (Drag & Drop supported)")
    file_a = st.file_uploader("Upload A (Excel to be matched)", type=["xlsx"])
    file_b = st.file_uploader("Upload B (Mapping: 2 columns)", type=["xlsx", "xls"])
    run = st.button("🚀 Run Matching (Preserve Format)", use_container_width=True)

def _clean_leading_qty(txt: str) -> str:
    return re.sub(r"^\\s*\\d+\\s*[x\\*]\\s*", "", txt, flags=re.IGNORECASE)

def _split_models(cell: str):
    parts = re.split(r"[,\，]", cell)
    return [p.strip() for p in parts if p.strip()]

def _norm_cols(cols):
    return [str(c).strip() for c in cols]

def build_mapping(b_df: pd.DataFrame):
    b_df = b_df.copy()
    b_df.columns = _norm_cols(b_df.columns)
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

def _parse_ref(ref: str):
    a, b = ref.split(":")
    m1 = re.match(r"([A-Z]+)(\d+)", a)
    m2 = re.match(r"([A-Z]+)(\d+)", b)
    if not (m1 and m2):
        return None
    return m1.group(1), int(m1.group(2)), m2.group(1), int(m2.group(2))

def widen_filter_columns_only(ws):
    try:
        if ws.auto_filter and ws.auto_filter.ref:
            parsed = _parse_ref(ws.auto_filter.ref)
            if not parsed: return
            sc, sr, ec, er = parsed
            new_end_col = get_column_letter(ws.max_column)
            ws.auto_filter.ref = f"{sc}{sr}:{new_end_col}{er}"
    except Exception:
        pass

def extend_table_if_needed(ws, header_row, model_col_idx):
    try:
        for tbl in list(getattr(ws, "_tables", [])):
            parsed = _parse_ref(tbl.ref)
            if not parsed: 
                continue
            sc, sr, ec, er = parsed
            start_col_idx = column_index_from_string(sc)
            end_col_idx = column_index_from_string(ec)
            if sr != header_row:
                continue
            if not (start_col_idx <= model_col_idx <= end_col_idx):
                continue
            new_end_col_idx = end_col_idx + 1
            tbl.ref = f"{sc}{sr}:{get_column_letter(new_end_col_idx)}{er}"
            max_id = 0
            for col in tbl.tableColumns._tableColumns:
                max_id = max(max_id, int(col.id))
            tbl.tableColumns._tableColumns.append(TableColumn(id=max_id+1, name="Items"))
            break
    except Exception:
        pass

# Formula shifting: bump any A1/R1C1-style references by +delta columns if their column >= insert_at
_cell_ref_re = re.compile(r"(\\$?)([A-Z]{1,3})(\\$?)(\\d+)")
_col_only_re = re.compile(r"(\\$?)([A-Z]{1,3})(\\$?)\\s*:\\s*(\\$?)([A-Z]{1,3})(\\$?)")  # e.g., R:S or $R:$S

def shift_formula_by_col(formula: str, insert_at_col: int, delta: int = 1) -> str:
    # Skip structured references/tables (contain '[' or ']')
    if "[" in formula and "]" in formula:
        return formula

    def shift_col(col_letters: str) -> str:
        idx = column_index_from_string(col_letters)
        if idx >= insert_at_col:
            return get_column_letter(idx + delta)
        return col_letters

    # First handle column-only ranges like R:S
    def repl_colrange(m):
        l1, c1, a1, l2, c2, a2 = m.groups()
        nc1 = shift_col(c1)
        nc2 = shift_col(c2)
        return f"{l1}{nc1}{a1}:{l2}{nc2}{a2}"

    formula2 = _col_only_re.sub(repl_colrange, formula)

    # Then handle full refs like R2, $R$10
    def repl_cell(m):
        abs_col, col, abs_row, row = m.groups()
        new_col = shift_col(col)
        return f"{abs_col}{new_col}{abs_row}{row}"

    formula3 = _cell_ref_re.sub(repl_cell, formula2)
    return formula3

def last_data_row_in_column(ws, col_idx: int, start_row: int = 2) -> int:
    # Scan upward from bottom until find a non-empty cell in 'Model Number' column
    for r in range(ws.max_row, start_row - 1, -1):
        val = ws.cell(row=r, column=col_idx).value
        if val not in (None, ""):
            return r
    return start_row - 1

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

        # Capture body style from the row below header, same column (if exists)
        body_style = None
        if ws.max_row >= header_row + 1:
            body_style = ws.cell(row=header_row + 1, column=model_col_idx)._style

        # Insert column
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

        # Adjust formulas on sheet (cells with '=' in value)
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                val = cell.value
                if isinstance(val, str) and val.startswith("="):
                    newf = shift_formula_by_col(val, insert_at_col=insert_at, delta=1)
                    if newf != val:
                        cell.value = newf

        # Process only up to the last data row in Model Number column
        last_row = last_data_row_in_column(ws, model_col_idx, start_row=header_row + 1)

        for r in range(header_row + 1, last_row + 1):
            raw = ws.cell(row=r, column=model_col_idx).value
            text = (str(raw).strip()) if raw is not None else ""
            if text == "":
                continue  # do not touch empty trailing rows
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
            tgt = ws.cell(row=r, column=insert_at)
            tgt.value = ",".join(out_items)
            # Copy body style so borders look consistent
            try:
                if body_style is not None:
                    tgt._style = body_style
            except Exception:
                pass
            if has_na:
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=c).fill = red_fill

        widen_filter_columns_only(ws)
        extend_table_if_needed(ws, header_row, model_col_idx)

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
        base_name = file_a.name.rsplit(".", 1)[0] if file_a.name else "A_matched"
        out_name = f"{base_name}_update.xlsx"
        st.success("Done! Formatting, totals and borders preserved as much as possible.")
        st.download_button("⬇️ Download result", data=processed,
                           file_name=out_name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.exception(e)
else:
    st.info("Drag & drop your A and B files on the left, then click the button.")
