
import re
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Model Number → Items 匹配工具（保留原格式版）", layout="wide")
st.title("🔗 Model Number → Items 匹配工具 · 保留原格式")

st.markdown("""
**说明**  
- 该版本在原文件 **A** 上“原地写入”（使用 openpyxl），尽量保留原有：单元格样式、列宽、筛选、冻结窗格等；只是在 **Model Number** 之后插入一列 **Items** 并填充值。  
- **匹配规则**：对 A 每个型号片段做“**包含匹配 → 取最长**”；匹配不到填 **N/A**；若某行含 N/A，整行标红。  
- 仍然会忽略开头数量（`49 x ...`、`289*...`）；支持中英文逗号分隔多个型号。
""")

with st.sidebar:
    st.header("上传文件")
    file_a = st.file_uploader("上传 A（待匹配 Excel）", type=["xlsx"])
    file_b = st.file_uploader("上传 B（对照表：两列）", type=["xlsx", "xls"])
    run = st.button("🚀 开始匹配（保留格式）", use_container_width=True)

# ---------- 工具函数 ----------
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
    # 尝试识别列名；否则前两列
    col1 = None; col2 = None
    for c in b_df.columns:
        if c.lower() == "model number" and col1 is None:
            col1 = c
        if c.lower() == "items" and col2 is None:
            col2 = c
    if col1 is None or col2 is None:
        col1, col2 = b_df.columns[:2]
    tmp = b_df[[col1, col2]].copy()
    tmp[col1] = tmp[col1].astype(str).strip()
    tmp[col2] = tmp[col2].astype(str).strip()
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
    # 尝试把自动筛选区域扩展到当前占用的最右列
    try:
        if ws.auto_filter and ws.auto_filter.ref:
            ref = ws.auto_filter.ref  # e.g., 'A1:Q999'
            first_cell = ref.split(":")[0].split("!")[0] if "!" in ref else ref.split(":")[0]
            # 以第一行作为表头
            max_col = ws.max_column
            last = f"{get_column_letter(max_col)}1:{get_column_letter(max_col)}{ws.max_row}"
            new_ref = f"A1:{get_column_letter(max_col)}{ws.max_row}"
            ws.auto_filter.ref = new_ref
    except Exception:
        pass

def process_workbook(a_bytes: bytes, b_df: pd.DataFrame) -> bytes:
    mapping, keys_sorted, keys_sorted_lower = build_mapping(b_df)
    red_fill = PatternFill(start_color="FFFF9999", end_color="FFFF9999", fill_type="solid")

    wb = load_workbook(BytesIO(a_bytes))
    for ws in wb.worksheets:
        # 找到表头行（默认第一行）并定位 Model Number 列
        header_row = 1
        headers = [ (c.value, c.column) for c in ws[header_row] ]
        model_col_idx = None
        for cell in ws[header_row]:
            if cell.value and str(cell.value).strip().lower() == "model number":
                model_col_idx = cell.column
                model_header_style = cell._style
                break
        if model_col_idx is None:
            continue  # 不处理该 sheet

        # 记录原 Model Number 列的列宽
        mn_letter = get_column_letter(model_col_idx)
        mn_width = ws.column_dimensions[mn_letter].width

        # 在其后插入一列
        insert_at = model_col_idx + 1
        ws.insert_cols(insert_at, 1)

        # 设置新列标题与样式、列宽
        header_cell = ws.cell(row=header_row, column=insert_at)
        header_cell.value = "Items"
        try:
            header_cell._style = model_header_style
        except Exception:
            pass
        # 列宽：若原值缺失，用 15 作为默认
        ws.column_dimensions[get_column_letter(insert_at)].width = mn_width if mn_width else 15

        # 遍历每一行填充值
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
                    out_items.append(mapping.get(mk, "N/A"))
                    if out_items[-1] == "N/A":
                        has_na = True
            ws.cell(row=r, column=insert_at).value = ",".join(out_items)

            if has_na:
                # 整行标红：从1到当前最大列
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=c).fill = red_fill

        # 尝试扩展自动筛选至最右列
        extend_autofilter_to_last_col(ws)

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

# ---------- 主逻辑 ----------
if run:
    if not file_a or not file_b:
        st.error("请同时上传 A 与 B。")
        st.stop()
    try:
        b_df = pd.read_excel(file_b)
        processed = process_workbook(file_a.read(), b_df)
        st.success("处理完成！已尽量保留原始格式。")
        st.download_button(
            "⬇️ 下载处理后的 A 文件（保留格式）",
            data=processed,
            file_name="A_matched_preserved.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.exception(e)
else:
    st.info("请在左侧上传文件并点击按钮开始处理。")
