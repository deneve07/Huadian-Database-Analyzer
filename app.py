"""
華典資料庫分析系統 (Streamlit 版)
----------------------------------
提供：
1. 自訂樞紐分析 - 讓使用者用拖拉的方式，從各資料集「所有欄位」中自由挑選、排序要放進報表的欄位，
   並自訂要合計(小計)的欄位（支援多層巢狀合計），數值欄位加總後可直接下載 Excel。
2. 處方釋出率分析 - 固定格式，維持原本的計算邏輯與版面，不提供自訂欄位。

需要的套件（requirements.txt）：
    streamlit
    pandas
    numpy
    openpyxl
    streamlit-sortables
"""

import re
import io
import pandas as pd
import numpy as np
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from streamlit_sortables import sort_items
    HAS_SORTABLES = True
except ImportError:
    HAS_SORTABLES = False


# ============================================================
# 基礎設定與資料載入
# ============================================================

st.set_page_config(page_title="華典資料庫分析系統", page_icon="📊", layout="wide")

DATA_SOURCES = {
    "HP+GP+DS 資料 (廠商 / 層級別 適用)": "data_hp_gp_ds.csv",
    "科別資料": "data_department.csv",
    "推估醫院資料": "data_hospital.csv",
}


@st.cache_data(show_spinner=False)
def load_data(path: str) -> pd.DataFrame:
    try:
        df = pd.read_csv(path, dtype=str)
    except Exception:
        return pd.DataFrame()

    # 欄位名稱正規化，讓不同資料集的欄位語意一致
    if "劑型小分類" in df.columns:
        df = df.rename(columns={"劑型小分類": "劑型"})
    rename_map = {}
    for y in ["2022", "2023", "2024"]:
        src = f"{y}年數量(顆)"
        if src in df.columns:
            rename_map[src] = f"{y}年申報量(顆)"
    if rename_map:
        df = df.rename(columns=rename_map)

    # 數值欄位轉型
    for col in df.columns:
        if ("申報量" in col) or ("金額" in col):
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def parse_dosage(d):
    m = re.search(r"[\d.]+", str(d))
    return float(m.group()) if m else 0.0


# ============================================================
# 通用樞紐分析引擎：任意欄位順序 + 任意層級巢狀合計
# ============================================================

def build_nested_rows(df: pd.DataFrame, row_fields: list, subtotal_fields: list, value_cols: list):
    """
    依 row_fields 的順序排列欄位，並依 subtotal_fields 的順序（由外而內）產生巢狀合計列。
    回傳一份 list[dict]，每個 dict 為 {'type': 'data'|'subtotal'|'total', 'values': {...}, 'sums': {...}}
    """
    year_cols = sorted([c for c in value_cols if "申報量" in c], reverse=True)
    sort_cols = list(row_fields) + year_cols
    sort_asc = [True] * len(row_fields) + [False] * len(year_cols)
    df_sorted = df.sort_values(by=sort_cols, ascending=sort_asc) if sort_cols else df

    # 除了小計欄位本身，其餘欄位若重複則留空（但只在「安全」的情況下：
    # 該欄位本身全域唯一值只有一個，或它排在所有小計層級最前面）
    if subtotal_fields:
        min_level_idx = min(row_fields.index(c) for c in subtotal_fields)
    else:
        min_level_idx = len(row_fields)

    group_cols_to_blank = []
    for c in row_fields:
        if c in subtotal_fields:
            continue
        if df[c].nunique() <= 1 or row_fields.index(c) < min_level_idx:
            group_cols_to_blank.append(c)

    rows = []
    last_vals = {}
    first_flags = {c: True for c in subtotal_fields}

    def emit_row(row):
        rec_vals = {}
        for f in row_fields:
            val = row[f]
            show = val
            if f in group_cols_to_blank:
                if val != last_vals.get(f):
                    last_vals[f] = val
                else:
                    show = ""
            elif f in first_flags:
                if not first_flags[f]:
                    show = ""
            rec_vals[f] = show
        rows.append({"type": "data", "values": rec_vals, "sums": {c: row[c] for c in value_cols}})
        for c in first_flags:
            first_flags[c] = False

    def emit_subtotal(field, label, totals):
        rec_vals = {f: "" for f in row_fields}
        rec_vals[field] = f"{label} 合計"
        rows.append({"type": "subtotal", "values": rec_vals, "sums": totals})

    def recurse(sub_df, level_idx):
        if level_idx >= len(subtotal_fields):
            for _, row in sub_df.iterrows():
                emit_row(row)
            return
        col = subtotal_fields[level_idx]
        for _, grp in sub_df.groupby(col, sort=False):
            totals = {c: grp[c].sum() for c in value_cols}
            label = grp[col].iloc[0]
            first_flags[col] = True
            recurse(grp, level_idx + 1)
            emit_subtotal(col, label, totals)

    if subtotal_fields:
        recurse(df_sorted, 0)
    else:
        for _, row in df_sorted.iterrows():
            emit_row(row)

    grand_totals = {c: df[c].sum() for c in value_cols}
    total_vals = {f: "" for f in row_fields}
    if row_fields:
        total_vals[row_fields[0]] = "總計"
    rows.append({"type": "total", "values": total_vals, "sums": grand_totals})

    return rows


def rows_to_preview_df(rows, row_fields, value_cols):
    records = []
    for r in rows:
        rec = dict(r["values"])
        for c in value_cols:
            v = r["sums"].get(c, "")
            rec[c] = f"{v:,.0f}" if v != "" else ""
        records.append(rec)
    return pd.DataFrame(records, columns=row_fields + value_cols)


def generate_excel_bytes(rows, row_fields, value_cols, report_title):
    wb = Workbook()
    ws = wb.active
    ws.title = "分析報表"

    ws.views.sheetView[0].showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True

    headers = row_fields + value_cols
    header_fill = PatternFill(start_color="00695C", end_color="00695C", fill_type="solid")
    subtotal_fill = PatternFill(start_color="E0F2F1", end_color="E0F2F1", fill_type="solid")
    total_fill = PatternFill(start_color="B2DFDB", end_color="B2DFDB", fill_type="solid")
    font_b_w = Font(name="微軟正黑體", bold=True, color="FFFFFF", size=12)
    font_norm = Font(name="微軟正黑體", size=11)
    font_bold = Font(name="微軟正黑體", bold=True, size=11)
    align_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_r = Alignment(horizontal="right", vertical="center")
    border_thin = Border(*[Side(style="thin", color="D9D9D9")] * 4)
    header_border = Border(*[Side(style="thin", color="FFFFFF")] * 4)

    ws.cell(row=1, column=1, value=report_title)
    ws.cell(row=1, column=1).font = Font(name="微軟正黑體", bold=True, size=16, color="004D40")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(len(headers), 1))
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center")

    header_row = 3
    ws.append([])
    ws.append([])
    ws.append(headers)
    ws.row_dimensions[header_row].height = 40
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=header_row, column=c)
        cell.font = font_b_w
        cell.fill = header_fill
        cell.border = header_border
        cell.alignment = align_c

    ws.print_title_rows = f"{header_row}:{header_row}"

    current_row = header_row + 1
    for r in rows:
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=current_row, column=i)
            if h in value_cols:
                cell.value = r["sums"].get(h, "")
                cell.number_format = "#,##0"
                cell.alignment = align_r
            else:
                cell.value = r["values"].get(h, "")
                cell.alignment = Alignment(horizontal="left" if i == 1 else "center", vertical="center", wrap_text=True)
            cell.font = font_bold if r["type"] != "data" else font_norm
            cell.border = border_thin
            if r["type"] == "subtotal":
                cell.fill = subtotal_fill
            elif r["type"] == "total":
                cell.fill = total_fill
        ws.row_dimensions[current_row].height = 30
        current_row += 1

    for col_idx, col_name in enumerate(headers, 1):
        col_letter = get_column_letter(col_idx)
        if col_idx == 1:
            ws.column_dimensions[col_letter].width = 28
        elif col_name in value_cols:
            ws.column_dimensions[col_letter].width = 16
        else:
            ws.column_dimensions[col_letter].width = 16

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ============================================================
# 處方釋出率分析（固定格式，不提供自訂欄位）
# ============================================================

def analysis_prescription(df_filtered, vendor_name=None):
    if df_filtered.empty:
        return None

    ds_hp_cond = (df_filtered["通路"] == "DS") & (df_filtered["層級別"].isin(["1.醫學中心", "2.區域醫院", "3.地區醫院"]))
    ds_hp_val = df_filtered[ds_hp_cond]["2024年申報量(顆)"].sum()
    hp_val = df_filtered[df_filtered["通路"] == "HP"]["2024年申報量(顆)"].sum()
    total_val = ds_hp_val + hp_val
    rate = (ds_hp_val / total_val * 100) if total_val > 0 else 0

    title = f"{vendor_name}_醫院處方釋出率(2024年)" if vendor_name else "整體醫院處方釋出率(2024年)"
    return {"title": title, "ds_hp_val": ds_hp_val, "hp_val": hp_val, "total_val": total_val, "rate": rate}


# ============================================================
# Streamlit 主畫面
# ============================================================

st.title("📊 華典資料庫分析系統")

tab_pivot, tab_prescription = st.tabs(["🧩 自訂樞紐分析", "💊 處方釋出率分析 (固定格式)"])

# ---------------- 自訂樞紐分析 ----------------
with tab_pivot:
    st.caption("依資料來源挑選欄位、拖拉排序，並自由決定要合計的欄位，效果如同 Excel 樞紐分析。")

    source_label = st.selectbox("第1步：選擇資料來源", list(DATA_SOURCES.keys()), key="pivot_source")
    df_raw = load_data(DATA_SOURCES[source_label])

    if df_raw.empty:
        st.warning(f"找不到資料檔案「{DATA_SOURCES[source_label]}」，請確認檔案已放置於工作目錄。")
    else:
        all_cols = list(df_raw.columns)
        value_cols_auto = [c for c in all_cols if ("申報量" in c) or ("金額" in c)]
        dim_cols_all = [c for c in all_cols if c not in value_cols_auto]

        st.markdown("### 🔍 第2步：篩選資料 (可複選；不選代表不篩選該欄位)")
        filters = {}
        filter_cols_ui = st.columns(3)
        for i, col in enumerate(dim_cols_all):
            options = sorted([v for v in df_raw[col].dropna().unique() if str(v).strip() != ""])
            if 1 < len(options) <= 300:
                with filter_cols_ui[i % 3]:
                    sel = st.multiselect(col, options, key=f"filter_{source_label}_{col}")
                    if sel:
                        filters[col] = sel

        df_filtered = df_raw.copy()
        for col, sel in filters.items():
            df_filtered = df_filtered[df_filtered[col].isin(sel)]

        st.info(f"篩選後共 **{len(df_filtered):,}** 筆資料")

        st.markdown("### 🧩 第3步：拖拉排序要放入報表的欄位 (由上到下＝報表由左到右)")
        if HAS_SORTABLES:
            ordered_all = sort_items(dim_cols_all, key=f"sortable_{source_label}")
        else:
            st.error("尚未安裝 streamlit-sortables 套件，暫以預設順序呈現，請於 requirements.txt 加入 streamlit-sortables 後重新部署。")
            ordered_all = dim_cols_all

        default_selection = ordered_all[: min(4, len(ordered_all))]
        row_fields_selected = st.multiselect(
            "勾選要放入報表的欄位（將依照上方拖曳後的順序自動排列，不需要再手動調整順序）",
            options=ordered_all,
            default=default_selection,
        )
        row_fields = [f for f in ordered_all if f in row_fields_selected]

        st.markdown("### Σ 第4步：選擇需要合計的欄位 (依上方順序由外而內巢狀合計)")
        subtotal_selected = st.multiselect("合計欄位 (可複選，可留空代表不需要任何小計)", options=row_fields)
        subtotal_fields = [f for f in row_fields if f in subtotal_selected]

        st.markdown("### 🔢 第5步：選擇要加總的數值欄位")
        value_cols = st.multiselect("數值欄位 (預設帶入所有申報量／金額欄位)", options=value_cols_auto, default=value_cols_auto)

        st.divider()
        if st.button("🚀 產生報表", type="primary", key="pivot_generate"):
            if not row_fields:
                st.error("⚠️ 請至少選擇一個要放入報表的欄位！")
            elif not value_cols:
                st.error("⚠️ 請至少選擇一個要加總的數值欄位！")
            elif df_filtered.empty:
                st.error("❌ 篩選後無資料，請放寬篩選條件。")
            else:
                rows = build_nested_rows(df_filtered, row_fields, subtotal_fields, value_cols)
                preview_df = rows_to_preview_df(rows, row_fields, value_cols)
                report_title = f"{source_label}_自訂分析"

                st.markdown("### 📄 報表預覽")
                st.dataframe(preview_df, use_container_width=True, height=min(700, 40 + 32 * len(preview_df)))

                excel_bytes = generate_excel_bytes(rows, row_fields, value_cols, report_title)
                st.download_button(
                    "📥 下載 Excel 報表",
                    data=excel_bytes,
                    file_name=f"{report_title}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# ---------------- 處方釋出率分析（固定格式） ----------------
with tab_prescription:
    st.caption("分析成分從醫院端流向藥局處方的釋出率，可看單一廠商。此分析維持固定格式，不提供自訂欄位。")

    df_hp = load_data("data_hp_gp_ds.csv")
    if df_hp.empty:
        st.warning("找不到資料檔案「data_hp_gp_ds.csv」，請確認檔案已放置於工作目錄。")
    else:
        comp_options = sorted([c for c in df_hp["成分簡稱"].dropna().unique() if c])
        search = st.text_input("第1步：輸入成分關鍵字", placeholder="例如: Levofloxacin", key="presc_search")
        filtered_comps = [c for c in comp_options if search.strip().lower() in c.lower()] if search.strip() else []

        comp = st.selectbox("第2步：選擇單一成分品項 (必選)", options=[""] + filtered_comps, key="presc_comp")

        combo, form, dose, vendor = None, None, None, None
        if comp:
            df_c = df_hp[df_hp["成分簡稱"] == comp]
            combo_options = sorted([v for v in df_c["單複方"].dropna().unique() if v]) if "單複方" in df_c.columns else []
            combo = st.selectbox("第3步：選擇單一單複方 (必選)", options=[""] + combo_options, key="presc_combo")

            if combo:
                df_cb = df_c[df_c["單複方"] == combo]
                form_options = sorted([v for v in df_cb["劑型"].dropna().unique() if v])
                form = st.selectbox("第4步：選擇單一劑型 (必選)", options=[""] + form_options, key="presc_form")

                if form:
                    df_cbf = df_cb[df_cb["劑型"] == form]
                    dose_options = sorted([v for v in df_cbf["含量"].dropna().unique() if v], key=parse_dosage)
                    dose = st.selectbox("第5步：選擇單一含量 (必選)", options=[""] + dose_options, key="presc_dose")

                    if dose:
                        df_cbfd = df_cbf[df_cbf["含量"] == dose]
                        show_vendor = st.checkbox("若需顯示單一廠商，請勾選", key="presc_show_vendor")
                        if show_vendor:
                            vendor_options = sorted([v for v in df_cbfd["廠商簡稱"].dropna().unique() if v])
                            vendor = st.selectbox("第6步：選擇單一廠商 (必選)", options=[""] + vendor_options, key="presc_vendor")

        st.divider()
        if st.button("🚀 產生報表", type="primary", key="presc_generate"):
            missing = []
            if not comp: missing.append("成分")
            if not combo: missing.append("單複方")
            if not form: missing.append("劑型")
            if not dose: missing.append("含量")
            if st.session_state.get("presc_show_vendor") and not vendor:
                missing.append("廠商")

            if missing:
                st.error(f"⚠️ 請務必選擇：{'、'.join(missing)}！")
            else:
                df_f = df_hp[
                    (df_hp["成分簡稱"] == comp) & (df_hp["單複方"] == combo) &
                    (df_hp["劑型"] == form) & (df_hp["含量"] == dose)
                ]
                if vendor:
                    df_f = df_f[df_f["廠商簡稱"] == vendor]

                summary = analysis_prescription(df_f, vendor)
                if not summary:
                    st.error("❌ 找不到符合條件的資料。")
                else:
                    st.markdown(f"## 📊 {summary['title']}")
                    c1, c2, c3 = st.columns(3)
                    c1.metric("DS from HP", f"{summary['ds_hp_val']:,.0f}")
                    c2.metric("HP", f"{summary['hp_val']:,.0f}")
                    c3.metric("DS from HP + HP", f"{summary['total_val']:,.0f}")
                    st.caption("公式：處方釋出率 ＝ DS from HP ÷ (DS from HP + HP) × 100%")
                    st.markdown(
                        f"<div style='font-size:32px; font-weight:bold; color:#D32F2F;'>{summary['title']}：{summary['rate']:.2f}%</div>",
                        unsafe_allow_html=True,
                    )
