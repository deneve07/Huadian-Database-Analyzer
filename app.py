"""
華典資料庫分析系統 (Streamlit 版)
----------------------------------
1. 自訂樞紐分析 - 用拖拉的方式，從各資料集「所有欄位」中自由挑選、排序要放進報表的欄位，
   每個欄位可直接篩選、決定是否合計，並可加入占比(%)、成長率(%) 計算欄位。
2. 處方釋出率分析 - 固定格式，維持原本的計算邏輯與版面，不提供自訂欄位。

需要的套件（requirements.txt）：
    streamlit
    pandas
    numpy
    openpyxl
    pyarrow
    streamlit-sortables
"""

import re
import io
import os
import pandas as pd
import numpy as np
import streamlit as st
import streamlit.components.v1 as components
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


@st.cache_data(show_spinner=False)
def load_data(path: str) -> pd.DataFrame:
    parquet_path = path.rsplit(".", 1)[0] + ".parquet"
    try:
        if os.path.exists(parquet_path):
            df = pd.read_parquet(parquet_path)
        else:
            # utf-8-sig 可自動去除 Excel 匯出 CSV 常見的 BOM 字元，
            # 否則第一欄欄名容易變成「\ufeff成分簡稱」導致 KeyError
            df = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
    except Exception:
        return pd.DataFrame()

    df.columns = [str(c).strip() for c in df.columns]

    if "劑型小分類" in df.columns:
        df = df.rename(columns={"劑型小分類": "劑型"})
    rename_map = {}
    for y in ["2022", "2023", "2024"]:
        src = f"{y}年數量(顆)"
        if src in df.columns:
            rename_map[src] = f"{y}年申報量(顆)"
    if rename_map:
        df = df.rename(columns=rename_map)

    for col in df.columns:
        if ("申報量" in col) or ("金額" in col):
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def parse_dosage(d):
    m = re.search(r"[\d.]+", str(d))
    return float(m.group()) if m else 0.0


def pretty_header(col: str):
    """把 '2022年申報量(顆)' 拆成 ('2022年','申報量')；'2022-2023年成長率(%)' 拆成 ('2022-2023年','成長率(%)')"""
    m = re.match(r"^(\d{4}(-\d{4})?年)(.+)$", col)
    if m:
        year = m.group(1)
        label = m.group(3).replace("(顆)", "")
        return year, label
    return None, col


# ============================================================
# 通用樞紐分析引擎：任意欄位順序 + 任意層級巢狀合計 + 占比/成長率
# ============================================================

def build_nested_rows(df: pd.DataFrame, row_fields: list, subtotal_fields: list, value_cols: list,
                       add_pct: bool = False, add_growth: bool = False):
    qty_cols = [c for c in value_cols if "申報量" in c]
    qty_years = sorted(set(c[:4] for c in qty_cols))

    pct_cols = [f"{y}年占比(%)" for y in qty_years] if add_pct else []
    growth_cols = [f"{qty_years[i]}-{qty_years[i+1]}年成長率(%)" for i in range(len(qty_years) - 1)] if add_growth else []

    def compute_extra_for_row(sums, top_totals):
        extra = {}
        for c in pct_cols:
            y = c[:4]
            base = f"{y}年申報量(顆)"
            extra[c] = sums.get(base, 0) / top_totals.get(base, 0) if top_totals.get(base, 0) else 0
        for c in growth_cols:
            y1 = c[:4]
            y2 = c[5:9]
            b1, b2 = f"{y1}年申報量(顆)", f"{y2}年申報量(顆)"
            extra[c] = (sums.get(b2, 0) - sums.get(b1, 0)) / sums.get(b1, 0) if sums.get(b1, 0) else 0
        return extra

    def compute_extra_for_group(totals):
        extra = {c: 1.0 for c in pct_cols}  # 小計/總計列本身佔比恆為 100%
        for c in growth_cols:
            y1 = c[:4]
            y2 = c[5:9]
            b1, b2 = f"{y1}年申報量(顆)", f"{y2}年申報量(顆)"
            extra[c] = (totals.get(b2, 0) - totals.get(b1, 0)) / totals.get(b1, 0) if totals.get(b1, 0) else 0
        return extra

    year_cols = sorted(qty_cols, reverse=True)
    sort_cols = list(row_fields) + year_cols
    sort_asc = [True] * len(row_fields) + [False] * len(year_cols)
    df_sorted = df.sort_values(by=sort_cols, ascending=sort_asc) if sort_cols else df

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

    def emit_row(row, top_totals):
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
        sums = {c: row[c] for c in value_cols}
        sums.update(compute_extra_for_row(sums, top_totals))
        rows.append({"type": "data", "values": rec_vals, "sums": sums})
        for c in first_flags:
            first_flags[c] = False

    def emit_subtotal(field, label, totals):
        rec_vals = {f: "" for f in row_fields}
        rec_vals[field] = f"{label} 合計"
        full = dict(totals)
        full.update(compute_extra_for_group(totals))
        rows.append({"type": "subtotal", "values": rec_vals, "sums": full})

    def recurse(sub_df, level_idx, top_totals):
        if level_idx >= len(subtotal_fields):
            for _, row in sub_df.iterrows():
                emit_row(row, top_totals)
            return
        col = subtotal_fields[level_idx]
        for _, grp in sub_df.groupby(col, sort=False):
            totals = {c: grp[c].sum() for c in value_cols}
            next_top = totals if level_idx == len(subtotal_fields) - 1 else top_totals
            label = grp[col].iloc[0]
            first_flags[col] = True
            recurse(grp, level_idx + 1, next_top)
            emit_subtotal(col, label, totals)

    grand_totals = {c: df[c].sum() for c in value_cols}
    if subtotal_fields:
        recurse(df_sorted, 0, grand_totals)
    else:
        for _, row in df_sorted.iterrows():
            emit_row(row, grand_totals)

    total_vals = {f: "" for f in row_fields}
    if row_fields:
        total_vals[row_fields[0]] = "總計"
    full_grand = dict(grand_totals)
    full_grand.update(compute_extra_for_group(grand_totals))
    rows.append({"type": "total", "values": total_vals, "sums": full_grand})

    return rows, pct_cols, growth_cols


# ============================================================
# 樣式化 HTML 預覽 (比照原系統的綠底標題、小計/總計配色)
# ============================================================

def build_html_table(rows, row_fields, value_cols, pct_cols, growth_cols, report_title):
    extra_cols = pct_cols + growth_cols
    headers = row_fields + value_cols + extra_cols

    th_style = "color:#FFFFFF;background-color:#00695C;text-align:center;border:1px solid #FFFFFF;padding:8px 10px;font-weight:bold;"
    td_style = "border:1px solid #D9D9D9;padding:6px 10px;"

    html = [
        "<div style='background:#FFFFFF;border-radius:10px;padding:18px;border:1px solid #eee;'>",
        f"<h3 style='text-align:center;color:#004D40;margin-top:0;'>{report_title}</h3>",
        "<div style='overflow-x:auto;'>",
        "<table id='report-table' style='border-collapse:separate;border-spacing:0;width:100%;font-size:14px;'>",
        "<tr>",
    ]
    for f in row_fields:
        html.append(f"<th style='{th_style}'>{f}</th>")
    for c in value_cols + extra_cols:
        year, label = pretty_header(c)
        if year:
            html.append(f"<th style='{th_style}'>{year}<br>{label}</th>")
        else:
            html.append(f"<th style='{th_style}'>{label}</th>")
    html.append("</tr>")

    for r in rows:
        bg = ""
        fw = "font-weight:normal;"
        if r["type"] == "subtotal":
            bg = "background-color:#E0F2F1;"
            fw = "font-weight:bold;"
        elif r["type"] == "total":
            bg = "background-color:#B2DFDB;"
            fw = "font-weight:bold;"
        html.append(f"<tr style='{bg}{fw}'>")
        for i, f in enumerate(row_fields):
            v = r["values"].get(f, "")
            align = "left" if i == 0 else "center"
            html.append(f"<td style='{td_style}text-align:{align};'>{v}</td>")
        for c in value_cols:
            v = r["sums"].get(c, "")
            html.append(f"<td style='{td_style}text-align:right;'>{v:,.0f}</td>" if v != "" else f"<td style='{td_style}'></td>")
        for c in extra_cols:
            v = r["sums"].get(c, 0)
            html.append(f"<td style='{td_style}text-align:right;'>{v:.1%}</td>")
        html.append("</tr>")

    html.append("</table></div></div>")
    return "".join(html)


CAPTURE_HTML_TEMPLATE = """
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<div id="capture-wrap">{table_html}</div>
<div style="text-align:center; margin-top:12px;">
  <button id="export-btn" style="background-color:#00695C;color:white;border:none;padding:10px 20px;
    border-radius:6px;font-size:14px;cursor:pointer;">🖼️ 匯出為 PNG 圖片</button>
</div>
<script>
document.getElementById('export-btn').addEventListener('click', function() {{
    const target = document.getElementById('capture-wrap');
    html2canvas(target, {{scale: 2, backgroundColor: '#ffffff'}}).then(function(canvas) {{
        const link = document.createElement('a');
        link.download = '{filename}.png';
        link.href = canvas.toDataURL('image/png');
        link.click();
    }});
}});
</script>
"""


# ============================================================
# Excel 匯出 (比照原系統的綠底標題、A4 橫式、標題列重複列印)
# ============================================================

def generate_excel_bytes(rows, row_fields, value_cols, pct_cols, growth_cols, report_title):
    extra_cols = pct_cols + growth_cols
    headers = row_fields + value_cols + extra_cols

    wb = Workbook()
    ws = wb.active
    ws.title = "分析報表"

    ws.views.sheetView[0].showGridLines = False
    ws.views.sheetView[0].zoomScale = 85
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True

    header_fill = PatternFill(start_color="00695C", end_color="00695C", fill_type="solid")
    subtotal_fill = PatternFill(start_color="E0F2F1", end_color="E0F2F1", fill_type="solid")
    total_fill = PatternFill(start_color="B2DFDB", end_color="B2DFDB", fill_type="solid")
    font_title = Font(name="微軟正黑體", bold=True, size=16, color="004D40")
    font_b_w = Font(name="微軟正黑體", bold=True, color="FFFFFF", size=12)
    font_norm = Font(name="微軟正黑體", size=11)
    font_bold = Font(name="微軟正黑體", bold=True, size=11)
    align_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_r = Alignment(horizontal="right", vertical="center")
    border_thin = Border(*[Side(style="thin", color="D9D9D9")] * 4)
    header_border = Border(*[Side(style="thin", color="FFFFFF")] * 4)

    ws.cell(row=1, column=1, value=report_title).font = font_title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(len(headers), 1))
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center")

    header_row = 3
    ws.append([])

    header_texts = []
    for f in row_fields:
        header_texts.append(f)
    for c in value_cols + extra_cols:
        year, label = pretty_header(c)
        header_texts.append(f"{year}\n{label}" if year else label)
    ws.append(header_texts)
    ws.row_dimensions[header_row].height = 48
    ws.print_title_rows = f"{header_row}:{header_row}"

    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=header_row, column=c)
        cell.font = font_b_w
        cell.fill = header_fill
        cell.border = header_border
        cell.alignment = align_c

    current_row = header_row + 1
    for r in rows:
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=current_row, column=i)
            if h in value_cols:
                v = r["sums"].get(h, "")
                cell.value = v
                cell.number_format = "#,##0"
                cell.alignment = align_r
            elif h in extra_cols:
                cell.value = r["sums"].get(h, 0)
                cell.number_format = "0.0%"
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
        else:
            ws.column_dimensions[col_letter].width = 16

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ============================================================
# 處方釋出率分析（固定格式）
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

ANALYSIS_TO_SOURCE = {
    "1. 廠商分析": "data_hp_gp_ds.csv",
    "2. 層級別分析": "data_hp_gp_ds.csv",
    "3. 科別分析": "data_department.csv",
    "4. 推估醫院分析": "data_hospital.csv",
}
ANA_OPTIONS = list(ANALYSIS_TO_SOURCE.keys()) + ["5. 處方釋出率分析"]

ana_choice = st.selectbox("第1步：選擇分析功能", ANA_OPTIONS, key="ana_choice")

# ------------------------------------------------------------
# 1~4：自訂樞紐分析
# ------------------------------------------------------------
if ana_choice in ANALYSIS_TO_SOURCE:
    st.caption("先選擇成分，接著把需要的欄位拖到「報表欄位」，即時預覽會隨著您的設定更新，效果如同 Excel 樞紐分析。")

    source_path = ANALYSIS_TO_SOURCE[ana_choice]
    df_raw = load_data(source_path)

    if df_raw.empty:
        st.warning(f"找不到資料檔案「{source_path}」，請確認檔案已放置於工作目錄。")
    elif "成分簡稱" not in df_raw.columns:
        st.error(f"資料檔案缺少「成分簡稱」欄位，實際欄位為：{list(df_raw.columns)}")
    else:
        comp_options = sorted([c for c in df_raw["成分簡稱"].dropna().unique() if c])
        # 單一元件即支援輸入時自動篩選（Streamlit multiselect 內建 type-to-search）
        comps_selected = st.multiselect(
            "第2步：輸入關鍵字並選擇成分品項 (必選)",
            options=comp_options,
            key=f"comps_{ana_choice}",
            placeholder="輸入關鍵字，例如 Levofloxacin",
        )

        if not comps_selected:
            st.info("請先選擇至少一個成分，才會顯示後續的欄位設定。")
        else:
            df_comp = df_raw[df_raw["成分簡稱"].isin(comps_selected)]

            other_cols = [c for c in df_raw.columns if c != "成分簡稱"]
            value_cols_auto = [c for c in other_cols if ("申報量" in c) or ("金額" in c)]
            dim_cols_all = [c for c in other_cols if c not in value_cols_auto]

            st.markdown("### 🧩 第3步：拖曳欄位到「報表欄位」，並排序")
            dnd_key = f"dnd_{ana_choice}_{'_'.join(sorted(comps_selected))}"

            if HAS_SORTABLES:
                containers = sort_items(
                    [
                        {"header": "📋 可用欄位 (拖曳到下方使用)", "items": dim_cols_all},
                        {"header": "📊 報表欄位 (由左到右排列；如同樞紐分析表的欄位區)", "items": []},
                    ],
                    multi_containers=True,
                    direction="horizontal",
                    key=dnd_key,
                )
                row_fields = containers[1]["items"] if containers and len(containers) > 1 else []
            else:
                st.error("尚未安裝 streamlit-sortables 套件，暫以勾選方式呈現，請於 requirements.txt 加入 streamlit-sortables 後重新部署即可拖曳。")
                row_fields = st.multiselect("選擇報表欄位", options=dim_cols_all, key=f"fallback_fields_{ana_choice}")

            if not row_fields:
                st.info("請至少拖曳一個欄位到「報表欄位」區塊。")
            else:
                st.markdown("### 🔍 第4步：逐欄篩選與合計 (點欄位下方的 🔽 展開篩選；如同 Excel 欄篩選)")
                filter_ui_cols = st.columns(len(row_fields))
                filters = {}
                subtotal_fields_selected = []
                for i, f in enumerate(row_fields):
                    with filter_ui_cols[i]:
                        st.markdown(f"**{f}**")
                        with st.popover("🔽 篩選 / 合計", use_container_width=True):
                            options = sorted([v for v in df_comp[f].dropna().unique() if str(v).strip() != ""])
                            sel = st.multiselect(f"篩選「{f}」(不選代表全選)", options=options, key=f"filt_{dnd_key}_{f}")
                            if sel:
                                filters[f] = sel
                            if st.checkbox(f"Σ 此欄要合計", key=f"sub_{dnd_key}_{f}"):
                                subtotal_fields_selected.append(f)
                subtotal_fields = [f for f in row_fields if f in subtotal_fields_selected]

                df_filtered = df_comp.copy()
                for col, sel in filters.items():
                    df_filtered = df_filtered[df_filtered[col].isin(sel)]

                st.caption(f"篩選後共 **{len(df_filtered):,}** 筆資料")

                c_val1, c_val2 = st.columns([2, 1])
                with c_val1:
                    value_cols = st.multiselect(
                        "第5步：選擇要加總的數值欄位 (預設帶入所有申報量／金額欄位)",
                        options=value_cols_auto, default=value_cols_auto, key=f"vals_{dnd_key}",
                    )
                with c_val2:
                    has_qty = any("申報量" in c for c in value_cols)
                    add_pct = st.checkbox("➕ 加入年度占比(%)", value=False, disabled=not has_qty, key=f"pct_{dnd_key}")
                    add_growth = st.checkbox("➕ 加入年度成長率(%)", value=False, disabled=not has_qty, key=f"growth_{dnd_key}")

                # 即時預覽（隨拖曳/篩選/勾選即時更新，不需按產生報表）
                full_row_fields = ["成分簡稱"] + row_fields
                report_title = f"{'_'.join(comps_selected)}_{ana_choice.split('.')[1].strip()}"

                if value_cols and not df_filtered.empty:
                    rows, pct_cols, growth_cols = build_nested_rows(
                        df_filtered, full_row_fields, subtotal_fields, value_cols, add_pct, add_growth
                    )
                    st.markdown("### 📄 報表即時預覽")
                    table_html = build_html_table(rows, full_row_fields, value_cols, pct_cols, growth_cols, report_title)
                    st.markdown(table_html, unsafe_allow_html=True)

                    st.markdown("### 📥 下載")
                    dl_col1, dl_col2 = st.columns(2)
                    with dl_col1:
                        excel_bytes = generate_excel_bytes(rows, full_row_fields, value_cols, pct_cols, growth_cols, report_title)
                        st.download_button(
                            "📄 下載 Excel 報表",
                            data=excel_bytes,
                            file_name=f"{report_title}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                    with dl_col2:
                        with st.expander("🖼️ 匯出為 PNG 圖片", expanded=False):
                            components.html(
                                CAPTURE_HTML_TEMPLATE.format(table_html=table_html, filename=report_title),
                                height=min(900, 200 + 40 * len(rows)),
                                scrolling=True,
                            )
                elif not value_cols:
                    st.warning("⚠️ 請至少選擇一個要加總的數值欄位。")
                else:
                    st.warning("❌ 篩選後無資料，請放寬篩選條件。")

# ------------------------------------------------------------
# 5：處方釋出率分析（固定格式）
# ------------------------------------------------------------
else:
    st.caption("分析成分從醫院端流向藥局處方的釋出率，可看單一廠商。此分析維持固定格式，不提供自訂欄位。")

    df_hp = load_data("data_hp_gp_ds.csv")
    if df_hp.empty:
        st.warning("找不到資料檔案「data_hp_gp_ds.csv」，請確認檔案已放置於工作目錄。")
    elif "成分簡稱" not in df_hp.columns:
        st.error(f"資料檔案缺少「成分簡稱」欄位，實際欄位為：{list(df_hp.columns)}")
    else:
        comp_options = sorted([c for c in df_hp["成分簡稱"].dropna().unique() if c])
        comp = st.selectbox(
            "第2步：輸入關鍵字並選擇單一成分品項 (必選)",
            options=[""] + comp_options, key="presc_comp",
        )

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
                    st.markdown(
                        f"""
                        <div style='background-color:#E0F2F1; padding:20px; border-radius:10px;
                            border-left: 6px solid #00695C;'>
                            <h3 style='margin-top:0; color:#004D40;'>📊 {summary['title']}</h3>
                            <ul style='font-size:18px; color:#004D40; list-style-type:none; padding-left:0; line-height:1.6;'>
                                <li><b>DS from HP：</b> {summary['ds_hp_val']:,.0f}</li>
                                <li><b>HP：</b> {summary['hp_val']:,.0f}</li>
                                <li><b>DS from HP + HP：</b> {summary['total_val']:,.0f}</li>
                                <li style='margin-top:6px; font-size:13px; color:#00695C;'>
                                    公式：處方釋出率 ＝ DS from HP ÷ (DS from HP + HP) × 100%</li>
                                <li style='margin-top:10px; font-size:26px;'><b>{summary['title']}：</b>
                                    <span style='color:#D32F2F; font-weight:bold;'>{summary['rate']:.2f}%</span></li>
                            </ul>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
