
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Merly - Tồn kho", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 2rem; max-width: 96%;}
h1, h2, h3 {color: #9d8479;}
small, .stCaption {color: #7d6a61 !important;}
.pivot-wrap {overflow-x: auto; background: white; border: 1px solid #e8dfdb; border-radius: 12px; padding: 8px;}
.pivot-table {border-collapse: collapse; width: 100%; min-width: 1000px; font-size: 14px;}
.pivot-table th, .pivot-table td {border: 1px solid #d9d9d9; padding: 6px 8px; text-align: center; white-space: nowrap;}
.pivot-table thead th {background: #f3ece8; color: #7c655b; font-weight: 700;}
.pivot-col-left {text-align: left !important;}
.ma-row {background: #f8f1ed; font-weight: 700; color: #6f5a51;}
.mau-row td:first-child {padding-left: 22px;}
.row-alert td {background: #ffd9d9 !important; color: #9b0000; font-weight: 700;}
.total-col {background: #fff2cc; font-weight: 700;}
.grand-total-row td {background: #ead7cf !important; font-weight: 700; color: #5d4d46;}
</style>
""", unsafe_allow_html=True)

col_logo, col_title = st.columns([1, 4])
with col_logo:
    st.image("Logo Merly.jpg", width=130)
with col_title:
    st.title("Merly - Hệ thống phân tích tồn kho")
    st.caption("Upload file tồn kho từ hệ thống để tách dữ liệu, xem pivot trực tiếp và ra quyết định nhập hàng / khuyến mãi.")

uploaded_file = st.file_uploader("Tải file Excel", type=["xlsx"])

def split_data(text):
    text = str(text).strip()
    if text == "":
        return pd.Series(["", "", ""])
    parts = text.split(" ")
    if len(parts) < 2:
        return pd.Series(["", "", ""])
    ma = parts[0]
    size = ""
    mau = ""
    for p in parts[1:]:
        if p.isdigit() and len(p) == 2:
            size = p
        else:
            mau += (" " if mau else "") + p
    return pd.Series([ma, size, mau.strip()])

def build_pivot_like_excel(df_clean):
    df2 = df_clean.copy()
    df2["Size"] = df2["Size"].astype(str)
    df2["Ton kho"] = pd.to_numeric(df2["Ton kho"], errors="coerce").fillna(0).astype(int)
    all_sizes = sorted([s for s in df2["Size"].dropna().unique() if str(s).isdigit()], key=lambda x: int(x))
    pivot = pd.pivot_table(df2, index=["Ma SP", "Mau"], columns="Size", values="Ton kho", aggfunc="sum", fill_value=0)
    for s in all_sizes:
        if s not in pivot.columns:
            pivot[s] = 0
    pivot = pivot[all_sizes]
    pivot["Grand Total"] = pivot.sum(axis=1)
    pivot = pivot.reset_index()
    ma_total = (df2.groupby("Ma SP")["Ton kho"].sum().reset_index().rename(columns={"Ton kho": "Grand Total"}))
    return pivot, ma_total, all_sizes

def render_pivot_html(pivot, ma_total, all_sizes):
    html = []
    html.append('<div class="pivot-wrap">')
    html.append('<table class="pivot-table">')
    html.append("<thead><tr>")
    html.append('<th class="pivot-col-left">Row Labels</th>')
    for s in all_sizes:
        html.append(f"<th>{s}</th>")
    html.append('<th class="total-col">Grand Total</th>')
    html.append("</tr></thead><tbody>")
    ma_list = pivot["Ma SP"].dropna().unique().tolist()
    grand_total_all = 0
    for ma in ma_list:
        block = pivot[pivot["Ma SP"] == ma].copy()
        ma_match = ma_total.loc[ma_total["Ma SP"] == ma, "Grand Total"]
        ma_sum = int(ma_match.iloc[0]) if not ma_match.empty else 0
        grand_total_all += ma_sum
        html.append('<tr class="ma-row">')
        html.append(f'<td class="pivot-col-left">◉ {ma}</td>')
        for s in all_sizes:
            v = int(block[s].sum()) if s in block.columns else 0
            html.append(f'<td>{"" if v == 0 else v}</td>')
        html.append(f'<td class="total-col">{ma_sum}</td>')
        html.append("</tr>")
        for _, row in block.iterrows():
            is_alert = False
            row_cells = []
            for s in all_sizes:
                v = int(row[s]) if s in row.index and pd.notnull(row[s]) else 0
                if v == 1:
                    is_alert = True
                row_cells.append("" if v == 0 else str(v))
            gt = int(row["Grand Total"]) if pd.notnull(row["Grand Total"]) else 0
            if gt == 1:
                is_alert = True
            tr_class = "mau-row row-alert" if is_alert else "mau-row"
            html.append(f'<tr class="{tr_class}">')
            html.append(f'<td class="pivot-col-left">{row["Mau"]}</td>')
            for cell in row_cells:
                html.append(f"<td>{cell}</td>")
            html.append(f'<td class="total-col">{gt if gt != 0 else ""}</td>')
            html.append("</tr>")
    html.append('<tr class="grand-total-row">')
    html.append('<td class="pivot-col-left">Grand Total</td>')
    for s in all_sizes:
        col_total = int(pivot[s].sum()) if s in pivot.columns else 0
        html.append(f'<td>{"" if col_total == 0 else col_total}</td>')
    html.append(f'<td class="total-col">{grand_total_all}</td>')
    html.append("</tr></tbody></table></div>")
    return "".join(html)

def to_excel_file(df_clean, pivot_detail, all_sizes):
    output = BytesIO()
    export_cols = ["Ma SP", "Mau"] + all_sizes + ["Grand Total"]
    pivot_export = pivot_detail[export_cols].copy()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_clean.to_excel(writer, sheet_name="DuLieuDaTach", index=False)
        pivot_export.to_excel(writer, sheet_name="PivotTonKho", index=False)
    output.seek(0)
    return output

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    if df.shape[1] < 8:
        st.error("File hiện không đúng cấu trúc tối thiểu. App đang cần ít nhất 8 cột, trong đó: A = Tên hàng hóa, C = Tồn kho, E = Giá bán, H = Group.")
        st.stop()

    df[["Ma SP", "Size", "Mau"]] = df.iloc[:, 0].apply(split_data)
    df["Gia Ban"] = df.iloc[:, 4]
    df["Ton kho"] = df.iloc[:, 2]
    df["Group"] = df.iloc[:, 7]

    df_clean = df[["Ma SP", "Size", "Mau", "Gia Ban", "Ton kho", "Group"]].copy()
    df_clean["Size"] = df_clean["Size"].astype(str)
    df_clean["Ton kho"] = pd.to_numeric(df_clean["Ton kho"], errors="coerce").fillna(0).astype(int)

    st.subheader("Bộ lọc")
    colf1, colf2 = st.columns(2)
    all_groups = sorted([g for g in df_clean["Group"].dropna().astype(str).unique().tolist() if g.strip() != ""])
    all_masp = sorted([m for m in df_clean["Ma SP"].dropna().astype(str).unique().tolist() if m.strip() != ""])

    with colf1:
        selected_groups = st.multiselect("Lọc theo Group", all_groups, default=all_groups)
    with colf2:
        selected_masp = st.multiselect("Lọc theo Mã SP", all_masp, default=all_masp)

    df_filtered = df_clean.copy()
    if all_groups:
        df_filtered = df_filtered[df_filtered["Group"].astype(str).isin(selected_groups)]
    if all_masp:
        df_filtered = df_filtered[df_filtered["Ma SP"].astype(str).isin(selected_masp)]

    st.subheader("Dữ liệu đã tách")
    st.dataframe(df_filtered, use_container_width=True, height=300)

    st.subheader("Pivot tồn kho")
    st.caption("Ẩn các ô = 0. Dòng màu nào có tồn = 1 sẽ được tô đỏ.")
    pivot_detail, ma_total, all_sizes = build_pivot_like_excel(df_filtered)
    pivot_html = render_pivot_html(pivot_detail, ma_total, all_sizes)
    st.markdown(pivot_html, unsafe_allow_html=True)

    st.subheader("Gợi ý hành động")
    col1, col2 = st.columns(2)
    with col1:
        st.write("Hàng tồn thấp (<= 2)")
        low_stock = df_filtered[df_filtered["Ton kho"] <= 2].sort_values(["Ton kho", "Ma SP", "Mau", "Size"])
        st.dataframe(low_stock, use_container_width=True, height=300)
    with col2:
        st.write("Hàng tồn cao (>= 10)")
        high_stock = df_filtered[df_filtered["Ton kho"] >= 10].sort_values(["Ton kho"], ascending=False)
        st.dataframe(high_stock, use_container_width=True, height=300)

    excel_file = to_excel_file(df_filtered, pivot_detail, all_sizes)
    st.download_button("Tải file Excel kết quả", data=excel_file, file_name="merly_tonkho_ketqua.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Hãy tải file Excel lên để bắt đầu.")
