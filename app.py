
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Merly - Tồn kho Pro", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 2rem; max-width: 96%;}
h1, h2, h3 {color: #9d8479;}
small, .stCaption {color: #7d6a61 !important;}
.filter-box {background: #f8f1ed; border: 1px solid #eadfd9; border-radius: 14px; padding: 12px 14px; margin-bottom: 10px;}
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
    st.title("Merly - Hệ thống phân tích tồn kho Pro")
    st.caption("Upload file tồn kho từ hệ thống để tách dữ liệu, xem dashboard, pivot trực tiếp và ra quyết định nhập hàng / khuyến mãi.")

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

def classify_ton_kho(qty):
    qty = int(qty) if pd.notnull(qty) else 0
    if qty <= 1:
        return "Cần nhập gấp"
    elif qty <= 3:
        return "Sắp hết"
    elif qty >= 10:
        return "Tồn cao"
    return "Bình thường"

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
    ma_total = df2.groupby("Ma SP")["Ton kho"].sum().reset_index().rename(columns={"Ton kho":"Grand Total"})
    ma_total = ma_total[ma_total["Grand Total"] > 0].copy()
    valid_ma = ma_total["Ma SP"].astype(str).tolist()
    pivot = pivot[pivot["Ma SP"].astype(str).isin(valid_ma)].copy()
    return pivot, ma_total, all_sizes

def render_pivot_html(pivot, ma_total, all_sizes):
    if pivot.empty:
        return '<div class="pivot-wrap"><p style="padding:12px;color:#7c655b;">Không có dữ liệu phù hợp.</p></div>'
    html = []
    html.append('<div class="pivot-wrap"><table class="pivot-table">')
    html.append('<thead><tr><th class="pivot-col-left">Row Labels</th>')
    for s in all_sizes:
        html.append(f"<th>{s}</th>")
    html.append('<th class="total-col">Grand Total</th></tr></thead><tbody>')
    ma_list = pivot["Ma SP"].dropna().astype(str).unique().tolist()
    grand_total_all = 0
    for ma in ma_list:
        block = pivot[pivot["Ma SP"].astype(str) == ma].copy()
        ma_match = ma_total.loc[ma_total["Ma SP"].astype(str) == ma, "Grand Total"]
        ma_sum = int(ma_match.iloc[0]) if not ma_match.empty else 0
        if ma_sum == 0:
            continue
        grand_total_all += ma_sum
        html.append('<tr class="ma-row">')
        html.append(f'<td class="pivot-col-left">◉ {ma}</td>')
        for s in all_sizes:
            v = int(block[s].sum()) if s in block.columns else 0
            html.append(f'<td>{"" if v == 0 else v}</td>')
        html.append(f'<td class="total-col">{ma_sum}</td></tr>')
        block = block[block["Grand Total"] > 0].copy()
        for _, row in block.iterrows():
            is_alert = False
            cells = []
            for s in all_sizes:
                v = int(row[s]) if s in row.index and pd.notnull(row[s]) else 0
                if v == 1:
                    is_alert = True
                cells.append("" if v == 0 else str(v))
            gt = int(row["Grand Total"]) if pd.notnull(row["Grand Total"]) else 0
            if gt == 1:
                is_alert = True
            tr_class = "mau-row row-alert" if is_alert else "mau-row"
            html.append(f'<tr class="{tr_class}"><td class="pivot-col-left">{row["Mau"]}</td>')
            for cell in cells:
                html.append(f"<td>{cell}</td>")
            html.append(f'<td class="total-col">{gt if gt != 0 else ""}</td></tr>')
    html.append('<tr class="grand-total-row"><td class="pivot-col-left">Grand Total</td>')
    for s in all_sizes:
        col_total = int(pivot[s].sum()) if s in pivot.columns else 0
        html.append(f'<td>{"" if col_total == 0 else col_total}</td>')
    html.append(f'<td class="total-col">{grand_total_all}</td></tr></tbody></table></div>')
    return "".join(html)

def to_excel_file(df_clean, pivot_detail, all_sizes, summary_ma, summary_group):
    output = BytesIO()
    export_cols = ["Ma SP", "Mau"] + all_sizes + ["Grand Total"]
    pivot_export = pivot_detail[export_cols].copy()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_clean.to_excel(writer, sheet_name="DuLieuDaTach", index=False)
        pivot_export.to_excel(writer, sheet_name="PivotTonKho", index=False)
        summary_ma.to_excel(writer, sheet_name="TongHopTheoMa", index=False)
        summary_group.to_excel(writer, sheet_name="TongHopTheoGroup", index=False)
    output.seek(0)
    return output

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    if df.shape[1] < 8:
        st.error("File cần ít nhất 8 cột: A = Tên hàng hóa, C = Tồn kho, E = Giá bán, H = Group.")
        st.stop()

    df[["Ma SP","Size","Mau"]] = df.iloc[:,0].apply(split_data)
    df["Gia Ban"] = pd.to_numeric(df.iloc[:,4], errors="coerce").fillna(0)
    df["Ton kho"] = pd.to_numeric(df.iloc[:,2], errors="coerce").fillna(0).astype(int)
    df["Group"] = df.iloc[:,7].astype(str).fillna("")
    df_clean = df[["Ma SP","Size","Mau","Gia Ban","Ton kho","Group"]].copy()
    df_clean["Phan loai"] = df_clean["Ton kho"].apply(classify_ton_kho)
    df_clean["Gia tri ton"] = df_clean["Gia Ban"] * df_clean["Ton kho"]

    st.markdown('<div class="filter-box">', unsafe_allow_html=True)
    st.subheader("Bộ lọc")
    col1, col2, col3 = st.columns(3)
    groups = sorted([g for g in df_clean["Group"].dropna().unique().tolist() if str(g).strip() != ""])
    masp = sorted([m for m in df_clean["Ma SP"].dropna().astype(str).unique().tolist() if m.strip() != ""])
    with col1:
        selected_groups = st.multiselect("Chọn nhóm sản phẩm", groups, default=groups)
    with col2:
        selected_masp = st.multiselect("Chọn Mã SP", masp, default=masp)
    with col3:
        selected_status = st.selectbox("Chế độ xem", ["Tất cả", "Cần nhập gấp", "Sắp hết", "Bình thường", "Tồn cao"])
    st.markdown('</div>', unsafe_allow_html=True)

    df_filtered = df_clean.copy()
    if selected_groups:
        df_filtered = df_filtered[df_filtered["Group"].isin(selected_groups)]
    else:
        df_filtered = df_filtered.iloc[0:0]
    if selected_masp:
        df_filtered = df_filtered[df_filtered["Ma SP"].astype(str).isin(selected_masp)]
    else:
        df_filtered = df_filtered.iloc[0:0]
    if selected_status != "Tất cả":
        df_filtered = df_filtered[df_filtered["Phan loai"] == selected_status]
    df_filtered = df_filtered[df_filtered["Ton kho"] > 0].copy()

    st.subheader("Tổng quan nhanh")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Tổng biến thể", f"{len(df_filtered):,}")
    m2.metric("Tổng mã SP", f"{df_filtered['Ma SP'].nunique():,}")
    m3.metric("Tổng tồn kho", f"{int(df_filtered['Ton kho'].sum()) if not df_filtered.empty else 0:,}")
    m4.metric("Giá trị tồn", f"{float(df_filtered['Gia tri ton'].sum()) if not df_filtered.empty else 0:,.0f}")
    m5, m6, m7 = st.columns(3)
    m5.metric("Cần nhập gấp", f"{len(df_filtered[df_filtered['Ton kho'] <= 1]):,}")
    m6.metric("Sắp hết", f"{len(df_filtered[(df_filtered['Ton kho'] > 1) & (df_filtered['Ton kho'] <= 3)]):,}")
    m7.metric("Tồn cao", f"{len(df_filtered[df_filtered['Ton kho'] >= 10]):,}")

    st.subheader("Dashboard phân tích")
    c1, c2 = st.columns(2)
    with c1:
        st.write("Tồn kho theo nhóm sản phẩm")
        group_chart = df_filtered.groupby("Group", as_index=False)["Ton kho"].sum().sort_values("Ton kho", ascending=False).set_index("Group")
        if not group_chart.empty:
            st.bar_chart(group_chart)
        else:
            st.info("Không có dữ liệu.")
    with c2:
        st.write("Top 10 mã có tồn cao nhất")
        top_ma_chart = df_filtered.groupby("Ma SP", as_index=False)["Ton kho"].sum().sort_values("Ton kho", ascending=False).head(10).set_index("Ma SP")
        if not top_ma_chart.empty:
            st.bar_chart(top_ma_chart)
        else:
            st.info("Không có dữ liệu.")

    summary_ma = df_filtered.groupby("Ma SP", as_index=False).agg({"Ton kho":"sum","Gia tri ton":"sum"}).sort_values(["Ton kho","Gia tri ton"], ascending=[False, False])
    summary_group = df_filtered.groupby("Group", as_index=False).agg({"Ton kho":"sum","Gia tri ton":"sum"}).sort_values(["Ton kho","Gia tri ton"], ascending=[False, False])

    st.subheader("Top mã cần chú ý")
    st.dataframe(summary_ma.head(20), use_container_width=True, height=320)

    st.subheader("Dữ liệu đã tách")
    st.dataframe(df_filtered[["Ma SP","Size","Mau","Gia Ban","Ton kho","Group","Phan loai","Gia tri ton"]], use_container_width=True, height=320)

    st.subheader("Pivot tồn kho")
    st.caption("Ẩn các ô = 0, ẩn toàn bộ sản phẩm có tổng tồn = 0, và tô đỏ dòng có tồn = 1.")
    pivot_detail, ma_total, all_sizes = build_pivot_like_excel(df_filtered)
    st.markdown(render_pivot_html(pivot_detail, ma_total, all_sizes), unsafe_allow_html=True)

    st.subheader("Gợi ý hành động")
    a1, a2 = st.columns(2)
    with a1:
        st.write("Danh sách cần nhập thêm")
        st.dataframe(df_filtered[df_filtered["Ton kho"] <= 2].sort_values(["Ton kho","Ma SP","Mau","Size"]), use_container_width=True, height=320)
    with a2:
        st.write("Danh sách nên cân nhắc chạy sale")
        st.dataframe(df_filtered[df_filtered["Ton kho"] >= 10].sort_values(["Ton kho","Gia tri ton"], ascending=[False, False]), use_container_width=True, height=320)

    st.subheader("Điểm ưu tiên xử lý")
    priority_df = df_filtered.copy()
    priority_df["Diem uu tien"] = priority_df.apply(lambda r: (100 if r["Ton kho"] <= 1 else 60 if r["Ton kho"] <= 3 else 10 if r["Ton kho"] >= 10 else 30) + (r["Gia tri ton"] / 1000000), axis=1)
    priority_df = priority_df.sort_values("Diem uu tien", ascending=False)
    st.dataframe(priority_df[["Ma SP","Size","Mau","Ton kho","Gia Ban","Gia tri ton","Group","Phan loai","Diem uu tien"]].head(50), use_container_width=True, height=320)

    excel_file = to_excel_file(df_filtered, pivot_detail, all_sizes, summary_ma, summary_group)
    st.download_button("Tải file Excel kết quả", data=excel_file, file_name="merly_tonkho_pro.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Hãy tải file Excel lên để bắt đầu.")
