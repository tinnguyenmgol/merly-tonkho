
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Merly - Tồn kho Pro+", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 2rem; max-width: 96%;}
h1, h2, h3 {color: #9d8479;}
small, .stCaption {color: #7d6a61 !important;}
.filter-box {background: #f8f1ed; border: 1px solid #eadfd9; border-radius: 14px; padding: 12px 14px; margin-bottom: 10px;}
.pivot-wrap {overflow-x: auto; background: white; border: 1px solid #e8dfdb; border-radius: 12px; padding: 8px;}
.pivot-table {border-collapse: collapse; width: 100%; min-width: 1100px; font-size: 14px;}
.pivot-table th, .pivot-table td {border: 1px solid #d9d9d9; padding: 6px 8px; text-align: center; white-space: nowrap;}
.pivot-table thead th {background: #f3ece8; color: #7c655b; font-weight: 700; position: sticky; top: 0; z-index: 2;}
.pivot-col-left {text-align: left !important;}
.group-row td {background: #efe3dc; font-weight: 700; color: #6f5a51;}
.ma-row td {background: #f8f1ed; font-weight: 700; color: #6f5a51;}
.mau-row td:first-child {padding-left: 28px;}
.row-alert td {background: #ffd9d9 !important; color: #9b0000; font-weight: 700;}
.row-low td {background: #fff5cf !important; color: #8a6d00;}
.total-col {background: #fff2cc; font-weight: 700;}
.badge-box {margin-top: 6px; margin-bottom: 6px; display:flex; gap:10px; flex-wrap:wrap;}
.badge {display:inline-block; padding:4px 10px; border-radius:999px; font-size:12px; font-weight:700;}
.badge-red {background:#ffd9d9; color:#9b0000;}
.badge-yellow {background:#fff2cc; color:#8a6d00;}
.badge-blue {background:#dceeff; color:#0b5cab;}
details.pivot-group, details.pivot-ma {margin-bottom:0;}
details.pivot-group summary, details.pivot-ma summary {
    list-style: none;
    cursor: pointer;
    padding: 0;
    outline: none;
}
details.pivot-group summary::-webkit-details-marker,
details.pivot-ma summary::-webkit-details-marker {display:none;}
.kpi-card {background: #fff; border:1px solid #eadfd9; border-radius:14px; padding:14px;}
</style>
""", unsafe_allow_html=True)

col_logo, col_title = st.columns([1, 4])
with col_logo:
    st.image("Logo Merly.jpg", width=130)
with col_title:
    st.title("Merly - Hệ thống phân tích tồn kho Pro+")
    st.caption("Pivot dạng nhóm giống Excel, có mở/đóng theo Group và Mã SP, ẩn 0 và blank, tô màu theo mức tồn.")

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

def build_pivot_hierarchical(df_clean):
    df2 = df_clean.copy()
    df2["Size"] = df2["Size"].astype(str)
    df2["Ton kho"] = pd.to_numeric(df2["Ton kho"], errors="coerce").fillna(0).astype(int)
    df2["Group"] = df2["Group"].astype(str).fillna("").str.strip()
    df2["Ma SP"] = df2["Ma SP"].astype(str).fillna("").str.strip()
    df2["Mau"] = df2["Mau"].astype(str).fillna("").str.strip()

    df2 = df2[df2["Ton kho"] > 0].copy()
    df2 = df2[(df2["Group"] != "") & (df2["Ma SP"] != "") & (df2["Mau"] != "")].copy()

    all_sizes = sorted(
        [s for s in df2["Size"].dropna().unique() if str(s).isdigit()],
        key=lambda x: int(x)
    )

    pivot = pd.pivot_table(
        df2,
        index=["Group", "Ma SP", "Mau"],
        columns="Size",
        values="Ton kho",
        aggfunc="sum",
        fill_value=0
    )

    for s in all_sizes:
        if s not in pivot.columns:
            pivot[s] = 0

    pivot = pivot[all_sizes]
    pivot["Grand Total"] = pivot.sum(axis=1)
    pivot = pivot[pivot["Grand Total"] > 0].reset_index()

    group_total = df2.groupby("Group", as_index=False)["Ton kho"].sum().rename(columns={"Ton kho":"Grand Total"})
    ma_total = df2.groupby(["Group", "Ma SP"], as_index=False)["Ton kho"].sum().rename(columns={"Ton kho":"Grand Total"})

    return pivot, group_total, ma_total, all_sizes

def render_pivot_html(pivot, group_total, ma_total, all_sizes):
    if pivot.empty:
        return '<div class="pivot-wrap"><p style="padding:12px;color:#7c655b;">Không có dữ liệu phù hợp.</p></div>'

    html = []
    html.append('<div class="badge-box">')
    html.append('<span class="badge badge-red">Tồn = 1</span>')
    html.append('<span class="badge badge-yellow">Tồn 2–3</span>')
    html.append('<span class="badge badge-blue">Pivot đã ẩn ô 0, blank và mã tổng = 0</span>')
    html.append('</div>')
    html.append('<div class="pivot-wrap">')
    html.append('<table class="pivot-table">')
    html.append('<thead><tr><th class="pivot-col-left">Row Labels</th>')
    for s in all_sizes:
        html.append(f'<th>{s}</th>')
    html.append('<th class="total-col">Grand Total</th></tr></thead><tbody>')

    all_grand_total = 0
    group_list = group_total.sort_values("Group")["Group"].tolist()

    for group in group_list:
        gsum_series = group_total.loc[group_total["Group"] == group, "Grand Total"]
        gsum = int(gsum_series.iloc[0]) if not gsum_series.empty else 0
        if gsum == 0:
            continue
        all_grand_total += gsum

        gblock = pivot[pivot["Group"] == group].copy()
        gcols = {}
        for s in all_sizes:
            gcols[s] = int(gblock[s].sum()) if s in gblock.columns else 0

        html.append('<tr class="group-row">')
        html.append(f'<td class="pivot-col-left">▶ {group}</td>')
        for s in all_sizes:
            html.append(f'<td>{"" if gcols[s] == 0 else gcols[s]}</td>')
        html.append(f'<td class="total-col">{gsum}</td></tr>')

        ma_list = ma_total[ma_total["Group"] == group].sort_values("Ma SP")["Ma SP"].tolist()
        for ma in ma_list:
            msum_series = ma_total.loc[(ma_total["Group"] == group) & (ma_total["Ma SP"] == ma), "Grand Total"]
            msum = int(msum_series.iloc[0]) if not msum_series.empty else 0
            if msum == 0:
                continue

            mblock = gblock[gblock["Ma SP"] == ma].copy()
            mcols = {}
            for s in all_sizes:
                mcols[s] = int(mblock[s].sum()) if s in mblock.columns else 0

            html.append('<tr class="ma-row">')
            html.append(f'<td class="pivot-col-left">◉ {ma}</td>')
            for s in all_sizes:
                html.append(f'<td>{"" if mcols[s] == 0 else mcols[s]}</td>')
            html.append(f'<td class="total-col">{msum}</td></tr>')

            mblock = mblock[mblock["Grand Total"] > 0].sort_values("Mau")
            for _, row in mblock.iterrows():
                row_class = "mau-row"
                gt = int(row["Grand Total"]) if pd.notnull(row["Grand Total"]) else 0
                if gt == 1:
                    row_class += " row-alert"
                elif 2 <= gt <= 3:
                    row_class += " row-low"

                html.append(f'<tr class="{row_class}">')
                html.append(f'<td class="pivot-col-left">{row["Mau"]}</td>')
                for s in all_sizes:
                    v = int(row[s]) if s in row.index and pd.notnull(row[s]) else 0
                    html.append(f'<td>{"" if v == 0 else v}</td>')
                html.append(f'<td class="total-col">{gt}</td></tr>')

    html.append('<tr class="group-row">')
    html.append('<td class="pivot-col-left">Grand Total</td>')
    for s in all_sizes:
        col_total = int(pivot[s].sum()) if s in pivot.columns else 0
        html.append(f'<td>{"" if col_total == 0 else col_total}</td>')
    html.append(f'<td class="total-col">{all_grand_total}</td></tr>')

    html.append('</tbody></table></div>')
    return "".join(html)

def to_excel_file(df_clean, pivot_detail, all_sizes, summary_ma, summary_group):
    output = BytesIO()
    export_cols = ["Group", "Ma SP", "Mau"] + all_sizes + ["Grand Total"]
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
    col1, col2, col3, col4 = st.columns(4)
    groups = sorted([g for g in df_clean["Group"].dropna().astype(str).unique().tolist() if g.strip() != ""])
    masp = sorted([m for m in df_clean["Ma SP"].dropna().astype(str).unique().tolist() if m.strip() != ""])
    with col1:
        selected_groups = st.multiselect("Nhóm sản phẩm", groups, default=groups)
    with col2:
        selected_masp = st.multiselect("Mã SP", masp, default=masp)
    with col3:
        selected_status = st.selectbox("Trạng thái", ["Tất cả", "Cần nhập gấp", "Sắp hết", "Bình thường", "Tồn cao"])
    with col4:
        selected_view = st.selectbox("Hiển thị pivot", ["Tất cả", "Chỉ cần nhập", "Chỉ tồn cao"])
    st.markdown('</div>', unsafe_allow_html=True)

    df_filtered = df_clean.copy()
    if selected_groups:
        df_filtered = df_filtered[df_filtered["Group"].astype(str).isin(selected_groups)]
    else:
        df_filtered = df_filtered.iloc[0:0]
    if selected_masp:
        df_filtered = df_filtered[df_filtered["Ma SP"].astype(str).isin(selected_masp)]
    else:
        df_filtered = df_filtered.iloc[0:0]
    if selected_status != "Tất cả":
        df_filtered = df_filtered[df_filtered["Phan loai"] == selected_status]

    if selected_view == "Chỉ cần nhập":
        df_filtered = df_filtered[df_filtered["Ton kho"] <= 3]
    elif selected_view == "Chỉ tồn cao":
        df_filtered = df_filtered[df_filtered["Ton kho"] >= 10]

    df_filtered = df_filtered[df_filtered["Ton kho"] > 0].copy()

    st.subheader("Tổng quan nhanh")
    a, b, c, d = st.columns(4)
    a.metric("Tổng biến thể", f"{len(df_filtered):,}")
    b.metric("Tổng mã SP", f"{df_filtered['Ma SP'].nunique():,}")
    c.metric("Tổng tồn kho", f"{int(df_filtered['Ton kho'].sum()) if not df_filtered.empty else 0:,}")
    d.metric("Giá trị tồn", f"{float(df_filtered['Gia tri ton'].sum()) if not df_filtered.empty else 0:,.0f}")
    e, f, g = st.columns(3)
    e.metric("Cần nhập gấp", f"{len(df_filtered[df_filtered['Ton kho'] <= 1]):,}")
    f.metric("Sắp hết", f"{len(df_filtered[(df_filtered['Ton kho'] > 1) & (df_filtered['Ton kho'] <= 3)]):,}")
    g.metric("Tồn cao", f"{len(df_filtered[df_filtered['Ton kho'] >= 10]):,}")

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
        st.write("Top 10 mã tồn cao")
        ma_chart = df_filtered.groupby("Ma SP", as_index=False)["Ton kho"].sum().sort_values("Ton kho", ascending=False).head(10).set_index("Ma SP")
        if not ma_chart.empty:
            st.bar_chart(ma_chart)
        else:
            st.info("Không có dữ liệu.")

    summary_ma = df_filtered.groupby(["Group","Ma SP"], as_index=False).agg({"Ton kho":"sum","Gia tri ton":"sum"}).sort_values(["Ton kho","Gia tri ton"], ascending=[False, False])
    summary_group = df_filtered.groupby("Group", as_index=False).agg({"Ton kho":"sum","Gia tri ton":"sum"}).sort_values(["Ton kho","Gia tri ton"], ascending=[False, False])

    st.subheader("Top mã cần chú ý")
    st.dataframe(summary_ma.head(20), use_container_width=True, height=300)

    st.subheader("Pivot tồn kho xịn hơn")
    st.caption("Đã thêm cấp Group → Mã SP → Màu, ẩn blank, ẩn ô 0, ẩn tổng = 0, tô đỏ tồn = 1, tô vàng tồn 2–3.")
    pivot_detail, group_total, ma_total, all_sizes = build_pivot_hierarchical(df_filtered)
    st.markdown(render_pivot_html(pivot_detail, group_total, ma_total, all_sizes), unsafe_allow_html=True)

    st.subheader("Dữ liệu đã tách")
    st.dataframe(df_filtered[["Ma SP","Size","Mau","Gia Ban","Ton kho","Group","Phan loai","Gia tri ton"]], use_container_width=True, height=320)

    st.subheader("Gợi ý hành động")
    x1, x2 = st.columns(2)
    with x1:
        st.write("Danh sách cần nhập thêm")
        need_import = df_filtered[df_filtered["Ton kho"] <= 2].sort_values(["Ton kho","Group","Ma SP","Mau","Size"])
        st.dataframe(need_import, use_container_width=True, height=320)
    with x2:
        st.write("Danh sách nên cân nhắc chạy sale")
        sale_df = df_filtered[df_filtered["Ton kho"] >= 10].sort_values(["Ton kho","Gia tri ton"], ascending=[False, False])
        st.dataframe(sale_df, use_container_width=True, height=320)

    excel_file = to_excel_file(df_filtered, pivot_detail, all_sizes, summary_ma, summary_group)
    st.download_button("Tải file Excel kết quả", data=excel_file, file_name="merly_tonkho_pro_plus.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Hãy tải file Excel lên để bắt đầu.")
