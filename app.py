
import os
from io import BytesIO
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Merly - Tồn kho + Bán hàng", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 2rem; max-width: 96%;}
h1, h2, h3 {color: #9d8479;}
.filter-box {background: #f8f1ed; border: 1px solid #eadfd9; border-radius: 14px; padding: 12px 14px; margin-bottom: 10px;}
.pivot-wrap {overflow-x: auto; background: white; border: 1px solid #e8dfdb; border-radius: 12px; padding: 8px;}
.pivot-table {border-collapse: collapse; width: 100%; min-width: 1100px; font-size: 14px;}
.pivot-table th, .pivot-table td {border: 1px solid #d9d9d9; padding: 6px 8px; text-align: center; white-space: nowrap;}
.pivot-table thead th {background: #f3ece8; color: #7c655b; font-weight: 700;}
.pivot-col-left {text-align: left !important;}
.group-row td {background: #efe3dc; font-weight: 700; color: #6f5a51;}
.ma-row td {background: #f8f1ed; font-weight: 700; color: #6f5a51;}
.mau-row td:first-child {padding-left: 28px;}
.row-alert td {background: #ffd9d9 !important; color: #9b0000; font-weight: 700;}
.row-low td {background: #fff5cf !important; color: #8a6d00;}
.total-col {background: #fff2cc; font-weight: 700;}
.ai-box {background:#fff; border:1px solid #eadfd9; border-radius:14px; padding:16px;}
.badge-box {margin-top:6px; margin-bottom:6px; display:flex; gap:10px; flex-wrap:wrap;}
.badge {display:inline-block; padding:4px 10px; border-radius:999px; font-size:12px; font-weight:700;}
.badge-red {background:#ffd9d9; color:#9b0000;}
.badge-yellow {background:#fff2cc; color:#8a6d00;}
.badge-blue {background:#dceeff; color:#0b5cab;}
</style>
""", unsafe_allow_html=True)

def split_data(text):
    text = str(text).strip()
    if text == "": return pd.Series(["", "", ""])
    parts = text.split(" ")
    if len(parts) < 2: return pd.Series(["", "", ""])
    ma = parts[0]; size = ""; mau = ""
    for p in parts[1:]:
        if p.isdigit() and len(p) == 2: size = p
        else: mau += (" " if mau else "") + p
    return pd.Series([ma, size, mau.strip()])

def process_inventory_file(file):
    df = pd.read_excel(file)
    if df.shape[1] < 8:
        raise ValueError("File cần ít nhất 8 cột: A tên hàng, C tồn kho, E giá bán, H group.")
    df[["Ma SP","Size","Mau"]] = df.iloc[:,0].apply(split_data)
    df["Gia Ban"] = pd.to_numeric(df.iloc[:,4], errors="coerce").fillna(0)
    df["Ton kho"] = pd.to_numeric(df.iloc[:,2], errors="coerce").fillna(0).astype(int)
    df["Group"] = df.iloc[:,7].astype(str).fillna("")
    return df[["Ma SP","Size","Mau","Gia Ban","Ton kho","Group"]].copy()

def classify_ton_kho(qty):
    qty = int(qty) if pd.notnull(qty) else 0
    if qty <= 0: return "Hết hàng"
    if qty <= 1: return "Cần nhập gấp"
    if qty <= 3: return "Sắp hết"
    if qty >= 10: return "Tồn cao"
    return "Bình thường"

def suggested_restock(row):
    qty = int(row["Ton kho"]) if pd.notnull(row["Ton kho"]) else 0
    return 5 if qty <= 0 else 4 if qty == 1 else 3 if qty == 2 else 2 if qty == 3 else 0

def sale_priority(row):
    qty = int(row["Ton kho"]) if pd.notnull(row["Ton kho"]) else 0
    value = float(row["Gia tri ton"]) if pd.notnull(row["Gia tri ton"]) else 0
    if qty >= 15: return "Ưu tiên sale mạnh"
    if qty >= 10 or value >= 5000000: return "Có thể chạy sale"
    return "Chưa cần sale"

def build_pivot_hierarchical(df_clean):
    df2 = df_clean.copy()
    df2["Size"] = df2["Size"].astype(str)
    df2["Ton kho"] = pd.to_numeric(df2["Ton kho"], errors="coerce").fillna(0).astype(int)
    df2["Group"] = df2["Group"].astype(str).fillna("").str.strip()
    df2["Ma SP"] = df2["Ma SP"].astype(str).fillna("").str.strip()
    df2["Mau"] = df2["Mau"].astype(str).fillna("").str.strip()
    df2 = df2[df2["Ton kho"] > 0].copy()
    df2 = df2[(df2["Group"] != "") & (df2["Ma SP"] != "") & (df2["Mau"] != "")].copy()
    all_sizes = sorted([s for s in df2["Size"].dropna().unique() if str(s).isdigit()], key=lambda x: int(x))
    pivot = pd.pivot_table(df2, index=["Group","Ma SP","Mau"], columns="Size", values="Ton kho", aggfunc="sum", fill_value=0)
    for s in all_sizes:
        if s not in pivot.columns: pivot[s] = 0
    pivot = pivot[all_sizes]
    pivot["Grand Total"] = pivot.sum(axis=1)
    pivot = pivot[pivot["Grand Total"] > 0].reset_index()
    group_total = df2.groupby("Group", as_index=False)["Ton kho"].sum().rename(columns={"Ton kho":"Grand Total"})
    ma_total = df2.groupby(["Group","Ma SP"], as_index=False)["Ton kho"].sum().rename(columns={"Ton kho":"Grand Total"})
    return pivot, group_total, ma_total, all_sizes

def render_pivot_html(pivot, group_total, ma_total, all_sizes):
    if pivot.empty:
        return '<div class="pivot-wrap"><p style="padding:12px;color:#7c655b;">Không có dữ liệu phù hợp.</p></div>'
    html = ['<div class="badge-box"><span class="badge badge-red">Tồn = 1</span><span class="badge badge-yellow">Tồn 2–3</span><span class="badge badge-blue">Ẩn ô 0, blank và mã tổng = 0</span></div>',
            '<div class="pivot-wrap"><table class="pivot-table"><thead><tr><th class="pivot-col-left">Row Labels</th>']
    for s in all_sizes: html.append(f'<th>{s}</th>')
    html.append('<th class="total-col">Grand Total</th></tr></thead><tbody>')
    all_total = 0
    for group in group_total.sort_values("Group")["Group"].tolist():
        gsum = int(group_total.loc[group_total["Group"] == group, "Grand Total"].iloc[0])
        if gsum == 0: continue
        all_total += gsum
        gblock = pivot[pivot["Group"] == group].copy()
        html.append(f'<tr class="group-row"><td class="pivot-col-left">▶ {group}</td>')
        for s in all_sizes:
            v = int(gblock[s].sum()) if s in gblock.columns else 0
            html.append(f'<td>{"" if v == 0 else v}</td>')
        html.append(f'<td class="total-col">{gsum}</td></tr>')
        for ma in ma_total[ma_total["Group"] == group].sort_values("Ma SP")["Ma SP"].tolist():
            msum_series = ma_total.loc[(ma_total["Group"] == group) & (ma_total["Ma SP"] == ma), "Grand Total"]
            msum = int(msum_series.iloc[0]) if not msum_series.empty else 0
            if msum == 0: continue
            mblock = gblock[gblock["Ma SP"] == ma].copy()
            html.append(f'<tr class="ma-row"><td class="pivot-col-left">◉ {ma}</td>')
            for s in all_sizes:
                v = int(mblock[s].sum()) if s in mblock.columns else 0
                html.append(f'<td>{"" if v == 0 else v}</td>')
            html.append(f'<td class="total-col">{msum}</td></tr>')
            for _, row in mblock[mblock["Grand Total"] > 0].sort_values("Mau").iterrows():
                gt = int(row["Grand Total"]) if pd.notnull(row["Grand Total"]) else 0
                row_class = "mau-row row-alert" if gt == 1 else "mau-row row-low" if 2 <= gt <= 3 else "mau-row"
                html.append(f'<tr class="{row_class}"><td class="pivot-col-left">{row["Mau"]}</td>')
                for s in all_sizes:
                    v = int(row[s]) if s in row.index and pd.notnull(row[s]) else 0
                    html.append(f'<td>{"" if v == 0 else v}</td>')
                html.append(f'<td class="total-col">{gt}</td></tr>')
    html.append('<tr class="group-row"><td class="pivot-col-left">Grand Total</td>')
    for s in all_sizes:
        col_total = int(pivot[s].sum()) if s in pivot.columns else 0
        html.append(f'<td>{"" if col_total == 0 else col_total}</td>')
    html.append(f'<td class="total-col">{all_total}</td></tr></tbody></table></div>')
    return "".join(html)

def to_excel_file(df_all, pivot_detail, all_sizes, summary_ma, summary_group, need_import, out_of_stock, best_color, best_size):
    output = BytesIO()
    export_cols = ["Group","Ma SP","Mau"] + all_sizes + ["Grand Total"]
    pivot_export = pivot_detail[export_cols].copy()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_all.to_excel(writer, sheet_name="DuLieuDaTach", index=False)
        pivot_export.to_excel(writer, sheet_name="PivotTonKho", index=False)
        summary_ma.to_excel(writer, sheet_name="TongHopTheoMa", index=False)
        summary_group.to_excel(writer, sheet_name="TongHopTheoGroup", index=False)
        need_import.to_excel(writer, sheet_name="CanNhapThem", index=False)
        out_of_stock.to_excel(writer, sheet_name="DaHetHang", index=False)
        best_color.to_excel(writer, sheet_name="BanChay_MaMau", index=False)
        best_size.to_excel(writer, sheet_name="BanChay_MaMauSize", index=False)
    output.seek(0)
    return output

def build_ai_prompt(summary_group, summary_ma, need_import, out_of_stock, sale_df, best_color):
    return f"""Bạn là trợ lý phân tích tồn kho và xu hướng bán hàng cho Merly.
Hãy viết khuyến nghị vận hành ngắn gọn, thực dụng, bằng tiếng Việt.

Dữ liệu tóm tắt theo Group:
{summary_group.head(10).to_dict(orient='records')}

Top mã theo tồn:
{summary_ma.head(15).to_dict(orient='records')}

Danh sách cần nhập thêm:
{need_import.head(20).to_dict(orient='records')}

Danh sách hết hàng:
{out_of_stock.head(20).to_dict(orient='records')}

Danh sách tồn cao/có thể sale:
{sale_df.head(20).to_dict(orient='records')}

Danh sách mã + màu bán chạy:
{best_color.head(15).to_dict(orient='records')}

Hãy trả ra đúng 5 phần:
1. Tóm tắt nhanh tình hình tồn kho.
2. 5 ưu tiên nhập hàng trước.
3. 5 gợi ý chạy sale hoặc đẩy bán.
4. 5 mã + màu đang có tín hiệu bán tốt.
5. Rủi ro tồn kho cần chú ý."""

def run_ai_analysis(prompt):
    api_key = None
    try: api_key = st.secrets.get("OPENAI_API_KEY", None)
    except Exception: api_key = None
    api_key = api_key or os.getenv("OPENAI_API_KEY")
    if not api_key:
        return None, "Chưa có OPENAI_API_KEY."
    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)
        response = client.responses.create(model="gpt-4.1-mini", input=prompt)
        return response.output_text, None
    except Exception as e:
        return None, f"Lỗi khi gọi AI: {e}"

def build_sales_compare(old_file, new_file):
    old_df = process_inventory_file(old_file)[["Ma SP","Mau","Size","Ton kho","Group"]].copy().rename(columns={"Ton kho":"Ton_cu"})
    new_df = process_inventory_file(new_file)[["Ma SP","Mau","Size","Ton kho","Group"]].copy().rename(columns={"Ton kho":"Ton_moi"})
    compare = old_df.merge(new_df, on=["Ma SP","Mau","Size"], how="outer", suffixes=("_old","_new"))
    compare["Group"] = compare["Group_old"].combine_first(compare["Group_new"])
    compare = compare.drop(columns=[c for c in ["Group_old","Group_new"] if c in compare.columns])
    compare["Ton_cu"] = pd.to_numeric(compare["Ton_cu"], errors="coerce").fillna(0).astype(int)
    compare["Ton_moi"] = pd.to_numeric(compare["Ton_moi"], errors="coerce").fillna(0).astype(int)
    compare["Da_ban"] = compare["Ton_cu"] - compare["Ton_moi"]
    sold = compare[compare["Da_ban"] > 0].copy()
    best_color = sold.groupby(["Group","Ma SP","Mau"], as_index=False)["Da_ban"].sum().sort_values("Da_ban", ascending=False)
    best_size = sold.groupby(["Group","Ma SP","Mau","Size"], as_index=False)["Da_ban"].sum().sort_values("Da_ban", ascending=False)
    return compare, sold, best_color, best_size

col_logo, col_title = st.columns([1,4])
with col_logo:
    st.image("Logo Merly.jpg", width=130)
with col_title:
    st.title("Merly - Tồn kho Pro AI-ready")
    st.caption("Có tab tồn kho, tab so sánh 2 file để suy ra bán chạy, pivot kiểu Excel, gợi ý nhập hàng, và khu vực AI-ready.")

tab1, tab2 = st.tabs(["Tồn kho & AI", "Phân tích bán hàng"])

with tab1:
    uploaded_file = st.file_uploader("Tải file Excel tồn kho", type=["xlsx"], key="inventory_file")
    if uploaded_file:
        df_clean = process_inventory_file(uploaded_file)
        df_clean["Phan loai"] = df_clean["Ton kho"].apply(classify_ton_kho)
        df_clean["Gia tri ton"] = df_clean["Gia Ban"] * df_clean["Ton kho"]
        df_clean["SL de xuat nhap"] = df_clean.apply(suggested_restock, axis=1)
        df_clean["Sale priority"] = df_clean.apply(sale_priority, axis=1)

        st.markdown('<div class="filter-box">', unsafe_allow_html=True)
        st.subheader("Bộ lọc")
        c1,c2,c3,c4 = st.columns(4)
        groups = sorted([g for g in df_clean["Group"].dropna().astype(str).unique().tolist() if g.strip() != ""])
        masp = sorted([m for m in df_clean["Ma SP"].dropna().astype(str).unique().tolist() if m.strip() != ""])
        with c1: selected_groups = st.multiselect("Nhóm sản phẩm", groups, default=groups, key="inv_groups")
        with c2: selected_masp = st.multiselect("Mã SP", masp, default=masp, key="inv_masp")
        with c3: selected_status = st.selectbox("Trạng thái", ["Tất cả","Hết hàng","Cần nhập gấp","Sắp hết","Bình thường","Tồn cao"], key="inv_status")
        with c4: selected_view = st.selectbox("Hiển thị pivot", ["Tất cả","Chỉ cần nhập","Chỉ tồn cao"], key="inv_view")
        st.markdown('</div>', unsafe_allow_html=True)

        df_base = df_clean.copy()
        df_base = df_base[df_base["Group"].astype(str).isin(selected_groups)] if selected_groups else df_base.iloc[0:0]
        df_base = df_base[df_base["Ma SP"].astype(str).isin(selected_masp)] if selected_masp else df_base.iloc[0:0]
        if selected_status != "Tất cả": df_base = df_base[df_base["Phan loai"] == selected_status]

        df_all = df_base.copy()
        df_pivot = df_base.copy()
        if selected_view == "Chỉ cần nhập": df_pivot = df_pivot[df_pivot["Ton kho"] <= 3]
        elif selected_view == "Chỉ tồn cao": df_pivot = df_pivot[df_pivot["Ton kho"] >= 10]
        df_pivot = df_pivot[df_pivot["Ton kho"] > 0].copy()

        a,b,c,d = st.columns(4)
        a.metric("Tổng biến thể", f"{len(df_all):,}")
        b.metric("Tổng mã SP", f"{df_all['Ma SP'].nunique():,}")
        c.metric("Tổng tồn kho", f"{int(df_all['Ton kho'].sum()) if not df_all.empty else 0:,}")
        d.metric("Giá trị tồn", f"{float(df_all['Gia tri ton'].sum()) if not df_all.empty else 0:,.0f}")

        e,f,g,h = st.columns(4)
        e.metric("Hết hàng", f"{len(df_all[df_all['Ton kho'] == 0]):,}")
        f.metric("Cần nhập gấp", f"{len(df_all[df_all['Ton kho'] <= 1]):,}")
        g.metric("Sắp hết", f"{len(df_all[(df_all['Ton kho'] > 1) & (df_all['Ton kho'] <= 3)]):,}")
        h.metric("Tồn cao", f"{len(df_all[df_all['Ton kho'] >= 10]):,}")

        cc1, cc2 = st.columns(2)
        with cc1:
            st.write("Tồn kho theo nhóm sản phẩm")
            group_chart = df_all.groupby("Group", as_index=False)["Ton kho"].sum().sort_values("Ton kho", ascending=False).set_index("Group")
            st.bar_chart(group_chart) if not group_chart.empty else st.info("Không có dữ liệu.")
        with cc2:
            st.write("Top 10 mã tồn cao")
            ma_chart = df_all.groupby("Ma SP", as_index=False)["Ton kho"].sum().sort_values("Ton kho", ascending=False).head(10).set_index("Ma SP")
            st.bar_chart(ma_chart) if not ma_chart.empty else st.info("Không có dữ liệu.")

        summary_ma = df_all.groupby(["Group","Ma SP"], as_index=False).agg({"Ton kho":"sum","Gia tri ton":"sum"}).sort_values(["Ton kho","Gia tri ton"], ascending=[False,False])
        summary_group = df_all.groupby("Group", as_index=False).agg({"Ton kho":"sum","Gia tri ton":"sum"}).sort_values(["Ton kho","Gia tri ton"], ascending=[False,False])
        st.subheader("Top mã cần chú ý")
        st.dataframe(summary_ma.head(20), use_container_width=True, height=280)

        st.subheader("Pivot tồn kho")
        pivot_detail, group_total, ma_total, all_sizes = build_pivot_hierarchical(df_pivot)
        st.markdown(render_pivot_html(pivot_detail, group_total, ma_total, all_sizes), unsafe_allow_html=True)

        need_import = df_all[df_all["Ton kho"] <= 2].sort_values(["Ton kho","Group","Ma SP","Mau","Size"])
        sale_df = df_all[df_all["Ton kho"] >= 10].sort_values(["Ton kho","Gia tri ton"], ascending=[False,False])
        out_of_stock = df_all[df_all["Ton kho"] == 0].sort_values(["Group","Ma SP","Mau","Size"])
        empty_best_color = pd.DataFrame(columns=["Group","Ma SP","Mau","Da_ban"])

        st.subheader("Dữ liệu đã tách")
        st.dataframe(df_all[["Ma SP","Size","Mau","Gia Ban","Ton kho","Group","Phan loai","Gia tri ton","SL de xuat nhap","Sale priority"]], use_container_width=True, height=320)

        x1,x2 = st.columns(2)
        with x1:
            st.write("Danh sách cần nhập thêm")
            st.dataframe(need_import[["Ma SP","Size","Mau","Ton kho","SL de xuat nhap","Group","Phan loai"]], use_container_width=True, height=320)
        with x2:
            st.write("Danh sách nên cân nhắc chạy sale")
            st.dataframe(sale_df[["Ma SP","Size","Mau","Ton kho","Gia tri ton","Group","Sale priority"]], use_container_width=True, height=320)

        st.subheader("SKU đã hết hàng")
        st.dataframe(out_of_stock[["Ma SP","Size","Mau","Gia Ban","Ton kho","Group","SL de xuat nhap"]], use_container_width=True, height=260)

        st.subheader("AI gợi ý hành động")
        st.markdown('<div class="ai-box">', unsafe_allow_html=True)
        ai_prompt = build_ai_prompt(summary_group, summary_ma, need_import, out_of_stock, sale_df, empty_best_color)
        extra_note = st.text_area("Ghi chú thêm cho AI (tuỳ chọn)", height=80, key="ai_note")
        if st.button("Phân tích bằng AI", key="ai_button"):
            final_prompt = ai_prompt + ("\n\nGhi chú thêm từ người dùng:\n" + extra_note if extra_note.strip() else "")

Ghi chú thêm từ người dùng:
" + extra_note if extra_note.strip() else "")
            with st.spinner("Đang phân tích..."):
                ai_text, ai_error = run_ai_analysis(final_prompt)
            if ai_error: st.warning(ai_error)
            else: st.markdown(ai_text)
        else:
            st.caption("App đã AI-ready. Chỉ cần thêm OPENAI_API_KEY vào Secrets của Streamlit.")
        st.markdown('</div>', unsafe_allow_html=True)

        excel_file = to_excel_file(df_all, pivot_detail, all_sizes, summary_ma, summary_group, need_import, out_of_stock, empty_best_color, pd.DataFrame())
        st.download_button("Tải file Excel kết quả", data=excel_file, file_name="merly_tonkho_sales_tab.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Hãy tải file Excel tồn kho lên để bắt đầu.")

with tab2:
    st.subheader("So sánh 2 file để suy ra mã + màu bán chạy")
    st.caption("Đã bán = Tồn file cũ - Tồn file mới. Chỉ giữ giá trị > 0. Nếu âm thì xem như có nhập thêm và bỏ qua.")
    c1,c2,c3 = st.columns(3)
    with c1: file_old = st.file_uploader("File tồn kho TRƯỚC", type=["xlsx"], key="sales_old")
    with c2: file_new = st.file_uploader("File tồn kho HIỆN TẠI", type=["xlsx"], key="sales_new")
    with c3: period_days = st.number_input("Số ngày giữa 2 file", min_value=1, value=7, step=1)

    if file_old and file_new:
        try:
            compare, sold, best_color, best_size = build_sales_compare(file_old, file_new)
            m1,m2,m3,m4 = st.columns(4)
            m1.metric("Tổng SKU có bán", f"{len(sold):,}")
            m2.metric("Tổng lượng bán suy ra", f"{int(sold['Da_ban'].sum()) if not sold.empty else 0:,}")
            m3.metric("Mã + màu bán chạy", f"{len(best_color):,}")
            m4.metric("Bán/ngày", f"{(float(sold['Da_ban'].sum())/period_days) if not sold.empty else 0:,.2f}")

            best_color_display = best_color.copy()
            best_color_display["Ban/ngay"] = (best_color_display["Da_ban"] / period_days).round(2)
            best_size_display = best_size.copy()
            best_size_display["Ban/ngay"] = (best_size_display["Da_ban"] / period_days).round(2)

            st.subheader("Top mã + màu bán chạy")
            st.dataframe(best_color_display.head(50), use_container_width=True, height=320)

            st.subheader("Top mã + màu + size bán chạy")
            st.dataframe(best_size_display.head(50), use_container_width=True, height=320)

            d1,d2 = st.columns(2)
            with d1:
                st.write("Top Group bán tốt")
                group_sold = sold.groupby("Group", as_index=False)["Da_ban"].sum().sort_values("Da_ban", ascending=False).set_index("Group")
                st.bar_chart(group_sold) if not group_sold.empty else st.info("Không có dữ liệu.")
            with d2:
                st.write("Top mã + màu bán tốt")
                chart_color = best_color.head(10).copy()
                if not chart_color.empty:
                    chart_color["Label"] = chart_color["Ma SP"].astype(str) + " - " + chart_color["Mau"].astype(str)
                    st.bar_chart(chart_color.set_index("Label")[["Da_ban"]])
                else:
                    st.info("Không có dữ liệu.")

            st.subheader("Dữ liệu so sánh chi tiết")
            st.dataframe(compare.sort_values("Da_ban", ascending=False), use_container_width=True, height=320)

            sales_xlsx = BytesIO()
            with pd.ExcelWriter(sales_xlsx, engine="openpyxl") as writer:
                compare.to_excel(writer, sheet_name="SoSanhChiTiet", index=False)
                sold.to_excel(writer, sheet_name="ChiTietDaBan", index=False)
                best_color_display.to_excel(writer, sheet_name="BanChay_MaMau", index=False)
                best_size_display.to_excel(writer, sheet_name="BanChay_MaMauSize", index=False)
            sales_xlsx.seek(0)
            st.download_button("Tải file phân tích bán hàng", data=sales_xlsx, file_name="merly_phan_tich_ban_hang.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Lỗi khi so sánh 2 file: {e}")
    else:
        st.info("Hãy tải đủ 2 file để bắt đầu so sánh.")
