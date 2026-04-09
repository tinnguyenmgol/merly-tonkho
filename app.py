
import os
from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Merly - Business AI", layout="wide")

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
.note {font-size: 12px; color: #7d6a61;}
</style>
""", unsafe_allow_html=True)

# ---------- helpers ----------
def split_data(text):
    text = str(text).strip()
    if text == "":
        return pd.Series(["", "", ""])
    parts = text.split(" ")
    if len(parts) < 2:
        return pd.Series(["", "", ""])
    ma = str(parts[0]).strip()
    size = ""
    mau = ""
    for p in parts[1:]:
        p = str(p).strip()
        if p.isdigit() and len(p) == 2:
            size = p
        else:
            mau += (" " if mau else "") + p
    return pd.Series([ma.strip(), size.strip(), mau.strip()])

def process_inventory_file(file):
    df = pd.read_excel(file)
    if df.shape[1] < 8:
        raise ValueError("File cần ít nhất 8 cột: A = Tên hàng hóa, C = Tồn kho, E = Giá bán, H = Group.")
    df[["Ma SP", "Size", "Mau"]] = df.iloc[:, 0].apply(split_data)
    df["Gia Ban"] = pd.to_numeric(df.iloc[:, 4], errors="coerce").fillna(0)
    df["Ton kho"] = pd.to_numeric(df.iloc[:, 2], errors="coerce").fillna(0).astype(int)
    df["Group"] = df.iloc[:, 7].astype(str).fillna("").str.strip()

    out = df[["Ma SP", "Size", "Mau", "Gia Ban", "Ton kho", "Group"]].copy()
    out["Ma SP"] = out["Ma SP"].astype(str).fillna("").str.strip()
    out["Size"] = out["Size"].astype(str).fillna("").str.strip()
    out["Mau"] = out["Mau"].astype(str).fillna("").str.strip()
    out["Group"] = out["Group"].astype(str).fillna("").str.strip()
    return out

def classify_ton_kho(qty):
    qty = int(qty) if pd.notnull(qty) else 0
    if qty <= 0:
        return "Hết hàng"
    if qty <= 1:
        return "Cần nhập gấp"
    if qty <= 3:
        return "Sắp hết"
    if qty >= 10:
        return "Tồn cao"
    return "Bình thường"

def suggested_restock(row):
    qty = int(row["Ton kho"]) if pd.notnull(row["Ton kho"]) else 0
    if qty <= 0:
        return 5
    if qty == 1:
        return 4
    if qty == 2:
        return 3
    if qty == 3:
        return 2
    return 0

def sale_priority(row):
    qty = int(row["Ton kho"]) if pd.notnull(row["Ton kho"]) else 0
    value = float(row["Gia tri ton"]) if pd.notnull(row["Gia tri ton"]) else 0
    if qty >= 15:
        return "Ưu tiên sale mạnh"
    if qty >= 10 or value >= 5000000:
        return "Có thể chạy sale"
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

    all_sizes = sorted(
        [s for s in df2["Size"].dropna().unique() if str(s).isdigit()],
        key=lambda x: int(x)
    )

    if df2.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), all_sizes

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

    group_total = df2.groupby("Group", as_index=False)["Ton kho"].sum().rename(columns={"Ton kho": "Grand Total"})
    ma_total = df2.groupby(["Group", "Ma SP"], as_index=False)["Ton kho"].sum().rename(columns={"Ton kho": "Grand Total"})
    return pivot, group_total, ma_total, all_sizes

def render_pivot_html(pivot, group_total, ma_total, all_sizes):
    if pivot.empty:
        return '<div class="pivot-wrap"><p style="padding:12px;color:#7c655b;">Không có dữ liệu phù hợp.</p></div>'

    html = []
    html.append('<div class="badge-box">')
    html.append('<span class="badge badge-red">Tồn = 1</span>')
    html.append('<span class="badge badge-yellow">Tồn 2–3</span>')
    html.append('<span class="badge badge-blue">Ẩn ô 0, blank và mã tổng = 0</span>')
    html.append('</div>')
    html.append('<div class="pivot-wrap"><table class="pivot-table">')
    html.append('<thead><tr><th class="pivot-col-left">Row Labels</th>')
    for s in all_sizes:
        html.append(f'<th>{s}</th>')
    html.append('<th class="total-col">Grand Total</th></tr></thead><tbody>')

    all_total = 0
    for group in group_total.sort_values("Group")["Group"].tolist():
        gsum_series = group_total.loc[group_total["Group"] == group, "Grand Total"]
        gsum = int(gsum_series.iloc[0]) if not gsum_series.empty else 0
        if gsum == 0:
            continue
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
            if msum == 0:
                continue
            mblock = gblock[gblock["Ma SP"] == ma].copy()

            html.append(f'<tr class="ma-row"><td class="pivot-col-left">◉ {ma}</td>')
            for s in all_sizes:
                v = int(mblock[s].sum()) if s in mblock.columns else 0
                html.append(f'<td>{"" if v == 0 else v}</td>')
            html.append(f'<td class="total-col">{msum}</td></tr>')

            for _, row in mblock[mblock["Grand Total"] > 0].sort_values("Mau").iterrows():
                gt = int(row["Grand Total"]) if pd.notnull(row["Grand Total"]) else 0
                row_class = "mau-row"
                if gt == 1:
                    row_class += " row-alert"
                elif 2 <= gt <= 3:
                    row_class += " row-low"

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

def safe_top_records(df, cols, n=10):
    if df is None or df.empty:
        return []
    available_cols = [c for c in cols if c in df.columns]
    return df[available_cols].head(n).to_dict(orient="records")

def build_ai_prompt_business(summary_group, summary_ma, need_import, out_of_stock, sale_df, best_color, best_size):
    group_data = safe_top_records(summary_group, ["Group", "Ton kho", "Gia tri ton"], 12)
    ma_data = safe_top_records(summary_ma, ["Group", "Ma SP", "Ton kho", "Gia tri ton"], 20)
    import_data = safe_top_records(need_import, ["Group", "Ma SP", "Mau", "Size", "Ton kho", "SL de xuat nhap", "Phan loai"], 30)
    out_data = safe_top_records(out_of_stock, ["Group", "Ma SP", "Mau", "Size", "Ton kho", "SL de xuat nhap"], 30)
    sale_data = safe_top_records(sale_df, ["Group", "Ma SP", "Mau", "Size", "Ton kho", "Gia tri ton", "Sale priority"], 30)
    best_color_data = safe_top_records(best_color, ["Group", "Ma SP", "Mau", "Da_ban", "Ban/ngay"], 20)
    best_size_data = safe_top_records(best_size, ["Group", "Ma SP", "Mau", "Size", "Da_ban", "Ban/ngay"], 20)

    return f"""
Bạn là trợ lý vận hành kinh doanh cho Merly, thương hiệu giày nữ big size.
Vai trò của bạn là hỗ trợ ra quyết định tồn kho, nhập hàng, sale, livestream và ưu tiên xử lý sản phẩm.

Nguyên tắc phân tích:
1. Ưu tiên tính thực dụng, ngắn gọn, hành động được ngay.
2. Không viết kiểu chung chung, không giải thích lý thuyết dài dòng.
3. Luôn ưu tiên:
   - tránh mất doanh thu do hết hàng
   - tránh tồn chết kéo dài
   - tập trung vào mã + màu đang bán được
   - không đề xuất sale bừa cho sản phẩm đang có tín hiệu bán tốt
4. Nếu một mã/màu đang bán tốt nhưng tồn thấp hoặc hết hàng, ưu tiên nhập lại.
5. Nếu tồn cao nhưng chưa có tín hiệu bán tốt, ưu tiên đề xuất sale / live / combo / nội dung trước khi giảm giá sâu.
6. Business rules của Merly:
   - ưu tiên giữ hàng cho các mẫu big size bán ổn
   - không sale mạnh các mã + màu đang bán chạy
   - nếu một mã đang hot nhưng thiếu nhiều size, ưu tiên nhập lại size bán nhanh trước
   - nếu một group đang có tín hiệu bán tốt, ưu tiên dồn nguồn lực nội dung và livestream cho group đó

Dữ liệu theo nhóm sản phẩm:
{group_data}

Dữ liệu top mã tồn kho:
{ma_data}

Danh sách cần nhập thêm:
{import_data}

Danh sách đã hết hàng:
{out_data}

Danh sách tồn cao / có thể sale:
{sale_data}

Danh sách mã + màu bán chạy:
{best_color_data}

Danh sách mã + màu + size bán chạy:
{best_size_data}

Hãy trả lời đúng theo 7 phần sau:

1. TÓM TẮT QUẢN TRỊ
2. ƯU TIÊN NHẬP HÀNG NGAY
3. ƯU TIÊN CHẠY SALE / XẢ HÀNG
4. GỢI Ý ĐẨY LIVESTREAM / NỘI DUNG
5. MẪU HOT CẦN GIỮ GIÁ / KHÔNG NÊN SALE MẠNH
6. RỦI RO TỒN KHO CẦN CHÚ Ý
7. KẾ HOẠCH HÀNH ĐỘNG 7 NGÀY

Yêu cầu văn phong:
- như trưởng bộ phận vận hành đang báo cáo cho chủ doanh nghiệp
- ngắn, sắc, thực tế
- không dùng icon
- không dùng câu quá dài
""".strip()

def run_ai_analysis(prompt):
    api_key = None
    try:
        api_key = st.secrets.get("OPENAI_API_KEY", None)
    except Exception:
        api_key = None
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
    old_df = process_inventory_file(old_file)[["Ma SP", "Mau", "Size", "Ton kho", "Group"]].copy()
    new_df = process_inventory_file(new_file)[["Ma SP", "Mau", "Size", "Ton kho", "Group"]].copy()

    # loại dòng rác/blank trước khi so sánh
    for df_ in (old_df, new_df):
        df_["Ma SP"] = df_["Ma SP"].astype(str).fillna("").str.strip()
        df_["Mau"] = df_["Mau"].astype(str).fillna("").str.strip()
        df_["Size"] = df_["Size"].astype(str).fillna("").str.strip()
        df_["Group"] = df_["Group"].astype(str).fillna("").str.strip()

    old_df = old_df[(old_df["Ma SP"] != "") & (old_df["Mau"] != "") & (old_df["Size"] != "")]
    new_df = new_df[(new_df["Ma SP"] != "") & (new_df["Mau"] != "") & (new_df["Size"] != "")]

    # gộp trùng SKU trước khi merge
    old_df = old_df.groupby(["Ma SP", "Mau", "Size", "Group"], as_index=False)["Ton kho"].sum()
    new_df = new_df.groupby(["Ma SP", "Mau", "Size", "Group"], as_index=False)["Ton kho"].sum()

    old_df = old_df.rename(columns={"Ton kho": "Ton_cu"})
    new_df = new_df.rename(columns={"Ton kho": "Ton_moi"})

    compare = old_df.merge(
        new_df,
        on=["Ma SP", "Mau", "Size"],
        how="outer",
        suffixes=("_old", "_new")
    )

    compare["Group"] = compare["Group_old"].combine_first(compare["Group_new"])
    compare = compare.drop(columns=[c for c in ["Group_old", "Group_new"] if c in compare.columns])

    compare["Group"] = compare["Group"].astype(str).fillna("").str.strip()
    compare["Ton_cu"] = pd.to_numeric(compare["Ton_cu"], errors="coerce").fillna(0).astype(int)
    compare["Ton_moi"] = pd.to_numeric(compare["Ton_moi"], errors="coerce").fillna(0).astype(int)
    compare["Da_ban"] = compare["Ton_cu"] - compare["Ton_moi"]

    compare = compare[(compare["Ma SP"] != "") & (compare["Mau"] != "") & (compare["Size"] != "")].copy()
    sold = compare[compare["Da_ban"] > 0].copy()

    best_color = (
        sold.groupby(["Group", "Ma SP", "Mau"], as_index=False)["Da_ban"]
        .sum()
        .sort_values("Da_ban", ascending=False)
    )

    best_size = (
        sold.groupby(["Group", "Ma SP", "Mau", "Size"], as_index=False)["Da_ban"]
        .sum()
        .sort_values("Da_ban", ascending=False)
    )

    return compare, sold, best_color, best_size

def to_excel_file(df_all, pivot_detail, all_sizes, summary_ma, summary_group, need_import, out_of_stock, best_color, best_size):
    output = BytesIO()
    export_cols = ["Group", "Ma SP", "Mau"] + all_sizes + ["Grand Total"]
    pivot_export = pivot_detail[export_cols].copy() if not pivot_detail.empty else pd.DataFrame()

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

# ---------- UI ----------
col_logo, col_title = st.columns([1, 4])
with col_logo:
    st.image("Logo Merly.jpg", width=130)
with col_title:
    st.title("Merly - Business AI")
    st.caption("Đã fix lỗi 2 tab: chart không còn render lỗi, tab bán hàng loại blank/rác trước khi suy ra bán chạy.")

tab1, tab2 = st.tabs(["Tồn kho & Business AI", "Phân tích bán hàng"])

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
        c1, c2, c3, c4 = st.columns(4)
        groups = sorted([g for g in df_clean["Group"].dropna().astype(str).unique().tolist() if g.strip() != ""])
        masp = sorted([m for m in df_clean["Ma SP"].dropna().astype(str).unique().tolist() if m.strip() != ""])

        with c1:
            selected_groups = st.multiselect("Nhóm sản phẩm", groups, default=groups, key="inv_groups")
        with c2:
            selected_masp = st.multiselect("Mã SP", masp, default=masp, key="inv_masp")
        with c3:
            selected_status = st.selectbox("Trạng thái", ["Tất cả", "Hết hàng", "Cần nhập gấp", "Sắp hết", "Bình thường", "Tồn cao"], key="inv_status")
        with c4:
            selected_view = st.selectbox("Hiển thị pivot", ["Tất cả", "Chỉ cần nhập", "Chỉ tồn cao"], key="inv_view")
        st.markdown('</div>', unsafe_allow_html=True)

        df_base = df_clean.copy()
        df_base = df_base[df_base["Group"].astype(str).isin(selected_groups)] if selected_groups else df_base.iloc[0:0]
        df_base = df_base[df_base["Ma SP"].astype(str).isin(selected_masp)] if selected_masp else df_base.iloc[0:0]
        if selected_status != "Tất cả":
            df_base = df_base[df_base["Phan loai"] == selected_status]

        df_all = df_base.copy()
        df_pivot = df_base.copy()
        if selected_view == "Chỉ cần nhập":
            df_pivot = df_pivot[df_pivot["Ton kho"] <= 3]
        elif selected_view == "Chỉ tồn cao":
            df_pivot = df_pivot[df_pivot["Ton kho"] >= 10]
        df_pivot = df_pivot[df_pivot["Ton kho"] > 0].copy()

        a, b, c, d = st.columns(4)
        a.metric("Tổng biến thể", f"{len(df_all):,}")
        b.metric("Tổng mã SP", f"{df_all['Ma SP'].nunique():,}")
        c.metric("Tổng tồn kho", f"{int(df_all['Ton kho'].sum()) if not df_all.empty else 0:,}")
        d.metric("Giá trị tồn", f"{float(df_all['Gia tri ton'].sum()) if not df_all.empty else 0:,.0f}")

        e, f, g, h = st.columns(4)
        e.metric("Hết hàng", f"{len(df_all[df_all['Ton kho'] == 0]):,}")
        f.metric("Cần nhập gấp", f"{len(df_all[df_all['Ton kho'] <= 1]):,}")
        g.metric("Sắp hết", f"{len(df_all[(df_all['Ton kho'] > 1) & (df_all['Ton kho'] <= 3)]):,}")
        h.metric("Tồn cao", f"{len(df_all[df_all['Ton kho'] >= 10]):,}")

        cc1, cc2 = st.columns(2)
        with cc1:
            st.write("Tồn kho theo nhóm sản phẩm")
            group_chart = df_all.groupby("Group", as_index=False)["Ton kho"].sum().sort_values("Ton kho", ascending=False).set_index("Group")
            if not group_chart.empty:
                st.bar_chart(group_chart)
            else:
                st.info("Không có dữ liệu.")
        with cc2:
            st.write("Top 10 mã tồn cao")
            ma_chart = df_all.groupby("Ma SP", as_index=False)["Ton kho"].sum().sort_values("Ton kho", ascending=False).head(10).set_index("Ma SP")
            if not ma_chart.empty:
                st.bar_chart(ma_chart)
            else:
                st.info("Không có dữ liệu.")

        summary_ma = df_all.groupby(["Group", "Ma SP"], as_index=False).agg({"Ton kho": "sum", "Gia tri ton": "sum"}).sort_values(["Ton kho", "Gia tri ton"], ascending=[False, False])
        summary_group = df_all.groupby("Group", as_index=False).agg({"Ton kho": "sum", "Gia tri ton": "sum"}).sort_values(["Ton kho", "Gia tri ton"], ascending=[False, False])

        st.subheader("Top mã cần chú ý")
        st.dataframe(summary_ma.head(20), use_container_width=True, height=280)

        st.subheader("Pivot tồn kho")
        pivot_detail, group_total, ma_total, all_sizes = build_pivot_hierarchical(df_pivot)
        st.markdown(render_pivot_html(pivot_detail, group_total, ma_total, all_sizes), unsafe_allow_html=True)

        need_import = df_all[df_all["Ton kho"] <= 2].sort_values(["Ton kho", "Group", "Ma SP", "Mau", "Size"])
        sale_df = df_all[df_all["Ton kho"] >= 10].sort_values(["Ton kho", "Gia tri ton"], ascending=[False, False])
        out_of_stock = df_all[df_all["Ton kho"] == 0].sort_values(["Group", "Ma SP", "Mau", "Size"])

        st.subheader("Dữ liệu đã tách")
        st.dataframe(df_all[["Ma SP", "Size", "Mau", "Gia Ban", "Ton kho", "Group", "Phan loai", "Gia tri ton", "SL de xuat nhap", "Sale priority"]], use_container_width=True, height=320)

        x1, x2 = st.columns(2)
        with x1:
            st.write("Danh sách cần nhập thêm")
            st.dataframe(need_import[["Ma SP", "Size", "Mau", "Ton kho", "SL de xuat nhap", "Group", "Phan loai"]], use_container_width=True, height=320)
        with x2:
            st.write("Danh sách nên cân nhắc chạy sale")
            st.dataframe(sale_df[["Ma SP", "Size", "Mau", "Ton kho", "Gia tri ton", "Group", "Sale priority"]], use_container_width=True, height=320)

        st.subheader("SKU đã hết hàng")
        st.dataframe(out_of_stock[["Ma SP", "Size", "Mau", "Gia Ban", "Ton kho", "Group", "SL de xuat nhap"]], use_container_width=True, height=260)

        st.subheader("Business AI gợi ý hành động")
        st.markdown('<div class="ai-box">', unsafe_allow_html=True)
        empty_best_color = pd.DataFrame(columns=["Group", "Ma SP", "Mau", "Da_ban", "Ban/ngay"])
        empty_best_size = pd.DataFrame(columns=["Group", "Ma SP", "Mau", "Size", "Da_ban", "Ban/ngay"])
        ai_prompt = build_ai_prompt_business(summary_group, summary_ma, need_import, out_of_stock, sale_df, empty_best_color, empty_best_size)

        extra_note = st.text_area(
            "Ghi chú thêm cho AI (tuỳ chọn)",
            placeholder="Ví dụ: ưu tiên nhóm dép sục trước, chưa muốn sale cao gót, tập trung hàng đi làm...",
            height=100,
            key="ai_note_business"
        )

        if st.button("Phân tích bằng AI", key="ai_button_business"):
            if extra_note.strip():
                final_prompt = f"""{ai_prompt}

Bối cảnh bổ sung từ người dùng:
{extra_note}
"""
            else:
                final_prompt = ai_prompt

            with st.spinner("Đang phân tích theo góc nhìn business..."):
                ai_text, ai_error = run_ai_analysis(final_prompt)

            if ai_error:
                st.warning(ai_error)
                st.code(final_prompt, language="text")
            else:
                st.markdown(ai_text)
        else:
            st.caption("App đã AI-ready. Chỉ cần thêm OPENAI_API_KEY vào Secrets của Streamlit.")
        st.markdown('</div>', unsafe_allow_html=True)

        excel_file = to_excel_file(df_all, pivot_detail, all_sizes, summary_ma, summary_group, need_import, out_of_stock, empty_best_color, empty_best_size)
        st.download_button("Tải file Excel kết quả", data=excel_file, file_name="merly_business_ai_fixed.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Hãy tải file Excel tồn kho lên để bắt đầu.")

with tab2:
    st.subheader("So sánh 2 file để suy ra mã + màu bán chạy")
    st.caption("Đã bán = Tồn file cũ - Tồn file mới. Đã loại bỏ dòng blank/rác trước khi suy ra bán chạy.")
    t1, t2, t3 = st.columns(3)
    with t1:
        file_old = st.file_uploader("File tồn kho TRƯỚC", type=["xlsx"], key="sales_old")
    with t2:
        file_new = st.file_uploader("File tồn kho HIỆN TẠI", type=["xlsx"], key="sales_new")
    with t3:
        period_days = st.number_input("Số ngày giữa 2 file", min_value=1, value=7, step=1)

    if file_old and file_new:
        try:
            compare, sold, best_color, best_size = build_sales_compare(file_old, file_new)

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Tổng SKU có bán", f"{len(sold):,}")
            m2.metric("Tổng lượng bán suy ra", f"{int(sold['Da_ban'].sum()) if not sold.empty else 0:,}")
            m3.metric("Mã + màu bán chạy", f"{len(best_color):,}")
            m4.metric("Bán/ngày", f"{(float(sold['Da_ban'].sum()) / period_days) if not sold.empty else 0:,.2f}")

            best_color_display = best_color.copy()
            best_color_display["Ban/ngay"] = (best_color_display["Da_ban"] / period_days).round(2)

            best_size_display = best_size.copy()
            best_size_display["Ban/ngay"] = (best_size_display["Da_ban"] / period_days).round(2)

            st.subheader("Top mã + màu bán chạy")
            st.dataframe(best_color_display.head(50), use_container_width=True, height=320)

            st.subheader("Top mã + màu + size bán chạy")
            st.dataframe(best_size_display.head(50), use_container_width=True, height=320)

            s1, s2 = st.columns(2)
            with s1:
                st.write("Top Group bán tốt")
                group_sold = sold.groupby("Group", as_index=False)["Da_ban"].sum().sort_values("Da_ban", ascending=False).set_index("Group")
                if not group_sold.empty:
                    st.bar_chart(group_sold)
                else:
                    st.info("Không có dữ liệu.")
            with s2:
                st.write("Top mã + màu bán tốt")
                chart_color = best_color_display.head(10).copy()
                if not chart_color.empty:
                    chart_color["Label"] = chart_color["Ma SP"].astype(str) + " - " + chart_color["Mau"].astype(str)
                    st.bar_chart(chart_color.set_index("Label")[["Da_ban"]])
                else:
                    st.info("Không có dữ liệu.")

            st.subheader("Dữ liệu so sánh chi tiết")
            st.dataframe(compare.sort_values("Da_ban", ascending=False), use_container_width=True, height=320)

            st.subheader("Business AI cho bán hàng")
            st.markdown('<div class="ai-box">', unsafe_allow_html=True)

            summary_group_sales = sold.groupby("Group", as_index=False)["Da_ban"].sum().rename(columns={"Da_ban": "Ton kho"})
            summary_group_sales["Gia tri ton"] = 0
            summary_ma_sales = sold.groupby(["Group", "Ma SP"], as_index=False)["Da_ban"].sum().rename(columns={"Da_ban": "Ton kho"})
            summary_ma_sales["Gia tri ton"] = 0
            need_import_sales = pd.DataFrame(columns=["Group", "Ma SP", "Mau", "Size", "Ton kho", "SL de xuat nhap", "Phan loai"])
            out_of_stock_sales = pd.DataFrame(columns=["Group", "Ma SP", "Mau", "Size", "Ton kho", "SL de xuat nhap"])
            sale_df_sales = pd.DataFrame(columns=["Group", "Ma SP", "Mau", "Size", "Ton kho", "Gia tri ton", "Sale priority"])

            ai_prompt_sales = build_ai_prompt_business(
                summary_group_sales,
                summary_ma_sales,
                need_import_sales,
                out_of_stock_sales,
                sale_df_sales,
                best_color_display,
                best_size_display
            )

            extra_note_sales = st.text_area(
                "Ghi chú thêm cho AI về tab bán hàng (tuỳ chọn)",
                placeholder="Ví dụ: ưu tiên tìm sản phẩm để live tuần này, hoặc chỉ tập trung nhóm dép sục...",
                height=100,
                key="sales_ai_note"
            )

            if st.button("Phân tích bán hàng bằng AI", key="sales_ai_button"):
                if extra_note_sales.strip():
                    final_prompt_sales = f"""{ai_prompt_sales}

Bối cảnh bổ sung từ người dùng:
{extra_note_sales}
"""
                else:
                    final_prompt_sales = ai_prompt_sales

                with st.spinner("Đang phân tích góc nhìn business cho bán hàng..."):
                    ai_text_sales, ai_error_sales = run_ai_analysis(final_prompt_sales)

                if ai_error_sales:
                    st.warning(ai_error_sales)
                    st.code(final_prompt_sales, language="text")
                else:
                    st.markdown(ai_text_sales)
            else:
                st.caption("Có thể dùng AI để chốt nhanh nhóm hot, mã hot, và hướng nhập hàng / livestream.")
            st.markdown('</div>', unsafe_allow_html=True)

            sales_xlsx = BytesIO()
            with pd.ExcelWriter(sales_xlsx, engine="openpyxl") as writer:
                compare.to_excel(writer, sheet_name="SoSanhChiTiet", index=False)
                sold.to_excel(writer, sheet_name="ChiTietDaBan", index=False)
                best_color_display.to_excel(writer, sheet_name="BanChay_MaMau", index=False)
                best_size_display.to_excel(writer, sheet_name="BanChay_MaMauSize", index=False)
            sales_xlsx.seek(0)

            st.download_button(
                "Tải file phân tích bán hàng",
                data=sales_xlsx,
                file_name="merly_phan_tich_ban_hang.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Lỗi khi so sánh 2 file: {e}")
    else:
        st.info("Hãy tải đủ 2 file để bắt đầu so sánh.")
