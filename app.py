
from pathlib import Path
import re
from io import BytesIO
from datetime import date, datetime
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Merly - Business AI V4", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 2rem; max-width: 96%;}
h1, h2, h3 {color: #9d8479;}
.ai-box {background:#fff; border:1px solid #eadfd9; border-radius:14px; padding:16px;}
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
.badge-box {margin-top:6px; margin-bottom:6px; display:flex; gap:10px; flex-wrap:wrap;}
.badge {display:inline-block; padding:4px 10px; border-radius:999px; font-size:12px; font-weight:700;}
.badge-red {background:#ffd9d9; color:#9b0000;}
.badge-yellow {background:#fff2cc; color:#8a6d00;}
.badge-blue {background:#dceeff; color:#0b5cab;}
.badge-green {background:#ddf6e8; color:#137a45;}
</style>
""", unsafe_allow_html=True)

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
HISTORY_CSV = DATA_DIR / "sales_compare_history.csv"

def split_data(text):
    text = str(text).strip()
    if text == "":
        return pd.Series(["", "", ""])
    parts = text.split(" ")
    if len(parts) < 2:
        return pd.Series(["", "", ""])
    ma = parts[0].strip()
    size = ""
    mau = ""
    for p in parts[1:]:
        p = str(p).strip()
        if p.isdigit() and len(p) == 2:
            size = p
        else:
            mau += (" " if mau else "") + p
    return pd.Series([ma, size, mau.strip()])

def process_inventory_file(file):
    df = pd.read_excel(file)
    if df.shape[1] < 8:
        raise ValueError("File cần ít nhất 8 cột: A = Tên hàng hóa, C = Tồn kho, E = Giá bán, H = Group.")
    df[["Ma SP", "Size", "Mau"]] = df.iloc[:, 0].apply(split_data)
    df["Gia Ban"] = pd.to_numeric(df.iloc[:, 4], errors="coerce").fillna(0)
    df["Ton kho"] = pd.to_numeric(df.iloc[:, 2], errors="coerce").fillna(0).astype(int)
    df["Group"] = df.iloc[:, 7].astype(str).fillna("").str.strip()
    out = df[["Ma SP", "Size", "Mau", "Gia Ban", "Ton kho", "Group"]].copy()
    for c in ["Ma SP", "Size", "Mau", "Group"]:
        out[c] = out[c].astype(str).fillna("").str.strip()
    return out

def parse_date_from_filename(filename):
    if not filename:
        return None
    name = Path(filename).stem
    candidates = []
    for d in re.findall(r"(\d{8})", name):
        for fmt in ("%d%m%Y", "%m%d%Y", "%Y%m%d"):
            try:
                dt = datetime.strptime(d, fmt).date()
                if 2020 <= dt.year <= 2100:
                    candidates.append(dt)
            except Exception:
                pass
    return candidates[0] if candidates else None

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
    for c in ["Group", "Ma SP", "Mau"]:
        df2[c] = df2[c].astype(str).fillna("").str.strip()
    df2 = df2[df2["Ton kho"] > 0].copy()
    df2 = df2[(df2["Group"] != "") & (df2["Ma SP"] != "") & (df2["Mau"] != "")].copy()
    all_sizes = sorted([s for s in df2["Size"].dropna().unique() if str(s).isdigit()], key=lambda x: int(x))
    if df2.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), all_sizes
    pivot = pd.pivot_table(df2, index=["Group", "Ma SP", "Mau"], columns="Size", values="Ton kho", aggfunc="sum", fill_value=0)
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
    html.append('<div class="badge-box"><span class="badge badge-red">Tồn = 1</span><span class="badge badge-yellow">Tồn 2–3</span><span class="badge badge-blue">Ẩn ô 0, blank và mã tổng = 0</span></div>')
    html.append('<div class="pivot-wrap"><table class="pivot-table"><thead><tr><th class="pivot-col-left">Row Labels</th>')
    for s in all_sizes:
        html.append(f'<th>{s}</th>')
    html.append('<th class="total-col">Grand Total</th></tr></thead><tbody>')
    all_total = 0
    for group in group_total.sort_values("Group")["Group"].tolist():
        gsum = int(group_total.loc[group_total["Group"] == group, "Grand Total"].iloc[0])
        if gsum == 0:
            continue
        all_total += gsum
        gblock = pivot[pivot["Group"] == group].copy()
        html.append(f'<tr class="group-row"><td class="pivot-col-left">▶ {group}</td>')
        for s in all_sizes:
            v = int(gblock[s].sum()) if s in gblock.columns else 0
            html.append(f'<td>{"" if v == 0 else v}</td>')
        html.append(f'<td class="total-col">{gsum}</td></tr>')
        ma_list = ma_total[ma_total["Group"] == group].sort_values("Ma SP")["Ma SP"].tolist()
        for ma in ma_list:
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

def build_sales_compare(old_file, new_file):
    old_df = process_inventory_file(old_file)[["Ma SP", "Mau", "Size", "Ton kho", "Group"]].copy()
    new_df = process_inventory_file(new_file)[["Ma SP", "Mau", "Size", "Ton kho", "Group"]].copy()
    for df_ in (old_df, new_df):
        for c in ["Ma SP", "Mau", "Size", "Group"]:
            df_[c] = df_[c].astype(str).fillna("").str.strip()
    old_df = old_df[(old_df["Ma SP"] != "") & (old_df["Mau"] != "") & (old_df["Size"] != "")]
    new_df = new_df[(new_df["Ma SP"] != "") & (new_df["Mau"] != "") & (new_df["Size"] != "")]
    old_df = old_df.groupby(["Ma SP", "Mau", "Size", "Group"], as_index=False)["Ton kho"].sum()
    new_df = new_df.groupby(["Ma SP", "Mau", "Size", "Group"], as_index=False)["Ton kho"].sum()
    old_df = old_df.rename(columns={"Ton kho": "Ton_cu"})
    new_df = new_df.rename(columns={"Ton kho": "Ton_moi"})
    compare = old_df.merge(new_df, on=["Ma SP", "Mau", "Size"], how="outer", suffixes=("_old", "_new"))
    compare["Group"] = compare["Group_old"].combine_first(compare["Group_new"])
    compare = compare.drop(columns=[c for c in ["Group_old", "Group_new"] if c in compare.columns])
    compare["Group"] = compare["Group"].astype(str).fillna("").str.strip()
    compare["Ton_cu"] = pd.to_numeric(compare["Ton_cu"], errors="coerce").fillna(0).astype(int)
    compare["Ton_moi"] = pd.to_numeric(compare["Ton_moi"], errors="coerce").fillna(0).astype(int)
    compare["Da_ban"] = compare["Ton_cu"] - compare["Ton_moi"]
    compare["Nhap_them"] = compare["Ton_moi"] - compare["Ton_cu"]
    compare = compare[(compare["Ton_cu"] != 0) | (compare["Ton_moi"] != 0) | (compare["Da_ban"] != 0)].copy()
    sold = compare[compare["Da_ban"] > 0].copy()
    best_color = sold.groupby(["Group", "Ma SP", "Mau"], as_index=False)["Da_ban"].sum().sort_values("Da_ban", ascending=False)
    best_size = sold.groupby(["Group", "Ma SP", "Mau", "Size"], as_index=False)["Da_ban"].sum().sort_values("Da_ban", ascending=False)
    return compare, sold, best_color, best_size

def classify_sale(x):
    x = int(x) if pd.notnull(x) else 0
    if x <= -3:
        return "Nhập hàng lại / tăng tồn"
    elif x in (-2, -1):
        return "Trả hàng / tăng tồn nhẹ"
    elif x == 0:
        return "Không bán"
    elif x == 1:
        return "Bán 1 đôi"
    elif x == 2:
        return "Bán 2 đôi"
    elif x == 3:
        return "Bán 3 đôi"
    return "Bán tốt"

def suggest_sale_action(row):
    sold = int(row["Da_ban"]) if pd.notnull(row["Da_ban"]) else 0
    if sold <= -3:
        return "Kiểm tra nhập thêm / hàng về lại"
    elif sold in (-2, -1):
        return "Theo dõi trả hàng / đổi size"
    elif sold == 0:
        return "Combo / voucher riêng / giảm giá"
    elif sold == 1:
        return "Voucher nhẹ + test ảnh"
    elif sold == 2:
        return "Đẩy live + seeding"
    elif sold == 3:
        return "Scale nhẹ / giữ quan sát"
    return "Giữ giá / đẩy mạnh"

def build_sales_report(compare):
    report = compare.groupby(["Group", "Ma SP", "Mau"], as_index=False)["Da_ban"].sum()
    report["Phân loại bán"] = report["Da_ban"].apply(classify_sale)
    report["Đề xuất"] = report.apply(suggest_sale_action, axis=1)
    report = report.sort_values(["Da_ban", "Group", "Ma SP", "Mau"])
    negative_light = report[report["Da_ban"].isin([-2, -1])]
    negative_heavy = report[report["Da_ban"] <= -3]
    zero_sale = report[report["Da_ban"] == 0]
    sale_1 = report[report["Da_ban"] == 1]
    sale_2 = report[report["Da_ban"] == 2]
    sale_3 = report[report["Da_ban"] == 3]
    return report, negative_light, negative_heavy, zero_sale, sale_1, sale_2, sale_3

def save_compare_history(compare_df, old_date, new_date, old_name, new_name):
    if compare_df.empty:
        return
    save_df = compare_df.copy()
    save_df["Ngay_file_truoc"] = str(old_date)
    save_df["Ngay_file_hien_tai"] = str(new_date)
    save_df["Ten_file_truoc"] = str(old_name)
    save_df["Ten_file_hien_tai"] = str(new_name)
    save_df["Ky_so_sanh"] = str(old_date) + " → " + str(new_date)
    save_df["Ngay_luu"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if HISTORY_CSV.exists():
        old = pd.read_csv(HISTORY_CSV)
        save_df = pd.concat([old, save_df], ignore_index=True)
    save_df.to_csv(HISTORY_CSV, index=False, encoding="utf-8-sig")

def load_compare_history():
    if HISTORY_CSV.exists():
        try:
            return pd.read_csv(HISTORY_CSV)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

def build_history_insights(hist_group):
    if hist_group.empty:
        return []
    out = []
    tmp = hist_group.copy()
    tmp["Da_ban"] = pd.to_numeric(tmp["Da_ban"], errors="coerce").fillna(0)
    tmp["Ky_so_sanh"] = tmp["Ngay_file_truoc"].astype(str) + " → " + tmp["Ngay_file_hien_tai"].astype(str)
    for grp, g in tmp.groupby("Group"):
        g = g.sort_values(["Ngay_file_truoc", "Ngay_file_hien_tai"]).tail(2)
        if len(g) >= 2:
            vals = g["Da_ban"].tolist()
            diff = vals[-1] - vals[-2]
            if diff >= 5:
                out.append(f"Nhóm {grp} đang tăng rõ ({vals[-2]} → {vals[-1]}).")
            elif diff <= -5:
                out.append(f"Nhóm {grp} đang giảm rõ ({vals[-2]} → {vals[-1]}).")
    return out[:10]

col_logo, col_title = st.columns([1, 4])
with col_logo:
    if Path("Logo Merly.jpg").exists():
        st.image("Logo Merly.jpg", width=120)
with col_title:
    st.title("Merly - Business AI V4")
    st.caption("Bản nâng cao: có tab Lịch sử so sánh PRO, biểu đồ xu hướng theo nhóm, bộ lọc group/mã và cảnh báo tăng giảm.")

tab1, tab2, tab3 = st.tabs(["Tồn kho", "Phân tích bán hàng", "Lịch sử so sánh PRO"])

with tab1:
    uploaded_file = st.file_uploader("Tải file Excel tồn kho", type=["xlsx"], key="inventory_file")
    if uploaded_file:
        df_clean = process_inventory_file(uploaded_file)
        df_clean["Phan loai"] = df_clean["Ton kho"].apply(classify_ton_kho)
        df_clean["Gia tri ton"] = df_clean["Gia Ban"] * df_clean["Ton kho"]
        df_clean["SL de xuat nhap"] = df_clean.apply(suggested_restock, axis=1)
        df_clean["Sale priority"] = df_clean.apply(sale_priority, axis=1)
        st.dataframe(df_clean, use_container_width=True, height=340)
        pivot_detail, group_total, ma_total, all_sizes = build_pivot_hierarchical(df_clean)
        st.markdown(render_pivot_html(pivot_detail, group_total, ma_total, all_sizes), unsafe_allow_html=True)
    else:
        st.info("Hãy tải file Excel tồn kho lên để bắt đầu.")

with tab2:
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        file_old = st.file_uploader("File tồn kho TRƯỚC", type=["xlsx"], key="sales_old")
    with c2:
        auto_old = parse_date_from_filename(file_old.name) if file_old else None
        date_old = st.date_input("Ngày file TRƯỚC", value=auto_old or date.today(), key="date_old")
    with c3:
        file_new = st.file_uploader("File tồn kho HIỆN TẠI", type=["xlsx"], key="sales_new")
    with c4:
        auto_new = parse_date_from_filename(file_new.name) if file_new else None
        date_new = st.date_input("Ngày file HIỆN TẠI", value=auto_new or date.today(), key="date_new")

    period_days = abs((date_new - date_old).days)
    if period_days == 0:
        period_days = 1
    st.caption(f"Số ngày giữa 2 file: {period_days} ngày")

    if file_old and file_new:
        compare, sold, best_color, best_size = build_sales_compare(file_old, file_new)
        report, negative_light, negative_heavy, zero_sale, sale_1, sale_2, sale_3 = build_sales_report(compare)

        tong_ban = int(sold["Da_ban"].sum()) if not sold.empty else 0
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Tổng SKU có bán", f"{len(sold):,}")
        m2.metric("Tổng lượng bán suy ra", f"{tong_ban:,}")
        m3.metric("Âm nhẹ (-1,-2)", f"{len(negative_light):,}")
        m4.metric("Âm lớn (<=-3)", f"{len(negative_heavy):,}")

        st.subheader("Biểu đồ bán hàng theo nhóm")
        st.caption("Biểu đồ này chỉ tính số bán dương.")
        group_sales_chart = sold.groupby("Group", as_index=False)["Da_ban"].sum().sort_values("Da_ban", ascending=False)
        if not group_sales_chart.empty:
            st.bar_chart(group_sales_chart.set_index("Group")[["Da_ban"]])
        else:
            st.info("Không có dữ liệu bán dương để vẽ biểu đồ.")

        st.subheader("Báo cáo bán hàng nhanh theo mã + màu")
        st.caption("Số âm không được xem là bán tốt.")
        st.dataframe(report, use_container_width=True, height=320)

        g1, g2 = st.columns(2)
        with g1:
            st.write("Âm nhẹ (-1, -2)")
            st.dataframe(negative_light, use_container_width=True, height=180)
            st.write("Âm lớn (<= -3)")
            st.dataframe(negative_heavy, use_container_width=True, height=180)
            st.write("Không bán được đôi nào")
            st.dataframe(zero_sale, use_container_width=True, height=180)
        with g2:
            st.write("Bán 1 đôi")
            st.dataframe(sale_1, use_container_width=True, height=180)
            st.write("Bán 2 đôi")
            st.dataframe(sale_2, use_container_width=True, height=180)
            st.write("Bán 3 đôi")
            st.dataframe(sale_3, use_container_width=True, height=180)

        if st.button("Lưu dữ liệu so sánh này", key="save_history"):
            save_compare_history(compare, date_old, date_new, file_old.name, file_new.name)
            st.success("Đã lưu dữ liệu so sánh.")

with tab3:
    st.subheader("Lịch sử so sánh PRO")
    history_df = load_compare_history()
    if history_df.empty:
        st.info("Chưa có dữ liệu lịch sử. Vào tab 'Phân tích bán hàng' rồi bấm 'Lưu dữ liệu so sánh này'.")
    else:
        history_df["Da_ban"] = pd.to_numeric(history_df["Da_ban"], errors="coerce").fillna(0)
        history_df["Ky_so_sanh"] = history_df["Ngay_file_truoc"].astype(str) + " → " + history_df["Ngay_file_hien_tai"].astype(str)

        f1, f2 = st.columns(2)
        groups = sorted(history_df["Group"].dropna().astype(str).unique().tolist())
        masp = sorted(history_df["Ma SP"].dropna().astype(str).unique().tolist())
        with f1:
            selected_group = st.selectbox("Lọc theo Group", ["Tất cả"] + groups)
        with f2:
            selected_masp = st.selectbox("Lọc theo Mã SP", ["Tất cả"] + masp)

        hist_filtered = history_df.copy()
        if selected_group != "Tất cả":
            hist_filtered = hist_filtered[hist_filtered["Group"].astype(str) == selected_group]
        if selected_masp != "Tất cả":
            hist_filtered = hist_filtered[hist_filtered["Ma SP"].astype(str) == selected_masp]

        h1, h2, h3 = st.columns(3)
        h1.metric("Dòng lịch sử", f"{len(hist_filtered):,}")
        h2.metric("Kỳ so sánh", f"{hist_filtered['Ky_so_sanh'].nunique():,}")
        h3.metric("Tổng bán dương", f"{int(hist_filtered[hist_filtered['Da_ban']>0]['Da_ban'].sum()):,}")

        hist_group = hist_filtered.groupby(["Ky_so_sanh", "Group"], as_index=False)["Da_ban"].sum().sort_values(["Ky_so_sanh","Da_ban"], ascending=[True,False])

        st.subheader("Biểu đồ xu hướng theo nhóm")
        if not hist_group.empty:
            group_for_chart = st.selectbox("Chọn nhóm để xem trend", sorted(hist_group["Group"].dropna().astype(str).unique().tolist()), key="trend_group")
            gchart = hist_group[hist_group["Group"].astype(str) == str(group_for_chart)].copy()
            if not gchart.empty:
                st.bar_chart(gchart.set_index("Ky_so_sanh")[["Da_ban"]])
        else:
            st.info("Không có đủ dữ liệu để vẽ trend.")

        st.subheader("Cảnh báo xu hướng")
        insights = build_history_insights(hist_group)
        if insights:
            for item in insights:
                st.write(f"- {item}")
        else:
            st.caption("Chưa đủ dữ liệu để sinh cảnh báo tăng/giảm rõ rệt.")

        st.subheader("So sánh 2 kỳ bất kỳ")
        periods = sorted(hist_filtered["Ky_so_sanh"].dropna().astype(str).unique().tolist())
        if len(periods) >= 2:
            s1, s2 = st.columns(2)
            with s1:
                p_old = st.selectbox("Kỳ 1", periods, key="period1")
            with s2:
                p_new = st.selectbox("Kỳ 2", periods, index=min(1, len(periods)-1), key="period2")
            df1 = hist_filtered[hist_filtered["Ky_so_sanh"].astype(str) == p_old].groupby(["Group","Ma SP","Mau"], as_index=False)["Da_ban"].sum().rename(columns={"Da_ban":"Ban_ky_1"})
            df2 = hist_filtered[hist_filtered["Ky_so_sanh"].astype(str) == p_new].groupby(["Group","Ma SP","Mau"], as_index=False)["Da_ban"].sum().rename(columns={"Da_ban":"Ban_ky_2"})
            merged = df1.merge(df2, on=["Group","Ma SP","Mau"], how="outer").fillna(0)
            merged["Chenhlech"] = merged["Ban_ky_2"] - merged["Ban_ky_1"]
            st.dataframe(merged.sort_values("Chenhlech", ascending=False), use_container_width=True, height=280)
        else:
            st.caption("Cần ít nhất 2 kỳ đã lưu để so sánh kỳ.")

        st.subheader("Dữ liệu lịch sử chi tiết")
        st.dataframe(hist_filtered, use_container_width=True, height=320)

        hist_xlsx = BytesIO()
        with pd.ExcelWriter(hist_xlsx, engine="openpyxl") as writer:
            hist_filtered.to_excel(writer, sheet_name="LichSuChiTiet", index=False)
            hist_group.to_excel(writer, sheet_name="TongHopTheoGroup", index=False)
        hist_xlsx.seek(0)
        st.download_button("Tải dữ liệu lịch sử", data=hist_xlsx, file_name="merly_lich_su_so_sanh_v4.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
