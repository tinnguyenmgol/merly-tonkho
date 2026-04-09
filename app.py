
from pathlib import Path
import re
from io import BytesIO
from datetime import date, datetime
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Merly - Business AI V3.1", layout="wide")

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

def summarize_history_by_group(history_df):
    if history_df.empty:
        return pd.DataFrame()
    df = history_df.copy()
    df["Da_ban"] = pd.to_numeric(df["Da_ban"], errors="coerce").fillna(0)
    return df.groupby(["Ngay_file_truoc", "Ngay_file_hien_tai", "Group"], as_index=False)["Da_ban"].sum()

col_logo, col_title = st.columns([1, 4])
with col_logo:
    if Path("Logo Merly.jpg").exists():
        st.image("Logo Merly.jpg", width=120)
with col_title:
    st.title("Merly - Business AI V3.1")
    st.caption("Đã sửa logic số âm: âm không còn bị xem là bán tốt. Âm nhẹ = trả hàng/tăng tồn nhẹ, âm lớn = hàng mới về lại.")

tab1, tab2 = st.tabs(["Tồn kho & Business AI", "Phân tích bán hàng"])

with tab1:
    st.info("Bản fix này tập trung sửa tab 2 theo phản hồi của anh. Tab 1 giữ như bản trước.")

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
        st.caption("Biểu đồ này chỉ tính số bán dương. Các dòng âm không được tính là bán.")
        group_sales_chart = sold.groupby("Group", as_index=False)["Da_ban"].sum().sort_values("Da_ban", ascending=False)
        if not group_sales_chart.empty:
            st.bar_chart(group_sales_chart.set_index("Group")[["Da_ban"]])
        else:
            st.info("Không có dữ liệu bán dương để vẽ biểu đồ.")

        st.subheader("Báo cáo bán hàng nhanh theo mã + màu")
        st.caption("Lưu ý: số âm không được xem là bán tốt. Âm nhẹ thường là trả hàng/đổi size, âm lớn thường là hàng mới về lại.")
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

        st.subheader("Đề xuất xử lý")
        st.markdown("""
- **Âm nhẹ (-1, -2)**: khả năng là trả hàng / đổi size. Chưa xem là bán chậm.  
- **Âm lớn (<= -3)**: thường là hàng mới về lại / nhập thêm tồn. Không xếp vào bán tốt hay bán chậm.  
- **Không bán**: combo, voucher riêng, giảm giá ngắn hạn.  
- **Bán 1–3 đôi**: test ảnh, caption, live, seeding rồi mới giảm sâu.  
- **Bán tốt**: giữ giá, đẩy mạnh.
""")

        if st.button("Lưu dữ liệu so sánh này", key="save_history"):
            save_compare_history(compare, date_old, date_new, file_old.name, file_new.name)
            st.success("Đã lưu dữ liệu so sánh.")

        hist = load_compare_history()
        if not hist.empty:
            st.subheader("Lịch sử so sánh đã lưu")
            hist_group = summarize_history_by_group(hist)
            if not hist_group.empty:
                group_options = sorted(hist_group["Group"].dropna().astype(str).unique().tolist())
                selected_group = st.selectbox("Chọn nhóm để xem lịch sử", group_options)
                gdf = hist_group[hist_group["Group"].astype(str) == str(selected_group)].copy()
                gdf["Ky_so_sanh"] = gdf["Ngay_file_truoc"].astype(str) + " → " + gdf["Ngay_file_hien_tai"].astype(str)
                if not gdf.empty:
                    st.bar_chart(gdf.set_index("Ky_so_sanh")[["Da_ban"]])
            st.dataframe(hist, use_container_width=True, height=260)
