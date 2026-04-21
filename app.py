
from pathlib import Path
from io import BytesIO
from datetime import datetime
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Merly Dashboard V3 Chiến Lược", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 2rem; max-width: 96%;}
h1, h2, h3 {color: #9d8479;}
.note-box {background: #f8f1ed; border: 1px solid #eadfd9; border-radius: 14px; padding: 14px 16px; margin-bottom: 12px;}
.ai-box {background: #ffffff; border: 1px solid #eadfd9; border-radius: 14px; padding: 16px;}
.kpi-note {font-size: 12px; color: #7d6a61;}
</style>
""", unsafe_allow_html=True)

def clean_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df

def parse_report_dates(raw_df):
    start_date = None
    end_date = None
    for i in range(min(10, len(raw_df))):
        left = str(raw_df.iloc[i, 0]).strip()
        right = raw_df.iloc[i, 1] if raw_df.shape[1] > 1 else None
        if left == "Từ ngày":
            start_date = pd.to_datetime(right, dayfirst=True, errors="coerce")
        elif left == "Đến ngày":
            end_date = pd.to_datetime(right, dayfirst=True, errors="coerce")
    return start_date, end_date

def parse_summary_metrics(raw_df):
    metrics = {}
    mapping = {
        "Tiền bán hàng": "gross_sales",
        "Hàng hóa bán": "gross_qty",
        "Tiền trả hàng": "return_sales",
        "Hàng hóa trả": "return_qty",
    }
    for i in range(min(10, len(raw_df))):
        left = str(raw_df.iloc[i, 0]).strip()
        val = raw_df.iloc[i, 1] if raw_df.shape[1] > 1 else None
        if left in mapping:
            metrics[mapping[left]] = pd.to_numeric(val, errors="coerce")
    return metrics

def parse_product_name(name):
    text = str(name).strip()
    tokens = text.split()
    model_code = ""
    size = ""
    color_tokens = tokens[:]

    for idx, tok in enumerate(tokens):
        if re.fullmatch(r"\d{3,5}(?:-[0-9A-Za-z]+)?", tok):
            model_code = tok
            color_tokens = tokens[:idx] + tokens[idx+1:]
            break

    if not model_code and tokens:
        model_code = tokens[0]
        color_tokens = tokens[1:]

    size_idx = None
    for idx, tok in enumerate(color_tokens):
        if tok.isdigit():
            n = int(tok)
            if 30 <= n <= 50:
                size = tok
                size_idx = idx
    if size_idx is not None:
        color_tokens = color_tokens[:size_idx] + color_tokens[size_idx+1:]

    color = " ".join(color_tokens).strip()
    return pd.Series([model_code.strip(), color, size])

def load_sales_file(file):
    raw = pd.read_excel(file, header=None)
    start_date, end_date = parse_report_dates(raw)
    summary = parse_summary_metrics(raw)
    df = pd.read_excel(file, header=7)
    df = clean_columns(df)

    expected = ['Mã hàng hóa', 'Tên hàng hóa', 'SL bán', 'Tiền bán hàng', 'SL trả', 'Tiền trả hàng']
    missing = [c for c in expected if c not in df.columns]
    if missing:
        raise ValueError(f"Thiếu cột trong file doanh số: {missing}")

    for c in ["SL bán", "Tiền bán hàng", "SL trả", "Tiền trả hàng"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df[["Ma SP", "Mau", "Size"]] = df["Tên hàng hóa"].apply(parse_product_name)
    df["SL thuần"] = df["SL bán"] - df["SL trả"]
    df["Doanh thu thuần"] = df["Tiền bán hàng"] - df["Tiền trả hàng"]

    if start_date is not None and end_date is not None:
        days = max((end_date - start_date).days + 1, 1)
    else:
        days = 1

    summary_final = {
        "start_date": start_date,
        "end_date": end_date,
        "days": days,
        "gross_sales": float(summary.get("gross_sales", df["Tiền bán hàng"].sum()) or 0),
        "gross_qty": float(summary.get("gross_qty", df["SL bán"].sum()) or 0),
        "return_sales": float(summary.get("return_sales", df["Tiền trả hàng"].sum()) or 0),
        "return_qty": float(summary.get("return_qty", df["SL trả"].sum()) or 0),
    }
    summary_final["net_sales"] = summary_final["gross_sales"] - summary_final["return_sales"]
    summary_final["net_qty"] = summary_final["gross_qty"] - summary_final["return_qty"]
    summary_final["return_rate_qty"] = (summary_final["return_qty"] / summary_final["gross_qty"]) if summary_final["gross_qty"] else 0
    summary_final["qty_per_day"] = summary_final["net_qty"] / days if days else 0
    summary_final["sales_per_day"] = summary_final["net_sales"] / days if days else 0
    return df, summary_final

def load_opening_stock_file(file):
    xl = pd.ExcelFile(file)
    sheet_names = xl.sheet_names

    # ưu tiên sheet ThongKe nếu có
    if "ThongKe" in sheet_names:
        df = pd.read_excel(file, sheet_name="ThongKe")
        df = clean_columns(df)
        needed = {"Ma SP", "Mau", "Size", "Ton kho", "Group"}
        if needed.issubset(set(df.columns)):
            stock = df.copy()
            stock["Ma SP"] = stock["Ma SP"].astype(str).str.strip()
            stock["Mau"] = stock["Mau"].astype(str).str.strip()
            stock["Size"] = stock["Size"].astype(str).str.strip()
            stock["Tồn đầu"] = pd.to_numeric(stock["Ton kho"], errors="coerce").fillna(0)
            stock["Group"] = stock["Group"].astype(str).str.strip()
            return stock[["Ma SP", "Mau", "Size", "Tồn đầu", "Group"]]

    # fallback sheet danh sách hàng hóa
    for sheet in sheet_names:
        df = pd.read_excel(file, sheet_name=sheet)
        df = clean_columns(df)
        needed = {"Tên hàng hóa", "Số lượng", "Nhóm hàng"}
        if needed.issubset(set(df.columns)):
            stock = df.copy()
            stock[["Ma SP", "Mau", "Size"]] = stock["Tên hàng hóa"].apply(parse_product_name)
            stock["Tồn đầu"] = pd.to_numeric(stock["Số lượng"], errors="coerce").fillna(0)
            stock["Group"] = stock["Nhóm hàng"].astype(str).str.strip()
            stock["Ma SP"] = stock["Ma SP"].astype(str).str.strip()
            stock["Mau"] = stock["Mau"].astype(str).str.strip()
            stock["Size"] = stock["Size"].astype(str).str.strip()
            return stock[["Ma SP", "Mau", "Size", "Tồn đầu", "Group"]]

    raise ValueError("Không đọc được file tồn đầu kỳ. Cần có sheet 'ThongKe' hoặc sheet có cột: Tên hàng hóa, Số lượng, Nhóm hàng.")

def classify_strategy(row):
    qty = row["SL thuần"]
    opening = row.get("Tồn đầu", 0)
    item_type = row.get("Loại hàng", "")

    if qty < 0:
        return "Theo dõi trả hàng"
    if item_type == "Hàng mới":
        if qty >= 5:
            return "Mẫu mới bán tốt"
        if qty >= 1:
            return "Mẫu mới cần theo dõi"
        return "Mẫu mới chưa phát sinh"
    # hàng cũ
    sell_through = (qty / opening) if opening and opening > 0 else 0
    if qty >= 10 or sell_through >= 0.5:
        return "Winner cần nhập thêm"
    if qty >= 5 or sell_through >= 0.25:
        return "Bán ổn giữ giá"
    if qty >= 2:
        return "Bán chậm cần kích"
    if qty == 1:
        return "Rất chậm cần ưu tiên"
    if qty == 0 and opening >= 5:
        return "Tồn cao không bán"
    if qty == 0:
        return "Không bán"
    return "Khác"

def suggest_action(row):
    strat = row["Chiến lược"]
    if strat == "Winner cần nhập thêm":
        return "Nhập thêm size hot + đẩy ads/live"
    if strat == "Bán ổn giữ giá":
        return "Giữ giá + duy trì hiển thị"
    if strat == "Bán chậm cần kích":
        return "Combo nhẹ + voucher nhóm"
    if strat == "Rất chậm cần ưu tiên":
        return "Test ảnh + caption + flash sale ngắn"
    if strat == "Tồn cao không bán":
        return "Xả hàng / combo / giảm giá"
    if strat == "Không bán":
        return "Voucher riêng / gom live xử lý"
    if strat == "Theo dõi trả hàng":
        return "Kiểm tra lỗi form / đổi size / hoàn"
    if strat == "Mẫu mới bán tốt":
        return "Theo dõi nhập thêm, chưa so tồn đầu"
    if strat == "Mẫu mới cần theo dõi":
        return "Quan sát thêm 1 kỳ trước khi quyết định"
    return "Theo dõi thêm"

def build_strategy_dashboard(sales_df, stock_df):
    work = sales_df.copy()
    stock = stock_df.copy()

    for c in ["Ma SP", "Mau", "Size"]:
        work[c] = work[c].astype(str).str.strip()
        stock[c] = stock[c].astype(str).str.strip()

    merged = work.merge(
        stock,
        on=["Ma SP", "Mau", "Size"],
        how="left",
        indicator=True,
        suffixes=("", "_stock")
    )

    merged["Loại hàng"] = merged["_merge"].map({
        "both": "Có từ đầu kỳ",
        "left_only": "Hàng mới",
        "right_only": "Không phát sinh"
    })
    merged["Tồn đầu"] = pd.to_numeric(merged["Tồn đầu"], errors="coerce").fillna(0)
    merged["Group"] = merged["Group"].fillna("Hàng mới chưa có nhóm")

    # nếu là hàng mới thì không so với đầu kỳ
    merged["So với đầu kỳ"] = merged.apply(
        lambda r: r["SL thuần"] if r["Loại hàng"] == "Hàng mới" else r["SL thuần"] - r["Tồn đầu"],
        axis=1
    )

    merged["Tốc độ/ngày"] = merged["SL thuần"] / max(1, merged.attrs.get("days", 1)) if "days" in merged.attrs else merged["SL thuần"]

    merged["Chiến lược"] = merged.apply(classify_strategy, axis=1)
    merged["Đề xuất"] = merged.apply(suggest_action, axis=1)

    group_summary = merged.groupby("Group", as_index=False).agg(
        SL_thuan=("SL thuần", "sum"),
        Doanh_thu_thuan=("Doanh thu thuần", "sum"),
        Ton_dau=("Tồn đầu", "sum"),
        So_SKU=("Mã hàng hóa", "count"),
    ).sort_values(["SL_thuan", "Doanh_thu_thuan"], ascending=[False, False])

    group_summary["Sell-through"] = group_summary.apply(
        lambda r: (r["SL_thuan"] / r["Ton_dau"]) if r["Ton_dau"] > 0 else 0, axis=1
    )

    model_summary = merged.groupby(["Group", "Ma SP"], as_index=False).agg(
        SL_thuan=("SL thuần", "sum"),
        Doanh_thu_thuan=("Doanh thu thuần", "sum"),
        Ton_dau=("Tồn đầu", "sum"),
        So_mau_size=("Mã hàng hóa", "count"),
    ).sort_values(["SL_thuan", "Doanh_thu_thuan"], ascending=[False, False])

    color_summary = merged.groupby(["Group", "Ma SP", "Mau"], as_index=False).agg(
        SL_thuan=("SL thuần", "sum"),
        Doanh_thu_thuan=("Doanh thu thuần", "sum"),
        Ton_dau=("Tồn đầu", "sum"),
        Loai_hang=("Loại hàng", "first"),
        Chien_luoc=("Chiến lược", "first"),
        De_xuat=("Đề xuất", "first"),
    ).sort_values(["SL_thuan", "Doanh_thu_thuan"], ascending=[False, False])

    new_items = color_summary[color_summary["Loai_hang"] == "Hàng mới"].copy()
    no_sale = color_summary[(color_summary["SL_thuan"] == 0) & (color_summary["Loai_hang"] != "Hàng mới")].copy()
    sold_1 = color_summary[color_summary["SL_thuan"] == 1].copy()
    sold_2 = color_summary[color_summary["SL_thuan"] == 2].copy()
    sold_3 = color_summary[color_summary["SL_thuan"] == 3].copy()
    winners = color_summary[color_summary["Chien_luoc"] == "Winner cần nhập thêm"].copy()
    slow = color_summary[color_summary["Chien_luoc"].isin(["Bán chậm cần kích", "Rất chậm cần ưu tiên", "Tồn cao không bán", "Không bán"])].copy()
    negative = color_summary[color_summary["SL_thuan"] < 0].copy()

    return merged, group_summary, model_summary, color_summary, new_items, no_sale, sold_1, sold_2, sold_3, winners, slow, negative

def build_insights(summary, group_summary, model_summary, new_items, winners, slow, negative):
    out = []
    if not group_summary.empty:
        top = group_summary.iloc[0]
        total = group_summary["SL_thuan"].sum()
        share = (top["SL_thuan"] / total) if total else 0
        out.append(f"Nhóm bán mạnh nhất hiện tại là **{top['Group']}**, bán thuần **{int(top['SL_thuan'])} đôi**, chiếm khoảng **{share:.1%}** lượng bán.")
    if not model_summary.empty:
        top_model = model_summary.iloc[0]
        out.append(f"Mã mẫu bán mạnh nhất là **{top_model['Ma SP']}**, bán thuần **{int(top_model['SL_thuan'])} đôi**.")
    if len(new_items) > 0:
        sold_new = int(new_items["SL_thuan"].sum())
        out.append(f"Có **{len(new_items)} mã+màu là hàng mới về sau**. Tổng lượng bán của hàng mới là **{sold_new} đôi**. Nhóm này chỉ theo dõi bán, không so tồn đầu kỳ.")
    if len(winners) > 0:
        out.append(f"Có **{len(winners)} mã+màu** đang ở trạng thái **Winner cần nhập thêm**.")
    if len(slow) > 0:
        out.append(f"Có **{len(slow)} mã+màu** thuộc nhóm bán chậm / không bán, nên gom chiến dịch combo, voucher hoặc giảm giá.")
    if len(negative) > 0:
        out.append(f"Có **{len(negative)} mã+màu** âm thuần, cần kiểm tra trả hàng, đổi size hoặc dữ liệu nhập bổ sung.")
    if summary["return_rate_qty"] > 0:
        out.append(f"Tỷ lệ trả hàng theo số lượng đang ở mức **{summary['return_rate_qty']:.1%}**.")
    return out[:8]

def export_excel(detail, group_summary, model_summary, color_summary, new_items, no_sale, sold_1, sold_2, sold_3, winners, slow, negative):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        detail.to_excel(writer, sheet_name="ChiTiet", index=False)
        group_summary.to_excel(writer, sheet_name="NhomBanChay", index=False)
        model_summary.to_excel(writer, sheet_name="MaMauBanChay", index=False)
        color_summary.to_excel(writer, sheet_name="MauBanChay", index=False)
        new_items.to_excel(writer, sheet_name="HangMoi", index=False)
        winners.to_excel(writer, sheet_name="WinnerCanNhap", index=False)
        slow.to_excel(writer, sheet_name="BanCham", index=False)
        no_sale.to_excel(writer, sheet_name="KhongBan", index=False)
        sold_1.to_excel(writer, sheet_name="Ban1Doi", index=False)
        sold_2.to_excel(writer, sheet_name="Ban2Doi", index=False)
        sold_3.to_excel(writer, sheet_name="Ban3Doi", index=False)
        negative.to_excel(writer, sheet_name="AmTraHang", index=False)
    output.seek(0)
    return output

st.title("Merly Dashboard V3 Chiến Lược")
st.caption("So sánh doanh số với tồn kho đầu kỳ. Hàng không có trong tồn đầu kỳ được xem là hàng mới về sau: chỉ thống kê bán, không so với đầu kỳ.")

st.markdown('<div class="note-box">', unsafe_allow_html=True)
st.write("""
**Cách dùng**
1. Upload file **doanh số theo hàng hóa**.
2. Upload file **danh sách hàng hóa đầu kỳ**.
3. Dashboard sẽ tự chia:
   - hàng có từ đầu kỳ
   - hàng mới về sau
4. Hàng mới chỉ thống kê lượt bán, **không so sánh với đầu kỳ**.
""")
st.markdown('</div>', unsafe_allow_html=True)

c1, c2 = st.columns(2)
with c1:
    sales_file = st.file_uploader("Upload file doanh số", type=["xlsx"], key="sales_file")
with c2:
    stock_file = st.file_uploader("Upload file tồn kho đầu kỳ", type=["xlsx"], key="stock_file")

if sales_file and stock_file:
    sales_df, summary = load_sales_file(sales_file)
    stock_df = load_opening_stock_file(stock_file)
    detail, group_summary, model_summary, color_summary, new_items, no_sale, sold_1, sold_2, sold_3, winners, slow, negative = build_strategy_dashboard(sales_df, stock_df)
    insights = build_insights(summary, group_summary, model_summary, new_items, winners, slow, negative)

    if summary["start_date"] is not None and summary["end_date"] is not None:
        st.subheader(f"Kỳ báo cáo: {summary['start_date'].strftime('%d/%m/%Y')} → {summary['end_date'].strftime('%d/%m/%Y')}")

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Doanh thu thuần", f"{summary['net_sales']:,.0f}")
    k2.metric("Số lượng thuần", f"{int(summary['net_qty']):,}")
    k3.metric("Bán/ngày", f"{summary['qty_per_day']:.1f}")
    k4.metric("Tỷ lệ trả hàng", f"{summary['return_rate_qty']:.1%}")

    k5, k6, k7, k8 = st.columns(4)
    k5.metric("Hàng mới", f"{len(new_items):,}")
    k6.metric("Winner cần nhập", f"{len(winners):,}")
    k7.metric("Bán chậm / cần kích", f"{len(slow):,}")
    k8.metric("Âm / trả hàng", f"{len(negative):,}")

    st.subheader("Insight chiến lược")
    st.markdown('<div class="ai-box">', unsafe_allow_html=True)
    for item in insights:
        st.write(f"- {item}")
    st.markdown('</div>', unsafe_allow_html=True)

    ch1, ch2 = st.columns(2)
    with ch1:
        st.subheader("Nhóm sản phẩm đang bán chạy")
        if not group_summary.empty:
            st.bar_chart(group_summary.set_index("Group")[["SL_thuan"]])
        else:
            st.info("Chưa có dữ liệu.")
    with ch2:
        st.subheader("Nhóm có sell-through cao")
        sell_chart = group_summary.copy()
        if not sell_chart.empty:
            sell_chart["Sell-through"] = sell_chart["Sell-through"].fillna(0)
            st.bar_chart(sell_chart.set_index("Group")[["Sell-through"]])
        else:
            st.info("Chưa có dữ liệu.")

    st.subheader("Top mã mẫu bán mạnh")
    st.dataframe(model_summary.head(30), use_container_width=True, height=300)

    st.subheader("Top mã + màu bán mạnh")
    st.dataframe(color_summary.head(50), use_container_width=True, height=320)

    p1, p2 = st.columns(2)
    with p1:
        st.write("Hàng mới về sau")
        st.dataframe(new_items, use_container_width=True, height=220)
        st.write("Winner cần nhập thêm")
        st.dataframe(winners, use_container_width=True, height=220)
    with p2:
        st.write("Nhóm bán chậm / cần kích")
        st.dataframe(slow, use_container_width=True, height=220)
        st.write("Nhóm âm / trả hàng")
        st.dataframe(negative, use_container_width=True, height=220)

    t1, t2 = st.columns(2)
    with t1:
        st.write("Không bán")
        st.dataframe(no_sale, use_container_width=True, height=200)
        st.write("Bán 1 đôi")
        st.dataframe(sold_1, use_container_width=True, height=200)
    with t2:
        st.write("Bán 2 đôi")
        st.dataframe(sold_2, use_container_width=True, height=200)
        st.write("Bán 3 đôi")
        st.dataframe(sold_3, use_container_width=True, height=200)

    st.subheader("Dữ liệu chi tiết sau khi đối chiếu tồn đầu kỳ")
    keep_cols = [
        "Mã hàng hóa", "Tên hàng hóa", "Group", "Ma SP", "Mau", "Size",
        "SL bán", "SL trả", "SL thuần", "Tồn đầu", "Loại hàng",
        "So với đầu kỳ", "Chiến lược", "Đề xuất"
    ]
    st.dataframe(detail[keep_cols], use_container_width=True, height=380)

    export_file = export_excel(detail, group_summary, model_summary, color_summary, new_items, no_sale, sold_1, sold_2, sold_3, winners, slow, negative)
    st.download_button(
        "Tải file phân tích chiến lược",
        data=export_file,
        file_name="merly_dashboard_v3_chien_luoc.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Hãy upload đủ 2 file: doanh số và tồn kho đầu kỳ.")
