
from pathlib import Path
from io import BytesIO
from datetime import datetime
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Merly Sales Dashboard Pro", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 2rem; max-width: 96%;}
h1, h2, h3 {color: #9d8479;}
.note-box {
    background: #f8f1ed;
    border: 1px solid #eadfd9;
    border-radius: 14px;
    padding: 14px 16px;
    margin-bottom: 12px;
}
.ai-box {
    background: #ffffff;
    border: 1px solid #eadfd9;
    border-radius: 14px;
    padding: 16px;
}
.kpi-note {
    font-size: 12px;
    color: #7d6a61;
}
</style>
""", unsafe_allow_html=True)

def parse_report_dates(raw_df):
    start_date = None
    end_date = None
    try:
        for i in range(min(10, len(raw_df))):
            left = str(raw_df.iloc[i, 0]).strip()
            right = str(raw_df.iloc[i, 1]).strip()
            if left == "Từ ngày":
                start_date = pd.to_datetime(right, dayfirst=True, errors="coerce")
            elif left == "Đến ngày":
                end_date = pd.to_datetime(right, dayfirst=True, errors="coerce")
    except Exception:
        pass
    return start_date, end_date

def parse_summary_metrics(raw_df):
    metrics = {}
    labels = {
        "Tiền bán hàng": "gross_sales",
        "Hàng hóa bán": "gross_qty",
        "Tiền trả hàng": "return_sales",
        "Hàng hóa trả": "return_qty",
    }
    for i in range(min(10, len(raw_df))):
        left = str(raw_df.iloc[i, 0]).strip()
        val = raw_df.iloc[i, 1]
        if left in labels:
            metrics[labels[left]] = pd.to_numeric(val, errors="coerce")
    return metrics

def parse_product_name(name):
    text = str(name).strip()
    tokens = text.split()
    model_code = ""
    size = ""
    color_tokens = tokens[:]

    # find model code: first token with 3-5 digits or pattern like 1597-7P
    for idx, tok in enumerate(tokens):
        if re.fullmatch(r"\d{3,5}(?:-[0-9A-Za-z]+)?", tok):
            model_code = tok
            color_tokens = tokens[:idx] + tokens[idx+1:]
            break

    # if not found, try first token as fallback
    if not model_code and tokens:
        model_code = tokens[0]
        color_tokens = tokens[1:]

    # find size token in remaining tokens: 30-50
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
    expected = ['Mã hàng hóa', 'Tên hàng hóa', 'SL bán', 'Tiền bán hàng', 'SL trả', 'Tiền trả hàng']
    missing = [c for c in expected if c not in df.columns]
    if missing:
        raise ValueError(f"Thiếu cột trong file: {missing}")
    df["SL bán"] = pd.to_numeric(df["SL bán"], errors="coerce").fillna(0)
    df["Tiền bán hàng"] = pd.to_numeric(df["Tiền bán hàng"], errors="coerce").fillna(0)
    df["SL trả"] = pd.to_numeric(df["SL trả"], errors="coerce").fillna(0)
    df["Tiền trả hàng"] = pd.to_numeric(df["Tiền trả hàng"], errors="coerce").fillna(0)

    df[["Mã mẫu", "Màu", "Size"]] = df["Tên hàng hóa"].apply(parse_product_name)
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
    summary_final["return_rate_sales"] = (summary_final["return_sales"] / summary_final["gross_sales"]) if summary_final["gross_sales"] else 0
    summary_final["qty_per_day"] = summary_final["net_qty"] / days if days else 0
    summary_final["sales_per_day"] = summary_final["net_sales"] / days if days else 0
    return df, summary_final

def classify_sku(net_qty):
    if net_qty >= 10:
        return "Bán mạnh"
    if net_qty >= 5:
        return "Bán ổn"
    if net_qty >= 2:
        return "Bán yếu"
    if net_qty == 1:
        return "Rất yếu"
    if net_qty == 0:
        return "Không bán"
    return "Âm / trả hàng"

def action_recommend(net_qty):
    if net_qty >= 10:
        return "Nhập thêm + đẩy ads/live"
    if net_qty >= 5:
        return "Giữ giá + duy trì hiển thị"
    if net_qty >= 2:
        return "Combo nhẹ + seeding"
    if net_qty == 1:
        return "Voucher nhẹ + test ảnh"
    if net_qty == 0:
        return "Combo / flash sale / voucher riêng"
    return "Kiểm tra trả hàng / đổi size"

def build_dashboard_tables(df, group_map_df=None):
    work = df.copy()

    if group_map_df is not None and not group_map_df.empty:
        group_map_df = group_map_df.copy()
        cols = {c.lower().strip(): c for c in group_map_df.columns}
        if "mã mẫu" in cols and "group" in cols:
            group_map_df = group_map_df[[cols["mã mẫu"], cols["group"]]].copy()
            group_map_df.columns = ["Mã mẫu", "Group"]
            group_map_df["Mã mẫu"] = group_map_df["Mã mẫu"].astype(str).str.strip()
            group_map_df["Group"] = group_map_df["Group"].astype(str).str.strip()
            work["Mã mẫu"] = work["Mã mẫu"].astype(str).str.strip()
            work = work.merge(group_map_df, on="Mã mẫu", how="left")
        else:
            work["Group"] = work["Mã mẫu"]
    else:
        work["Group"] = work["Mã mẫu"]

    work["Phân loại"] = work["SL thuần"].apply(classify_sku)
    work["Đề xuất"] = work["SL thuần"].apply(action_recommend)

    group_summary = work.groupby("Group", as_index=False).agg(
        SL_thuan=("SL thuần", "sum"),
        Doanh_thu_thuan=("Doanh thu thuần", "sum"),
        SL_ban=("SL bán", "sum"),
        SL_tra=("SL trả", "sum"),
        So_SKU=("Mã hàng hóa", "count"),
    ).sort_values(["SL_thuan", "Doanh_thu_thuan"], ascending=[False, False])

    model_summary = work.groupby(["Group", "Mã mẫu"], as_index=False).agg(
        SL_thuan=("SL thuần", "sum"),
        Doanh_thu_thuan=("Doanh thu thuần", "sum"),
        SL_ban=("SL bán", "sum"),
        SL_tra=("SL trả", "sum"),
        So_bien_the=("Mã hàng hóa", "count"),
    ).sort_values(["SL_thuan", "Doanh_thu_thuan"], ascending=[False, False])

    color_summary = work.groupby(["Group", "Mã mẫu", "Màu"], as_index=False).agg(
        SL_thuan=("SL thuần", "sum"),
        Doanh_thu_thuan=("Doanh thu thuần", "sum"),
        SL_ban=("SL bán", "sum"),
        SL_tra=("SL trả", "sum"),
    ).sort_values(["SL_thuan", "Doanh_thu_thuan"], ascending=[False, False])

    size_summary = work.groupby(["Size"], as_index=False).agg(
        SL_thuan=("SL thuần", "sum"),
        Doanh_thu_thuan=("Doanh thu thuần", "sum"),
    ).sort_values("SL_thuan", ascending=False)

    no_sale = color_summary[color_summary["SL_thuan"] == 0].copy()
    weak_1 = color_summary[color_summary["SL_thuan"] == 1].copy()
    weak_2 = color_summary[color_summary["SL_thuan"] == 2].copy()
    weak_3 = color_summary[color_summary["SL_thuan"] == 3].copy()
    returned = color_summary[color_summary["SL_thuan"] < 0].copy()

    color_summary["Phân loại"] = color_summary["SL_thuan"].apply(classify_sku)
    color_summary["Đề xuất"] = color_summary["SL_thuan"].apply(action_recommend)

    return work, group_summary, model_summary, color_summary, size_summary, no_sale, weak_1, weak_2, weak_3, returned

def build_rule_insights(summary, group_summary, model_summary, color_summary, no_sale, returned):
    insights = []
    if not group_summary.empty:
        top_group = group_summary.iloc[0]
        total_qty = group_summary["SL_thuan"].sum()
        share = (top_group["SL_thuan"] / total_qty) if total_qty else 0
        insights.append(f"Nhóm dẫn đầu hiện tại là **{top_group['Group']}**, bán thuần **{int(top_group['SL_thuan'])} đôi**, chiếm khoảng **{share:.1%}** tổng lượng bán.")
    if not model_summary.empty:
        top_model = model_summary.iloc[0]
        insights.append(f"Mã mẫu bán mạnh nhất là **{top_model['Mã mẫu']}**, bán thuần **{int(top_model['SL_thuan'])} đôi**. Nên ưu tiên giữ hàng và tăng hiển thị.")
    if len(no_sale) > 0:
        insights.append(f"Có **{len(no_sale)} mã+màu** chưa bán được đôi nào. Nên gom thành danh sách xử lý riêng bằng combo, voucher hoặc flash sale.")
    if len(returned) > 0:
        returned_qty = int(returned['SL_thuan'].sum())
        insights.append(f"Có **{len(returned)} mã+màu** đang âm thuần ({returned_qty} đôi). Nhóm này cần kiểm tra trả hàng, đổi size hoặc nhập về lại.")
    if summary["return_rate_qty"] >= 0.1:
        insights.append(f"Tỷ lệ trả hàng theo số lượng đang ở mức **{summary['return_rate_qty']:.1%}**, khá cao. Nên soi kỹ nhóm hàng có SL trả lớn.")
    elif summary["return_rate_qty"] > 0:
        insights.append(f"Tỷ lệ trả hàng theo số lượng ở mức **{summary['return_rate_qty']:.1%}**. Cần theo dõi nhưng chưa quá xấu.")
    return insights[:6]

def to_excel_export(summary_df, group_summary, model_summary, color_summary, size_summary, no_sale, weak_1, weak_2, weak_3, returned):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="DuLieuGoc", index=False)
        group_summary.to_excel(writer, sheet_name="NhomBanChay", index=False)
        model_summary.to_excel(writer, sheet_name="MaMauBanChay", index=False)
        color_summary.to_excel(writer, sheet_name="MauBanChay", index=False)
        size_summary.to_excel(writer, sheet_name="SizeBanChay", index=False)
        no_sale.to_excel(writer, sheet_name="KhongBan", index=False)
        weak_1.to_excel(writer, sheet_name="Ban1Doi", index=False)
        weak_2.to_excel(writer, sheet_name="Ban2Doi", index=False)
        weak_3.to_excel(writer, sheet_name="Ban3Doi", index=False)
        returned.to_excel(writer, sheet_name="AmTraHang", index=False)
    output.seek(0)
    return output

st.title("Merly Sales Dashboard Pro")
st.caption("Dashboard thông minh cho file doanh số theo hàng hóa: tốc độ bán, nhóm bán chạy, mã + màu mạnh, size hot, nhóm bán chậm và gợi ý hành động.")

st.markdown('<div class="note-box">', unsafe_allow_html=True)
st.write("""
**Cách dùng**
1. Tải file **Doanh số theo hàng hóa** từ hệ thống.
2. Nếu có file mapping thật cho nhóm sản phẩm, tải thêm file đó với 2 cột: **Mã mẫu** và **Group**.
3. Dashboard sẽ tự tính:
   - doanh thu thuần
   - số lượng bán thuần
   - tốc độ bán/ngày
   - nhóm đang bán chạy
   - danh sách bán chậm / không bán / trả hàng
""")
st.markdown('</div>', unsafe_allow_html=True)

col1, col2 = st.columns([2, 1])
with col1:
    sales_file = st.file_uploader("Tải file doanh số theo hàng hóa", type=["xlsx"], key="sales_file")
with col2:
    group_map_file = st.file_uploader("Tải file mapping Group (tuỳ chọn)", type=["xlsx", "csv"], key="group_map")

if sales_file:
    df, summary = load_sales_file(sales_file)

    group_map_df = None
    if group_map_file:
        if group_map_file.name.lower().endswith(".csv"):
            group_map_df = pd.read_csv(group_map_file)
        else:
            group_map_df = pd.read_excel(group_map_file)

    work, group_summary, model_summary, color_summary, size_summary, no_sale, weak_1, weak_2, weak_3, returned = build_dashboard_tables(df, group_map_df)
    insights = build_rule_insights(summary, group_summary, model_summary, color_summary, no_sale, returned)

    if summary["start_date"] is not None and summary["end_date"] is not None:
        st.subheader(f"Kỳ báo cáo: {summary['start_date'].strftime('%d/%m/%Y')} → {summary['end_date'].strftime('%d/%m/%Y')}")

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Doanh thu thuần", f"{summary['net_sales']:,.0f}")
    k2.metric("Số lượng thuần", f"{int(summary['net_qty']):,}")
    k3.metric("Bán/ngày", f"{summary['qty_per_day']:.1f}")
    k4.metric("Doanh thu/ngày", f"{summary['sales_per_day']:,.0f}")

    k5, k6, k7, k8 = st.columns(4)
    k5.metric("SL bán", f"{int(summary['gross_qty']):,}")
    k6.metric("SL trả", f"{int(summary['return_qty']):,}")
    k7.metric("Tỷ lệ trả hàng", f"{summary['return_rate_qty']:.1%}")
    k8.metric("SKU có bán", f"{int((work['SL thuần'] > 0).sum()):,}")

    st.subheader("Insight nhanh")
    st.markdown('<div class="ai-box">', unsafe_allow_html=True)
    if insights:
        for item in insights:
            st.write(f"- {item}")
    else:
        st.caption("Chưa đủ dữ liệu để rút insight.")
    st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Nhóm sản phẩm đang bán chạy")
        if not group_summary.empty:
            st.bar_chart(group_summary.set_index("Group")[["SL_thuan"]])
        else:
            st.info("Chưa có dữ liệu nhóm.")
    with c2:
        st.subheader("Size đang bán chạy")
        if not size_summary.empty:
            size_chart = size_summary.copy()
            size_chart["Size"] = size_chart["Size"].replace("", "Không rõ")
            st.bar_chart(size_chart.set_index("Size")[["SL_thuan"]])
        else:
            st.info("Chưa có dữ liệu size.")

    st.subheader("Top mã mẫu bán mạnh")
    st.dataframe(model_summary.head(30), use_container_width=True, height=300)

    st.subheader("Top mã + màu bán mạnh")
    st.dataframe(color_summary.head(50), use_container_width=True, height=320)

    s1, s2 = st.columns(2)
    with s1:
        st.write("Không bán được")
        st.dataframe(no_sale, use_container_width=True, height=220)
        st.write("Bán 1 đôi")
        st.dataframe(weak_1, use_container_width=True, height=220)
    with s2:
        st.write("Bán 2 đôi")
        st.dataframe(weak_2, use_container_width=True, height=220)
        st.write("Bán 3 đôi")
        st.dataframe(weak_3, use_container_width=True, height=220)

    st.subheader("Nhóm âm / trả hàng / về lại")
    st.dataframe(returned, use_container_width=True, height=240)

    st.subheader("Dữ liệu chi tiết đã xử lý")
    detail_cols = ["Mã hàng hóa", "Tên hàng hóa", "Group", "Mã mẫu", "Màu", "Size", "SL bán", "SL trả", "SL thuần", "Tiền bán hàng", "Tiền trả hàng", "Doanh thu thuần", "Phân loại", "Đề xuất"]
    st.dataframe(work[detail_cols], use_container_width=True, height=360)

    export_file = to_excel_export(work, group_summary, model_summary, color_summary, size_summary, no_sale, weak_1, weak_2, weak_3, returned)
    st.download_button(
        "Tải file dashboard phân tích",
        data=export_file,
        file_name="merly_sales_dashboard_pro.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Hãy tải file doanh số theo hàng hóa lên để bắt đầu.")
