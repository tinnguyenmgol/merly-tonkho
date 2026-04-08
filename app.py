
import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Merly - Tồn kho nội bộ",
    page_icon="📦",
    layout="wide"
)

# ---------- Theme helpers ----------
BRAND = {
    "bg": "#F8F5F2",
    "card": "#FFFFFF",
    "accent": "#B89D8C",
    "accent_dark": "#9D8475",
    "text": "#5C4B42",
    "soft": "#EFE7E1",
    "danger": "#C96F6F",
    "warning": "#D7A65A",
    "success": "#8AA17D",
}

LOGO_PATH = Path(__file__).parent / "Logo Merly.jpg"

st.markdown(f"""
<style>
    .stApp {{
        background: linear-gradient(180deg, {BRAND["bg"]} 0%, #ffffff 100%);
        color: {BRAND["text"]};
    }}
    .hero {{
        background: linear-gradient(135deg, rgba(184,157,140,.18), rgba(239,231,225,.55));
        border: 1px solid rgba(184,157,140,.25);
        border-radius: 20px;
        padding: 22px 24px;
        margin-bottom: 18px;
    }}
    .metric-card {{
        background: {BRAND["card"]};
        border: 1px solid rgba(184,157,140,.2);
        border-radius: 18px;
        padding: 14px 16px;
        box-shadow: 0 8px 30px rgba(120, 91, 73, 0.06);
    }}
    .section-title {{
        font-size: 1.05rem;
        font-weight: 700;
        color: {BRAND["text"]};
        margin: .25rem 0 .75rem 0;
    }}
    div[data-testid="stFileUploader"] {{
        background: rgba(255,255,255,.75);
        border-radius: 18px;
        padding: 8px 12px 2px 12px;
        border: 1px dashed rgba(184,157,140,.45);
    }}
    .small-note {{
        color: #7A675B;
        font-size: .92rem;
    }}
</style>
""", unsafe_allow_html=True)

# ---------- parsing ----------
SIZE_PATTERN = re.compile(r"^\d{2}$")

def split_product_name(name: str):
    text = str(name).strip()
    if not text:
        return "", "", ""

    parts = text.split()
    if len(parts) < 2:
        return "", "", ""

    ma_sp = parts[0]
    size_val = ""
    color_parts = []

    for p in parts[1:]:
        if SIZE_PATTERN.match(p):
            size_val = p
        else:
            color_parts.append(p)

    mau = " ".join(color_parts).strip()
    return ma_sp, size_val, mau


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = list(df.columns)

    def find_col(candidates, fallback_idx=None):
        for c in cols:
            if c in candidates:
                return c
        if fallback_idx is not None and fallback_idx < len(cols):
            return cols[fallback_idx]
        return None

    ten_col = find_col(["Tên hàng hóa"], 0)
    ton_col = find_col(["Số lượng"], 2)
    gia_col = find_col(["Giá bán"], 4)
    group_col = find_col(["Nhóm hàng"], 7)
    sku_col = find_col(["Mã hàng hóa"], 1)

    if ten_col is None:
        raise ValueError("Không tìm thấy cột Tên hàng hóa.")

    data = df.copy()
    parsed = data[ten_col].apply(split_product_name)
    data["Ma SP"] = parsed.apply(lambda x: x[0])
    data["Size"] = parsed.apply(lambda x: x[1])
    data["Mau"] = parsed.apply(lambda x: x[2])

    data["Gia Ban"] = data[gia_col] if gia_col else None
    data["Ton kho"] = pd.to_numeric(data[ton_col], errors="coerce").fillna(0) if ton_col else 0
    data["Group"] = data[group_col] if group_col else ""
    data["Ma hang hoa"] = data[sku_col] if sku_col else ""

    clean = data[["Ma SP", "Size", "Mau", "Gia Ban", "Ton kho", "Group", "Ma hang hoa"]].copy()
    clean["Size"] = clean["Size"].astype(str).str.strip()
    clean["Ma SP"] = clean["Ma SP"].astype(str).str.strip()
    clean["Mau"] = clean["Mau"].astype(str).str.strip()

    # Bỏ dòng tổng hợp / dòng không phải biến thể
    clean = clean[
        (clean["Ma SP"] != "") &
        (clean["Size"].str.match(r"^\d{2}$")) &
        (clean["Mau"] != "")
    ].copy()

    clean["Size"] = clean["Size"].astype(int)
    return clean


def make_pivot(clean: pd.DataFrame) -> pd.DataFrame:
    pivot = pd.pivot_table(
        clean,
        index=["Ma SP", "Mau"],
        columns="Size",
        values="Ton kho",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="Grand Total"
    )
    return pivot


def make_summary_by_ma(clean: pd.DataFrame) -> pd.DataFrame:
    return (
        clean.groupby("Ma SP", as_index=False)["Ton kho"]
        .sum()
        .sort_values("Ton kho", ascending=False)
    )


def build_excel(processed_df: pd.DataFrame, pivot_df: pd.DataFrame, low_df: pd.DataFrame, high_df: pd.DataFrame):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        processed_df.to_excel(writer, index=False, sheet_name="ThongKe")
        pivot_df.reset_index().to_excel(writer, index=False, sheet_name="Pivot")
        low_df.to_excel(writer, index=False, sheet_name="TonThap")
        high_df.to_excel(writer, index=False, sheet_name="TonCao")
    output.seek(0)
    return output


# ---------- UI ----------
left, right = st.columns([1, 5])
with left:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=150)
with right:
    st.markdown("""
    <div class="hero">
        <div style="font-size:1.6rem;font-weight:800;margin-bottom:4px;">Merly - Công cụ tồn kho nội bộ</div>
        <div class="small-note">
            Upload file Excel xuất từ hệ thống bán hàng để tự động tách <b>Mã SP - Size - Màu - Giá bán - Tồn kho - Group</b>,
            đồng thời tạo bảng tổng hợp phục vụ đặt hàng thêm và lên chương trình khuyến mãi.
        </div>
    </div>
    """, unsafe_allow_html=True)

with st.expander("Định dạng file đầu vào đang hỗ trợ", expanded=False):
    st.write("""
    Ứng dụng ưu tiên đọc các cột chuẩn từ file xuất hệ thống:
    - **Tên hàng hóa**
    - **Mã hàng hóa**
    - **Số lượng**
    - **Giá bán**
    - **Nhóm hàng**

    Nếu tên cột đúng như trên, app sẽ tự nhận diện. Nếu không, app sẽ thử fallback theo vị trí cột giống file hiện tại của Merly.
    """)

uploaded = st.file_uploader("Tải file Excel (.xlsx)", type=["xlsx"])

if uploaded is not None:
    try:
        raw_df = pd.read_excel(uploaded)
        clean_df = standardize_columns(raw_df)

        if clean_df.empty:
            st.warning("Không tìm thấy dòng biến thể hợp lệ sau khi tách. Kiểm tra lại file nguồn.")
            st.stop()

        pivot_df = make_pivot(clean_df)
        by_ma_df = make_summary_by_ma(clean_df)

        low_df = clean_df[clean_df["Ton kho"] <= 2].sort_values(["Ton kho", "Ma SP", "Size"], ascending=[True, True, True])
        high_df = clean_df[clean_df["Ton kho"] >= 10].sort_values(["Ton kho"], ascending=False)

        total_variants = int(len(clean_df))
        total_stock = int(clean_df["Ton kho"].sum())
        low_count = int((clean_df["Ton kho"] <= 2).sum())
        high_count = int((clean_df["Ton kho"] >= 10).sum())

        c1, c2, c3, c4 = st.columns(4)
        metrics = [
            ("Biến thể hợp lệ", total_variants),
            ("Tổng tồn kho", total_stock),
            ("Biến thể tồn thấp ≤ 2", low_count),
            ("Biến thể tồn cao ≥ 10", high_count),
        ]
        for col, (label, value) in zip([c1, c2, c3, c4], metrics):
            with col:
                st.markdown(f'<div class="metric-card"><div class="small-note">{label}</div><div style="font-size:1.8rem;font-weight:800;">{value}</div></div>', unsafe_allow_html=True)

        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "Dữ liệu đã tách", "Pivot tồn kho", "Tồn thấp", "Tồn cao", "Tổng theo mã"
        ])

        with tab1:
            st.markdown('<div class="section-title">Bảng chuẩn hóa để lọc / pivot</div>', unsafe_allow_html=True)
            st.dataframe(clean_df, use_container_width=True, height=520)

        with tab2:
            st.markdown('<div class="section-title">Pivot theo Mã SP - Màu - Size</div>', unsafe_allow_html=True)
            st.dataframe(pivot_df, use_container_width=True, height=620)

        with tab3:
            st.markdown('<div class="section-title">Biến thể tồn thấp - ưu tiên nhập thêm</div>', unsafe_allow_html=True)
            st.dataframe(low_df, use_container_width=True, height=520)

        with tab4:
            st.markdown('<div class="section-title">Biến thể tồn cao - cân nhắc đẩy khuyến mãi</div>', unsafe_allow_html=True)
            st.dataframe(high_df, use_container_width=True, height=520)

        with tab5:
            st.markdown('<div class="section-title">Tổng tồn kho theo mã sản phẩm</div>', unsafe_allow_html=True)
            st.dataframe(by_ma_df, use_container_width=True, height=520)

        excel_data = build_excel(clean_df, pivot_df, low_df, high_df)

        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "⬇️ Tải file Excel kết quả",
                data=excel_data,
                file_name="merly_tonkho_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with d2:
            st.download_button(
                "⬇️ Tải CSV dữ liệu đã tách",
                data=clean_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="merly_tonkho_clean.csv",
                mime="text/csv",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"Lỗi khi xử lý file: {e}")
else:
    st.info("Hãy upload file Excel xuất từ hệ thống để bắt đầu.")
