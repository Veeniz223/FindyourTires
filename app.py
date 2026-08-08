import pandas as pd
import streamlit as st
import re

# 1. Cấu hình trang Web
st.set_page_config(
    page_title="Tư Vấn Lốp Xe Ô TÔ",
    page_icon="🛞",
    layout="centered"
)

# 2. Tùy chỉnh Giao diện CSS Chuyên Nghiệp
st.markdown("""
    <style>
    /* Nền chính & Font chữ */
    .stApp {
        background-color: #0e1117;
    }
    
    /* Tiêu đề chính */
    .main-title {
        text-align: center;
        color: #FF4B4B;
        font-weight: 800;
        font-size: 2.2rem;
        margin-bottom: 5px;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .sub-title {
        text-align: center;
        color: #9CA3AF;
        font-size: 1rem;
        margin-bottom: 25px;
    }

    /* Thẻ Sản Phẩm (Card) */
    .tire-card {
        background: #1e2638;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 20px;
        border: 1px solid #2e3b52;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3);
        transition: transform 0.2s ease;
    }
    .tire-card:hover {
        border-color: #FF4B4B;
    }

    /* Tên Sản Phẩm */
    .tire-name {
        color: #FFFFFF;
        font-size: 1.3rem;
        font-weight: 700;
        margin-bottom: 10px;
    }

    /* Huy hiệu Phân Khúc */
    .badge-cao-cap {
        background-color: #f59e0b;
        color: #000000;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: bold;
    }
    .badge-tam-trung {
        background-color: #3b82f6;
        color: #ffffff;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: bold;
    }
    .badge-tiet-kiem {
        background-color: #10b981;
        color: #ffffff;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: bold;
    }

    /* Khung Giá */
    .price-box {
        background-color: #064e3b;
        border: 1px solid #10b981;
        border-radius: 8px;
        padding: 10px;
        text-align: center;
        margin-top: 12px;
    }
    .price-title {
        color: #a7f3d0;
        font-size: 0.85rem;
        margin-bottom: 2px;
    }
    .price-amount {
        color: #34d399;
        font-size: 1.4rem;
        font-weight: 800;
    }
    
    /* Chi tiết giá lẻ */
    .detail-price {
        color: #d1d5db;
        font-size: 0.9rem;
        margin: 4px 0;
    }
    </style>
""", unsafe_allow_html=True)

# Hàm chuyển đổi định dạng tiền thành số
def chuyen_thanh_so(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val)
    numbers_only = re.sub(r'[^\d]', '', s)
    if numbers_only:
        return float(numbers_only)
    return 0.0

# 3. Đọc dữ liệu từ file Excel
@st.cache_data
def load_data():
    file_path = "danh_muc_lop.xlsx"
    df = pd.read_excel(file_path, sheet_name=0)
    
    for col in df.columns:
        if 'size' in str(col).lower() or 'cỡ' in str(col).lower():
            df['Size_Chuan'] = df[col].astype(str).str.strip().str.upper()
            break
    if 'Size_Chuan' not in df.columns:
        df['Size_Chuan'] = df.iloc[:, 0].astype(str).str.strip().str.upper()
        
    return df

try:
    df = load_data()
except Exception as e:
    st.error(f"Lỗi đọc file Excel: {e}")
    st.stop()

# Header Giao diện
st.markdown('<div class="main-title">🛞 TRỢ LÝ TƯ VẤN LỐP XE Ô TÔ</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Tra cứu báo giá trọn gói & gợi ý phân khúc chuẩn xác 24/7</div>', unsafe_allow_html=True)

# Khung Tìm Kiếm
with st.container():
    col1, col2 = st.columns([2, 1])
    with col1:
        size_nhap = st.text_input("🔍 Mã Size lốp:", value="205/55R16", placeholder="Nhập ví dụ: 205/55R16")
    with col2:
        nhu_cau = st.selectbox("🎯 Nhu cầu sử dụng:", ["Tất cả nhu cầu", "Chạy dịch vụ / Taxi", "Đi gia đình (Cần êm)"])

st.markdown("<br>", unsafe_allow_html=True)

# Xử lý Kết Quả Tra Cứu
if st.button("🚀 TÌM LỐP & XEM BÁO GIÁ TRỌN GÓI", type="primary", use_container_width=True):
    size_chuan = size_nhap.strip().upper()
    ket_qua = df[df['Size_Chuan'] == size_chuan]
    
    if ket_qua.empty:
        st.warning(f"Rất tiếc, mã size '{size_chuan}' hiện chưa có trong danh mục!")
    else:
        st.success(f"✅ Tìm thấy **{len(ket_qua)}** lựa chọn phù hợp cho mã size **{size_chuan}**")
        
        for index, row in ket_qua.iterrows():
            thuong_hieu = ""
            ten_sp = ""
            dong_xe = ""
            val_gia_lop = 0.0
            val_gia_keo = 0.0
            
            for col in df.columns:
                col_name = str(col).strip()
                col_lower = col_name.lower()
                val = row[col]
                
                if 'thương hiệu' in col_lower or 'hãng' in col_lower or 'brand' in col_lower:
                    thuong_hieu = str(val) if pd.notna(val) else ""
                elif 'sản phẩm' in col_lower or 'gai' in col_lower or 'tên' in col_lower:
                    ten_sp = str(val) if pd.notna(val) else ""
                elif 'xe' in col_lower or 'dòng xe' in col_lower:
                    dong_xe = str(val) if pd.notna(val) else ""
                elif 'keo' in col_lower or 'đinh' in col_lower:
                    val_gia_keo = chuyen_thanh_so(val)
                elif 'giá' in col_lower or 'đơn giá' in col_lower or 'niêm yết' in col_lower:
                    if 'keo' not in col_lower and 'đinh' not in col_lower:
                        val_gia_lop = chuyen_thanh_so(val)

            tong_tien = val_gia_lop + val_gia_keo
            
            # Phân loại & Gắn badge
            str_th = thuong_hieu.lower()
            if any(h in str_th for h in ['michelin', 'bridgestone', 'continental', 'goodyear']):
                tu_van = "🌟 **Cao Cấp:** Cách âm siêu vượt trội, bám đường mưa tốt, cực kỳ êm ái cho xe gia đình."
                badge_html = '<span class="badge-cao-cap">GIA ĐÌNH / CAO CẤP</span>'
                phan_khuc = "Gia đình / Cao cấp"
            elif any(h in str_th for h in ['hankook', 'kumho', 'yokohama', 'dunlop', 'toyota']):
                tu_van = "⚖️ **Cân Bằng:** Độ êm tốt, gai lâu mòn, chi phí hợp lý cho cả xe gia đình & dịch vụ."
                badge_html = '<span class="badge-tam-trung">TẦM TRUNG / CÂN BẰNG</span>'
                phan_khuc = "Tầm trung"
            else:
                tu_van = "💰 **Tiết Kiệm:** Gai cực bền, chịu tải tốt, giá rẻ giúp tối ưu vốn nhanh cho xe dịch vụ/Grab."
                badge_html = '<span class="badge-tiet-kiem">DỊCH VỤ / TỐI ƯU CHÍPHÍ</span>'
                phan_khuc = "Dịch vụ / Tiết kiệm"

            # Lọc nhu cầu
            if nhu_cau == "Chạy dịch vụ / Taxi" and phan_khuc == "Gia đình / Cao cấp":
                continue
            if nhu_cau == "Đi gia đình (Cần êm)" and phan_khuc == "Dịch vụ / Tiết kiệm":
                continue

            # Render Giao diện Thẻ Sản Phẩm
            st.markdown(f"""
            <div class="tire-card">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div class="tire-name">🔹 {thuong_hieu if thuong_hieu else 'Lốp xe'} {ten_sp}</div>
                    {badge_html}
                </div>
                {"<p style='color: #9CA3AF; margin-bottom: 8px;'>🚗 <b>Xe tương thích:</b> " + dong_xe + "</p>" if dong_xe else ""}
                <div class="detail-price">🛞 <b>Giá lốp:</b> {val_gia_lop:,.0f} VNĐ/quả</div>
                <div class="detail-price">🛡️ <b>Giá tráng keo chống đinh:</b> {val_gia_keo:,.0f} VNĐ/quả</div>
                <div class="price-box">
                    <div class="price-title">TỔNG CHI PHÍ LĂN BÁNH TRỌN GÓI</div>
                    <div class="price-amount">{tong_tien:,.0f} VNĐ / quả</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            st.info(tu_van)