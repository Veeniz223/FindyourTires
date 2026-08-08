import pandas as pd
import streamlit as st
import re
import urllib.parse

# ==========================================
# THÔNG TIN CỬA HÀNG & NHÂN VIÊN KINH DOANH
# ==========================================
TEN_CUAHANG = "THẾ GIỚI LỐP XE Ô TÔ CẦN THƠ"
TEN_NHANVIEN = "Dương Hoàng Vinh"
CHUC_DANH = "Nhân viên kinh doanh"
HOTLINE = "0937971684"               # Ông sửa lại SĐT của ông tại đây
ZALO_PHONE = "0937971684"            # Ông sửa lại SĐT Zalo tại đây
ZALO_LINK = f"https://zalo.me/{ZALO_PHONE}"
DIA_CHI = "176, Phạm Hùng, Cái Răng, TP Cần Thơ" 
# ==========================================

# 1. Cấu hình trang Web
st.set_page_config(
    page_title=f"Tư Vấn Lốp Xe - {TEN_CUAHANG}",
    page_icon="🛞",
    layout="centered"
)

# 2. Tùy chỉnh Giao diện CSS
st.markdown("""
    <style>
    .stApp { background-color: #0e1117; }
    
    .main-title {
        text-align: center; color: #FF4B4B; font-weight: 800;
        font-size: 2rem; margin-bottom: 5px; text-transform: uppercase;
    }
    .sub-title { text-align: center; color: #9CA3AF; font-size: 0.95rem; margin-bottom: 15px; }

    /* Khung Liên hệ */
    .contact-box {
        background: linear-gradient(135deg, #1e2638, #111827);
        border: 1px solid #FF4B4B; border-radius: 12px; padding: 18px;
        text-align: center; margin-bottom: 25px; box-shadow: 0 4px 15px rgba(255, 75, 75, 0.15);
    }
    .store-name { color: #FACC15; font-size: 1.2rem; font-weight: 800; text-transform: uppercase; margin-bottom: 8px; }
    .staff-info { color: #E5E7EB; font-size: 1rem; margin-bottom: 8px; }
    .contact-info { color: #9CA3AF; font-size: 0.9rem; margin-bottom: 6px; }
    .btn-contact-container { display: flex; justify-content: center; gap: 12px; margin-top: 12px; }
    .btn-call { background-color: #EF4444; color: white !important; padding: 8px 18px; border-radius: 20px; text-decoration: none; font-weight: bold; font-size: 0.88rem; }
    .btn-zalo { background-color: #0068FF; color: white !important; padding: 8px 18px; border-radius: 20px; text-decoration: none; font-weight: bold; font-size: 0.88rem; }

    /* Thẻ Sản Phẩm */
    .tire-card {
        background: #1e2638; border-radius: 12px; padding: 20px;
        margin-bottom: 15px; border: 1px solid #2e3b52; box-shadow: 0 4px 12px rgba(0,0,0,0.3);
    }
    .tire-name { color: #FFFFFF; font-size: 1.3rem; font-weight: 700; margin-bottom: 10px; }

    /* Badges */
    .badge-cao-cap { background-color: #f59e0b; color: #000; padding: 3px 10px; border-radius: 20px; font-size: 0.8rem; font-weight: bold; }
    .badge-tam-trung { background-color: #3b82f6; color: #fff; padding: 3px 10px; border-radius: 20px; font-size: 0.8rem; font-weight: bold; }
    .badge-tiet-kiem { background-color: #10b981; color: #fff; padding: 3px 10px; border-radius: 20px; font-size: 0.8rem; font-weight: bold; }

    /* Khung Giá */
    .price-box { background-color: #064e3b; border: 1px solid #10b981; border-radius: 8px; padding: 10px; text-align: center; margin-top: 12px; }
    .price-title { color: #a7f3d0; font-size: 0.85rem; }
    .price-amount { color: #34d399; font-size: 1.4rem; font-weight: 800; }
    .detail-price { color: #d1d5db; font-size: 0.9rem; margin: 4px 0; }
    
    .btn-zalo-quote {
        display: inline-block; background-color: #0068FF; color: white !important;
        padding: 10px 18px; border-radius: 8px; font-weight: bold;
        text-decoration: none; font-size: 0.9rem; margin-top: 10px; text-align: center;
    }
    </style>
""", unsafe_allow_html=True)

# Hàm chuyển đổi số tiền
def chuyen_thanh_so(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    s = str(val)
    numbers_only = re.sub(r'[^\d]', '', s)
    return float(numbers_only) if numbers_only else 0.0

# Hàm làm sạch chuỗi tìm kiếm (Loại bỏ dấu gạch ngang, khoảng trắng)
def clean_text(text):
    if pd.isna(text):
        return ""
    return re.sub(r'[\s\-_.\/\\]+', '', str(text)).upper()

# 3. Đọc dữ liệu Excel
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

# Khung Thông Tin Cửa Hàng
st.markdown(f"""
<div class="contact-box">
    <div class="store-name">🏢 {TEN_CUAHANG}</div>
    <div class="staff-info">👨‍💼 <b>{CHUC_DANH}:</b> <span style="color: #6EE7B7; font-weight: bold;">{TEN_NHANVIEN}</span></div>
    <div class="contact-info">📞 Hotline/Zalo: <strong style="color: #FACC15; font-size: 1.05rem;">{HOTLINE}</strong></div>
    <div class="contact-info">📍 Địa chỉ: <strong>{DIA_CHI}</strong></div>
    <div class="btn-contact-container">
        <a href="tel:{HOTLINE}" class="btn-call">📞 BẤM GỌI NGAY</a>
        <a href="{ZALO_LINK}" target="_blank" class="btn-zalo">💬 NHẮN ZALO</a>
    </div>
</div>
""", unsafe_allow_html=True)

# TẠO TABS CHỨC NĂNG
tab1, tab2 = st.tabs(["🔍 TRA CỨU BÁO GIÁ LỐP", "📅 ĐẶT LỊCH HẸN TẠI GARA"])

with tab1:
    col1, col2 = st.columns([2, 1])
    with col1:
        tim_kiem = st.text_input("🔍 Mã Size lốp HOẶC Tên Dòng Xe:", value="205/55R16", placeholder="Ví dụ: 205/55R16, CX5, Xpander...")
    with col2:
        nhu_cau = st.selectbox("🎯 Nhu cầu sử dụng:", ["Tất cả nhu cầu", "Chạy dịch vụ / Taxi", "Đi gia đình (Cần êm)"])

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("🚀 TÌM LỐP & XEM BÁO GIÁ TRỌN GÓI", type="primary", use_container_width=True):
        raw_input = tim_kiem.strip()
        cleaned_input = clean_text(raw_input)
        
        words = [clean_text(w) for w in re.split(r'\s+', raw_input) if clean_text(w)]

        def match_row(row):
            row_cleaned_str = clean_text(" ".join(row.astype(str)))
            if cleaned_input in row_cleaned_str:
                return True
            if words and all(w in row_cleaned_str for w in words):
                return True
            return False

        ket_qua = df[df.apply(match_row, axis=1)]
        
        if ket_qua.empty:
            st.warning(f"Rất tiếc, không tìm thấy kết quả cho từ khóa '{raw_input}'!")
        else:
            st.success(f"✅ Tìm thấy **{len(ket_qua)}** lựa chọn phù hợp cho từ khóa **'{raw_input}'**")
            
            for index, row in ket_qua.iterrows():
                thuong_hieu, ten_sp, dong_xe = "", "", ""
                val_gia_lop, val_gia_keo = 0.0, 0.0
                
                for col in df.columns:
                    col_lower = str(col).strip().lower()
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
                    badge_html = '<span class="badge-tiet-kiem">DỊCH VỤ / TỐI ƯU CHÍ PHÍ</span>'
                    phan_khuc = "Dịch vụ / Tiết kiệm"

                if nhu_cau == "Chạy dịch vụ / Taxi" and phan_khuc == "Gia đình / Cao cấp": continue
                if nhu_cau == "Đi gia đình (Cần êm)" and phan_khuc == "Dịch vụ / Tiết kiệm": continue

                text_zalo = f"Chào bạn, {TEN_CUAHANG} xin gửi báo giá lốp {thuong_hieu} {ten_sp}:\n- Giá lốp: {val_gia_lop:,.0f}đ/quả\n- Tráng keo chống đinh: {val_gia_keo:,.0f}đ/quả\n👉 TỔNG TRỌN GÓI: {tong_tien:,.0f}đ/quả.\nTư vấn trực tiếp: {TEN_NHANVIEN} ({HOTLINE})"
                zalo_quote_url = f"https://zalo.me/{ZALO_PHONE}?text={urllib.parse.quote(text_zalo)}"

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
                st.markdown(f'<a href="{zalo_quote_url}" target="_blank" class="btn-zalo-quote">📲 GỬI BÁO GIÁ NÀY QUA ZALO CHO KHÁCH</a><br><br>', unsafe_allow_html=True)

# TAB 2: ĐẶT LỊCH HẸN
with tab2:
    st.subheader("📝 Đặt Lịch Hẹn Làm Lốp / Tráng Keo Tận Nơi")
    st.write("Khách hàng điền thông tin bên dưới để ưu tiên xếp lịch phục vụ nhanh nhất:")
    
    with st.form("form_dat_lich"):
        ten_kh = st.text_input("Họ và tên khách hàng:")
        sdt_kh = st.text_input("Số điện thoại liên hệ:")
        dongxe_kh = st.text_input("Dòng xe đang đi (Ví dụ: Xpander, Camry, CX5...):")
        dich_vu = st.multiselect("Dịch vụ cần làm:", ["Thay lốp mới", "Tráng keo chống đinh", "Cân bằng động / Cân chỉnh thước lái", "Vá lốp lưu động"])
        ngay_hen = st.date_input("Ngày dự định ghé gara:")
        ghi_chu = st.text_area("Ghi chú thêm:")
        
        btn_submit = st.form_submit_button("🚀 ĐẶT LỊCH HẸN NGAY", type="primary", use_container_width=True)
        
        if btn_submit:
            if not ten_kh or not sdt_kh:
                st.error("Vui lòng điền Họ tên và Số điện thoại!")
            else:
                msg = f"ĐẶT LỊCH HẸN TẠI GARA:\n- Khách hàng: {ten_kh}\n- SĐT: {sdt_kh}\n- Loại xe: {dongxe_kh}\n- Dịch vụ: {', '.join(dich_vu)}\n- Ngày hẹn: {ngay_hen}\n- Ghi chú: {ghi_chu}"
                zalo_send_url = f"https://zalo.me/{ZALO_PHONE}?text={urllib.parse.quote(msg)}"
                st.success("✅ Thông tin đã sẵn sàng! Bấm nút bên dưới để gửi lịch hẹn trực tiếp sang Zalo:")
                st.markdown(f'<a href="{zalo_send_url}" target="_blank" style="display:inline-block; background:#10B981; color:white; padding:10px 20px; border-radius:8px; font-weight:bold; text-decoration:none;">💬 BẤM ĐỂ GỬI ĐẶT LỊCH QUA ZALO</a>', unsafe_allow_html=True)