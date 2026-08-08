import pandas as pd
import streamlit as st
import re

# 1. Cấu hình giao diện Web
st.set_page_config(page_title="Tư Vấn Lốp Xe Cần Thơ", layout="centered")

st.title("🚗 TRỢ LÝ TƯ VẤN LỐP XE Ô TÔ")
st.write("Tra cứu báo giá trọn gói & gợi ý phân khúc thông minh")

# Hàm tự động chuyển đổi định dạng tiền thành số
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

# 2. Đọc dữ liệu từ file Excel
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
    st.error(f"Lỗi đọc file Excel 'danh_muc_lop.xlsx': {e}")
    st.stop()

# 3. Khung tìm kiếm trên Web
st.subheader("🔍 Nhập thông tin tra cứu")
col1, col2 = st.columns([2, 1])

with col1:
    size_nhap = st.text_input("Mã Size lốp (Ví dụ: 205/55R16, 185/65R15...):", value="205/55R16")
with col2:
    nhu_cau = st.selectbox("Nhu cầu vận hành:", ["Tất cả nhu cầu", "Chạy dịch vụ / Taxi", "Đi gia đình (Cần êm)"])

# 4. Hiển thị kết quả tra cứu
if st.button("XEM BÁO GIÁ & GỢI Ý TRỌN GÓI", type="primary", key="btn_tra_cuu"):
    size_chuan = size_nhap.strip().upper()
    ket_qua = df[df['Size_Chuan'] == size_chuan]
    
    if ket_qua.empty:
        st.warning(f"Rất tiếc, mã size '{size_chuan}' hiện chưa có trong danh mục!")
    else:
        st.success(f"✅ Tìm thấy {len(ket_qua)} lựa chọn cho mã size: {size_chuan}")
        
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
            
            # Logic AI tư vấn tự động theo thương hiệu & giá tiền
            str_th = thuong_hieu.lower()
            if any(h in str_th for h in ['michelin', 'bridgestone', 'continental', 'goodyear']):
                tu_van = "🌟 **Dòng Cao Cấp:** Khả năng cách âm siêu vượt trội, bám đường mưa tuyệt đối, cực kỳ êm ái cho xe gia đình."
                phan_khuc = "Gia đình / Cao cấp"
            elif any(h in str_th for h in ['hankook', 'kumho', 'yokohama', 'dunlop', 'toyota']):
                tu_van = "⚖️ **Dòng Cân Bằng:** Độ êm tốt, gai lâu mòn, chi phí hợp lý cho cả xe gia đình và xe chạy dịch vụ cao cấp."
                phan_khuc = "Tầm trung"
            else:
                tu_van = "💰 **Dòng Tối Ưu Chi Phí:** Gai cực bền, chịu tải tốt, giá thành rẻ giúp tối ưu vốn nhanh cho xe chạy dịch vụ / Grab."
                phan_khuc = "Dịch vụ / Tiết kiệm"

            # Lọc theo nhu cầu người dùng chọn
            if nhu_cau == "Chạy dịch vụ / Taxi" and phan_khuc == "Gia đình / Cao cấp":
                continue
            if nhu_cau == "Đi gia đình (Cần êm)" and phan_khuc == "Dịch vụ / Tiết kiệm":
                continue

            # Hiển thị thông tin
            st.markdown(f"### 🔹 {thuong_hieu if thuong_hieu else 'Lốp xe'} {ten_sp}")
            if dong_xe:
                st.markdown(f"- 🚗 **Xe tương thích:** {dong_xe}")
                
            if val_gia_lop > 0:
                st.markdown(f"- 🛞 **Giá lốp nguyên bản:** `{val_gia_lop:,.0f} VNĐ/quả`")
            if val_gia_keo > 0:
                st.markdown(f"- 🛡️ **Giá tráng keo chống đinh:** `{val_gia_keo:,.0f} VNĐ/quả`")
                
            if tong_tien > 0:
                st.markdown(f"- 💰 **TỔNG CHI PHÍ TRỌN GÓI:** **`{tong_tien:,.0f} VNĐ/quả`**")
            
            st.info(tu_van)
            st.markdown("---")