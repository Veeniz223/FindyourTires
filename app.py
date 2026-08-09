import pandas as pd
import streamlit as st
import re
import urllib.parse
import requests
import os
import streamlit.components.v1 as components
# ==========================================
# THÔNG TIN CỬA HÀNG & MÃ AI
# ==========================================
TEN_CUAHANG = "THẾ GIỚI LỐP XE Ô TÔ CẦN THƠ"
TEN_NHANVIEN = "Dương Hoàng Vinh"
CHUC_DANH = "Nhân viên kinh doanh"
HOTLINE = "0937971684"               # Nhập SĐT của ông
ZALO_PHONE = "0937971684"            # Nhập SĐT Zalo của ông
ZALO_LINK = f"https://zalo.me/{ZALO_PHONE}"
DIA_CHI = "176, Phạm Hùng, Cái Răng, TP Cần Thơ" 

# DÁN MÃ API CỦA ÔNG VÀO TRONG DẤU NGOẶC KÉP DƯỚI ĐÂY:
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "DAN_API_KEY_MOI_VAO_DAY").strip()
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
    .main-title { text-align: center; color: #FF4B4B; font-weight: 800; font-size: 2rem; margin-bottom: 5px; text-transform: uppercase; }
    .sub-title { text-align: center; color: #9CA3AF; font-size: 0.95rem; margin-bottom: 15px; }
    .contact-box { background: linear-gradient(135deg, #1e2638, #111827); border: 1px solid #FF4B4B; border-radius: 12px; padding: 18px; text-align: center; margin-bottom: 25px; box-shadow: 0 4px 15px rgba(255, 75, 75, 0.15); }
    .store-name { color: #FACC15; font-size: 1.2rem; font-weight: 800; text-transform: uppercase; margin-bottom: 8px; }
    .staff-info { color: #E5E7EB; font-size: 1rem; margin-bottom: 8px; }
    .contact-info { color: #9CA3AF; font-size: 0.9rem; margin-bottom: 6px; }
    .btn-contact-container { display: flex; justify-content: center; gap: 12px; margin-top: 12px; }
    .btn-call { background-color: #EF4444; color: white !important; padding: 8px 18px; border-radius: 20px; text-decoration: none; font-weight: bold; font-size: 0.88rem; }
    .btn-zalo { background-color: #0068FF; color: white !important; padding: 8px 18px; border-radius: 20px; text-decoration: none; font-weight: bold; font-size: 0.88rem; }
    .tire-card { background: #1e2638; border-radius: 12px; padding: 20px; margin-bottom: 15px; border: 1px solid #2e3b52; box-shadow: 0 4px 12px rgba(0,0,0,0.3); }
    .tire-name { color: #FFFFFF; font-size: 1.25rem; font-weight: 700; margin-bottom: 6px; }
    .car-version-box { background-color: #111827; border-left: 4px solid #38BDF8; padding: 8px 12px; border-radius: 4px; margin-bottom: 12px; color: #38BDF8; font-weight: 600; font-size: 0.95rem; }
    .badge-cao-cap { background-color: #f59e0b; color: #000; padding: 3px 10px; border-radius: 20px; font-size: 0.8rem; font-weight: bold; }
    .badge-tam-trung { background-color: #3b82f6; color: #fff; padding: 3px 10px; border-radius: 20px; font-size: 0.8rem; font-weight: bold; }
    .badge-tiet-kiem { background-color: #10b981; color: #fff; padding: 3px 10px; border-radius: 20px; font-size: 0.8rem; font-weight: bold; }
    .price-box { background-color: #064e3b; border: 1px solid #10b981; border-radius: 8px; padding: 10px; text-align: center; margin-top: 12px; }
    .price-title { color: #a7f3d0; font-size: 0.85rem; }
    .price-amount { color: #34d399; font-size: 1.4rem; font-weight: 800; }
    .detail-price { color: #d1d5db; font-size: 0.9rem; margin: 4px 0; }
    .btn-zalo-quote { display: inline-block; background-color: #0068FF; color: white !important; padding: 10px 18px; border-radius: 8px; font-weight: bold; text-decoration: none; font-size: 0.9rem; margin-top: 8px; text-align: center; }
    </style>
""", unsafe_allow_html=True)

# Hàm chuyển đổi số tiền & làm sạch chuỗi
def chuyen_thanh_so(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    numbers_only = re.sub(r'[^\d]', '', str(val))
    return float(numbers_only) if numbers_only else 0.0

def clean_text(text):
    if pd.isna(text): return ""
    return re.sub(r'[\s\-_.\/\\]+', '', str(text)).upper()

# 3. Đọc dữ liệu Excel
@st.cache_data
def load_data():
    df = pd.read_excel("danh_muc_lop.xlsx", sheet_name=0)
    for col in df.columns:
        if 'size' in str(col).lower() or 'cỡ' in str(col).lower():
            df['Size_Chuan'] = df[col].astype(str).str.strip().str.upper()
            break
    if 'Size_Chuan' not in df.columns: df['Size_Chuan'] = df.iloc[:, 0].astype(str).str.strip().str.upper()
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
tab1, tab2, tab3 = st.tabs(["🔍 TRA CỨU", "🤖 CHAT VỚI AI 24/7", "📅 ĐẶT LỊCH HẸN"])

with tab1:
    col1, col2 = st.columns([2, 1])
    with col1: tim_kiem = st.text_input("🔍 Nhập Mã Size lốp HOẶC Tên Dòng Xe:", value="CX5")
    with col2: nhu_cau = st.selectbox("🎯 Nhu cầu sử dụng:", ["Tất cả nhu cầu", "Chạy dịch vụ / Taxi", "Đi gia đình (Cần êm)"])
    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("🚀 TÌM LỐP & XEM BÁO GIÁ", type="primary", use_container_width=True):
        raw_input = tim_kiem.strip()
        cleaned_input = clean_text(raw_input)
        words = [clean_text(w) for w in re.split(r'\s+', raw_input) if clean_text(w)]

        def match_row(row):
            row_str = clean_text(" ".join(row.astype(str)))
            if cleaned_input in row_str: return True
            if words and all(w in row_str for w in words): return True
            return False

        ket_qua = df[df.apply(match_row, axis=1)]
        
        if ket_qua.empty: st.warning(f"Rất tiếc, không tìm thấy kết quả nào cho từ khóa '{raw_input}'!")
        else:
            st.success(f"✅ Tìm thấy **{len(ket_qua)}** lựa chọn phù hợp.")
            for index, row in ket_qua.iterrows():
                thuong_hieu, ten_sp, dong_xe, size_lop = "", "", "", ""
                val_gia_lop, val_gia_keo = 0.0, 0.0
                for col in df.columns:
                    col_lower = str(col).strip().lower()
                    val = row[col]
                    if 'size' in col_lower or 'cỡ' in col_lower: size_lop = str(val) if pd.notna(val) else ""
                    elif 'thương hiệu' in col_lower or 'hãng' in col_lower: thuong_hieu = str(val) if pd.notna(val) else ""
                    elif 'sản phẩm' in col_lower or 'gai' in col_lower: ten_sp = str(val) if pd.notna(val) else ""
                    elif 'xe' in col_lower or 'dòng xe' in col_lower: dong_xe = str(val) if pd.notna(val) else ""
                    elif 'keo' in col_lower or 'đinh' in col_lower: val_gia_keo = chuyen_thanh_so(val)
                    elif 'giá' in col_lower or 'đơn giá' in col_lower:
                        if 'keo' not in col_lower and 'đinh' not in col_lower: val_gia_lop = chuyen_thanh_so(val)

                tong_tien = val_gia_lop + val_gia_keo
                str_th = thuong_hieu.lower()
                if any(h in str_th for h in ['michelin', 'bridgestone', 'continental', 'goodyear']):
                    tu_van, badge_html, phan_khuc = "🌟 Cao Cấp: Siêu êm ái.", '<span class="badge-cao-cap">CAO CẤP</span>', "Gia đình"
                elif any(h in str_th for h in ['hankook', 'kumho', 'yokohama', 'dunlop']):
                    tu_van, badge_html, phan_khuc = "⚖️ Cân Bằng: Chi phí hợp lý.", '<span class="badge-tam-trung">TẦM TRUNG</span>', "Tầm trung"
                else:
                    tu_van, badge_html, phan_khuc = "💰 Tiết Kiệm: Gai siêu bền.", '<span class="badge-tiet-kiem">TIẾT KIỆM</span>', "Dịch vụ"

                if nhu_cau == "Chạy dịch vụ / Taxi" and phan_khuc == "Gia đình": continue
                if nhu_cau == "Đi gia đình (Cần êm)" and phan_khuc == "Dịch vụ": continue

                zalo_txt = f"Chào bạn, tôi báo giá lốp {thuong_hieu} {ten_sp} (Size {size_lop}):\n- Lốp: {val_gia_lop:,.0f}đ\n- Tráng keo: {val_gia_keo:,.0f}đ\n👉 TỔNG: {tong_tien:,.0f}đ/quả.\nTư vấn: {TEN_NHANVIEN} ({HOTLINE})"
                
                st.markdown(f"""
                <div class="tire-card">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                        <div class="tire-name">🔹 {thuong_hieu} {ten_sp} (Size: <span style="color: #FACC15;">{size_lop}</span>)</div>
                        {badge_html}
                    </div>
                    {"<div class='car-version-box'>🚗 Dòng xe tương thích: " + dong_xe + "</div>" if dong_xe else ""}
                    <div class="detail-price">🛞 Giá lốp: {val_gia_lop:,.0f}đ &nbsp;|&nbsp; 🛡️ Giá keo: {val_gia_keo:,.0f}đ</div>
                    <div class="price-box"><div class="price-title">TỔNG TRỌN GÓI</div><div class="price-amount">{tong_tien:,.0f} VNĐ / quả</div></div>
                </div>
                """, unsafe_allow_html=True)
                st.info(tu_van)
                st.markdown(f'<a href="https://zalo.me/{ZALO_PHONE}?text={urllib.parse.quote(zalo_txt)}" target="_blank" class="btn-zalo-quote">📲 GỬI BÁO GIÁ QUA ZALO</a><br><br>', unsafe_allow_html=True)

# TAB 2: CHATBOT AI TƯ VẤN 24/7
with tab2:
    st.subheader("🤖 Trợ Lý AI Tư Vấn Lốp Xe 24/7")
    st.write("Hỏi tôi về các loại lốp hoặc quy trình tráng keo nhé!")

    if "messages" not in st.session_state:
        st.session_state.messages = [{"role": "assistant", "content": f"Xin chào! Tôi là Trợ lý AI của **{TEN_NHANVIEN}**. Tôi có thể giúp gì cho bạn?"}]

    for msg in st.session_state.messages: st.chat_message(msg["role"]).write(msg["content"])

    if user_prompt := st.chat_input("Nhập câu hỏi của bạn tại đây...", key="ai_chat"):
        st.session_state.messages.append({"role": "user", "content": user_prompt})
        st.chat_message("user").write(user_prompt)

        if not GEMINI_API_KEY or GEMINI_API_KEY == "DÁN_MÃ_API_CỦA_ÔNG_VÀO_ĐÂY":
            response_text = "⚠️ Chưa có GEMINI_API_KEY. Hãy đặt API key mới vào biến GEMINI_API_KEY trong app.py hoặc biến môi trường GEMINI_API_KEY."
        else:
            try:
                system_instruction = (
                    f"Bạn là tư vấn viên kỹ thuật lốp xe tên {TEN_NHANVIEN} "
                    f"làm việc tại {TEN_CUAHANG}. Lịch sự, am hiểu lốp xe. "
                    "Không tự bịa giá, size hoặc mã gai. Nếu thiếu dữ liệu thì nói rõ."
                )

                # Gọi Gemini REST API trực tiếp, không dùng SDK cũ.
                # Ưu tiên model Flash ổn định; nếu tài khoản không có model ưu tiên,
                # tự kiểm tra danh sách model có quyền generateContent.
                api_base = "https://generativelanguage.googleapis.com/v1beta"
                headers = {
                    "Content-Type": "application/json",
                    "x-goog-api-key": GEMINI_API_KEY
                }

                preferred_models = ["gemini-2.5-flash", "gemini-2.0-flash"]
                available_models = []

                try:
                    models_res = requests.get(
                        f"{api_base}/models",
                        headers={"x-goog-api-key": GEMINI_API_KEY},
                        timeout=15
                    )
                    if models_res.ok:
                        for model_info in models_res.json().get("models", []):
                            name = model_info.get("name", "")
                            methods = model_info.get("supportedGenerationMethods", [])
                            if "generateContent" in methods:
                                available_models.append(name.split("/")[-1])
                except requests.RequestException:
                    pass

                model_name = next(
                    (m for m in preferred_models if m in available_models),
                    None
                )

                if model_name is None:
                    if available_models:
                        flash_models = [
                            m for m in available_models
                            if "flash" in m.lower() and "embedding" not in m.lower()
                        ]
                        model_name = flash_models[0] if flash_models else available_models[0]
                    else:
                        model_name = "gemini-2.5-flash"

                api_url = f"{api_base}/models/{model_name}:generateContent"

                payload = {
                    "systemInstruction": {
                        "parts": [{"text": system_instruction}]
                    },
                    "contents": [
                        {
                            "role": "user",
                            "parts": [{"text": user_prompt}]
                        }
                    ],
                    "generationConfig": {
                        "temperature": 0.3,
                        "maxOutputTokens": 1024
                    }
                }

                res = requests.post(
                    api_url,
                    json=payload,
                    headers=headers,
                    timeout=30
                )

                try:
                    data = res.json()
                except ValueError:
                    data = {}

                if res.ok:
                    candidates = data.get("candidates", [])
                    if candidates:
                        parts = candidates[0].get("content", {}).get("parts", [])
                        response_text = "".join(
                            part.get("text", "")
                            for part in parts
                            if part.get("text")
                        ).strip()

                        if not response_text:
                            response_text = "⚠️ Gemini đã phản hồi nhưng không có nội dung."
                    else:
                        response_text = "⚠️ Gemini không trả về nội dung tư vấn."
                else:
                    err = data.get("error", {})
                    code = err.get("code", res.status_code)
                    message = err.get("message", res.text)

                    if code == 404:
                        response_text = (
                            f"⛔ Model Gemini '{model_name}' không khả dụng cho API Key này. "
                            "Hãy kiểm tra Gemini API/Google AI Studio và model được cấp quyền."
                        )
                    elif code in (400, 401, 403):
                        response_text = (
                            "⛔ API Key Gemini không hợp lệ, chưa được kích hoạt hoặc "
                            "chưa có quyền gọi Gemini API. Hãy tạo/cập nhật API Key mới."
                        )
                    elif code == 429:
                        response_text = (
                            "⏳ Gemini đang giới hạn số lượt gọi (429). "
                            "Vui lòng thử lại sau."
                        )
                    else:
                        response_text = f"⛔ Gemini trả về lỗi HTTP {code}: {message}"

            except requests.Timeout:
                response_text = "⏱️ Gemini phản hồi quá lâu. Vui lòng thử lại."
            except requests.RequestException as e:
                response_text = f"🌐 Không kết nối được Gemini API: {e}"
            except Exception as e:
                response_text = f"Lỗi hệ thống: {e}"

        st.session_state.messages.append({"role": "assistant", "content": response_text})
        st.chat_message("assistant").write(response_text)

# TAB 3: ĐẶT LỊCH HẸN
with tab3:
    st.subheader("📝 Đặt Lịch Hẹn")
    with st.form("form_dat_lich"):
        ten_kh = st.text_input("Họ tên:")
        sdt_kh = st.text_input("Số điện thoại:")
        dongxe_kh = st.text_input("Xe đang đi:")
        dich_vu = st.multiselect("Dịch vụ:", ["Thay lốp mới", "Tráng keo", "Cân bằng động"])
        btn_submit = st.form_submit_button("🚀 ĐẶT LỊCH NGAY", type="primary", use_container_width=True)
        if btn_submit and ten_kh and sdt_kh:
            msg = f"ĐẶT LỊCH:\n- Khách: {ten_kh}\n- SĐT: {sdt_kh}\n- Xe: {dongxe_kh}\n- DV: {', '.join(dich_vu)}"
            st.markdown(f'<a href="https://zalo.me/{ZALO_PHONE}?text={urllib.parse.quote(msg)}" target="_blank" class="btn-zalo-quote">💬 BẤM GỬI ZALO</a>', unsafe_allow_html=True)
            import streamlit.components.v1 as components

# Đoạn mã lấy từ nút Installation của Coze
# Hiển thị khung chat AI lên web
st.markdown("---")
st.subheader("💬 Trợ Lý Tư Vấn Lốp Xe AI")

coze_html = """
<script src="https://sf-cdn.coze.com/obj/unpkg-va/flow-platform/chat-app-sdk/1.2.0-beta.6/libs/oversea/index.js"></script>
<script>
  new CozeWebSDK.WebChatClient({
    config: {
      bot_id: '7672001578140024885',
    },
    componentProps: {
      title: 'Trợ lý Lốp xe Cần Thơ',
    },
    auth: {
      type: 'token',
      token: 'pat_GZzNnhGUnT1t8J5o9P1G1...',  # (giữ nguyên token thật của ông)
      onRefreshToken: function () {
        return 'pat_GZzNnhGUnT1t8J5o9P1G1...'
      }
    }
  });
</script>
"""

components.html(coze_html, height=650)
# Hiển thị khung chat AI lên web
st.markdown("---")
st.subheader("💬 Trợ Lý Tư Vấn Lốp Xe AI")
components.html(coze_html, height=650)