import time
import streamlit as st
from streamlit_cookies_controller import CookieController
from google import genai
from gemini_service import GeminiService
from word_processor import WordProcessor
from pptx_processor import PPTXProcessor

# --- CẤU HÌNH BẢN QUYỀN & GIỚI HẠN DÙNG THỬ ---
VALID_ACCOUNTS = {
    "admin": "GIAOVIEN2026",
    "thayhung": "123456",
    "giaovien": "hoctap2026"
}
MAX_FREE_TRIALS = 2
COOKIE_KEY = "user_khbd_trial_usage"

st.set_page_config(
    page_title="Tích hợp Năng lực số & AI vào KHBD / PowerPoint",
    page_icon="📝",
    layout="wide"
)

# Khởi tạo bộ điều khiển Cookie
cookie_controller = CookieController()

# Đọc số lượt dùng đã lưu từ Cookie trình duyệt
cookie_val = cookie_controller.get(COOKIE_KEY)
try:
    current_usage = int(cookie_val) if cookie_val is not None else 0
except (ValueError, TypeError):
    current_usage = 0

# Đồng bộ với session_state
if "usage_count" not in st.session_state:
    st.session_state["usage_count"] = current_usage
else:
    st.session_state["usage_count"] = max(st.session_state["usage_count"], current_usage)

# CSS giao diện
st.markdown(
    """
    <style>
    div[data-testid="stPageLink"] a {
        background-color: #0284C7 !important;
        color: white !important;
        border-radius: 8px !important;
        padding: 8px 16px !important;
        text-decoration: none !important;
        font-weight: bold !important;
        border: 1px solid #0369A1 !important;
        display: inline-flex !important;
        align-items: center !important;
        transition: all 0.3s ease !important;
    }
    div[data-testid="stPageLink"] a:hover {
        background-color: #0369A1 !important;
        color: #F0F9FF !important;
        border-color: #075985 !important;
        transform: translateY(-2px);
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown("## 🤖 Tích hợp Năng lực số và AI tự động vào KHBD / PowerPoint")
st.info("Hỗ trợ tích hợp Năng lực số (Thông tư 02/2025/TT-BGDĐT) và Năng lực AI (QĐ 2422/QĐ-BGDĐT) vào file Word (.docx) hoặc Slide Notes của PowerPoint (.pptx).")

# --- CẤU HÌNH HỆ THỐNG & TÀI KHOẢN ---
with st.expander("⚙️ **CẤU HÌNH HỆ THỐNG & KÍCH HOẠT BẢN QUYỀN:**", expanded=True):
    col_cfg1, col_cfg2 = st.columns([1.1, 0.9])
    
    with col_cfg1:
        st.markdown(
            "🔑 **Google Gemini API Key:** "
            "([👉 Nhấn vào đây để lấy API Key miễn phí](https://aistudio.google.com/app/apikey))",
            unsafe_allow_html=True
        )
        
        default_key = st.session_state.get("gemini_api_key", "")
        col_key_input, col_key_btn = st.columns([3, 1])
        with col_key_input:
            api_key_input = st.text_input(
                "API Key",
                value=default_key,
                type="password",
                placeholder="Dán API Key (AIzaSy...)",
                label_visibility="collapsed"
            )
        with col_key_btn:
            check_key_btn = st.button("Kiểm tra", use_container_width=True)

        if api_key_input:
            st.session_state["gemini_api_key"] = api_key_input.strip()
            api_key = api_key_input.strip()
        else:
            api_key = ""

        if check_key_btn:
            if not api_key:
                st.warning("⚠️ Vui lòng dán mã API Key trước khi kiểm tra.")
            else:
                with st.spinner("Đang kiểm tra kết nối API..."):
                    try:
                        client_test = genai.Client(api_key=api_key)
                        client_test.models.generate_content(
                            model="gemini-2.5-flash",
                            contents="ping"
                        )
                        st.success("✅ Kết nối thành công! API Key hợp lệ.")
                    except Exception as ex:
                        msg_vi = GeminiService._format_api_error(ex)
                        st.error(f"❌ Kết nối thất bại: {msg_vi}")

    with col_cfg2:
        st.markdown("🔐 **Đăng nhập bản quyền (Không giới hạn lượt dùng):**")
        col_acc1, col_acc2 = st.columns(2)
        with col_acc1:
            input_user = st.text_input("Tên tài khoản", placeholder="Nhập tên tài khoản...", label_visibility="collapsed")
        with col_acc2:
            input_pwd = st.text_input("Mật khẩu", type="password", placeholder="Nhập mật khẩu...", label_visibility="collapsed")

        # Kiểm tra trạng thái tài khoản
        is_authenticated = (
            input_user.strip() in VALID_ACCOUNTS and 
            VALID_ACCOUNTS.get(input_user.strip()) == input_pwd.strip()
        )

        remaining_trials = max(0, MAX_FREE_TRIALS - st.session_state["usage_count"])

        if is_authenticated:
            st.success(f"🎉 Đã kích hoạt bản quyền đầy đủ (Tài khoản: **{input_user}**). Không giới hạn lượt dùng!")
        else:
            if remaining_trials > 0:
                st.info(f"🎁 Chế độ dùng thử: Còn **{remaining_trials}/{MAX_FREE_TRIALS}** lượt tích hợp miễn phí trên trình duyệt này.")
            else:
                st.error("⛔ Đã hết 2 lượt dùng thử! Vui lòng nhập đúng Tên tài khoản và Mật khẩu.")

    col_sub1, col_sub2 = st.columns(2)
    with col_sub1:
        cap_hoc = st.selectbox(
            "**Cấp học mục tiêu:**",
            ["Tự động nhận diện", "Tiểu học", "THCS", "THPT"]
        )
    with col_sub2:
        integration_type = st.selectbox(
            "**Loại tích hợp:**",
            ["Cả hai", "Năng lực số", "Năng lực AI"]
        )

# Quyết định quyền được phép chạy tiếp
can_use_app = is_authenticated or (remaining_trials > 0)

# --- MÀN HÌNH CHÍNH ---
col_left, col_right = st.columns([2, 1])

with col_left:
    with st.container(border=True):
        st.markdown(
            """
            <div style="background-color: #E0F2FE; padding: 4px; border-left: 5px solid #0284C7; border-radius: 4px; margin-bottom: 10px;">
                <h4 style="margin: 0; color: #0369A1;">📂 1. Tải lên tệp bài giảng / kế hoạch bài dạy</h4>
            </div>
            """, 
            unsafe_allow_html=True
        )
        
        uploaded_file = st.file_uploader(
            "**Chọn file Word (.docx) hoặc PowerPoint (.pptx):**", 
            type=["docx", "pptx"],
            help="Hệ thống hỗ trợ cả file Word (.docx) và Slide PowerPoint (.pptx)."
        )

        if uploaded_file is not None:
            file_bytes = uploaded_file.read()
            file_ext = uploaded_file.name.split('.')[-1].lower()
            st.success(f"✔️ Đã tải lên file: **{uploaded_file.name}**")
            st.session_state['original_filename'] = uploaded_file.name
            st.session_state['file_ext'] = file_ext
            
            if st.button("🚀 Bắt đầu tích hợp", type="primary", use_container_width=True):
                # 1. Kiểm tra quyền sử dụng
                if not can_use_app:
                    st.error("⛔ Bạn đã sử dụng hết 2 lượt dùng thử. Vui lòng nhập đúng Tên tài khoản và Mật khẩu để tiếp tục.")
                    st.stop()

                # 2. Kiểm tra API Key
                if not api_key:
                    st.error("⚠️ Vui lòng nhập **Google Gemini API Key** ở khung cấu hình phía trên trước khi tiếp tục.")
                    st.stop()

                with st.spinner("🔄 Đang phân tích dữ liệu và tích hợp năng lực..."):
                    try:
                        ai_handler = GeminiService(api_key=api_key)
                        
                        if file_ext == "pptx":
                            progress_bar = st.progress(20, text="Đang trích xuất nội dung các Slide...")
                            slides_data = PPTXProcessor.extract_slides_text(file_bytes)
                            
                            if not slides_data:
                                st.error("❌ Không tìm thấy văn bản trong file PowerPoint.")
                                st.stop()
                                
                            progress_bar.progress(50, text="AI đang thiết kế ghi chú sư phạm cho từng slide...")
                            doc_text = PPTXProcessor.format_doc_text_for_ai(slides_data)
                            ai_result = ai_handler.analyze_pptx_and_integrate(doc_text, cap_hoc, integration_type)
                            st.session_state['ai_result'] = ai_result
                            
                            progress_bar.progress(80, text="Đang cập nhật nội dung vào Slide Notes...")
                            processed_file = PPTXProcessor.integrate_into_notes(file_bytes, ai_result)
                            st.session_state['processed_file'] = processed_file
                            
                        else:
                            progress_bar = st.progress(20, text="Đang đọc nội dung file Word...")
                            doc_text = WordProcessor.extract_text(file_bytes)
                            
                            if not doc_text.strip():
                                st.error("❌ File Word trống hoặc không tìm thấy nội dung hợp lệ.")
                                st.stop()
                                
                            progress_bar.progress(50, text="AI đang phân tích và thiết kế nội dung tích hợp...")
                            ai_result = ai_handler.analyze_and_integrate(doc_text, cap_hoc, integration_type)
                            st.session_state['ai_result'] = ai_result
                            
                            progress_bar.progress(80, text="Đang chèn nội dung vào file Word...")
                            processed_file = WordProcessor.integrate_digital_capacity(file_bytes, ai_result, integration_type)
                            st.session_state['processed_file'] = processed_file
                        
                        # Cập nhật số lượt vào Session State và ghi bền vững vào Cookie trình duyệt (hạn 365 ngày)
                        if not is_authenticated:
                            new_count = st.session_state["usage_count"] + 1
                            st.session_state["usage_count"] = new_count
                            cookie_controller.set(COOKIE_KEY, str(new_count), max_age=365*24*60*60)

                        progress_bar.progress(100, text="Hoàn tất xử lý!")
                        st.success("🎉 Tích hợp thành công!")
                        
                        # Cho cookie kịp đồng bộ trước khi refresh lại UI
                        time.sleep(0.5)
                        st.rerun()
                        
                    except Exception as e:
                        st.error(f"❌ Đã xảy ra lỗi trong quá trình xử lý: {str(e)}")

    # --- Hiển thị kết quả và nút tải xuống ---
    if 'ai_result' in st.session_state and 'processed_file' in st.session_state:
        with st.container(border=True):
            st.markdown(
                """
                <div style="background-color: #E0F2FE; padding: 4px; border-left: 5px solid #0284C7; border-radius: 4px; margin-bottom: 10px;">
                    <h4 style="margin: 0; color: #0369A1;">📋 2. Kết quả tích hợp</h4>
                </div>
                """, 
                unsafe_allow_html=True
            )
            res = st.session_state['ai_result']
            sua_doi_list = res.get('sua_doi', [])
            file_ext = st.session_state.get('file_ext', 'docx')
            
            if not sua_doi_list:
                st.warning("AI không tìm thấy hoặc không đề xuất vị trí tích hợp nào.")
            else:
                for idx, item in enumerate(sua_doi_list):
                    content = item.get('insert_content', 'Không có nội dung')
                    loai = item.get('loai', 'Năng lực số')
                    icon = "🧠" if loai == "Năng lực AI" else "💻"
                    color = "#D97706" if loai == "Năng lực AI" else "#0066CC"
                    
                    if file_ext == "pptx":
                        slide_num = item.get('slide_number', 'Chưa rõ')
                        with st.expander(f"{icon} Slide {slide_num}: ({loai})", expanded=True):
                            st.markdown(f"**Ghi chú vào Slide Notes:** <span style='color:{color}; font-weight:bold;'>{content}</span>", unsafe_allow_html=True)
                    else:
                        anchor = item.get('anchor_text', 'Không rõ vị trí')
                        with st.expander(f"{icon} Vị trí {idx+1}: Sau \"{anchor}\" ({loai})", expanded=True):
                            st.markdown(f"**Văn bản gốc:** `{anchor}`")
                            st.markdown(f"**Nội dung chèn:** <span style='color:{color}; font-weight:bold;'>{content}</span>", unsafe_allow_html=True)
            
            st.markdown("---")
            
            orig_name = st.session_state.get('original_filename', 'KHBD_TichHop')
            base_name = orig_name.rsplit('.', 1)[0]
            
            if file_ext == "pptx":
                download_filename = f"{base_name}_TichHop_Notes.pptx"
                mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            else:
                download_filename = f"{base_name}_Tichhop_So_AI.docx"
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            
            st.download_button(
                label=f"💾 TẢI XUỐNG TỆP ĐÃ TÍCH HỢP (.{file_ext.upper()})",
                data=st.session_state['processed_file'],
                file_name=download_filename,
                type="primary",
                mime=mime_type,
                use_container_width=True
            )

with col_right:
    st.markdown("### ℹ️ Hướng dẫn & Chính sách sử dụng")
    st.markdown("""
    - **Bước 1:** Nhập **API key** và bấm **Kiểm tra**
    - **Bước 2:** Tải lên file KHBD **Word(.docx)** hoặc Bài giảng **PowerPoint(.pptx)**
    - **Bước 3:** Bấm vào **Bắt đầu tích hợp**
    -----------------------
    - **Dùng thử miễn phí:** Tối đa **2 lần** tích hợp trên mỗi trình duyệt.
    - **Bản quyền đầy đủ:** Nhập đúng **Tên tài khoản & Mật khẩu** được cấp để sử dụng không giới hạn.
    - **Lấy API Key:** Nhận miễn phí tại [Google AI Studio](https://aistudio.google.com/app/apikey).
    """)
    st.markdown("#### 📌 Khung năng lực áp dụng:")
    st.markdown("""
    - **Năng lực số:** Thông tư số 02/2025/TT-BGDĐT.
    - **Năng lực AI:** Quyết định số 2422/QĐ-BGDĐT.
    -----------------------
    - **Zalo: 0913117321**
    """)

st.divider()
col_left_f, col_right_f = st.columns(2)
with col_left_f:
    st.caption("Phát triển bởi Ngo Thanh Hung © 2026")
with col_right_f:
    st.markdown(
        "<div style='text-align: right; color: gray; font-size: 0.85em;'>"
        "AI có thể mắc lỗi. Cần kiểm tra kỹ các thông tin quan trọng."
        "</div>", 
        unsafe_allow_html=True
    )
