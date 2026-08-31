import streamlit as st
from google import genai
from gemini_service import GeminiService
from word_processor import WordProcessor
from pptx_processor import PPTXProcessor

# Mật khẩu mở khóa tính năng PowerPoint (Bạn có thể đổi mật khẩu này tùy ý)
PREMIUM_PASSWORD = "GIAOVIEN2026"

st.set_page_config(
    page_title="Tích hợp Năng lực số & AI vào KHBD / PowerPoint",
    page_icon="📝",
    layout="wide"
)

# CSS tùy chỉnh
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

# --- CẤU HÌNH HỆ THỐNG ---
with st.expander("⚙️ **CẤU HÌNH HỆ THỐNG & XÁC THỰC BẢN QUYỀN:**", expanded=True):
    col_cfg1, col_cfg2 = st.columns([1, 1])
    
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
            check_key_btn = st.button("Kiểm tra", type="primary", use_container_width=True)

        if api_key_input:
            st.session_state["gemini_api_key"] = api_key_input.strip()
            api_key = api_key_input.strip()
        else:
            api_key = ""

        # Xử lý khi nhấn nút Kiểm tra API Key
        if check_key_btn:
            if not api_key:
                st.warning("⚠️ Vui lòng dán mã API Key trước khi kiểm tra.")
            else:
                try:
                    client_test = genai.Client(api_key=api_key)
                    # Gọi thử một câu lệnh tối thiểu để xác thực key
                    client_test.models.generate_content(
                        model="gemini-2.5-flash",
                        contents="ping"
                    )
                    st.success("✅ API Key hợp lệ và sẵn sàng sử dụng!")
                except Exception as ex:
                    st.error(f"❌ API Key không hợp lệ hoặc đã hết hạn. Chi tiết lỗi: {str(ex)}")

    with col_cfg2:
        st.markdown("🔐 **Mật khẩu mở khóa tính năng nâng cao (PowerPoint):**")
        input_password = st.text_input(
            "Mật khẩu",
            type="password",
            placeholder="Nhập mật khẩu để mở khóa PPTX...",
            label_visibility="collapsed"
        )
        
        # Kiểm tra trạng thái mật khẩu
        is_premium = (input_password == PREMIUM_PASSWORD)
        if input_password:
            if is_premium:
                st.success("🎉 Mở khóa thành công! Cho phép tích hợp cả Word (.docx) và PowerPoint (.pptx).")
            else:
                st.error("❌ Mật khẩu chưa đúng. Hệ thống đang ở chế độ cơ bản (Chỉ tích hợp file Word .docx).")
        else:
            st.caption("ℹ️ *Chưa có mật khẩu: Chỉ hỗ trợ xử lý file Word (.docx).*")

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

# Thiết lập định dạng file cho phép tải lên dựa theo trạng thái mật khẩu
allowed_types = ["docx", "pptx"] if is_premium else ["docx"]

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
        
        help_text = "Hỗ trợ cả .docx và .pptx" if is_premium else "Chế độ cơ bản: Chỉ hỗ trợ tệp Word (.docx)."
        uploaded_file = st.file_uploader(
            f"**Chọn file ({', '.join(allowed_types)}):**", 
            type=allowed_types,
            help=help_text
        )

        if uploaded_file is not None:
            file_bytes = uploaded_file.read()
            file_ext = uploaded_file.name.split('.')[-1].lower()
            st.success(f"✔️ Đã tải lên file: **{uploaded_file.name}**")
            st.session_state['original_filename'] = uploaded_file.name
            st.session_state['file_ext'] = file_ext
            
            if st.button("🚀 Bắt đầu tích hợp", type="primary", use_container_width=True):
                if not api_key:
                    st.error("⚠️ Vui lòng nhập **Google Gemini API Key** ở khung cấu hình phía trên trước khi tiếp tục.")
                    st.stop()

                # Kiểm tra bảo mật nếu file là PowerPoint
                if file_ext == "pptx" and not is_premium:
                    st.error("⛔ Bạn cần nhập đúng mật khẩu kích hoạt để sử dụng tính năng tích hợp PowerPoint.")
                    st.stop()

                with st.spinner("🔄 Đang phân tích dữ liệu và tích hợp năng lực..."):
                    try:
                        ai_handler = GeminiService(api_key=api_key)
                        
                        if file_ext == "pptx":
                            # Luồng xử lý file PowerPoint
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
                            # Luồng xử lý file Word
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
                        
                        progress_bar.progress(100, text="Hoàn tất xử lý!")
                        st.success("🎉 Tích hợp thành công!")
                        
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
    st.markdown("### ℹ️ Hướng dẫn sử dụng")
    st.markdown("""
    - **Bước 1:** Nhận API Key tại link [Google AI Studio](https://aistudio.google.com/app/apikey) và dán vào ô nhập.
    - **Bước 2:** Bấm **"Kiểm tra"** để xác minh API Key.
    - **Bước 3:** Nhập mật khẩu nếu muốn sử dụng tính năng tích hợp PowerPoint.
    - **Bước 4:** Tải file lên, chọn cấp học và nhấn **"Bắt đầu tích hợp"**.
    """)
    st.markdown("#### 📌 Khung năng lực áp dụng:")
    st.markdown("""
    - **Năng lực số:** Thông tư số 02/2025/TT-BGDĐT.
    - **Năng lực AI:** Quyết định số 2422/QĐ-BGDĐT.
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
