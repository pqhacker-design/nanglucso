import streamlit as st
from gemini_service import GeminiService
from word_processor import WordProcessor
from pptx_processor import PPTXProcessor

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

st.set_page_config(
    page_title="Tích hợp Năng lực số & AI vào KHBD / PowerPoint",
    page_icon="📝",
    layout="wide"
)

st.markdown("## 🤖 Tích hợp Năng lực số và AI tự động vào KHBD / PowerPoint")
st.info("Hỗ trợ tích hợp Năng lực số (Thông tư 02/2025/TT-BGDĐT) và Năng lực AI (QĐ 2422/QĐ-BGDĐT) vào file Word (.docx) hoặc Slide Notes của PowerPoint (.pptx).")

# --- CẤU HÌNH HỆ THỐNG ---
with st.expander("⚙️ **CẤU HÌNH HỆ THỐNG:**", expanded=False):
    col_cfg1, col_cfg2, col_cfg3 = st.columns([2, 1, 1])
    
    with col_cfg1:
        if "gemini_api_key" in st.session_state and st.session_state["gemini_api_key"].strip() != "":
            api_key = st.session_state["gemini_api_key"]
            st.success("🔑 **Trạng thái API Key:** Đã nhận diện thành công.")
        else:
            st.warning("⚠️ **Chưa tìm thấy API Key:** Vui lòng quay lại **Trang chủ** để nhập Google Gemini API Key.")
            st.page_link("🏠_Trang_Chủ.py", label="Nhấn vào đây để Quay lại Trang chủ", icon="🔄")
            st.stop()

    with col_cfg2:
        cap_hoc = st.selectbox(
            "**Chọn cấp học mục tiêu:**",
            ["Tự động nhận diện", "Tiểu học", "THCS", "THPT"]
        )

    with col_cfg3:
        integration_type = st.selectbox(
            "**Loại tích hợp:**",
            ["Cả hai", "Năng lực số", "Năng lực AI"]
        )

# --- MÀN HÌNH CHÍNH: 2 CỘT ---
col_left, col_right = st.columns([2, 1])

with col_left:
    with st.container(border=True):
        st.markdown(
            """
            <div style="background-color: #E0F2FE; padding: 4px; border-left: 5px solid #0284C7; border-radius: 4px; margin-bottom: 10px;">
                <h4 style="margin: 0; color: #0369A1;">📂 1. Tải lên tệp Giáo án (.docx) hoặc Bài giảng (.pptx)</h4>
            </div>
            """, 
            unsafe_allow_html=True
        )
        uploaded_file = st.file_uploader(
            "**Chọn file Word (.docx) hoặc PowerPoint (.pptx):**", 
            type=["docx", "pptx"],
            help="Hệ thống tự động nhận diện định dạng file tải lên."
        )

        if uploaded_file is not None:
            file_bytes = uploaded_file.read()
            file_ext = uploaded_file.name.split('.')[-1].lower()
            st.success(f"✔️ Đã tải lên file thành công: **{uploaded_file.name}**")
            st.session_state['original_filename'] = uploaded_file.name
            st.session_state['file_ext'] = file_ext
            
            if st.button("🚀 Bắt đầu tích hợp", type="primary", use_container_width=True):
                with st.spinner("🔄 Đang đọc dữ liệu và gửi phân tích tới Gemini AI..."):
                    try:
                        ai_handler = GeminiService(api_key=api_key)
                        
                        if file_ext == "pptx":
                            # Luồng xử lý file PowerPoint
                            progress_bar = st.progress(15, text="Đang đọc nội dung các Slide PowerPoint...")
                            slides_data = PPTXProcessor.extract_slides_text(file_bytes)
                            
                            if not slides_data:
                                st.error("❌ Không tìm thấy văn bản trong file PowerPoint.")
                                st.stop()
                                
                            progress_bar.progress(40, text="AI đang phân tích các slide và thiết kế nội dung ghi chú...")
                            doc_text = PPTXProcessor.format_doc_text_for_ai(slides_data)
                            ai_result = ai_handler.analyze_pptx_and_integrate(doc_text, cap_hoc, integration_type)
                            st.session_state['ai_result'] = ai_result
                            
                            progress_bar.progress(80, text="Đang chèn nội dung vào Slide Notes...")
                            processed_file = PPTXProcessor.integrate_into_notes(file_bytes, ai_result)
                            st.session_state['processed_file'] = processed_file
                            
                        else:
                            # Luồng xử lý file Word
                            progress_bar = st.progress(15, text="Đang đọc nội dung file Word...")
                            doc_text = WordProcessor.extract_text(file_bytes)
                            
                            if not doc_text.strip():
                                st.error("❌ File Word trống hoặc không tìm thấy nội dung văn bản hợp lệ.")
                                st.stop()
                                
                            progress_bar.progress(40, text="AI đang phân tích và thiết kế nội dung tích hợp...")
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
                st.warning("AI không tìm thấy hoặc không đề xuất vị trí tích hợp nào phù hợp.")
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
                            st.markdown(f"**Văn bản gốc tìm thấy:** `{anchor}`")
                            st.markdown(f"**Nội dung được chèn:** <span style='color:{color}; font-weight:bold;'>{content}</span>", unsafe_allow_html=True)
            
            st.markdown("---")
            
            # Tên file tải về
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
    - **Tải lên** file giáo án Word (`.docx`) hoặc bài giảng PowerPoint (`.pptx`).
    - Chọn **cấp học mục tiêu** và **loại năng lực** cần tích hợp.
    - Nhấn **"Bắt đầu tích hợp"** để AI tự động phân tích và xử lý.
    - **Cơ chế hoạt động:**
      + **Với file Word (.docx):** Tự động chèn trực tiếp mục tiêu và hoạt động số/AI vào đúng vị trí của từng bài.
      + **Với file PowerPoint (.pptx):** Tự động thêm các hướng dẫn sư phạm số/AI vào phần **Ghi chú diễn giả (Slide Notes)** mà không làm ảnh hưởng đến bố cục trình chiếu của slide.
    """)
    st.markdown("#### 📌 Khung chuẩn tham chiếu:")
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
