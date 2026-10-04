import json
from pydantic import BaseModel, Field
from typing import List, Literal, Optional
from google import genai
from google.genai import types
from google.genai.errors import APIError

# --- ĐỊNH NGHĨA SCHEMA DỮ LIỆU ĐẦU RA CỦA GEMINI ---
class SuaDoiItem(BaseModel):
    slide_number: Optional[int] = Field(
        default=None, 
        description="Số thứ tự của Slide trong file PowerPoint (đối với tệp PPTX)."
    )
    anchor_text: Optional[str] = Field(
        default="", 
        description="Điểm neo: Trích câu văn nguyên bản (nếu là KHBD) hoặc Tên bài học chính xác (nếu là bảng KHGD)."
    )
    insert_content: str = Field(
        description="Nội dung sư phạm tích hợp ngắn gọn, khả thi, bắt đầu bằng động từ hành động của học sinh hoặc hướng dẫn của GV."
    )
    loai: Literal["Năng lực số", "Năng lực AI", "Giáo dục STEM"] = Field(
        description="Phân loại: 'Năng lực số', 'Năng lực AI' hoặc 'Giáo dục STEM'."
    )

class TichHopResult(BaseModel):
    sua_doi: List[SuaDoiItem]


# --- LỚP DỊCH VỤ GEMINI XỬ LÝ SƯ PHẠM ---
class GeminiService:
    def __init__(self, api_key: str):
        if not api_key:
            raise ValueError("Mã API Key không được để trống. Vui lòng nhập API Key.")
        self.client = genai.Client(api_key=api_key)
        self.model_name = "gemini-2.5-flash"

    @staticmethod
    def _format_api_error(error: Exception) -> str:
        """Chuẩn hóa và chuyển đổi các mã lỗi kỹ thuật từ Google sang tiếng Việt thân thiện."""
        err_str = str(error)
        if "API_KEY_INVALID" in err_str or "API key not valid" in err_str or "PERMISSION_DENIED" in err_str:
            return "Mã API Key không chính xác hoặc đã bị vô hiệu hóa. Vui lòng kiểm tra lại trên Google AI Studio."
        if "RESOURCE_EXHAUSTED" in err_str or "429" in err_str or "quota" in err_str.lower():
            return "Đã đạt giới hạn yêu cầu của Google (hết lượt dùng thử hoặc gửi liên tục). Vui lòng đợi 30 - 60 giây rồi thử lại."
        if "INVALID_ARGUMENT" in err_str:
            return "Nội dung tệp gửi đi chứa ký tự đặc biệt hoặc định dạng chưa phù hợp."
        if "DEADLINE_EXCEEDED" in err_str or "timeout" in err_str.lower():
            return "Quá thời gian kết nối (Timeout) do mạng yếu hoặc tệp quá dài. Vui lòng thử lại."
        if "UNAVAILABLE" in err_str or "503" in err_str:
            return "Máy chủ Google Gemini đang bảo trì hoặc quá tải tạm thời. Vui lòng thử lại sau vài phút."
        return f"Lỗi phản hồi từ Google: {err_str}"

    def _get_tt02_framework_prompt(self, cap_hoc: str) -> str:
        """Khung Năng lực số Thông tư 02/2025/TT-BGDĐT."""
        base_framework = """
* Khung Năng lực số (TT 02/2025/TT-BGDĐT) gồm 6 miền:
  1. Vận hành thiết bị và phần mềm
  2. Khai thác thông tin và dữ liệu
  3. Giao tiếp và hợp tác trong môi trường số
  4. Sáng tạo nội dung số
  5. An toàn và an ninh số
  6. Giải quyết vấn đề trong môi trường số
"""
        level_guide = {
            "Tiểu học": "\n- Tiểu học: Thao tác đơn giản, tìm kiếm cơ bản, ý thức bảo vệ tư thế, mắt và bảo mật thông tin cá nhân.",
            "THCS": "\n- THCS: Khai thác phần mềm môn học/mô phỏng, xử lý số liệu bảng tính, làm việc nhóm trực tuyến an toàn, tôn trọng bản quyền số.",
            "THPT": "\n- THPT: Phân tích dữ liệu chuyên sâu, thiết kế sản phẩm số tương tác, an ninh mạng, tư duy máy tính và pháp luật số.",
            "Tự động nhận diện": "\n- Tự động nhận diện cấp học theo từng nội dung để tích hợp vừa sức học sinh."
        }
        return base_framework + level_guide.get(cap_hoc, level_guide["Tự động nhận diện"])

    def _get_stem_framework_prompt(self, cap_hoc: str) -> str:
        """Khung Giáo dục STEM theo Công văn 3089/BGDĐT-GDTrH và 909/BGDĐT-GDTH."""
        stem_base = """
* Khung Giáo dục STEM (Công văn 3089/BGDĐT-GDTrH và 909/BGDĐT-GDTH):
  - Định hướng: Dạy học giải quyết vấn đề thực tiễn gắn với liên môn Khoa học (S), Công nghệ (T), Kỹ thuật (E), Toán học (M).
  - Quy trình kỹ thuật 5 bước cốt lõi:
    1. Xác định vấn đề thực tiễn hoặc tiêu chí sản phẩm.
    2. Nghiên cứu kiến thức nền tảng và đề xuất giải pháp.
    3. Lựa chọn giải pháp thiết kế khả thi.
    4. Chế tạo mẫu, thử nghiệm và đánh giá chất lượng.
    5. Chia sẻ, thảo luận và cải tiến sản phẩm.
"""
        level_stem = {
            "Tiểu học": "\n- Tiểu học: Trải nghiệm thực hành đơn giản từ vật liệu tái chế, lắp ráp mô hình trực quan gần gũi.",
            "THCS": "\n- THCS: Chế tạo dụng cụ học tập/thí nghiệm, giải quyết bài toán môi trường, ứng dụng nguyên lý khoa học và kỹ thuật đơn giản.",
            "THPT": "\n- THPT: Thiết kế dự án nghiên cứu, lập trình mô phỏng, phân tích tối ưu hóa kỹ thuật và chi phí thực hiện.",
            "Tự động nhận diện": "\n- Điều chỉnh độ phức tạp của bài toán kỹ thuật theo đúng đối tượng học sinh."
        }
        return stem_base + level_stem.get(cap_hoc, level_stem["Tự động nhận diện"])

    def _build_focus_instruction(self, cap_hoc: str, integration_type: str) -> str:
        """Xây dựng chỉ dẫn tập trung theo loại tích hợp."""
        tt02_info = self._get_tt02_framework_prompt(cap_hoc)
        stem_info = self._get_stem_framework_prompt(cap_hoc)
        ai_framework_info = """
* Khung Năng lực AI (QĐ 2422/QĐ-BGDĐT):
  - Nhận thức AI: Nhận diện công nghệ AI quanh ta, hiểu nguyên lý cơ bản của AI.
  - Ứng dụng AI: Dùng AI hỗ trợ tra cứu, tóm tắt bài, gợi mở ý tưởng, hỗ trợ giải bài tập, tạo hình ảnh minh họa.
  - Tư duy phản biện & Đạo đức AI: Đánh giá độ tin cậy của AI, kiểm chứng nguồn tin, chống gian lận học thuật, tôn trọng bản quyền.
"""
        if integration_type == "Năng lực số":
            return f"YÊU CẦU: CHỈ TÍCH HỢP NĂNG LỰC SỐ (TT 02/2025/TT-BGDĐT).\n{tt02_info}\nTất cả các mục trả về đều có 'loai': 'Năng lực số'."
        elif integration_type == "Năng lực AI":
            return f"YÊU CẦU: CHỈ TÍCH HỢP NĂNG LỰC AI (QĐ 2422/QĐ-BGDĐT).\n{ai_framework_info}\nTất cả các mục trả về đều có 'loai': 'Năng lực AI'."
        elif integration_type == "Giáo dục STEM":
            return f"YÊU CẦU: CHỈ TÍCH HỢP GIÁO DỤC STEM (CV 3089/BGDĐT & CV 909/BGDĐT).\n{stem_info}\nTất cả các mục trả về đều có 'loai': 'Giáo dục STEM'."
        else:  # "Tất cả" hoặc "Cả hai"
            return f"""
YÊU CẦU BẮT BUỘC: TÍCH HỢP TOÀN DIỆN CẢ 3 NHÓM:
1. Năng lực số (TT 02/2025/TT-BGDĐT): {tt02_info}
2. Năng lực AI (QĐ 2422/QĐ-BGDĐT): {ai_framework_info}
3. Hoạt động/Bài học STEM (CV 3089/BGDĐT, CV 909/BGDĐT): {stem_info}

Quy tắc phân bổ:
- Kết quả trả về PHẢI phân bổ hài hòa cả 3 giá trị 'Năng lực số', 'Năng lực AI', 'Giáo dục STEM'.
"""

    def analyze_and_integrate(self, doc_text: str, cap_hoc: str, integration_type: str) -> dict:
        """
        Phân tích tài liệu Word: Tự động phân loại KHBD (giáo án tiến trình) hoặc KHGD (bảng năm học)
        và quét toàn bộ từ đầu đến cuối không giới hạn số lượng.
        """
        focus_instruction = self._build_focus_instruction(cap_hoc, integration_type)
        
        text_lower = doc_text.lower()
        is_khgd_table = ("số tiết" in text_lower or "so tiet" in text_lower) and \
                        ("tuần" in text_lower or "tuan" in text_lower) and \
                        ("yêu cầu cần đạt" in text_lower)

        if is_khgd_table:
            context_guide = """
TÀI LIỆU LÀ BẢNG PHÂN PHỐI CHƯƠNG TRÌNH / KẾ HOẠCH GIÁO DỤC (GỒM CÁC CỘT: STT, BÀI HỌC, SỐ TIẾT, TUẦN, YÊU CẦU CẦN ĐẠT).
QUY TẮC QUAN TRỌNG:
1. Bạn PHẢI duyệt qua TẤT CẢ các bài học có trong bảng từ bài đầu tiên đến bài cuối cùng.
2. Với MỖI bài học, phải đề xuất nội dung tích hợp bổ sung vào cột 'Yêu cầu cần đạt'.
3. `anchor_text`: Trích chính xác TÊN BÀI HỌC trong bảng (ví dụ: "Bài 1. Đơn thức", "Bài 2. Đa thức", "Luyện tập chung").
4. `insert_content`: Nội dung năng lực bổ sung ngắn gọn, chuẩn sư phạm.
"""
        else:
            context_guide = """
TÀI LIỆU LÀ KẾ HOẠCH BÀI DẠY (GIÁO ÁN TIẾN TRÌNH LÊN LỚP, CÓ THỂ CÓ BẢNG HOẠT ĐỘNG GV - HS HOẶC NHIỀU BÀI DẠY).
QUY TẮC PHỦ KÍN TOÀN BỘ GIÁO ÁN (QUÉT SÂU - KHÔNG BỎ SÓT - KHÔNG GIỚI HẠN VỊ TRÍ):
1. Bạn PHẢI đọc và xử lý từ đầu đến dòng cuối cùng của tài liệu. Tuyệt đối không được chỉ dừng lại ở các bài đầu tiên hay chỉ lấy 4-5 vị trí.
2. Với MỖI bài dạy/tiết dạy xuất hiện trong file, PHẢI tích hợp rải đều qua các mục:
   - 1 vị trí ở phần "Mục tiêu / Năng lực"
   - 1 vị trí ở "Hoạt động 1: Khởi động" (câu hỏi, dẫn dắt, video số, AI gợi mở)
   - 1-2 vị trí ở "Hoạt động 2: Hình thành kiến thức" (thao tác số, thực hành công cụ, giải pháp STEM)
   - 1 vị trí ở "Hoạt động 3: Luyện tập" (bảng tính, đối chiếu kết quả với AI, bài tập thực hành)
   - 1 vị trí ở "Hoạt động 4: Vận dụng" (dự án STEM mini, bài toán thực tiễn)
3. `anchor_text`: Trích NGUYÊN VĂN một câu hoặc dòng lệnh có thật tại bước hoạt động đó (từ 6 - 15 từ, plain text, không thêm định dạng markdown).
4. KHÔNG chọn các anchor_text nằm sát nhau để tránh dồn ứ nội dung vào một chỗ.
"""

        prompt = f"""
Bạn là chuyên gia sư phạm và chuyển đổi số trong giáo dục phổ thông Việt Nam.
Hãy phân tích tài liệu dưới đây và xác định TẤT CẢ các vị trí tích hợp phù hợp từ đầu đến cuối tệp.

Cấp học chỉ định: {cap_hoc}

{context_guide}

{focus_instruction}

Nội dung tài liệu gốc:
----------------------------------
{doc_text}
----------------------------------
"""

        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=TichHopResult,
                    temperature=0.35,
                    max_output_tokens=8192
                )
            )
            return json.loads(response.text)
        except APIError as ae:
            raise RuntimeError(self._format_api_error(ae))
        except json.JSONDecodeError:
            raise RuntimeError("Dữ liệu phản hồi từ AI chưa đúng cấu trúc JSON. Vui lòng bấm thử lại.")
        except Exception as e:
            raise RuntimeError(self._format_api_error(e))

    def analyze_pptx_and_integrate(self, slides_text: str, cap_hoc: str, integration_type: str) -> dict:
        """Phân tích các slide PowerPoint và đề xuất lời nhắc sư phạm vào Slide Notes cho toàn bộ bài giảng."""
        focus_instruction = self._build_focus_instruction(cap_hoc, integration_type)

        prompt = f"""
Bạn là chuyên gia sư phạm và thiết kế bài giảng điện tử.
Hãy phân tích danh sách các Slide PowerPoint dưới đây để đưa ra các GHI CHÚ SƯ PHẠM vào phần Slide Notes (Ghi chú diễn giả cho giáo viên).

Cấp học chỉ định: {cap_hoc}

{focus_instruction}

QUY TẮC QUÉT TOÀN BỘ SLIDE:
1. Rà soát lần lượt từ Slide 1 đến Slide cuối cùng của tệp.
2. Với các Slide có hoạt động học (thảo luận, bài tập, câu hỏi, thực hành, vận dụng), hãy xác định chính xác `slide_number` (số nguyên) để đề xuất nội dung.
3. Nội dung `insert_content`: Lời nhắc sư phạm thiết thực cho GV khi đang trình chiếu (câu hỏi gợi mở, ứng dụng phần mềm số, dùng AI phản biện, hoạt động giải quyết vấn đề STEM).
4. Phủ đều khắp các slide của bài giảng, không dừng lại ở những slide đầu tiên.

Danh sách Slide:
----------------------------------
{slides_text}
----------------------------------
"""

        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=TichHopResult,
                    temperature=0.35,
                    max_output_tokens=8192
                )
            )
            return json.loads(response.text)
        except APIError as ae:
            raise RuntimeError(self._format_api_error(ae))
        except json.JSONDecodeError:
            raise RuntimeError("Dữ liệu phản hồi từ AI chưa đúng cấu trúc. Vui lòng bấm thử lại.")
        except Exception as e:
            raise RuntimeError(self._format_api_error(e))
