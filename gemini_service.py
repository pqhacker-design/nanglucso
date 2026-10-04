import json
from pydantic import BaseModel, Field
from typing import List, Literal, Optional
from google import genai
from google.genai import types
from google.genai.errors import APIError

class SuaDoiItem(BaseModel):
    slide_number: Optional[int] = Field(default=None, description="Số thứ tự Slide PowerPoint.")
    anchor_text: Optional[str] = Field(default="", description="Câu văn/dòng neo có thật trong giáo án Word.")
    insert_content: str = Field(description="Nội dung tích hợp ngắn gọn, chuẩn thuật ngữ sư phạm.")
    loai: Literal["Năng lực số", "Năng lực AI", "Giáo dục STEM"] = Field(
        description="Loại tích hợp: 'Năng lực số', 'Năng lực AI' hoặc 'Giáo dục STEM'."
    )

class TichHopResult(BaseModel):
    sua_doi: List[SuaDoiItem]


class GeminiService:
    def __init__(self, api_key: str):
        if not api_key:
            raise ValueError("Mã API Key không được để trống. Vui lòng nhập API Key.")
        self.client = genai.Client(api_key=api_key)
        self.model_name = "gemini-2.5-flash"

    @staticmethod
    def _format_api_error(error: Exception) -> str:
        err_str = str(error)
        if "API_KEY_INVALID" in err_str or "API key not valid" in err_str or "PERMISSION_DENIED" in err_str:
            return "Mã API Key không chính xác hoặc đã bị vô hiệu hóa. Vui lòng kiểm tra lại trên Google AI Studio."
        if "RESOURCE_EXHAUSTED" in err_str or "429" in err_str or "quota" in err_str.lower():
            return "Đã đạt giới hạn yêu cầu của Google. Vui lòng đợi 30 - 60 giây rồi thử lại."
        if "INVALID_ARGUMENT" in err_str:
            return "Nội dung tệp gửi đi có định dạng chưa phù hợp."
        if "DEADLINE_EXCEEDED" in err_str or "timeout" in err_str.lower():
            return "Thời gian xử lý quá lâu do đường truyền Internet hoặc tệp quá dài."
        if "UNAVAILABLE" in err_str or "503" in err_str:
            return "Máy chủ Google Gemini đang bảo trì hoặc quá tải tạm thời."
        return f"Lỗi từ Google: {err_str}"

    def _get_tt02_framework_prompt(self, cap_hoc: str) -> str:
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
            "Tiểu học": "\n- Tiểu học: Thao tác đơn giản, tìm kiếm cơ bản, ý thức an toàn thiết bị, bảo vệ mắt và dữ liệu cá nhân.",
            "THCS": "\n- THCS: Sử dụng ứng dụng môn học/mô phỏng, xử lý số liệu bảng tính, làm việc nhóm trực tuyến an toàn, tôn trọng bản quyền.",
            "THPT": "\n- THPT: Phân tích số liệu chuyên sâu, thiết kế sản phẩm số đa phương tiện, an ninh mạng và pháp luật số.",
            "Tự động nhận diện": "\n- Tự động nhận diện cấp học theo nội dung bài dạy để tích hợp vừa sức học sinh."
        }
        return base_framework + level_guide.get(cap_hoc, level_guide["Tự động nhận diện"])

    def _get_stem_framework_prompt(self, cap_hoc: str) -> str:
        stem_base = """
* Khung Giáo dục STEM (Công văn 3089/BGDĐT-GDTrH và Công văn 909/BGDĐT-GDTH):
  - Định hướng: Dạy học giải quyết vấn đề thực tiễn gắn liền với liên môn Khoa học (S), Công nghệ (T), Kỹ thuật (E), Toán học (M).
  - Quy trình kỹ thuật 5 bước cốt lõi:
    1. Xác định vấn đề/tiêu chí sản phẩm thực tiễn.
    2. Nghiên cứu kiến thức nền tảng và đề xuất giải pháp.
    3. Lựa chọn giải pháp thiết kế khả thi.
    4. Chế tạo mẫu, thử nghiệm và đánh giá chất lượng.
    5. Chia sẻ, thảo luận và cải tiến sản phẩm.
"""
        level_stem = {
            "Tiểu học": "\n- Tiểu học: Trải nghiệm thực hành đơn giản từ vật liệu tái chế, lắp ghép mô hình trực quan gần gũi.",
            "THCS": "\n- THCS: Chế tạo dụng cụ học tập/thí nghiệm, giải quyết bài toán môi trường, ứng dụng nguyên lý khoa học và kỹ thuật đơn giản.",
            "THPT": "\n- THPT: Thiết kế dự án nghiên cứu, lập trình mô phỏng, phân tích tối ưu hóa kỹ thuật và chi phí thực hiện.",
            "Tự động nhận diện": "\n- Điều chỉnh độ phức tạp của bài toán kỹ thuật theo đúng đối tượng học sinh."
        }
        return stem_base + level_stem.get(cap_hoc, level_stem["Tự động nhận diện"])

    def _build_focus_instruction(self, cap_hoc: str, integration_type: str) -> str:
        tt02_info = self._get_tt02_framework_prompt(cap_hoc)
        stem_info = self._get_stem_framework_prompt(cap_hoc)
        ai_framework_info = """
* Khung Năng lực AI (QĐ 2422/QĐ-BGDĐT):
  - Nhận thức AI: Hiểu khái niệm cơ bản, nhận diện công nghệ AI quanh ta.
  - Ứng dụng AI: Sử dụng AI để tra cứu, tóm tắt, tìm ý tưởng, hỗ trợ làm bài tập.
  - Tư duy phản biện & Đạo đức AI: Đánh giá độ tin cậy kết quả của AI, chống gian lận học thuật.
"""
        if integration_type == "Năng lực số":
            return f"YÊU CẦU: CHỈ TÍCH HỢP NĂNG LỰC SỐ (TT 02/2025/TT-BGDĐT).\n{tt02_info}\nTất cả các mục có 'loai': 'Năng lực số'."
        elif integration_type == "Năng lực AI":
            return f"YÊU CẦU: CHỈ TÍCH HỢP NĂNG LỰC AI (QĐ 2422/QĐ-BGDĐT).\n{ai_framework_info}\nTất cả các mục có 'loai': 'Năng lực AI'."
        elif integration_type == "Giáo dục STEM":
            return f"YÊU CẦU: CHỈ TÍCH HỢP BÀI HỌC/HOẠT ĐỘNG STEM (CV 3089/BGDĐT & CV 909/BGDĐT).\n{stem_info}\nTất cả các mục có 'loai': 'Giáo dục STEM'."
        else:  # "Tất cả"
            return f"""
YÊU CẦU BẮT BUỘC KHI CHỌN 'TẤT CẢ' (QUÉT SÂU TOÀN DIỆN):
Tích hợp đồng thời:
1. Năng lực số (TT 02/2025/TT-BGDĐT): {tt02_info}
2. Năng lực AI (QĐ 2422/QĐ-BGDĐT): {ai_framework_info}
3. Hoạt động/Bài học STEM (CV 3089/BGDĐT, CV 909/BGDĐT): {stem_info}

Quy tắc:
- Mục Tiêu: Bổ sung chỉ tiêu về Năng lực số/AI và phẩm chất giải quyết vấn đề STEM.
- Tiến trình học: Hoạt động Khởi động/Hình thành kiến thức/Luyện tập/Vận dụng phải phân bổ hài hòa cả 3 nhóm 'Năng lực số', 'Năng lực AI', 'Giáo dục STEM'.
"""

    def analyze_and_integrate(self, doc_text: str, cap_hoc: str, integration_type: str) -> dict:
        focus_instruction = self._build_focus_instruction(cap_hoc, integration_type)

        prompt = f"""
Bạn là chuyên gia sư phạm và chuyển đổi số trong giáo dục Việt Nam.
Hãy đọc KỸ LƯỠNG và PHÂN TÍCH TỪNG DÒNG của tài liệu giáo án dưới đây để tìm các vị trí tích hợp phù hợp.

Cấp học chỉ định: {cap_hoc}

{focus_instruction}

QUY TẮC ANCHOR TEXT:
- `anchor_text` PHẢI là câu văn nguyên văn (plain text, không thêm dấu `**` hay markdown) có trong giáo án.
- Chọn cụm từ dài 6 - 15 từ mang ngữ cảnh riêng biệt của từng bài/hoạt động.
- Trả về từ 4 đến 8 vị trí tích hợp phân bổ đều theo tiến trình bài học.

Nội dung giáo án:
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
                    temperature=0.45
                )
            )
            return json.loads(response.text)
        except APIError as ae:
            raise RuntimeError(self._format_api_error(ae))
        except json.JSONDecodeError:
            raise RuntimeError("Dữ liệu phản hồi chưa đúng cấu trúc JSON. Vui lòng thử lại.")
        except Exception as e:
            raise RuntimeError(self._format_api_error(e))

    def analyze_pptx_and_integrate(self, slides_text: str, cap_hoc: str, integration_type: str) -> dict:
        focus_instruction = self._build_focus_instruction(cap_hoc, integration_type)

        prompt = f"""
Bạn là chuyên gia sư phạm và bài giảng điện tử.
Hãy phân tích danh sách các Slide PowerPoint dưới đây để đưa các gợi ý sư phạm vào phần Slide Notes.

Cấp học chỉ định: {cap_hoc}

{focus_instruction}

QUY TẮC:
- Trả về danh sách kèm `slide_number` cụ thể.
- Lời nhắc sư phạm thực tế cho GV khi đang trình chiếu (ví dụ: đặt câu hỏi nghiên cứu giải pháp STEM, hướng dẫn tra cứu số, dùng AI phản biện).
- Chọn từ 4 đến 8 slide hoạt động trọng tâm để tích hợp.

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
                    temperature=0.45
                )
            )
            return json.loads(response.text)
        except APIError as ae:
            raise RuntimeError(self._format_api_error(ae))
        except json.JSONDecodeError:
            raise RuntimeError("Dữ liệu phản hồi chưa đúng cấu trúc. Vui lòng thử lại.")
        except Exception as e:
            raise RuntimeError(self._format_api_error(e))
