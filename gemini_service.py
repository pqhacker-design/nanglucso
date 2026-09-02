import json
from pydantic import BaseModel, Field
from typing import List, Literal, Optional
from google import genai
from google.genai import types
from google.genai.errors import APIError

class SuaDoiItem(BaseModel):
    slide_number: Optional[int] = Field(default=None, description="Số thứ tự Slide PowerPoint.")
    anchor_text: Optional[str] = Field(default="", description="Câu văn/dòng neo có thật trong giáo án Word.")
    insert_content: str = Field(description="Nội dung tích hợp ngắn gọn, chuẩn sư phạm.")
    loai: Literal["Năng lực số", "Năng lực AI"] = Field(description="Loại năng lực: 'Năng lực số' hoặc 'Năng lực AI'.")

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
            "Tiểu học": "\n- Tiểu học: Thao tác đơn giản, tìm kiếm cơ bản, ý thức bảo vệ mắt, tư thế và an toàn thông tin cá nhân.",
            "THCS": "\n- THCS: Sử dụng phần mềm học tập/mô phỏng, xử lý số liệu, làm việc nhóm trực tuyến an toàn, tôn trọng bản quyền số.",
            "THPT": "\n- THPT: Phân tích dữ liệu chuyên sâu, sáng tạo sản phẩm số tương tác, tư duy máy tính và pháp luật số.",
            "Tự động nhận diện": "\n- Tự động nhận diện lớp/cấp học theo từng bài để tích hợp vừa sức với đối tượng học sinh."
        }
        return base_framework + level_guide.get(cap_hoc, level_guide["Tự động nhận diện"])

    def analyze_and_integrate(self, doc_text: str, cap_hoc: str, integration_type: str) -> dict:
        tt02_info = self._get_tt02_framework_prompt(cap_hoc)
        ai_framework_info = """
* Khung Năng lực AI (QĐ 2422/QĐ-BGDĐT):
  - Nhận thức về AI: Nhận biết AI trong học tập và đời sống, hiểu nguyên lý cơ bản.
  - Ứng dụng AI: Sử dụng AI để tra cứu, tóm tắt, tìm ý tưởng, hỗ trợ làm bài tập, dịch thuật, mô phỏng.
  - Tư duy phản biện & Đạo đức AI: Đánh giá độ tin cậy kết quả của AI, kiểm chứng nguồn tin, chống gian lận học thuật.
"""

        if integration_type == "Năng lực số":
            focus_instruction = f"YÊU CẦU: TÍCH HỢP NĂNG LỰC SỐ (TT 02/2025/TT-BGDĐT).\n{tt02_info}\nTất cả các mục có 'loai': 'Năng lực số'."
        elif integration_type == "Năng lực AI":
            focus_instruction = f"YÊU CẦU: TÍCH HỢP NĂNG LỰC AI (QĐ 2422/QĐ-BGDĐT).\n{ai_framework_info}\nTất cả các mục có 'loai': 'Năng lực AI'."
        else:
            focus_instruction = f"""
YÊU CẦU BẮT BUỘC KHI CHỌN 'CẢ HAI' (QUÉT SÂU TOÀN DIỆN):
Bạn PHẢI tích hợp đồng thời cả NĂNG LỰC SỐ (TT 02/2025/TT-BGDĐT) và NĂNG LỰC AI (QĐ 2422/QĐ-BGDĐT).
{tt02_info}
{ai_framework_info}

QUY TẮC PHÂN BỔ BẮT BUỘC ĐỂ KHÔNG BỎ SÓT VỊ TRÍ:
1. Đối với MỤC TIÊU BÀI DẠY: Bắt buộc đề xuất TỐI THIỂU 1 chỉ tiêu Năng lực số VÀ 1 chỉ tiêu Năng lực AI.
2. Đối với TIẾN TRÌNH HOẠT ĐỘNG: Phải rà soát lần lượt TỪNG HOẠT ĐỘNG:
   - Hoạt động Khởi động: Tích hợp công cụ số (trò chơi trắc nghiệm/video) hoặc AI (chatbot gợi mở bài toán).
   - Hoạt động Hình thành kiến thức: Tích hợp phần mềm mô phỏng/tra cứu số liệu hoặc ứng dụng AI tóm tắt/giải thích khái niệm.
   - Hoạt động Luyện tập: Tích hợp bảng tính, vẽ đồ thị, nộp bài số hóa hoặc đối chiếu bài làm với AI để phản biện lỗi sai.
   - Hoạt động Vận dụng: Tích hợp sáng tạo sản phẩm số (infographic/video ngắn) hoặc khai thác AI mở rộng liên hệ thực tiễn.
3. Số lượng đề xuất: Với mỗi bài học, hãy tìm ít nhất từ 4 đến 8 vị trí tích hợp phù hợp, cân bằng giữa hai loại.
"""

        prompt = f"""
Bạn là chuyên gia sư phạm và chuyển đổi số trong giáo dục Việt Nam.
Hãy đọc KỸ LƯỠNG và PHÂN TÍCH TỪNG DÒNG của tài liệu giáo án dưới đây (tài liệu có thể gồm nhiều bài).

Cấp học: {cap_hoc}

{focus_instruction}

QUY TẮC ANCHOR TEXT:
- `anchor_text` PHẢI là câu văn nguyên văn (plain text, không thêm định dạng markdown) có sẵn trong giáo án.
- Chọn cụm từ dài 6 - 15 từ mang ngữ cảnh riêng biệt của từng bài/hoạt động để không bị nhầm lẫn giữa các vị trí.

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
            raise RuntimeError("Dữ liệu phản hồi chưa đúng cấu trúc. Vui lòng thử lại.")
        except Exception as e:
            raise RuntimeError(self._format_api_error(e))

    def analyze_pptx_and_integrate(self, slides_text: str, cap_hoc: str, integration_type: str) -> dict:
        tt02_info = self._get_tt02_framework_prompt(cap_hoc)
        ai_framework_info = """
* Khung Năng lực AI (QĐ 2422/QĐ-BGDĐT):
  - Nhận thức AI: Nhận diện công nghệ AI liên quan đến chủ đề bài học.
  - Ứng dụng AI: Dùng AI gợi mở câu hỏi, tìm ý tưởng, tóm tắt bài.
  - Phản biện & Đạo đức: Nhắc nhở HS kiểm chứng độ chính xác của AI.
"""

        if integration_type == "Năng lực số":
            focus_instruction = f"TÍCH HỢP NĂNG LỰC SỐ (TT 02/2025/TT-BGDĐT):\n{tt02_info}"
        elif integration_type == "Năng lực AI":
            focus_instruction = f"TÍCH HỢP NĂNG LỰC AI (QĐ 2422/QĐ-BGDĐT):\n{ai_framework_info}"
        else:
            focus_instruction = f"""
TÍCH HỢP ĐỒNG THỜI CẢ HAI NĂNG LỰC (QUÉT SÂU MỌI SLIDE):
{tt02_info}
{ai_framework_info}
- Hãy rà soát toàn bộ các slide, tìm ít nhất từ 4 đến 8 slide phù hợp để đưa ghi chú diễn giả.
- Slide bài học kiến thức: Đề xuất công cụ số hoặc câu hỏi tư duy AI.
- Slide bài tập/thực hành: Hướng dẫn HS dùng AI/phần mềm để kiểm tra kết quả và phản biện.
"""

        prompt = f"""
Bạn là chuyên gia sư phạm và bài giảng điện tử.
Hãy phân tích danh sách các Slide bài giảng PowerPoint dưới đây để đưa ra các gợi ý sư phạm vào phần Slide Notes.

Cấp học: {cap_hoc}

{focus_instruction}

QUY TẮC:
- Trả về danh sách với `slide_number` cụ thể.
- Đề xuất lời nhắc sư phạm chi tiết, thực tế cho giáo viên khi đang trình chiếu slide đó.

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
