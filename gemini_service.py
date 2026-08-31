import json
from pydantic import BaseModel, Field
from typing import List, Literal, Optional
from google import genai
from google.genai import types
from google.genai.errors import APIError

# Định nghĩa Schema cho kết quả trả về
class SuaDoiItem(BaseModel):
    slide_number: Optional[int] = Field(default=None, description="Số thứ tự của Slide trong file PowerPoint.")
    anchor_text: Optional[str] = Field(default="", description="Câu văn/dòng neo có thật trong giáo án Word.")
    insert_content: str = Field(description="Nội dung tích hợp ngắn gọn, chuẩn sư phạm.")
    loai: Literal["Năng lực số", "Năng lực AI"] = Field(description="Loại năng lực: 'Năng lực số' hoặc 'Năng lực AI'.")

class TichHopResult(BaseModel):
    sua_doi: List[SuaDoiItem]


class GeminiService:
    def __init__(self, api_key: str):
        if not api_key:
            raise ValueError("API Key không được để trống.")
        self.client = genai.Client(api_key=api_key)
        self.model_name = "gemini-2.5-flash"

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
            "Tiểu học": "\n- Tiểu học: Thao tác đơn giản, tìm kiếm cơ bản, ý thức bảo vệ mắt, tư thế và an toàn riêng tư.",
            "THCS": "\n- THCS: Khai thác phần mềm môn học, đánh giá thông tin, làm việc nhóm trực tuyến an toàn, tôn trọng bản quyền.",
            "THPT": "\n- THPT: Phân tích dữ liệu nâng cao, sáng tạo sản phẩm số đa phương tiện, an toàn thông tin và giải quyết vấn đề thực tế.",
            "Tự động nhận diện": "\n- Tự nhận diện cấp học theo từng nội dung để tích hợp phù hợp."
        }
        return base_framework + level_guide.get(cap_hoc, level_guide["Tự động nhận diện"])

    def analyze_and_integrate(self, doc_text: str, cap_hoc: str, integration_type: str) -> dict:
        """Phân tích và tích hợp cho file giáo án Word (.docx)."""
        tt02_info = self._get_tt02_framework_prompt(cap_hoc)
        ai_framework_info = """
* Khung Năng lực AI (QĐ 2422/QĐ-BGDĐT):
  - Nhận thức về AI: Hiểu khái niệm cơ bản, nhận biết ứng dụng AI.
  - Ứng dụng AI: Dùng AI hỗ trợ tra cứu, gợi ý ý tưởng, tóm tắt, dịch thuật.
  - Tư duy phản biện & Đạo đức AI: Đánh giá độ tin cậy, trách nhiệm và tính trung thực khi dùng AI.
"""

        if integration_type == "Năng lực số":
            focus_instruction = f"YÊU CẦU: CHỈ TÍCH HỢP NĂNG LỰC SỐ (TT 02/2025/TT-BGDĐT).\n{tt02_info}\nTất cả các mục có 'loai': 'Năng lực số'."
        elif integration_type == "Năng lực AI":
            focus_instruction = f"YÊU CẦU: CHỈ TÍCH HỢP NĂNG LỰC AI (QĐ 2422/QĐ-BGDĐT).\n{ai_framework_info}\nTất cả các mục có 'loai': 'Năng lực AI'."
        else:
            focus_instruction = f"""
YÊU CẦU BẮT BUỘC KHI CHỌN 'CẢ HAI':
Phải tích hợp cả NĂNG LỰC SỐ (TT 02/2025/TT-BGDĐT) và NĂNG LỰC AI (QĐ 2422/QĐ-BGDĐT):
{tt02_info}
{ai_framework_info}
Danh sách `sua_doi` trả về PHẢI CÓ CẢ 2 LOẠI ('Năng lực số' và 'Năng lực AI').
"""

        prompt = f"""
Bạn là chuyên gia giáo dục và chuyển đổi số trong giáo dục phổ thông Việt Nam.
Hãy đọc toàn bộ tài liệu giáo án bên dưới và đề xuất các vị trí tích hợp.

Cấp học chỉ định: {cap_hoc}

{focus_instruction}

QUY TẮC QUAN TRỌNG VỀ anchor_text:
1. `anchor_text` PHẢI trích dẫn NGUYÊN VĂN một câu/dòng chữ có thật trong tài liệu giáo án (Plain text, không thêm dấu `**` hay markdown).
2. Trích đoạn dài từ 5 - 15 từ đặc trưng cho bài học để tìm kiếm chính xác vị trí.

Nội dung giáo án gốc:
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
                    temperature=0.3
                )
            )
            return json.loads(response.text)
        except APIError as ae:
            raise RuntimeError(f"Lỗi kết nối Gemini API: {str(ae)}")
        except json.JSONDecodeError:
            raise RuntimeError("Gemini phản hồi sai cấu trúc JSON.")
        except Exception as e:
            raise RuntimeError(f"Đã xảy ra lỗi: {str(e)}")

    def analyze_pptx_and_integrate(self, slides_text: str, cap_hoc: str, integration_type: str) -> dict:
        """Phân tích các slide PowerPoint và đề xuất nội dung ghi chú diễn giả (Slide Notes)."""
        tt02_info = self._get_tt02_framework_prompt(cap_hoc)
        ai_framework_info = """
* Khung Năng lực AI (QĐ 2422/QĐ-BGDĐT):
  - Nhận thức về AI: Nhận diện ứng dụng AI trong đời sống và môn học.
  - Ứng dụng AI: Dùng AI hỗ trợ gợi ý ý tưởng, tóm tắt, tra cứu thông tin, tạo minh họa.
  - Tư duy phản biện & Đạo đức AI: Đánh giá độ tin cậy kết quả của AI, tôn trọng bản quyền và sử dụng có trách nhiệm.
"""

        if integration_type == "Năng lực số":
            focus_instruction = f"YÊU CẦU: CHỈ TÍCH HỢP NĂNG LỰC SỐ (TT 02/2025/TT-BGDĐT).\n{tt02_info}\nTất cả các mục có 'loai': 'Năng lực số'."
        elif integration_type == "Năng lực AI":
            focus_instruction = f"YÊU CẦU: CHỈ TÍCH HỢP NĂNG LỰC AI (QĐ 2422/QĐ-BGDĐT).\n{ai_framework_info}\nTất cả các mục có 'loai': 'Năng lực AI'."
        else:
            focus_instruction = f"""
YÊU CẦU BẮT BUỘC KHI CHỌN 'CẢ HAI':
Phải tích hợp cả NĂNG LỰC SỐ (TT 02/2025/TT-BGDĐT) và NĂNG LỰC AI (QĐ 2422/QĐ-BGDĐT).
Danh sách trả về phải có sự phân bổ cả 2 loại ('Năng lực số' và 'Năng lực AI') phù hợp với từng slide.
"""

        prompt = f"""
Bạn là chuyên gia sư phạm và chuyển đổi số trong giáo dục phổ thông Việt Nam.
Hãy đọc danh sách các Slide bài giảng PowerPoint dưới đây. Hãy phân tích nội dung từng slide và đề xuất các GHI CHÚ SƯ PHẠM (để đưa vào phần Slide Notes cho giáo viên) nhằm tích hợp Năng lực số / Năng lực AI vào hoạt động dạy học của slide đó.

Cấp học chỉ định: {cap_hoc}

{focus_instruction}

QUY TẮC QUAN TRỌNG:
1. Xác định chính xác `slide_number` (số nguyên) của slide cần tích hợp.
2. Nội dung `insert_content` là lời nhắc/hướng dẫn sư phạm ngắn gọn, thiết thực cho giáo viên (ví dụ: "GV hướng dẫn HS tra cứu dữ liệu...", "GV nhắc HS dùng AI tạo gợi ý nhưng cần đối chiếu SGK...", "Tổ chức cho HS nộp sản phẩm số qua link nhóm...").
3. Không cần tích hợp trên tất cả mọi slide, chỉ chọn những slide hoạt động trọng tâm, slide thảo luận hoặc bài tập.

Danh sách nội dung các Slide:
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
                    temperature=0.3
                )
            )
            return json.loads(response.text)
        except APIError as ae:
            raise RuntimeError(f"Lỗi kết nối Gemini API: {str(ae)}")
        except json.JSONDecodeError:
            raise RuntimeError("Gemini phản hồi sai cấu trúc JSON.")
        except Exception as e:
            raise RuntimeError(f"Đã xảy ra lỗi: {str(e)}")
