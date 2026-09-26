import io
import re
from docx import Document
from docx.shared import RGBColor, Pt
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

class WordProcessor:
    @staticmethod
    def extract_text(file_bytes: bytes) -> str:
        doc = Document(io.BytesIO(file_bytes))
        full_text = []
        for table in doc.tables:
            for row in table.rows:
                cells_text = [c.text.strip().replace('\n', ' ') for c in row.cells if c.text.strip()]
                if cells_text:
                    full_text.append(" | ".join(cells_text))
        for para in doc.paragraphs:
            if para.text.strip():
                full_text.append(para.text.strip())
        return "\n".join(full_text)

    @staticmethod
    def insert_paragraph_after(paragraph, text, color_rgb, prefix=""):
        new_p = OxmlElement('w:p')
        paragraph._p.addnext(new_p)
        new_para = Paragraph(new_p, paragraph._parent)
        new_para.paragraph_format.space_before = Pt(3)
        new_para.paragraph_format.space_after = Pt(3)
        new_para.paragraph_format.line_spacing = 1.15
        
        if prefix:
            run_p = new_para.add_run(f"{prefix} ")
            run_p.font.color.rgb = color_rgb
            run_p.bold = True
            run_p.italic = True
            run_p.font.size = Pt(11)

        run_t = new_para.add_run(text)
        run_t.font.color.rgb = color_rgb
        run_t.italic = True
        run_t.font.size = Pt(11)
        return new_para

    @staticmethod
    def integrate_digital_capacity(file_bytes: bytes, ai_data: dict, integration_type: str) -> io.BytesIO:
        doc = Document(io.BytesIO(file_bytes))
        sua_doi_list = ai_data.get('sua_doi', [])
        
        color_digital = RGBColor(0, 102, 204)  # Xanh dương
        color_ai = RGBColor(214, 107, 0)       # Cam
        color_stem = RGBColor(16, 124, 65)     # Xanh lá

        for item in sua_doi_list:
            lesson_name = item.get('anchor_text', '').strip()
            content = item.get('insert_content', '').strip()
            loai = item.get('loai', 'Năng lực số')

            if not content:
                continue

            if loai == "Giáo dục STEM":
                prefix, color = "[Giáo dục STEM]:", color_stem
            elif loai == "Năng lực AI":
                prefix, color = "[Năng lực AI]:", color_ai
            else:
                prefix, prefix_color = "[Năng lực số]:", color_digital
                color = color_digital

            # Quét các bảng trong tệp Kế hoạch giáo dục
            for table in doc.tables:
                for row in table.rows:
                    row_text = " ".join([c.text for c in row.cells]).lower()
                    # Kiểm tra xem dòng này có chứa đúng tên Bài học cần tích hợp không
                    if lesson_name.lower() in row_text:
                        # Cột Yêu cầu cần đạt thường nằm ở ô cuối cùng của hàng
                        target_cell = row.cells[-1]
                        
                        # Tìm vị trí thích hợp trong ô: ưu tiên sau mục 2. Năng lực hoặc cuối ô
                        target_para = target_cell.paragraphs[-1]
                        for p in target_cell.paragraphs:
                            if "năng lực" in p.text.lower():
                                target_para = p
                        
                        WordProcessor.insert_paragraph_after(target_para, content, color, prefix)
                        break

        output_stream = io.BytesIO()
        doc.save(output_stream)
        output_stream.seek(0)
        return output_stream
