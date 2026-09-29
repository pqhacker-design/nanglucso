import io
import re
from difflib import SequenceMatcher
from docx import Document
from docx.shared import RGBColor, Pt
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

class WordProcessor:
    @staticmethod
    def _clean_str(text: str) -> str:
        if not text:
            return ""
        text = text.replace('\xa0', ' ').replace('\t', ' ').replace('\r', '')
        text = re.sub(r'[*_#`.:\-,]', ' ', text)
        return re.sub(r'\s+', ' ', text).strip().lower()

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

        # Thu thập TẤT CẢ paragraphs (cả đoạn văn ngoài lẫn đoạn văn nằm trong ô bảng)
        all_doc_paras = list(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if p.text.strip():
                            all_doc_paras.append(p)

        # Lưu vết các đoạn đã chèn để không chèn trùng
        used_paras = set()

        for item in sua_doi_list:
            raw_anchor = item.get('anchor_text', '').strip()
            content = item.get('insert_content', '').strip()
            loai = item.get('loai', 'Năng lực số')

            if not content or not raw_anchor:
                continue

            clean_anchor = WordProcessor._clean_str(raw_anchor)

            if loai == "Giáo dục STEM":
                prefix, color = "[Giáo dục STEM]:", color_stem
            elif loai == "Năng lực AI":
                prefix, color = "[Năng lực AI]:", color_ai
            else:
                prefix, color = "[Năng lực số]:", color_digital

            inserted = False
            best_match_para = None
            best_ratio = 0.0

            # 1. Tìm chính xác đoạn văn chứa câu neo (anchor)
            for para in all_doc_paras:
                clean_p = WordProcessor._clean_str(para.text)
                if not clean_p:
                    continue

                # Nếu câu neo nằm trong đoạn văn này
                if clean_anchor in clean_p and para not in used_paras:
                    WordProcessor.insert_paragraph_after(para, content, color, prefix)
                    used_paras.add(para)
                    inserted = True
                    break

                # Tính độ tương đồng phòng ngừa sai khác khoảng trắng/dấu câu
                ratio = SequenceMatcher(None, clean_anchor, clean_p).ratio()
                if ratio > best_ratio and para not in used_paras:
                    best_ratio = ratio
                    best_match_para = para

            # 2. Nếu không khớp 100%, dùng đoạn văn có độ tương đồng cao nhất (>= 60%)
            if not inserted and best_match_para is not None and best_ratio >= 0.60:
                WordProcessor.insert_paragraph_after(best_match_para, content, color, prefix)
                used_paras.add(best_match_para)
                inserted = True

        output_stream = io.BytesIO()
        doc.save(output_stream)
        output_stream.seek(0)
        return output_stream
