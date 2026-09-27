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
        """Làm sạch chuỗi: loại bỏ khoảng trắng thừa, dấu câu đặc biệt, đưa về chữ thường."""
        if not text:
            return ""
        text = text.replace('\xa0', ' ').replace('\t', ' ').replace('\r', '')
        # Bỏ dấu chấm, hai chấm, gạch ngang để so sánh tên bài chuẩn xác
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
        """Chèn đoạn văn mới liền sau đoạn văn paragraph chỉ định."""
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

        # Kiểm tra xem tài liệu có bảng Kế hoạch giáo dục không
        has_tables = len(doc.tables) > 0

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

            # --- TRƯỜNG HỢP 1: TÀI LIỆU DẠNG BẢNG (KẾ HOẠCH GIÁO DỤC) ---
            if has_tables:
                for table in doc.tables:
                    for row in table.rows:
                        # Xác định văn bản của toàn bộ hàng
                        row_cells_text = [WordProcessor._clean_str(c.text) for c in row.cells]
                        full_row_text = " ".join(row_cells_text)

                        # Kiểm tra xem dòng này có chứa tên bài học (hoặc độ tương đồng cao)
                        match_found = False
                        if clean_anchor in full_row_text:
                            match_found = True
                        else:
                            # So khớp gần đúng với ô tên bài học (thường là ô thứ 2 hoặc 1)
                            for cell_text in row_cells_text[:3]:
                                if len(cell_text) > 3 and SequenceMatcher(None, clean_anchor, cell_text).ratio() >= 0.65:
                                    match_found = True
                                    break

                        if match_found:
                            # Tìm đúng ô chứa cột "Yêu cầu cần đạt" (thường là ô có chữ 'kiến thức' hoặc 'năng lực' hoặc ô cuối)
                            target_cell = None
                            for cell in row.cells:
                                c_lower = cell.text.lower()
                                if "năng lực" in c_lower or "kiến thức" in c_lower or "phẩm chất" in c_lower:
                                    target_cell = cell
                                    break
                            
                            # Nếu không nhận diện được từ khóa, mặc định lấy ô dài nhất hoặc ô cuối cùng
                            if not target_cell:
                                target_cell = max(row.cells, key=lambda c: len(c.text))

                            # Chọn vị trí chèn trong ô: ưu tiên sau đoạn có chữ 'năng lực'
                            target_para = target_cell.paragraphs[-1]
                            for p in target_cell.paragraphs:
                                if "năng lực" in p.text.lower():
                                    target_para = p
                            
                            WordProcessor.insert_paragraph_after(target_para, content, color, prefix)
                            inserted = True
                            break
                    if inserted:
                        break

            # --- TRƯỜNG HỢP 2: NẾU CHƯA CHÈN ĐƯỢC HOẶC TÀI LIỆU LÀ KHBD VĂN BẢN THƯỜNG ---
            if not inserted:
                # Quét tất cả các đoạn văn trong tài liệu
                all_paragraphs = list(doc.paragraphs)
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                all_paragraphs.append(p)

                best_match_para = None
                best_ratio = 0.0

                for para in all_paragraphs:
                    clean_p = WordProcessor._clean_str(para.text)
                    if not clean_p:
                        continue
                    
                    if clean_anchor in clean_p or clean_p in clean_anchor:
                        WordProcessor.insert_paragraph_after(para, content, color, prefix)
                        inserted = True
                        break

                    ratio = SequenceMatcher(None, clean_anchor, clean_p).ratio()
                    if ratio > best_ratio:
                        best_ratio = ratio
                        best_match_para = para

                if not inserted and best_match_para is not None and best_ratio >= 0.55:
                    WordProcessor.insert_paragraph_after(best_match_para, content, color, prefix)
                    inserted = True

        output_stream = io.BytesIO()
        doc.save(output_stream)
        output_stream.seek(0)
        return output_stream
