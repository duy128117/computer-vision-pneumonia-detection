# -*- coding: utf-8 -*-
"""
Script to generate the complete Vietnamese mid-term report in Word format (.docx).
Output: d:\CODE\tgmt\baocao_giuaki_viemphoi.docx
"""

import os
import sys
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

OUTPUT_PATH = r"d:\CODE\tgmt\baocao_giuaki_viemphoi.docx"
RESULTS_DIR = r"d:\CODE\tgmt\results"

def create_report():
    doc = Document()

    # 1. Page Setup (A4, standard margin: Left 3cm, Right 2cm, Top 2cm, Bottom 2cm)
    for section in doc.sections:
        section.page_width = Pt(595.3)   # 210mm
        section.page_height = Pt(841.9)  # 297mm
        section.left_margin = Pt(85.05)  # 30mm
        section.right_margin = Pt(56.7)  # 20mm
        section.top_margin = Pt(56.7)    # 20mm
        section.bottom_margin = Pt(56.7) # 20mm

    # Normal Style configuration
    style_normal = doc.styles['Normal']
    font_normal = style_normal.font
    font_normal.name = 'Times New Roman'
    font_normal.size = Pt(13)
    font_normal.color.rgb = RGBColor(0x1F, 0x24, 0x21) # dark charcoal

    # Helpers
    def set_cell_background(cell, fill_hex):
        shading_xml = f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'
        cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

    def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    def set_table_borders(table, color="B0BEC5", sz="4"):
        tblPr = table._tbl.tblPr
        borders_xml = f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:left w:val="none"/>
            <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:right w:val="none"/>
            <w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="none"/>
        </w:tblBorders>
        '''
        tblPr.append(parse_xml(borders_xml))

    def add_para(text="", bold=False, italic=False, size=13, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=6, color=None, font_name="Times New Roman"):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.25
        if text:
            r = p.add_run(text)
            r.font.name = font_name
            r.font.size = Pt(size)
            r.bold = bold
            r.italic = italic
            if color:
                r.font.color.rgb = color
        return p

    def add_h1(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(15)
        r.bold = True
        r.font.color.rgb = RGBColor(0x00, 0x33, 0x66) # Dark Navy
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(13.5)
        r.bold = True
        r.font.color.rgb = RGBColor(0x0A, 0x4B, 0x78)
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(13)
        r.bold = True
        r.italic = True
        r.font.color.rgb = RGBColor(0x1B, 0x3A, 0x4B)
        return p

    def add_bullet(text, bold_prefix="", level=0):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.2
        if bold_prefix:
            r1 = p.add_run(bold_prefix)
            r1.font.name = 'Times New Roman'
            r1.font.size = Pt(13)
            r1.bold = True
        r2 = p.add_run(text)
        r2.font.name = 'Times New Roman'
        r2.font.size = Pt(13)
        return p

    def add_image_safe(img_path, caption, width=Inches(5.5)):
        if os.path.exists(img_path):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(4)
            p.add_run().add_picture(img_path, width=width)
            
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(0)
            p_cap.paragraph_format.space_after = Pt(10)
            r = p_cap.add_run(f"Hình: {caption}")
            r.font.name = 'Times New Roman'
            r.font.size = Pt(11)
            r.italic = True
            r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    def style_table(tbl, col_widths, headers, data):
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(tbl, "CCCCCC", "6")

        # Header
        hdr_cells = tbl.rows[0].cells
        for i, h in enumerate(headers):
            hdr_cells[i].text = ""
            p = hdr_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            r = p.add_run(h)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(11.5)
            r.bold = True
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            set_cell_background(hdr_cells[i], "003366")
            set_cell_margins(hdr_cells[i], 120, 120, 140, 140)

        # Rows
        for row_idx, row_data in enumerate(data):
            row_cells = tbl.rows[row_idx + 1].cells
            bg_color = "F7F9FB" if row_idx % 2 == 1 else "FFFFFF"
            for col_idx, cell_value in enumerate(row_data):
                row_cells[col_idx].text = ""
                p = row_cells[col_idx].paragraphs[0]
                if any(char.isdigit() for char in cell_value) and len(cell_value) < 15:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)
                r = p.add_run(cell_value)
                r.font.name = 'Times New Roman'
                r.font.size = Pt(11)
                if col_idx == 0 or "ResNet50" in cell_value:
                    r.bold = True
                set_cell_background(row_cells[col_idx], bg_color)
                set_cell_margins(row_cells[col_idx], 80, 80, 120, 120)

        # Set column widths
        for row in tbl.rows:
            for idx, width in enumerate(col_widths):
                row.cells[idx].width = width

    # ==========================================
    # 1. TRANG BÌA (COVER PAGE)
    # ==========================================
    add_para("TRƯỜNG ĐẠI HỌC BÁCH KHOA - ĐẠI HỌC ĐÀ NẴNG", bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=10, space_after=3)
    add_para("KHOA CÔNG NGHỆ THÔNG TIN", bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=20, color=RGBColor(0x00, 0x33, 0x66))

    # Decorative line
    p_line = doc.add_paragraph()
    p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_line = p_line.add_run("______________________________")
    r_line.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    p_line.paragraph_format.space_after = Pt(30)

    add_para("BÁO CÁO GIỮA KỲ HỌC PHẦN", bold=True, size=16, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=10, space_after=5, color=RGBColor(0x00, 0x33, 0x66))
    add_para("XỬ LÝ ẢNH & THỊ GIÁC MÁY TÍNH", bold=True, size=18, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=35, color=RGBColor(0xC0, 0x39, 0x2B))

    add_para("ĐỀ TÀI:", bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=10, space_after=8)
    add_para("PHÂN LOẠI VIÊM PHỔI TỪ ẢNH X-QUANG NGỰC BẰNG DEEP LEARNING VÀ GIẢI THÍCH KẾT QUẢ BẰNG GRAD-CAM", bold=True, size=16, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=8, color=RGBColor(0x00, 0x2B, 0x49))
    add_para("(Pneumonia Classification from Chest X-Ray Images using Transfer Learning and Grad-CAM Explainability)", italic=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=45, color=RGBColor(0x55, 0x55, 0x55))

    # Box thông tin SV & GVHD
    table_cover = doc.add_table(rows=3, cols=2)
    table_cover.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_cover.rows[0].cells[0].paragraphs[0].add_run("Giảng viên hướng dẫn:").bold = True
    table_cover.rows[0].cells[1].paragraphs[0].add_run("TS. Huỳnh Hữu Hưng")
    table_cover.rows[1].cells[0].paragraphs[0].add_run("Sinh viên thực hiện:").bold = True
    table_cover.rows[1].cells[1].paragraphs[0].add_run("Mai Phước Thịnh\nNguyễn Đình Duy")
    table_cover.rows[2].cells[0].paragraphs[0].add_run("Lớp học phần:").bold = True
    table_cover.rows[2].cells[1].paragraphs[0].add_run("Xử lý ảnh - Nhóm 11")
    for r in table_cover.rows:
        for c in r.cells:
            for p in c.paragraphs:
                p.paragraph_format.line_spacing = 1.2
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(12.5)

    add_para("Đà Nẵng, năm 2026", italic=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=60, space_after=0)
    doc.add_page_break()

    # ==========================================
    # 2. LỜI NÓI ĐẦU
    # ==========================================
    add_h1("LỜI NÓI ĐẦU")
    add_para("Trong bối cảnh trí tuệ nhân tạo (AI) đang phát triển mạnh mẽ, các hệ thống hỗ trợ chẩn đoán y tế bằng hình ảnh ngày càng đóng vai trò quan trọng trong thực hành lâm sàng. Đặc biệt, trong lĩnh vực chẩn đoán hình ảnh, việc ứng dụng các kỹ thuật xử lý ảnh và học sâu để phân loại bệnh lý từ ảnh X-quang đã trở thành một hướng nghiên cứu trọng điểm, góp phần nâng cao hiệu quả sàng lọc và giảm tải cho bác sĩ chuyên khoa.")
    add_para("Viêm phổi là một trong những bệnh lý đường hô hấp phổ biến và nguy hiểm, đặc biệt ở trẻ em và người cao tuổi. Việc phát hiện sớm viêm phổi từ ảnh X-quang ngực đóng vai trò quyết định trong việc chỉ định phác đồ điều trị kịp thời. Tuy nhiên, việc đọc ảnh X-quang đòi hỏi bác sĩ chuyên khoa giàu kinh nghiệm và quá trình này có thể chịu ảnh hưởng bởi yếu tố mệt mỏi, chủ quan. Do đó, một hệ thống tự động hỗ trợ sàng lọc ảnh X-quang có tiềm năng ứng dụng rất lớn trong thực tiễn y tế.")
    add_para("Xuất phát từ những yêu cầu thực tiễn đó, nhóm chúng em lựa chọn đề tài \"Phân loại viêm phổi từ ảnh X-quang ngực bằng Deep Learning và giải thích kết quả bằng Grad-CAM\" với mục tiêu tìm hiểu, nghiên cứu và triển khai ba mô hình học sâu tiên tiến — ResNet50, DenseNet121 và InceptionV3 — trên bộ dữ liệu X-quang ngực nhi khoa công khai, đồng thời sử dụng Grad-CAM để trực quan hóa vùng ảnh mà mô hình tập trung khi đưa ra quyết định.")
    add_para("Trong quá trình thực hiện đề tài, mặc dù đã cố gắng hết sức, nhưng do kiến thức và thời gian còn hạn chế nên không tránh khỏi những thiếu sót. Nhóm chúng em rất mong nhận được sự góp ý quý báu từ thầy Huỳnh Hữu Hưng và các bạn để đề tài được hoàn thiện hơn.")
    add_para("Chúng em xin chân thành cảm ơn!", italic=True)
    doc.add_page_break()

    # ==========================================
    # 3. MỤC LỤC
    # ==========================================
    add_h1("MỤC LỤC")
    toc_items = [
        ("MỞ ĐẦU", "1"),
        ("  1. Lý do chọn đề tài", "1"),
        ("  2. Mục tiêu đề tài", "1"),
        ("  3. Phạm vi và đối tượng nghiên cứu", "2"),
        ("CHƯƠNG I. CƠ SỞ LÝ THUYẾT", "3"),
        ("  1.1. Tổng quan về phân loại ảnh y tế", "3"),
        ("  1.2. Bộ dữ liệu Chest X-Ray Pneumonia", "3"),
        ("  1.3. Cơ sở lý thuyết về xử lý ảnh và học sâu", "4"),
        ("    1.3.1. Chuẩn hóa ảnh (Image Normalization)", "4"),
        ("    1.3.2. Mạng nơ-ron tích chập (CNN)", "4"),
        ("    1.3.3. ResNet50 — Residual Learning", "5"),
        ("    1.3.4. DenseNet121 — Dense Connectivity", "5"),
        ("    1.3.5. InceptionV3 — Multi-Scale Convolutions", "6"),
        ("  1.4. Kỹ thuật Transfer Learning", "6"),
        ("  1.5. Grad-CAM — Giải thích quyết định mô hình", "7"),
        ("  1.6. Hàm mất mát và xử lý mất cân bằng dữ liệu", "7"),
        ("  1.7. Các chỉ số đánh giá", "8"),
        ("  1.8. Kết luận chương", "8"),
        ("CHƯƠNG II. PHƯƠNG PHÁP NGHIÊN CỨU", "9"),
        ("  2.1. Tổng quan pipeline hệ thống", "9"),
        ("  2.2. Cài đặt và chuẩn bị thực nghiệm", "9"),
        ("  2.3. Bộ dữ liệu và tiền xử lý chống rò rỉ", "10"),
        ("  2.4. Kiến trúc và cấu hình mô hình", "11"),
        ("  2.5. Chiến lược huấn luyện hai giai đoạn", "12"),
        ("  2.6. Đánh giá và kiểm định thống kê", "13"),
        ("CHƯƠNG III. KẾT QUẢ THỰC NGHIỆM VÀ THẢO LUẬN", "14"),
        ("  3.1. Phân bố dữ liệu sau làm sạch", "14"),
        ("  3.2. Kết quả chi tiết trên tập kiểm thử độc lập", "15"),
        ("  3.3. Kiểm định thống kê McNemar", "17"),
        ("  3.4. Phân tích trực quan hóa Grad-CAM", "18"),
        ("  3.5. Kết luận lựa chọn mô hình tối ưu", "20"),
        ("CHƯƠNG IV. KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN", "21"),
        ("  4.1. Đánh giá kết quả đạt được", "21"),
        ("  4.2. Hạn chế của đề tài", "21"),
        ("  4.3. Hướng phát triển tiếp theo", "22"),
        ("TÀI LIỆU THAM KHẢO", "23")
    ]
    for title, page in toc_items:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        r_title = p.add_run(title)
        r_title.font.name = 'Times New Roman'
        r_title.font.size = Pt(12)
        if not title.startswith(" "):
            r_title.bold = True
        
        # Dots
        dots_count = max(5, 75 - len(title) * 2)
        r_dots = p.add_run(" " + "." * dots_count + " ")
        r_dots.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
        
        r_page = p.add_run(page)
        r_page.font.name = 'Times New Roman'
        r_page.font.size = Pt(12)
        if not title.startswith(" "):
            r_page.bold = True

    doc.add_page_break()

    # ==========================================
    # 4. MỞ ĐẦU
    # ==========================================
    add_h1("MỞ ĐẦU")
    
    add_h2("1. Lý do chọn đề tài")
    add_para("Viêm phổi (Pneumonia) là bệnh nhiễm trùng đường hô hấp cấp tính gây viêm phế nang ở một hoặc cả hai bên phổi. Theo báo cáo của Tổ chức Y tế Thế giới (WHO), viêm phổi là nguyên nhân truyền nhiễm gây tử vong đơn lẻ lớn nhất ở trẻ em dưới 5 tuổi trên toàn cầu, cướp đi sinh mạng của hàng trăm nghìn trẻ em mỗi năm. Chẩn đoán viêm phổi từ ảnh chụp X-quang ngực (Chest X-Ray) là phương pháp hình ảnh chuẩn và nhanh chóng nhất hiện nay. Tuy nhiên, việc đọc ảnh đòi hỏi bác sĩ chuyên khoa chẩn đoán hình ảnh có chuyên môn cao và mất nhiều thời gian, đặc biệt tại các bệnh viện tuyến dưới hoặc các khu vực thiếu thốn nhân lực y tế.")
    add_para("Bộ dữ liệu Chest X-Ray Images (Pneumonia) của Kermany et al. (2018) là benchmark phổ biến nhất trong cộng đồng nghiên cứu thị giác máy tính, cung cấp ảnh X-quang ngực nhi khoa thực tế với hai nhãn NORMAL (bình thường) và PNEUMONIA (viêm phổi). Đây là cơ sở thực nghiệm lý tưởng để nghiên cứu giải pháp tự động hóa sàng lọc bệnh lý.")
    add_para("Tuy nhiên, một rào cản lớn khi áp dụng mạng học sâu vào y tế là tính chất \"hộp đen\" (black-box). Các bác sĩ không thể tin cậy một mô hình nếu không hiểu rõ lý do đằng sau dự đoán đó. Vì vậy, bên cạnh việc tối ưu độ chính xác và kiểm soát âm tính giả (bỏ sót bệnh), việc tích hợp kỹ thuật giải thích mô hình như Grad-CAM (Gradient-weighted Class Activation Mapping) là yêu cầu tối quan trọng để hỗ trợ bác sĩ xác nhận vùng tổn thương trên phổi có thực sự khớp với dự đoán hay không.")
    add_para("Câu hỏi nghiên cứu cốt lõi của đề tài: Deep Learning có thể phân loại ảnh X-quang ngực thành Normal/Pneumonia chính xác đến mức nào? Ba kiến trúc ResNet50, DenseNet121 và InceptionV3 khác nhau ra sao về mặt đánh đổi giữa bỏ sót ca bệnh và báo động giả? Và mô hình đang thực sự dựa vào vùng nào trên ảnh để đưa ra quyết định?")

    add_h2("2. Mục tiêu đề tài")
    add_para("Xây dựng một hệ thống hoàn chỉnh có khả năng phân loại ảnh X-quang ngực thành hai lớp NORMAL và PNEUMONIA một cách tự động, chuẩn xác và có khả năng giải thích, so sánh đối sánh ba kiến trúc học sâu kinh điển (ResNet50, DenseNet121, InceptionV3) pretrained trên ImageNet, kết hợp Grad-CAM để trực quan hóa vùng quan tâm của mạng và triển khai ứng dụng demo tương tác bằng Streamlit.")
    add_para("Các mục tiêu cụ thể gồm:")
    add_bullet("Thiết kế và triển khai pipeline tiền xử lý ảnh và chống rò rỉ dữ liệu (data leakage) nghiêm ngặt theo cấp độ bệnh nhân (patient-level).", bold_prefix="• ")
    add_bullet("So sánh đối sánh khách quan ba kiến trúc ResNet50, DenseNet121, InceptionV3 qua chiến lược huấn luyện 2 giai đoạn (freeze backbone → fine-tuning toàn mạng) và hàm mất mát có trọng số cân bằng lớp.", bold_prefix="• ")
    add_bullet("Kiểm định thống kê McNemar chính xác (exact binomial test) để khẳng định sự khác biệt giữa các mô hình có ý nghĩa thống kê hay chỉ do ngẫu nhiên.", bold_prefix="• ")
    add_bullet("Áp dụng thuật toán Grad-CAM để trực quan hóa bản đồ nhiệt (heatmap) giải thích quyết định cho các trường hợp TP, TN, FP, FN.", bold_prefix="• ")
    add_bullet("Xây dựng giao diện web demo trực quan bằng Streamlit phục vụ thử nghiệm phân loại ảnh đơn lẻ và kiểm tra heatmap.", bold_prefix="• ")

    add_h2("3. Phạm vi và đối tượng nghiên cứu")
    add_para("a) Đối tượng nghiên cứu:", bold=True)
    add_bullet("Các kiến trúc mạng tích chập sâu: ResNet50 (Residual Learning), DenseNet121 (Dense Connectivity), InceptionV3 (Multi-Scale Convolutions).", bold_prefix="• ")
    add_bullet("Kỹ thuật Transfer Learning sử dụng trọng số khởi tạo pretrained từ tập ImageNet.", bold_prefix="• ")
    add_bullet("Phương pháp giải thích thị giác Grad-CAM cho mạng nơ-ron tích chập.", bold_prefix="• ")
    add_bullet("Chiến lược cân bằng lớp bằng hàm mất mát Binary Cross-Entropy có trọng số (weighted BCE loss).", bold_prefix="• ")
    add_bullet("Kiểm định thống kê McNemar phân tích sự khác biệt giữa các bộ phân loại trên cùng tập dữ liệu.", bold_prefix="• ")

    add_para("b) Phạm vi nghiên cứu:", bold=True)
    add_bullet("Sử dụng bộ dữ liệu Chest X-Ray Images của Kermany et al. (2018) gồm 5.856 ảnh X-quang ngực nhi khoa (1–5 tuổi).", bold_prefix="• ")
    add_bullet("Bài toán phân loại nhị phân: NORMAL (Bình thường) vs PNEUMONIA (Viêm phổi bao gồm vi khuẩn và virus).", bold_prefix="• ")
    add_bullet("Chuẩn hóa kích thước ảnh về 224×224 pixel, chuyển thang xám thành 3 kênh RGB theo phân phối ImageNet.", bold_prefix="• ")
    add_bullet("Đánh giá toàn diện trên tập kiểm thử độc lập gồm 618 ảnh đã lọc trùng lặp với 6 chỉ số: Accuracy, Precision, Sensitivity, Specificity, F1-Score, AUC.", bold_prefix="• ")

    doc.add_page_break()

    # ==========================================
    # 5. CHƯƠNG I: CƠ SỞ LÝ THUYẾT
    # ==========================================
    add_h1("CHƯƠNG I. CƠ SỞ LÝ THUYẾT")
    
    add_h2("1.1. Tổng quan về phân loại ảnh y tế")
    add_para("Phân loại ảnh y tế (Medical Image Classification) là một trong những ứng dụng quan trọng hàng đầu của thị giác máy tính và học máy trong y tế hiện đại. Khác với ảnh tự nhiên (chụp phong cảnh, đồ vật), ảnh y tế — đặc biệt là ảnh X-quang lồng ngực (CXR) — mang các đặc thù khắt khe:")
    add_bullet("Độ tương phản thấp: Vùng mô mềm, xương sườn và nhu mô phổi có ranh giới xám mờ nhạt, đòi hỏi bộ lọc đặc trưng cực kỳ nhạy bén.", bold_prefix="• ")
    add_bullet("Đặc thù giải phẫu: Cấu trúc cơ thể người có tính đối xứng giải phẫu (trái/phải); vị trí tim, bóng hoành và nhu mô phổi cố định, do đó không thể lật ảnh ngẫu nhiên ngang (horizontal flip) một cách bừa bãi như ảnh thường.", bold_prefix="• ")
    add_bullet("Cái giá của sai số: Trong y học, sai lầm Âm tính giả (False Negative - FN, tức bỏ sót người mắc bệnh) nguy hiểm hơn rất nhiều so với Dương tính giả (False Positive - FP, báo động giả cần kiểm tra lại). Do đó, chỉ số Sensitivity (Độ nhạy) luôn được đặt lên hàng đầu.", bold_prefix="• ")

    add_h2("1.2. Bộ dữ liệu Chest X-Ray Pneumonia")
    add_para("Bộ dữ liệu thực nghiệm được sử dụng là Chest X-Ray Images (Pneumonia) do Kermany, Goldbaum và cộng sự công bố trên tạp chí y sinh danh tiếng Cell (2018), chia sẻ qua Mendeley Data V2 dưới giấy phép Creative Commons CC BY 4.0. Toàn bộ ảnh X-quang ngực tư thế sau - trước (AP) được thu thập từ các bệnh nhân nhi khoa từ 1 đến 5 tuổi tại Trung tâm Y tế Phụ nữ và Trẻ em Quảng Châu.")
    add_para("Bộ dữ liệu gốc gồm 5.856 ảnh thuộc hai lớp chính:")
    add_bullet("NORMAL (Bình thường): Hình ảnh phổi trong suốt, không có vùng mờ bất thường.", bold_prefix="• ")
    add_bullet("PNEUMONIA (Viêm phổi): Bao gồm cả viêm phổi do vi khuẩn (thường xuất hiện tổn thương đông đặc phân thùy khu trú) và viêm phổi do virus (thường biểu hiện thâm nhiễm mô kẽ lan tỏa khắp hai phế trường).", bold_prefix="• ")
    add_para("Đặc điểm phân bố gốc: Tập Train có 5.216 ảnh, tập Test có 624 ảnh và tập Val chỉ có 16 ảnh (8 Normal, 8 Pneumonia). Việc tập Val gốc chỉ có 16 ảnh là quá nhỏ để theo dõi độ hội tụ và chọn mô hình một cách tin cậy, đồng thời việc chia ảnh ngẫu nhiên nếu không phân tách theo mã bệnh nhân sẽ gây ra hiện tượng rò rỉ dữ liệu (data leakage). Vấn đề này đã được nhóm xử lý triệt để trong Chương II.")

    add_h2("1.3. Cơ sở lý thuyết về xử lý ảnh và học sâu")
    
    add_h3("1.3.1. Chuẩn hóa ảnh (Image Normalization)")
    add_para("Ảnh X-quang ngực trong bộ dữ liệu gốc là ảnh grayscale (1 kênh), kích thước không đồng đều. Để tận dụng các trọng số khởi tạo của mạng sâu huấn luyện trên ImageNet, toàn bộ ảnh được chuẩn hóa theo quy trình:")
    add_bullet("Resize kích thước: Đưa tất cả ảnh về kích thước đồng nhất 224 × 224 pixel bằng nội suy song tuyến (Bilinear Interpolation).", bold_prefix="• ")
    add_bullet("Chuyển đổi kênh màu: Nhân bản kênh xám thành 3 kênh giống nhau (R = G = B) để tương thích với cấu trúc đầu vào của các backbone ImageNet chuẩn.", bold_prefix="• ")
    add_bullet("Chuẩn hóa thống kê: Chuẩn hóa từng kênh ảnh theo trung bình và độ lệch chuẩn của ImageNet: Mean = [0.485, 0.456, 0.406] và Std = [0.229, 0.224, 0.225]. Phép biến đổi này giúp bề mặt hàm mất mát trơn hơn, tránh gradient dao động mạnh và tăng tốc độ hội tụ.", bold_prefix="• ")

    add_h3("1.3.2. Mạng nơ-ron tích chập (CNN)")
    add_para("Mạng nơ-ron tích chập (Convolutional Neural Network - CNN) là kiến trúc mạng nơ-ron nhân tạo chuyên biệt để xử lý dữ liệu dạng lưới 2D như hình ảnh. Ba nguyên lý nền tảng của CNN gồm: trường tiếp nhận cục bộ (local receptive fields), chia sẻ trọng số (shared weights) và lấy mẫu giảm kích thước (pooling). Cấu trúc điển hình bao gồm các tầng tích chập trích xuất đặc trưng không gian, tầng kích hoạt phi tuyến ReLU, tầng gộp (MaxPooling/AvgPooling) và các tầng kết nối đầy đủ (Fully Connected) ở đầu ra.")

    add_h3("1.3.3. ResNet50 — Residual Learning")
    add_para("Mô hình ResNet (Residual Network) do He et al. (2016) đề xuất đã tạo ra bước đột phá khi giải quyết được hiện tượng suy biến mạng (degradation problem) khi tăng độ sâu. Thay vì ép các lớp tích chập trực tiếp học ánh xạ mục tiêu H(x), ResNet bổ sung các đường kết nối tắt (skip connection hay identity shortcut) để mạng học hàm thặng dư F(x) = H(x) - x, từ đó đầu ra trở thành:")
    add_para("y = F(x, {W_i}) + x", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para("Đường kết nối tắt cho phép gradient truyền ngược trực tiếp từ các tầng cuối về các tầng nông mà không bị triệt tiêu (vanishing gradient). Kiến trúc ResNet50 gồm 50 tầng sâu với các khối Bottleneck Residual 3 tầng (1×1 conv giảm kênh, 3×3 conv trích xuất không gian, 1×1 conv tăng kênh), với tổng cộng 23.510.081 tham số huấn luyện.")

    add_h3("1.3.4. DenseNet121 — Dense Connectivity")
    add_para("DenseNet (Densely Connected Convolutional Networks) do Huang et al. (2017) đề xuất đưa nguyên lý kết nối tắt lên một tầm cao mới: mỗi tầng nhận toàn bộ bản đồ đặc trưng (feature maps) của tất cả các tầng trước đó làm đầu vào theo phép ghép nối chuỗi (concatenation):")
    add_para("x_l = H_l([x_0, x_1, ..., x_{l-1}])", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para("Ưu điểm nổi bật của DenseNet121 là tối đa hóa khả năng tái sử dụng đặc trưng (feature reuse), giảm thiểu dư thừa tham số. Nhờ vậy, DenseNet121 sở hữu 121 tầng sâu nhưng chỉ có 6.954.881 tham số (chỉ bằng khoảng 30% kích thước của ResNet50), giúp tiết kiệm tài nguyên bộ nhớ và hạn chế overfitting trên các tập dữ liệu nhỏ.")

    add_h3("1.3.5. InceptionV3 — Multi-Scale Convolutions")
    add_para("Kiến trúc InceptionV3 do Szegedy et al. (2016) phát triển dựa trên triết lý trích xuất đặc trưng đa quy mô (multi-scale feature extraction). Thay vì chọn một kích thước bộ lọc cố định, mỗi khối Inception chứa các nhánh tích chập song song với nhiều kích thước kernel khác nhau: 1×1, 3×3 và tích chập phân rã (factorized convolutions, ví dụ phân rã bộ lọc 3×3 thành 1×3 và 3×1).")
    add_para("Trong bài toán ảnh X-quang phổi, vùng tổn thương đông đặc do viêm phổi có kích thước rất đa dạng: từ những chấm mờ nhỏ phân tán của phế quản phế viêm cho đến các mảng mờ đông đặc lớn chiếm trọn một thùy phổi. Kiến trúc InceptionV3 (với 21.787.617 tham số) là một ứng viên tiềm năng để kiểm chứng giả thuyết đa quy mô này.")

    add_h2("1.4. Kỹ thuật Transfer Learning")
    add_para("Huấn luyện một mạng nơ-ron sâu từ đầu (from scratch) trên dữ liệu ảnh y tế gặp hai thách thức lớn: số lượng ảnh y tế thường hạn chế và nguy cơ overfitting cực cao. Transfer Learning (Học chuyển giao) giải quyết bài toán này bằng cách tái sử dụng các trọng số đã được học từ cơ sở dữ liệu ImageNet (gồm hơn 1,4 triệu ảnh thuộc 1.000 danh mục).")
    add_para("Các bộ lọc ở các tầng đầu của CNN ImageNet đã nắm bắt rất tốt các đặc trưng thị giác cấp thấp cơ bản như đường viền (edges), kết cấu bề mặt (textures) và góc cạnh (corners). Khi áp dụng vào bài toán phân loại viêm phổi:")
    add_bullet("Giữ nguyên phần thân mạng (backbone) đóng vai trò bộ trích xuất đặc trưng.", bold_prefix="• ")
    add_bullet("Thay thế tầng phân loại cuối cùng (classifier head) thành một tầng Linear với 1 neuron đầu ra duy nhất cho bài toán nhị phân.", bold_prefix="• ")
    add_bullet("Đầu ra logit được đưa qua hàm Sigmoid để thu được xác suất ca bệnh: p = Sigmoid(z) = 1 / (1 + exp(-z)).", bold_prefix="• ")

    add_h2("1.5. Grad-CAM — Giải thích quyết định mô hình")
    add_para("Grad-CAM (Gradient-weighted Class Activation Mapping) do Selvaraju et al. (2017) đề xuất là kỹ thuật giải thích trực quan hàng đầu cho mạng CNN. Grad-CAM sử dụng dòng gradient của điểm số lớp mục tiêu y_c đối với các bản đồ đặc trưng A^k của tầng tích chập cuối cùng để tính toán trọng số quan trọng alpha_k:")
    add_para("alpha_k^c = (1 / Z) * sum_i sum_j (d y_c / d A_{i,j}^k)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para("Bản đồ nhiệt kích hoạt L_{Grad-CAM}^c sau đó được tính bằng tổ hợp tuyến tính có trọng số, đi qua hàm kích hoạt ReLU để chỉ giữ lại các đặc trưng có đóng góp tích cực vào việc tăng điểm số của lớp mục tiêu:")
    add_para("L_{Grad-CAM}^c = ReLU(sum_k alpha_k^c * A^k)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para("Bản đồ nhiệt được nội suy về kích thước ảnh gốc 224×224, áp dụng bảng màu JET và phủ (overlay) lên ảnh X-quang xám với tỷ lệ 60% ảnh gốc + 40% heatmap. Trong đề tài:")
    add_bullet("Với lớp PNEUMONIA (target = 1): sử dụng điểm số logit (+z).", bold_prefix="• ")
    add_bullet("Với lớp NORMAL (target = 0): sử dụng điểm số đối nghịch (-z), giúp bản đồ phản ánh chính xác vùng khiến mô hình xếp ảnh vào lớp bình thường.", bold_prefix="• ")

    add_h2("1.6. Hàm mất mát và xử lý mất cân bằng dữ liệu")
    add_para("Trong tập dữ liệu thực tế, số lượng ca PNEUMONIA chiếm khoảng 70.6% tổng số mẫu (tỷ lệ xấp xỉ 2.5 : 1 so với NORMAL). Nếu huấn luyện bằng hàm mất mát tiêu chuẩn, mô hình có xu hướng thiên lệch về lớp đa số để tối đa hóa độ chính xác tổng thể.")
    add_para("Để khắc phục, đề tài sử dụng BCEWithLogitsLoss tích hợp hệ số trọng số pos_weight:")
    add_para("Loss = - (1/N) * sum [ w * y_i * log(sigma(z_i)) + (1 - y_i) * log(1 - sigma(z_i)) ]", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para("Trong đó, hệ số pos_weight được tính toán chính xác từ tập huấn luyện:")
    add_para("w = N_{NORMAL} / N_{PNEUMONIA} = 1078 / 2737 xấp xỉ 0.394", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para("Ý nghĩa: Hệ số w < 1 giúp làm giảm bớt sức nặng mất mát của các ca PNEUMONIA, cân bằng đóng góp gradient giữa hai lớp, ngăn chặn tình trạng mô hình dự đoán toàn bộ thành dương tính.")

    add_h2("1.7. Các chỉ số đánh giá")
    add_para("Quy ước chuẩn y tế: Dương tính (Positive) = PNEUMONIA, Âm tính (Negative) = NORMAL. Với ngưỡng phân loại cố định 0.5, hiệu năng được đánh giá bằng 6 chỉ số:")
    add_bullet("Accuracy (Độ chính xác tổng thể): (TP + TN) / (TP + TN + FP + FN)", bold_prefix="• ")
    add_bullet("Precision (Độ chuẩn xác): TP / (TP + FP) — tỷ lệ các ca được chẩn đoán viêm phổi thực sự mắc bệnh.", bold_prefix="• ")
    add_bullet("Sensitivity / Recall (Độ nhạy): TP / (TP + FN) — tỷ lệ phát hiện thành công người bệnh. Trong y tế, chỉ số này cực kỳ then chốt để không bỏ sót bệnh nhân.", bold_prefix="• ")
    add_bullet("Specificity (Độ đặc hiệu): TN / (TN + FP) — tỷ lệ nhận diện đúng người khỏe mạnh, hạn chế việc chỉ định xét nghiệm xâm lấn không cần thiết.", bold_prefix="• ")
    add_bullet("F1-Score: Trung bình điều hòa giữa Precision và Recall: 2 * (P * R) / (P + R).", bold_prefix="• ")
    add_bullet("AUC-ROC: Diện tích dưới đường cong ROC (Receiver Operating Characteristic), đo lường năng lực phân tách tổng thể của mô hình qua mọi ngưỡng cắt có thể có.", bold_prefix="• ")

    add_h2("1.8. Kết luận chương")
    add_para("Chương I đã làm rõ bức tranh tổng thể về cơ sở lý thuyết của xử lý ảnh y tế, đặc điểm bộ dữ liệu thực nghiệm, nguyên lý hoạt động của 3 kiến trúc mạng nơ-ron tích chập (ResNet50, DenseNet121, InceptionV3), kỹ thuật giải thích quyết định Grad-CAM, cùng phương pháp cân bằng trọng số mất mát. Đây là kim chỉ nam cho toàn bộ thiết kế thực nghiệm được trình bày ở Chương II.")

    doc.add_page_break()

    # ==========================================
    # 6. CHƯƠNG II: PHƯƠNG PHÁP NGHIÊN CỨU
    # ==========================================
    add_h1("CHƯƠNG II. PHƯƠNG PHÁP NGHIÊN CỨU")

    add_h2("2.1. Tổng quan pipeline hệ thống")
    add_para("Quy trình thực nghiệm được thiết kế khép kín và tự động hóa cao từ khâu dữ liệu thô đến xuất kết quả, gồm 7 bước chính:")
    add_bullet("Bước 1: Làm sạch dữ liệu, kiểm tra mã băm SHA-256 pixel và tách nhóm mã bệnh nhân (patient grouping).", bold_prefix="1. ")
    add_bullet("Bước 2: Phân chia tập dữ liệu chống rò rỉ theo phương pháp StratifiedGroupKFold theo ID bệnh nhân.", bold_prefix="2. ")
    add_bullet("Bước 3: Chuẩn hóa ảnh, tiền xử lý và áp dụng kỹ thuật tăng cường dữ liệu an toàn y tế.", bold_prefix="3. ")
    add_bullet("Bước 4: Thiết lập 3 mô hình backbone (ResNet50, DenseNet121, InceptionV3) khởi tạo từ ImageNet.", bold_prefix="4. ")
    add_bullet("Bước 5: Thực hiện chiến lược huấn luyện 2 giai đoạn (Head warmup → Full network fine-tuning) với ReduceLROnPlateau và Early Stopping.", bold_prefix="5. ")
    add_bullet("Bước 6: Đánh giá độc lập trên tập kiểm thử 618 ảnh, xuất ma trận nhầm lẫn, đường cong ROC và tính kiểm định thống kê McNemar.", bold_prefix="6. ")
    add_bullet("Bước 7: Trích xuất bản đồ kích hoạt Grad-CAM phân tích định tính lỗi mô hình và đóng gói ứng dụng Streamlit.", bold_prefix="7. ")

    add_h2("2.2. Cài đặt và chuẩn bị thực nghiệm")
    add_para("a) Môi trường phần cứng và thư viện phần mềm:", bold=True)
    add_bullet("Phần cứng: Hệ thống máy trạm chạy GPU NVIDIA chuyên dụng với hỗ trợ Mixed Precision (CUDA AMP GradScaler).", bold_prefix="• ")
    add_bullet("PyTorch 2.6.0 & torchvision 0.21.0: Thư viện nền tảng để xây dựng mạng học sâu, nạp pretrained model và tối ưu gradient.", bold_prefix="• ")
    add_bullet("Scikit-learn: Tính toán toàn bộ các chỉ số ma trận nhầm lẫn, đường cong ROC, AUROC và StratifiedGroupKFold.", bold_prefix="• ")
    add_bullet("Pillow (PIL) & Matplotlib: Xử lý đọc ảnh X-quang và vẽ đồ thị chuyên nghiệp.", bold_prefix="• ")
    add_bullet("Streamlit: Framework xây dựng ứng dụng web demo tương tác.", bold_prefix="• ")

    add_para("b) Bảng cấu hình siêu tham số đồng nhất (Unified Hyperparameters):", bold=True)
    param_headers = ["Tham số", "Giá trị", "Ý nghĩa trong thực nghiệm"]
    param_data = [
        ["random_seed", "42", "Cố định seed ngẫu nhiên cho Torch, Numpy để tái lập kết quả"],
        ["image_size", "224 × 224", "Kích thước ảnh đầu vào chuẩn hóa cho 3 backbone"],
        ["batch_size", "8", "Kích thước mini-batch huấn luyện"],
        ["head_epochs", "3", "Số chu kỳ làm ấm tầng phân loại (đóng băng backbone)"],
        ["finetune_epochs", "12", "Số chu kỳ tinh chỉnh toàn mạng tối đa"],
        ["head_lr", "1e-3", "Tốc độ học AdamW ở giai đoạn làm ấm đầu phân loại"],
        ["finetune_lr", "1e-4", "Tốc độ học tinh chỉnh toàn mạng (giảm 10 lần)"],
        ["weight_decay", "1e-4", "Hệ số suy giảm trọng số kiểm soát overfitting"],
        ["patience", "4", "Số chu kỳ kiên nhẫn cho Early Stopping theo dõi val_loss"],
        ["lr_patience", "2", "Giảm 50% LR khi val_loss không cải thiện sau 2 chu kỳ"],
        ["pos_weight", "0.394", "Trọng số cân bằng lớp trong BCEWithLogitsLoss"],
        ["amp", "True", "Kích hoạt Automatic Mixed Precision (FP16/FP32)"]
    ]
    tbl_param = doc.add_table(rows=len(param_data) + 1, cols=3)
    style_table(tbl_param, [Inches(1.8), Inches(1.2), Inches(3.5)], param_headers, param_data)

    add_h2("2.3. Bộ dữ liệu và tiền xử lý chống rò rỉ")
    add_h3("2.3.1. Quy trình chống rò rỉ dữ liệu (Anti-Leakage Protocol)")
    add_para("Hiện tượng rò rỉ dữ liệu (data leakage) trong phân loại X-quang thường xảy ra khi các ảnh chụp từ cùng một bệnh nhân (ví dụ chụp ở các góc độ hoặc thời điểm khác nhau) vô tình bị phân tán ở cả tập train và tập test. Khi đó mô hình sẽ \"học vẹt\" đặc điểm cá nhân của bệnh nhân thay vì đặc trưng bệnh lý.")
    add_para("Đề tài áp dụng quy trình kiểm soát 4 bước nghiêm ngặt:")
    add_bullet("Giữ nguyên cấu trúc tập Test gốc (624 ảnh) để đảm bảo tính so sánh khách quan với các công bố quốc tế, chỉ loại bỏ 6 ảnh bị trùng lặp pixel hoàn toàn (SHA-256 hash collision) -> còn lại 618 ảnh Test độc lập tuyệt đối.", bold_prefix="1. ")
    add_bullet("Gộp toàn bộ ảnh trong thư mục train và val gốc thành tập phát triển (Development set).", bold_prefix="2. ")
    add_bullet("Trích xuất ID bệnh nhân qua phân tích tên tệp bằng biểu thức chính quy (regular expressions): nhận diện các mẫu như person\\d+_*, IM-\\d+, NORMAL2-IM-\\d+.", bold_prefix="3. ")
    add_bullet("Loại bỏ 457 ảnh trong tập Development có cùng ID bệnh nhân với tập Test; loại bỏ tiếp 22 ảnh trùng lặp pixel nội bộ -> tổng cộng loại bỏ 485 ảnh rò rỉ và trùng lặp.", bold_prefix="4. ")
    add_bullet("Phân chia tập Development sạch theo tỷ lệ 80/20 thành Train (3.815 ảnh) và Validation (938 ảnh) bằng thuật toán StratifiedGroupKFold theo nhóm ID bệnh nhân.", bold_prefix="5. ")

    add_h3("2.3.2. Tiền xử lý và tăng cường dữ liệu (Data Augmentation)")
    add_para("Để mô hình học được tính bất biến với các biến động chụp X-quang thực tế, nhóm áp dụng kỹ thuật tăng cường dữ liệu thận trọng:")
    add_bullet("Biến đổi hình học nhẹ: RandomAffine với góc xoay ngẫu nhiên giới hạn trong ±7 độ, độ dịch chuyển tịnh tiến ±3%, tỷ lệ co giãn 95% – 105%.", bold_prefix="• ")
    add_bullet("Biến đổi quang học: ColorJitter thay đổi độ sáng (brightness) ±8% và độ tương phản (contrast) ±8%, mô phỏng sự khác biệt về điện áp phát tia X giữa các máy chụp.", bold_prefix="• ")
    add_bullet("Không sử dụng lật ngang (Horizontal Flip): Vì trong phim chụp X-quang lồng ngực, bóng tim nằm lệch trái và phân bố thùy phổi hai bên khác nhau; việc lật ngang làm mất tính chuẩn mực giải phẫu y khoa.", bold_prefix="• ")

    add_h2("2.4. Kiến trúc và cấu hình mô hình")
    add_para("Cả ba mô hình đều được khởi tạo từ trọng số ImageNet công bố chính thức trong torchvision.models:")
    add_bullet("ResNet50: Thay thế tầng phân loại cuối cùng model.fc = nn.Linear(2048, 1). Lớp trích xuất Grad-CAM là model.layer4[-1].", bold_prefix="• ")
    add_bullet("DenseNet121: Thay thế tầng phân loại model.classifier = nn.Linear(1024, 1). Lớp trích xuất Grad-CAM là model.features.norm5.", bold_prefix="• ")
    add_bullet("InceptionV3: Tải pretrained với aux_logits=True theo quy chuẩn torchvision, sau đó gỡ bỏ hoàn toàn nhánh phụ AuxLogits, thiết lập transform_input=True và thay thế model.fc = nn.Linear(2048, 1). Lớp trích xuất Grad-CAM là model.Mixed_7c.", bold_prefix="• ")

    add_h2("2.5. Chiến lược huấn luyện hai giai đoạn")
    add_h3("Giai đoạn 1: Khởi động tầng phân loại (Head Warmup - 3 Epochs)")
    add_para("Đóng băng toàn bộ các tầng trọng số của phần thân mạng (backbone). Đặt backbone ở chế độ eval() để khóa chặt các tham số thống kê running mean và running variance của các lớp Batch Normalization. Chỉ tối ưu hóa duy nhất tầng Linear mới bằng AdamW với learning rate 1e-3. Giai đoạn này giúp tầng đầu ra nhanh chóng điều chỉnh về khoảng giá trị tối ưu mà không làm xáo trộn các đặc trưng chất lượng cao đã học từ ImageNet.")

    add_h3("Giai đoạn 2: Tinh chỉnh toàn bộ mạng (Full Network Fine-Tuning - Tối đa 12 Epochs)")
    add_para("Mở khóa toàn bộ tham số của mạng, khôi phục lại trọng số tốt nhất đã ghi nhận từ Giai đoạn 1. Tiến hành huấn luyện toàn mạng với tốc độ học nhỏ hơn 10 lần (lr = 1e-4) để tinh chỉnh thích ứng sâu với đặc thù ảnh X-quang. Bộ điều phối ReduceLROnPlateau tự động giảm 50% tốc độ học khi hàm mất mát trên tập kiểm định không giảm sau 2 chu kỳ. Thuật toán Early Stopping sẽ ngắt tiến trình huấn luyện nếu sau 4 chu kỳ liên tiếp không có sự cải thiện val_loss, giúp bảo vệ mô hình tuyệt đối khỏi tình trạng học vẹt.")

    add_h2("2.6. Đánh giá và kiểm định thống kê")
    add_para("Mô hình cuối cùng được lựa chọn theo tiêu chí khách quan trên tập validation: Ưu tiên Validation AUC -> Sensitivity -> Specificity. Sau đó, mô hình được đưa vào đánh giá độc lập duy nhất một lần trên tập Test 618 ảnh.")
    add_para("Kiểm định thống kê McNemar chính xác (Exact Binomial Test) được thực hiện để so sánh các cặp mô hình (Pairwise Comparison). Với hai mô hình A và B, gọi b là số mẫu A đúng nhưng B đoán sai, c là số mẫu A sai nhưng B đoán đúng. Kiểm định nhị thức tính toán xác suất p-value để chứng minh sự chênh lệch có ý nghĩa thống kê thực sự ở mức ý nghĩa alpha = 0.05.")

    doc.add_page_break()

    # ==========================================
    # 7. CHƯƠNG III: KẾT QUẢ THỰC NGHIỆM
    # ==========================================
    add_h1("CHƯƠNG III. KẾT QUẢ THỰC NGHIỆM VÀ THẢO LUẬN")

    add_h2("3.1. Phân bố dữ liệu sau làm sạch")
    add_para("Sau khi áp dụng quy trình lọc trùng lặp và loại bỏ rò rỉ bệnh nhân, cơ cấu tập dữ liệu thực nghiệm cuối cùng được tổng hợp trong bảng dưới đây:")
    
    split_headers = ["Phân tập", "Số ca NORMAL", "Số ca PNEUMONIA", "Tổng số ảnh", "Tỷ lệ (%)"]
    split_data = [
        ["Tập Huấn luyện (Train)", "1.078", "2.737", "3.815", "71,03%"],
        ["Tập Xác thực (Validation)", "270", "668", "938", "17,46%"],
        ["Tập Kiểm thử (Test)", "231", "387", "618", "11,51%"],
        ["Toàn bộ tập dữ liệu sạch", "1.579", "3.792", "5.371", "100,00%"]
    ]
    tbl_split = doc.add_table(rows=len(split_data) + 1, cols=5)
    style_table(tbl_split, [Inches(2.0), Inches(1.2), Inches(1.3), Inches(1.1), Inches(1.0)], split_headers, split_data)

    img_dist = os.path.join(RESULTS_DIR, "dataset_distribution.png")
    add_image_safe(img_dist, "Biểu đồ phân bố số lượng ảnh hai lớp NORMAL và PNEUMONIA qua các phân tập", width=Inches(5.2))

    add_h2("3.2. Kết quả chi tiết trên tập kiểm thử độc lập")
    add_para("Hiệu năng của 3 mô hình trên tập kiểm thử độc lập gồm 618 ảnh được đối sánh chi tiết theo các chỉ số phân loại và số đếm ma trận nhầm lẫn:")

    metrics_headers = ["Chỉ số đánh giá", "ResNet50 (Mô hình chọn)", "DenseNet121", "InceptionV3"]
    metrics_data = [
        ["Độ chính xác (Accuracy)", "90,61%", "86,57%", "86,08%"],
        ["Độ chuẩn xác (Precision)", "87,30%", "82,48%", "81,82%"],
        ["Độ nhạy (Sensitivity / Recall)", "99,48%", "99,74%", "100,00%"],
        ["Độ đặc hiệu (Specificity)", "75,76%", "64,50%", "62,77%"],
        ["Điểm F1-Score", "93,00%", "90,29%", "90,00%"],
        ["Diện tích dưới đường cong (AUC)", "96,79%", "95,78%", "96,29%"],
        ["Âm tính thật (TN - Ca thường đúng)", "175", "149", "145"],
        ["Dương tính giả (FP - Báo động giả)", "56", "82", "86"],
        ["Âm tính giả (FN - Bỏ sót bệnh)", "2", "1", "0"],
        ["Dương tính thật (TP - Bệnh đúng)", "385", "386", "387"],
        ["Số lượng tham số (Parameters)", "23.510.081", "6.954.881", "21.787.617"],
        ["Dung lượng Checkpoint lưu trữ", "89,99 MB", "27,11 MB", "83,44 MB"],
        ["Thời gian huấn luyện thực tế", "15 phút 15 giây", "18 phút 46 giây", "23 phút 41 giây"]
    ]
    tbl_metrics = doc.add_table(rows=len(metrics_data) + 1, cols=4)
    style_table(tbl_metrics, [Inches(2.5), Inches(1.6), Inches(1.2), Inches(1.2)], metrics_headers, metrics_data)

    img_comp = os.path.join(RESULTS_DIR, "plots", "model_comparison.png")
    add_image_safe(img_comp, "Biểu đồ cột so sánh tổng thể các chỉ số thực nghiệm của 3 kiến trúc", width=Inches(5.4))

    img_roc = os.path.join(RESULTS_DIR, "roc", "model_comparison_roc.png")
    add_image_safe(img_roc, "Đường cong ROC đối sánh khả năng phân loại giữa ResNet50, DenseNet121 và InceptionV3", width=Inches(5.0))

    add_h3("Phân tích đánh đổi giữa Âm tính giả (FN) và Dương tính giả (FP)")
    add_para("Phân tích sâu vào ma trận nhầm lẫn mang lại phát hiện y khoa quan trọng về sự đánh đổi:")
    add_bullet("InceptionV3 đạt Sensitivity tuyệt đối 100% (FN = 0, phát hiện trọn vẹn 387/387 ca viêm phổi). Tuy nhiên, cái giá phải trả là Specificity sụt giảm nghiêm trọng xuống chỉ còn 62.77%, dẫn đến tới 86 ca báo động giả (FP).", bold_prefix="• ")
    add_bullet("DenseNet121 bỏ sót 1 ca bệnh (FN = 1) nhưng có tới 82 ca báo động giả (Specificity = 64.50%).", bold_prefix="• ")
    add_bullet("ResNet50 chỉ để lọt 2 ca bệnh (Sensitivity = 99.48%), nhưng kiểm soát báo động giả vượt trội với Specificity đạt 75.76% (chỉ 56 ca FP, giảm được 30 ca báo giả so với InceptionV3 và 26 ca so với DenseNet121).", bold_prefix="• ")
    add_para("Trong thực hành lâm sàng, việc giảm 30 ca dương tính giả giúp tránh cho 30 bệnh nhân không phải chịu các can thiệp kháng sinh hoặc chụp cắt lớp vi tính (CT) không đáng có, đồng thời tiết kiệm đáng kể chi phí xét nghiệm và giảm tải cho bệnh viện.")

    # Thêm ma trận nhầm lẫn
    img_cm_res = os.path.join(RESULTS_DIR, "confusion_matrix", "resnet50_confusion_matrix.png")
    add_image_safe(img_cm_res, "Ma trận nhầm lẫn của mô hình ResNet50 trên 618 ảnh kiểm thử", width=Inches(4.5))

    add_h2("3.3. Kiểm định thống kê McNemar")
    add_para("Để khẳng định ưu thế của ResNet50 có ý nghĩa thống kê hay chỉ do ngẫu nhiên, kiểm định McNemar chính xác (Exact Binomial Test) được tiến hành đối sánh từng cặp:")

    mcnemar_headers = ["Cặp so sánh", "A đúng / B sai", "A sai / B đúng", "Exact p-value", "Kết luận thống kê (alpha = 0.05)"]
    mcnemar_data = [
        ["ResNet50 vs DenseNet121", "32", "7", "0.000070", "Bác bỏ H0: ResNet50 vượt trội có ý nghĩa rất cao (p < 0.001)"],
        ["ResNet50 vs InceptionV3", "36", "8", "0.000025", "Bác bỏ H0: ResNet50 vượt trội có ý nghĩa rất cao (p < 0.001)"],
        ["DenseNet121 vs InceptionV3", "20", "17", "0.742829", "Chấp nhận H0: Không có sự khác biệt có ý nghĩa thống kê"]
    ]
    tbl_mcnemar = doc.add_table(rows=len(mcnemar_data) + 1, cols=5)
    style_table(tbl_mcnemar, [Inches(1.8), Inches(1.0), Inches(1.0), Inches(1.1), Inches(2.0)], mcnemar_headers, mcnemar_data)

    add_para("Kết luận thống kê: Giá trị p-value cực nhỏ (< 0.0001) khẳng định một cách vững chắc rằng ưu thế dẫn đầu về độ chính xác của ResNet50 trên tập dữ liệu này là hoàn toàn có thật về mặt thống kê toán học, không phải do phân chia mẫu ngẫu nhiên.")

    add_h2("3.4. Phân tích trực quan hóa Grad-CAM")
    add_para("Grad-CAM được áp dụng cho cả 4 nhóm kết quả dự đoán (TP, TN, FP, FN) nhằm kiểm tra xem mạng học sâu có thực sự nhìn vào tổn thương phổi hay bị đánh lừa bởi các đặc trưng ngoại lai:")
    add_bullet("Trường hợp Dương tính thật (TP): Bản đồ nhiệt Grad-CAM của ResNet50 tập trung rõ nét vào các đám mờ đông đặc ở vùng rốn phổi và đáy phổi, trùng khớp với đặc điểm tổn thương của bệnh học viêm phổi.", bold_prefix="• ")
    add_bullet("Trường hợp Âm tính thật (TN): Bản đồ kích hoạt phân tán đều hoặc có cường độ kích hoạt rất thấp ở hai phế trường, phản ánh phổi thông khí bình thường.", bold_prefix="• ")
    add_bullet("Trường hợp Dương tính giả (FP): Một số ca chụp có bóng tim lớn (cardiomegaly) hoặc có dị vật/ống thở/marker chữ R-L làm mạng nhầm lẫn và kích hoạt ngoài phế trường — đây là dấu hiệu của shortcut learning mà các hệ thống AI cần đặc biệt lưu tâm.", bold_prefix="• ")
    add_bullet("Trường hợp Âm tính giả (FN): Chỉ có 2 ca trong ResNet50. Xem xét kỹ ảnh gốc cho thấy tổn thương là các đốm mờ mô kẽ rất mỏng, độ tương phản cực kỳ thấp, rất khó nhận biết ngay cả với mắt thường.", bold_prefix="• ")

    # Chèn ảnh Grad-CAM mẫu của ResNet50
    gcam_tp_overlay = os.path.join(RESULTS_DIR, "gradcam", "resnet50", "TP_001_overlay.png")
    add_image_safe(gcam_tp_overlay, "Minh họa Grad-CAM Overlay trường hợp Dương tính thật (TP) trên mô hình ResNet50", width=Inches(4.2))

    gcam_fp_overlay = os.path.join(RESULTS_DIR, "gradcam", "resnet50", "FP_001_overlay.png")
    add_image_safe(gcam_fp_overlay, "Minh họa Grad-CAM Overlay trường hợp Dương tính giả (FP) phân tích hiện tượng shortcut learning", width=Inches(4.2))

    add_h2("3.5. Kết luận lựa chọn mô hình tối ưu")
    add_para("Dựa trên toàn bộ kết quả phân tích định lượng và định tính:")
    add_bullet("ResNet50 được lựa chọn là mô hình sản xuất tối ưu của đề tài nhờ đạt điểm số cao nhất ở hầu hết các chỉ số (Accuracy 90.61%, Precision 87.30%, F1 93.00%, AUC 96.79%), đồng thời duy trì sự cân bằng mẫu mực giữa Sensitivity (99.48%) và Specificity (75.76%).", bold_prefix="1. ")
    add_bullet("Thời gian huấn luyện nhanh nhất: ResNet50 hoàn thành hội tụ chỉ trong 15 phút 15 giây (nhanh hơn InceptionV3 gần 8.5 phút).", bold_prefix="2. ")
    add_bullet("Được bảo chứng vững chắc bởi kiểm định thống kê McNemar p < 0.001.", bold_prefix="3. ")

    doc.add_page_break()

    # ==========================================
    # 8. CHƯƠNG IV: KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN
    # ==========================================
    add_h1("CHƯƠNG IV. KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN")

    add_h2("4.1. Đánh giá kết quả đạt được")
    add_para("Qua quá trình thực hiện đề tài giữa kỳ, nhóm đã đạt được các kết quả nổi bật sau:")
    add_bullet("Xây dựng thành công quy trình xử lý dữ liệu chuẩn khoa học: Loại bỏ triệt để nguy cơ rò rỉ bệnh nhân (leakage) và trùng lặp ảnh, chuẩn hóa tập dữ liệu 5.371 ảnh chất lượng cao.", bold_prefix="• ")
    add_bullet("Triển khai và huấn luyện thành công 3 kiến trúc Deep Learning tiên tiến: Áp dụng chiến lược huấn luyện hai giai đoạn và cân bằng trọng số mất mát hiệu quả.", bold_prefix="• ")
    add_bullet("Đánh giá toàn diện và đối soát khách quan: Không chỉ so sánh số liệu kiểm thử mà còn thực hiện kiểm định thống kê McNemar chính xác và phân tích chuyên sâu về sự đánh đổi FN/FP.", bold_prefix="• ")
    add_bullet("Trực quan hóa giải thích mô hình với Grad-CAM: Cung cấp bằng chứng minh bạch về vùng kích hoạt của mạng nơ-ron, nhận diện được nguy cơ shortcut learning trong chẩn đoán y tế.", bold_prefix="• ")
    add_bullet("Đóng gói ứng dụng demo tương tác bằng Streamlit: Cho phép người dùng tải ảnh X-quang bất kỳ, nhận kết quả chẩn đoán tức thì cùng bản đồ nhiệt Grad-CAM trực quan.", bold_prefix="• ")

    add_h2("4.2. Hạn chế của đề tài")
    add_bullet("Tính thuần nhất của nguồn dữ liệu: Dữ liệu chỉ được thu thập từ một bệnh viện nhi khoa duy nhất (Guangzhou Women and Children's Medical Center), chưa kiểm chứng năng lực tổng quát hóa trên bệnh nhân người lớn hoặc các máy chụp khác nhau.", bold_prefix="• ")
    add_bullet("Thực nghiệm trên một seed ngẫu nhiên duy nhất (seed 42): Cần mở rộng lặp lại trên nhiều seed để tính khoảng tin cậy 95% (Confidence Intervals).", bold_prefix="• ")
    add_bullet("Phân loại nhị phân tổng quát: Chưa phân tách chi tiết giữa viêm phổi do vi khuẩn (Bacterial) và virus (Viral), điều này rất quan trọng trong quyết định dùng kháng sinh trên lâm sàng.", bold_prefix="• ")
    add_bullet("Tuyên bố miễn trừ trách nhiệm y tế: Hệ thống được xây dựng nhằm mục đích học thuật và nghiên cứu công nghệ, không thể thay thế chẩn đoán chuyên môn của bác sĩ có chứng chỉ hành nghề.", bold_prefix="• ")

    add_h2("4.3. Hướng phát triển tiếp theo")
    add_bullet("Kiểm định ngoại kiểm (External Validation): Thử nghiệm mô hình trên các kho dữ liệu đa trung tâm quy mô lớn như NIH ChestX-ray14, CheXpert hoặc MIMIC-CXR.", bold_prefix="• ")
    add_bullet("Hiệu chuẩn xác suất (Probability Calibration): Áp dụng phương pháp Platt Scaling hoặc Isotonic Regression để xác suất đầu ra phản ánh đúng mức độ tin cậy thực tế của ca bệnh.", bold_prefix="• ")
    add_bullet("Mở rộng phân loại đa lớp (Multi-class Classification): Phân biệt rõ NORMAL, BACTERIAL PNEUMONIA và VIRAL PNEUMONIA.", bold_prefix="• ")
    add_bullet("Kết hợp mô hình Vision-Language (VLM): Nghiên cứu tích hợp các mô hình đa phương thức hiện đại để tự động sinh báo cáo chẩn đoán hình ảnh sơ bộ dạng văn bản.", bold_prefix="• ")

    doc.add_page_break()

    # ==========================================
    # 9. TÀI LIỆU THAM KHẢO
    # ==========================================
    add_h1("TÀI LIỆU THAM KHẢO")
    
    refs = [
        "[1] Kermany, D. S., Goldbaum, M., Cai, W., et al. (2018). Identifying Medical Diagnoses and Treatable Diseases by Image-Based Deep Learning. Cell, 172(5), 1122-1131. DOI: 10.1016/j.cell.2018.02.010 (CC BY 4.0).",
        "[2] He, K., Zhang, X., Ren, S., & Sun, J. (2016). Deep Residual Learning for Image Recognition. Proceedings of the IEEE Conference on Computer Vision and Pattern Recognition (CVPR 2016), pp. 770-778.",
        "[3] Huang, G., Liu, Z., Van Der Maaten, L., & Weinberger, K. Q. (2017). Densely Connected Convolutional Networks. Proceedings of the IEEE Conference on Computer Vision and Pattern Recognition (CVPR 2017), pp. 4700-4708.",
        "[4] Szegedy, C., Vanhoucke, V., Ioffe, S., Shlens, J., & Wojna, Z. (2016). Rethinking the Inception Architecture for Computer Vision. Proceedings of the IEEE Conference on Computer Vision and Pattern Recognition (CVPR 2016), pp. 2818-2826.",
        "[5] Selvaraju, R. R., Cogswell, M., Das, A., Vedantam, R., Parikh, D., & Batra, D. (2017). Grad-CAM: Visual Explanations from Deep Networks via Gradient-based Localization. Proceedings of the IEEE International Conference on Computer Vision (ICCV 2017), pp. 618-626.",
        "[6] Slide bài giảng môn học Xử lý ảnh và Thị giác máy tính của TS. Huỳnh Hữu Hưng, Trường Đại học Bách Khoa - Đại học Đà Nẵng, 2025 - 2026.",
        "[7] Kaggle Dataset — Chest X-Ray Images (Pneumonia): https://www.kaggle.com/datasets/paultimothymooney/chest-xray-pneumonia",
        "[8] McNemar, Q. (1947). Note on the sampling error of the difference between correlated proportions or percentages. Psychometrika, 12(2), 153-157."
    ]
    for ref in refs:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.2
        r = p.add_run(ref)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(11.5)

    # Lưu tài liệu
    doc.save(OUTPUT_PATH)
    print(f"SUCCESS: Word report generated at {OUTPUT_PATH}".encode('ascii', 'ignore').decode('ascii'))

if __name__ == '__main__':
    create_report()
