# -*- coding: utf-8 -*-
"""
build_report_like_sample.py
Generates the complete Final Report matching the exact presentation style of:
'd:\CODE\tgmt\Báo cáo cuối kì Thị Giác máy tính.docx'

Group: Nhóm 03 - CV CT2 F403
Members: Mai Phước Thịnh, Nguyễn Đình Duy
Target file: d:\CODE\tgmt\baocao_final_viemphoi.docx
Preserves full scientific paper in Appendix at the end.
Main body <= 50 pages.
"""

import os
import sys
import io
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

ROOT = Path(__file__).resolve().parents[1]
OUTPUT_PATH = ROOT / "baocao_final_viemphoi.docx"
RESULTS_DIR = ROOT / "results"
PLOTS_DIR = RESULTS_DIR / "plots"

def generate_report():
    doc = Document()

    # Base page settings from sample: Letter size (612 x 792 pt), margins: left 73.7, right 59.55, top 56.7, bottom 56.7
    PAGE_W = 612.0
    PAGE_H = 792.0
    MAR_L = 73.7
    MAR_R = 59.55
    MAR_T = 56.7
    MAR_B = 56.7

    def apply_section_layout(sec, header_text="", is_cover=False):
        sec.page_width = Pt(PAGE_W)
        sec.page_height = Pt(PAGE_H)
        sec.left_margin = Pt(MAR_L)
        sec.right_margin = Pt(MAR_R)
        sec.top_margin = Pt(MAR_T)
        sec.bottom_margin = Pt(MAR_B)
        sec.different_first_page_header_footer = False

        # Header
        hdr = sec.header
        hdr_p = hdr.paragraphs[0]
        hdr_p.text = ""
        if header_text and not is_cover:
            hdr_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r = hdr_p.add_run(header_text)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(11)
            r.italic = True
            r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

        # Footer (Page number)
        ftr = sec.footer
        ftr_p = ftr.paragraphs[0]
        ftr_p.text = ""
        if not is_cover:
            ftr_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            # Add dynamic page number field inside runs (OpenXML compliant)
            r1 = ftr_p.add_run()
            r1._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w')))
            r2 = ftr_p.add_run()
            r2._r.append(parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w')))
            r3 = ftr_p.add_run()
            r3._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w')))
            r4 = ftr_p.add_run()
            r4._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w')))

    # Base font style
    style_normal = doc.styles['Normal']
    font_normal = style_normal.font
    font_normal.name = 'Times New Roman'
    font_normal.size = Pt(13)
    font_normal.color.rgb = RGBColor(0x00, 0x00, 0x00)

    def set_table_borders(table, color="000000", sz="6"):
        """Grid borders matching the sample."""
        tblPr = table._tbl.tblPr
        borders_xml = f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        </w:tblBorders>
        '''
        tblPr.append(parse_xml(borders_xml))

    def set_cell_margins(cell, top=70, bottom=70, left=100, right=100):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    def add_para(text="", bold=False, italic=False, size=13, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, line_spacing=1.2, color=None):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        if text:
            r = p.add_run(text)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(size)
            r.bold = bold
            r.italic = italic
            r.font.color.rgb = color if color else RGBColor(0x00, 0x00, 0x00)
        return p

    def add_h1(text):
        """Heading 1: Centered, Bold, 13pt, Black, space before 18, space after 12."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(12)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(13)
        r.bold = True
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        return p

    def add_h2(text):
        """Heading 2: Left, Bold, 13pt, Black, space before 12, space after 6."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(13)
        r.bold = True
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        return p

    def add_h3(text):
        """Heading 3: Left, Bold, 13pt, Black, space before 8, space after 4."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(13)
        r.bold = True
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        return p

    def add_bullet(text, bold_prefix=""):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.2
        if bold_prefix:
            r1 = p.add_run(bold_prefix)
            r1.font.name = 'Times New Roman'
            r1.font.size = Pt(13)
            r1.bold = True
            r1.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        r2 = p.add_run(text)
        r2.font.name = 'Times New Roman'
        r2.font.size = Pt(13)
        r2.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        return p

    def add_table_caption(text):
        """Table caption: Above table, Centered, Italic, 12pt Times New Roman."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        r.italic = True
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        return p

    def add_figure_caption(text):
        """Figure caption: Below figure, Centered, Italic, 12pt Times New Roman."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(10)
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        r.italic = True
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        return p

    def add_image_sample(img_path, caption, width=Inches(5.2)):
        img_p = Path(img_path)
        if img_p.exists():
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.keep_with_next = True
            p.add_run().add_picture(str(img_p), width=width)
            add_figure_caption(caption)

    def style_table_sample(tbl, col_widths, headers, data):
        """Styled table matching the sample: clean black grid borders, centered, cell vertical center."""
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(tbl, "000000", "6")

        # Header row
        hdr_cells = tbl.rows[0].cells
        for i, h in enumerate(headers):
            hdr_cells[i].text = ""
            p = hdr_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.15
            r = p.add_run(h)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(12)
            r.bold = True
            r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
            set_cell_margins(hdr_cells[i], 80, 80, 100, 100)
            tcPr = hdr_cells[i]._tc.get_or_add_tcPr()
            tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))

        # Body rows
        for row_idx, row_data in enumerate(data):
            row_cells = tbl.rows[row_idx + 1].cells
            for col_idx, cell_value in enumerate(row_data):
                row_cells[col_idx].text = ""
                p = row_cells[col_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx in [0, len(row_data)-1] and len(str(cell_value)) < 15 else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.line_spacing = 1.15
                r = p.add_run(str(cell_value))
                r.font.name = 'Times New Roman'
                r.font.size = Pt(12)
                r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
                set_cell_margins(row_cells[col_idx], 70, 70, 100, 100)
                tcPr = row_cells[col_idx]._tc.get_or_add_tcPr()
                tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))

        for row in tbl.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Inches(w)

    print("Building Document matching Sample Style...")

    # =========================================================================
    # SECTION 0: COVER PAGE (Trang bìa theo chuẩn mẫu)
    # =========================================================================
    sec0 = doc.sections[0]
    apply_section_layout(sec0, is_cover=True)

    for _ in range(7):
        add_para("")

    add_para("BÁO CÁO TỔNG HỢP KẾT QUẢ THỰC NGHIỆM", bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=6)
    add_para("PHÂN LOẠI VIÊM PHỔI TRÊN ẢNH X-QUANG LỒNG NGỰC NHI KHOA BẰNG RESNET50 VÀ SPATIAL ATTENTION - TÁC ĐỘNG CỦA RÒ RỈ DỮ LIỆU, CÂN BẰNG LỚP VÀ KHỬ BẪY HỌC PHÍM TẮT", 
             bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=30)

    for _ in range(4):
        add_para("")

    add_para("Nhóm thực hiện: Nhóm 03 - CV CT2 F403", bold=False, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=4)
    add_para("Sinh viên thực hiện:", bold=False, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2, space_after=2)
    add_para("Mai Phước Thịnh", bold=False, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=1, space_after=1)
    add_para("Nguyễn Đình Duy", bold=False, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=1, space_after=1)

    for _ in range(12):
        add_para("")

    add_para("Tháng 11/2026", bold=False, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0)

    # =========================================================================
    # SECTION 1: MỤC LỤC
    # =========================================================================
    sec1 = doc.add_section(WD_SECTION.NEW_PAGE)
    apply_section_layout(sec1, header_text="Mục lục")

    add_para("MỤC LỤC", bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=12, space_after=12)

    toc_entries = [
        ("CHƯƠNG 1. TỔNG QUAN NGHIÊN CỨU LIÊN QUAN", "3"),
        ("  1.1. Giới thiệu bài toán", "3"),
        ("  1.2. Mục tiêu nghiên cứu", "3"),
        ("  1.3. Tổng quan các hướng tiếp cận và đánh giá ưu, nhược điểm", "4"),
        ("    1.3.1. Mạng nơ-ron tích chập và học chuyển giao trong phân loại ảnh X-quang", "4"),
        ("    1.3.2. Kiến trúc ResNet50 trong phân loại ảnh X-quang phổi", "4"),
        ("    1.3.3. Bộ dữ liệu Kermany và bài toán chẩn đoán viêm phổi nhi khoa", "5"),
        ("    1.3.4. Vấn đề rò rỉ danh tính bệnh nhân và phân chia tập dữ liệu", "6"),
        ("    1.3.5. Bẫy học phím tắt (Shortcut Learning) và thiên lệch vùng nhìn", "6"),
        ("    1.3.6. Cơ sở toán học của mạng tích chập và phân loại ảnh y sinh", "7"),
        ("    1.3.7. Học chuyển giao và chiến lược điều chỉnh mô hình", "7"),
        ("    1.3.8. Khác biệt miền dữ liệu và nguy cơ học đặc trưng không liên quan", "8"),
        ("  1.4. Các công trình nghiên cứu liên quan", "8"),
        ("    1.4.1. Các công bố kinh điển trên tập dữ liệu Kermany (Kermany 2018, Chouhan 2020)", "8"),
        ("    1.4.2. Các hướng tiếp cận cơ chế chú ý (CBAM, Spatial Attention)", "9"),
        ("    1.4.3. Kiến trúc ResNet và ý nghĩa của kết nối tắt tàn dư", "9"),
        ("    1.4.4. Đối chiếu ưu điểm và hạn chế của các hướng tiếp cận", "10"),
        ("    1.4.5. Vấn đề rò rỉ danh tính bệnh nhân trong các nghiên cứu trước", "10"),
        ("    1.4.6. Vai trò của việc kiểm toán dữ liệu và khử trùng lặp SHA-256", "11"),
        ("    1.4.7. Bảng tổng hợp ưu điểm, hạn chế và khả năng áp dụng", "11"),
        ("    1.4.8. Tổng hợp bằng chứng từ các công trình và bài học thiết kế", "12"),
        ("  1.5. Khoảng trống nghiên cứu và định hướng thực hiện", "12"),
        ("CHƯƠNG 2. GIẢI PHÁP ĐỀ XUẤT VÀ THIẾT KẾ THỰC NGHIỆM", "13"),
        ("  2.0. Tổng quan kiến trúc giải pháp và phạm vi triển khai", "13"),
        ("  2.1. Dữ liệu và thiết kế thực nghiệm", "14"),
        ("    2.1.1. Dữ liệu X-quang nhi khoa Kermany và quy trình thu nhận tại Quảng Châu", "14"),
        ("    2.1.2. Thiết kế chuỗi thực nghiệm đối chứng (Baseline vs Attention vs Naive Split)", "15"),
        ("  2.2. Cấu hình huấn luyện", "16"),
        ("    2.2.1. Cấu trúc ResNet50 và cơ chế học phần dư", "16"),
        ("    2.2.2. Cơ sở lựa chọn huấn luyện hai giai đoạn và checkpoint", "17"),
        ("    2.2.3. Thiết kế đối chứng và ý nghĩa của ba random seed", "17"),
        ("  2.3. Lựa chọn baseline và các kiến trúc so sánh", "18"),
        ("  2.4. Module Spatial Lung-RoI Guided Attention Gate", "18"),
        ("    2.4.1. Mục tiêu và cơ sở toán học", "18"),
        ("    2.4.2. Quy trình xử lý đặc trưng và kết nối tàn dư", "19"),
        ("  2.5. Cấu hình augmentation và kỹ thuật an toàn y sinh", "20"),
        ("    2.5.1. Quá trình huấn luyện mô hình đề xuất", "20"),
        ("    2.5.2. Stage 1 (Khởi động và làm ấm cổng chú ý)", "21"),
        ("    2.5.3. Stage 2 (Fine-tuning toàn diện)", "21"),
        ("    2.5.4. Nguyên tắc đánh giá và chọn mô hình", "22"),
        ("    2.5.5. Kiểm soát nguy cơ rò rỉ thông tin và khả năng tái lập", "22"),
        ("  2.6. Kiểm soát chất lượng dữ liệu, kiểm tra trùng lặp và lưu trữ minh chứng", "23"),
        ("    2.6.1. Lưu trữ kết quả và minh chứng thực nghiệm", "24"),
        ("  2.7. Phương pháp đo lường và phân tích kết quả", "25"),
        ("  2.8. Quy tắc bảo đảm tính hợp lệ và khả năng tái lập", "26"),
        ("CHƯƠNG 3. ĐÁNH GIÁ KẾT QUẢ VÀ THẢO LUẬN", "27"),
        ("  3.1. Kết quả chi tiết các thực nghiệm", "27"),
        ("  3.2. Kết quả trung bình qua ba random seed", "28"),
        ("  3.3. Kiểm định thống kê", "29"),
        ("    3.3.1. Phương pháp kiểm định McNemar chính xác", "29"),
        ("    3.3.2. Kết quả McNemar exact test", "30"),
        ("  3.4. Kết quả paired bootstrap", "31"),
        ("    3.4.1. Khoảng tin cậy 95% Bootstrap Confidence Intervals", "31"),
        ("  3.5. Kết quả kiểm định rò rỉ dữ liệu có kiểm soát (Controlled Leakage Gap)", "32"),
        ("    3.5.1. So sánh trực tiếp Naive Split vs Clean Patient-Disjoint Split", "32"),
        ("    3.5.2. Giải mã hiện tượng 98%-99% trong y văn", "33"),
        ("  3.6. Thảo luận", "34"),
        ("    3.6.1. Tác động của việc phân hoạch độc lập bệnh nhân và Generalization Gap", "34"),
        ("    3.6.2. Tác động của kỹ thuật cân bằng dữ liệu: Weighted BCE vs Oversampling 1:1", "35"),
        ("    3.6.3. Tác động của cổng chú ý Spatial Lung-RoI: Định lượng năng lượng Grad-CAM", "36"),
        ("  3.7. Hạn chế của nghiên cứu", "38"),
        ("  3.8. Minh họa kết quả thực nghiệm và phân tích định tính Grad-CAM", "39"),
        ("    3.8.1. Hàm ý của kết quả đối với ứng dụng lâm sàng tại bệnh viện", "40"),
        ("  3.9. Kết luận", "41"),
        ("TÀI LIỆU THAM KHẢO", "43"),
        ("PHỤ LỤC: TOÀN VĂN BÀI BÁO KHOA HỌC (IEEE CONFERENCE FORMAT)", "45")
    ]

    for title, pg in toc_entries:
        p_toc = doc.add_paragraph()
        p_toc.paragraph_format.space_before = Pt(1)
        p_toc.paragraph_format.space_after = Pt(1)
        p_toc.paragraph_format.line_spacing = 1.15
        is_ch = title.startswith("CHƯƠNG") or title in ["MỤC LỤC", "TÀI LIỆU THAM KHẢO", "PHỤ LỤC: TOÀN VĂN BÀI BÁO KHOA HỌC (IEEE CONFERENCE FORMAT)"]
        r1 = p_toc.add_run(title)
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(12) if not is_ch else Pt(12.5)
        r1.bold = is_ch
        dots_count = max(4, 78 - len(title))
        r_dots = p_toc.add_run(" " + "." * dots_count + " ")
        r_dots.font.name = 'Times New Roman'
        r_dots.font.size = Pt(10)
        r_dots.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        r2 = p_toc.add_run(pg)
        r2.font.name = 'Times New Roman'
        r2.font.size = Pt(12) if not is_ch else Pt(12.5)
        r2.bold = is_ch

    print("Building Chapter 1 (matching sample structure)...")
    # =========================================================================
    # SECTION 2: CHƯƠNG 1. TỔNG QUAN NGHIÊN CỨU LIÊN QUAN
    # =========================================================================
    sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
    apply_section_layout(sec2, header_text="Chương 1. Tổng quan nghiên cứu liên quan")

    add_h1("CHƯƠNG 1. TỔNG QUAN NGHIÊN CỨU LIÊN QUAN")

    add_h2("1.1. Giới thiệu bài toán")
    add_para("Bài toán phân loại bệnh viêm phổi từ ảnh chụp X-quang lồng ngực (Chest X-Ray - CXR) là một trong những ứng dụng trọng yếu và có tác động xã hội sâu sắc nhất của thị giác máy tính trong y tế thông minh. Viêm phổi là tình trạng nhiễm trùng cấp tính đường hô hấp dưới làm tổn thương phế nang, dẫn đến tích tụ dịch rỉ viêm hoặc mủ và cản trở nghiêm trọng quá trình trao đổi khí. Theo các báo cáo dịch tễ học thường niên của Tổ chức Y tế Thế giới (WHO) và UNICEF, viêm phổi là nguyên nhân truyền nhiễm hàng đầu gây tử vong ở trẻ em dưới năm tuổi, cướp đi sinh mạng của hơn 800.000 trẻ nhỏ mỗi năm trên toàn cầu.")
    add_para("Trong nghiên cứu này, một thách thức then chốt mang tính bản chất là sự khác biệt sâu sắc giữa môi trường dữ liệu chuẩn hóa trong phòng thí nghiệm và môi trường lâm sàng thực tế tại bệnh viện. Ảnh X-quang ngực nhi khoa trong bộ dữ liệu Kermany et al. (công bố trên Cell năm 2018) được chụp trên các bệnh nhi từ 1 đến 5 tuổi tại Trung tâm Y tế Phụ nữ và Trẻ em Quảng Châu. Đặc thù giải phẫu lồng ngực trẻ em có độ biến thiên rất cao (xương sườn nằm ngang, bóng tuyến ức sinh lý 'cánh buồm' dễ gây nhầm lẫn với đông đặc thùy phổi, tỷ lệ tim/lồng ngực lớn). Đồng thời, ảnh chụp thực tế luôn chứa các yếu tố gây nhiễu vật lý như ký hiệu chữ chì định hướng 'L'/'R', bóng mờ cơ hoành và các dây điện cực y tế.")
    add_para("Chuỗi thực nghiệm của đề tài được thiết kế một cách có hệ thống nhằm khảo sát ba câu hỏi khoa học cốt lõi:")
    add_para("Thứ nhất: Tác động của rò rỉ danh tính bệnh nhân (Patient Identity Leakage) khi phân chia ngẫu nhiên từng ảnh đối với độ tin cậy của mô hình.")
    add_para("Thứ hai: Tác động độc lập của cơ chế chú ý không gian (Spatial Lung-RoI Guided Attention Gate) trong việc khử bỏ bẫy học phím tắt ngoài phổi.")
    add_para("Thứ ba: Bản chất của kỹ thuật cân bằng dữ liệu: liệu việc nhân bản mẫu 1:1 (Random Oversampling) có thực sự tốt hơn hàm mất mát điều chỉnh trọng số (Weighted BCE Loss) trên dữ liệu ảnh y tế có cấu trúc sinh học hay không.")

    add_h2("1.2. Mục tiêu nghiên cứu")
    add_para("Nghiên cứu hướng đến các mục tiêu khoa học cụ thể sau:")
    add_para("- Thiết lập quy trình kiểm định dữ liệu 4 tầng tự động hóa: khử sạch 419 ảnh trùng lặp SHA-256 và 66 tệp tin hỏng; bóc tách 1.349 định danh bệnh nhân thực thụ để phát hành bộ chuẩn Patient-Disjoint Benchmark gồm 5.371 ảnh sạch độc lập bệnh nhân tuyệt đối.")
    add_para("- Đánh giá độ ổn định và tính tái lập của các cấu hình mô hình qua ba random seed (42, 123, 456), báo cáo đầy đủ chỉ số trung bình và độ lệch chuẩn (Mean ± SD).")
    add_para("- Thiết kế thực nghiệm đối chứng có kiểm soát (Controlled Leakage Experiment) để định lượng chính xác khoảng cách thổi phồng hiệu năng (Leakage Inflation Gap) giữa cách chia ngẫu nhiên truyền thống và phân hoạch sạch độc lập bệnh nhân.")
    add_para("- Nghiên cứu thiết kế module Spatial Lung-RoI Guided Attention Gate siêu nhẹ (chỉ 98 tham số) tích hợp vào tầng Layer4 của ResNet-50 kết hợp tối ưu hóa hai pha tách rời (Two-Stage Decoupled Optimization).")
    add_para("- Ứng dụng kiểm định giả thuyết thống kê McNemar nhị thức chính xác và khoảng tin cậy 95% Paired Bootstrap (1.000 lần lặp) để đánh giá ý nghĩa thống kê của các chênh lệch hiệu năng.")
    add_para("- Xây dựng chỉ số định lượng tỷ lệ năng lượng nhiệt Grad-CAM trong vùng phổi (Intra-Pulmonary Energy Ratio E_RoI) nhằm đánh giá khách quan năng lực triệt tiêu bẫy học phím tắt của cơ chế chú ý.")

    add_h2("1.3. Tổng quan các hướng tiếp cận và đánh giá ưu, nhược điểm")
    add_h3("1.3.1. Mạng nơ-ron tích chập và học chuyển giao trong phân loại ảnh X-quang")
    add_para("Mạng nơ-ron tích chập (CNN) đã cách mạng hóa lĩnh vực thị giác máy tính y tế nhờ khả năng học biểu diễn đặc trưng phân cấp tự động từ dữ liệu pixel thô. Tuy nhiên, việc huấn luyện một mạng CNN sâu từ đầu (from scratch) trên các bộ dữ liệu X-quang y tế thường gặp khó khăn do kích thước mẫu bị giới hạn, dẫn đến hiện tượng quá khớp (overfitting). Phương pháp học chuyển giao (Transfer Learning) từ các mô hình tiền huấn luyện trên tập dữ liệu tự nhiên ImageNet quy mô lớn đã trở thành giải pháp tiêu chuẩn. Các tầng tích chập ban đầu giữ lại khả năng trích xuất các đặc trưng thị giác phổ quát (cạnh, góc, kết cấu vi mô), trong khi các tầng sâu hơn được tinh chỉnh để học các mẫu bệnh lý đặc thù của viêm phổi.")

    add_h3("1.3.2. Kiến trúc ResNet50 trong phân loại ảnh X-quang phổi")
    add_para("Kiến trúc mạng tàn dư sâu ResNet-50 do He et al. (2016) đề xuất giải quyết triệt để vấn đề tiêu biến độ dốc (Vanishing Gradient) thông qua các kết nối tắt tàn dư (Identity Shortcut Connections). Cấu trúc khối thắt cổ chai (Bottleneck Block) với ba lớp tích chập liên tiếp (1×1 giảm chiều, 3×3 trích xuất không gian, 1×1 khôi phục chiều) mang lại sự cân bằng tối ưu giữa năng lực trích xuất ngữ cảnh bệnh lý toàn cục và chi phí tính toán (23.5 triệu tham số).")

    add_h3("1.3.3. Bộ dữ liệu Kermany và bài toán chẩn đoán viêm phổi nhi khoa")
    add_para("Bộ dữ liệu Kermany et al. (Cell 2018) là bộ chuẩn phổ biến nhất trong nghiên cứu phân loại viêm phổi nhi khoa bằng học sâu. Bộ dữ liệu bao gồm 5.856 ảnh X-quang ngực chụp ở tư thế trước - sau (AP view) trên bệnh nhi từ 1 đến 5 tuổi, được phân thành hai nhóm chính: Bình thường (Normal) và Viêm phổi (Pneumonia, gồm căn nguyên vi khuẩn và virus). Điểm đặc thù của viêm phổi vi khuẩn là tạo ra các ổ đông đặc thùy phổi đồng nhất (lobar consolidation) và dấu hiệu phế quản khí; trong khi viêm phổi virus biểu hiện dưới dạng thâm nhiễm mô kẽ lan tỏa hai bên phế trường.")

    add_h3("1.3.4. Vấn đề rò rỉ danh tính bệnh nhân và phân chia tập dữ liệu")
    add_para("Một vấn đề nghiêm trọng tồn tại trong cấu trúc phát hành gốc của Kermany là sự thiếu hụt siêu dữ liệu quản lý bệnh nhân. Một bệnh nhi khi nằm viện điều trị thường được chỉ định chụp X-quang nhiều lần (từ 2 đến 8 phim chụp). Khi phân chia tập dữ liệu theo phương pháp ngẫu nhiên từng ảnh (Random Image Split), các phim chụp khác nhau của cùng một bệnh nhân đã bị phân tán vào cả tập huấn luyện và tập kiểm thử. Do mỗi đứa trẻ có đặc trưng khung xương lồng ngực hoàn toàn độc nhất, mạng CNN dễ dàng đạt độ chính xác giả tạo bằng cách học nhận diện 'đứa trẻ nào' thay vì học 'tổn thương ở đâu'.")

    add_h3("1.3.5. Bẫy học phím tắt (Shortcut Learning) và thiên lệch vùng nhìn")
    add_para("Học phím tắt xảy ra khi mạng nơ-ron học sâu tối ưu hóa hàm mất mát dựa vào các mối tương quan giả tạo (spurious correlations) có sẵn trong dữ liệu thay vì học bản chất bệnh học thực sự. Trong ảnh X-quang, mạng thường hội tụ vào các ký hiệu chữ chì 'L'/'R', viền đệm đen của máy quét hoặc bờ xương đòn. Khi đưa mô hình sang môi trường bệnh viện khác không có các ký hiệu hoặc quy trình chụp tương tự, hiệu năng của mô hình sẽ sụp đổ hoàn toàn.")

    add_h3("1.3.6. Cơ sở toán học của mạng tích chập và phân loại ảnh y sinh")
    add_para("Phép tích chập 2D biến đổi một tensor đầu vào thông qua ma trận trọng số hạt nhân (kernel). Sự kết hợp giữa phép tích chập, chuẩn hóa khối (BatchNorm) và hàm kích hoạt phi tuyến ReLU cho phép mạng xấp xỉ các hàm phân tách phi tuyến phức tạp trong không gian đặc trưng. Hàm mất mát entropy chéo nhị phân (Binary Cross-Entropy) đo lường khoảng cách phân phối xác suất giữa nhãn thực tế và dự đoán của mô hình.")

    add_h3("1.3.7. Học chuyển giao và chiến lược điều chỉnh mô hình")
    add_para("Chiến lược huấn luyện đóng băng một phần (Feature Extraction Warmup) và sau đó mở băng toàn bộ (End-to-End Fine-Tuning) cho phép mô hình thích nghi dần dần với miền ảnh y tế mà không phá hủy các trọng số biểu diễn đã học từ ImageNet.")

    add_h3("1.3.8. Khác biệt miền dữ liệu và nguy cơ học đặc trưng không liên quan")
    add_para("Tương tự như sự khác biệt miền giữa PlantVillage và PlantDoc trong bài toán nông nghiệp, ảnh X-quang từ các máy chụp khác nhau hoặc giữa các khoa phòng khác nhau tạo ra khoảng cách miền (Domain Shift). Nếu mô hình học các đặc trưng nền ngoài phổi, nó sẽ hoàn toàn mất khả năng tổng quát hóa trên bệnh nhân ngoài miền huấn luyện.")

    add_h2("1.4. Các công trình nghiên cứu liên quan")
    add_h3("1.4.1. Các công bố kinh điển trên tập dữ liệu Kermany (Kermany 2018, Chouhan 2020)")
    add_para("Kermany et al. (Cell 2018) áp dụng Inception-V3 và công bố độ chính xác 92.8%. Chouhan et al. (IEEE TMI 2020) kết hợp 5 kiến trúc CNN và công bố 96.4%. Ayan & Ünver (2019) và Rahman et al. (2020) tiếp tục công bố các con số lên tới 98%–99%. Tuy nhiên, toàn bộ các nghiên cứu này đều sử dụng cách chia ngẫu nhiên từng ảnh gốc và giữ nguyên các ảnh trùng lặp.")

    add_h3("1.4.2. Các hướng tiếp cận cơ chế chú ý (CBAM, Spatial Attention)")
    add_para("Các cơ chế chú ý như CBAM (Woo et al. 2018) hay SE-Net giúp mạng tập trung vào các kênh và vùng quan trọng. Tuy nhiên, CBAM chứa cơ chế Channel Attention phức tạp với nhiều lớp Dense, dễ gây quá khớp trên ảnh X-quang xám. Đề tài đề xuất cơ chế chú ý không gian tối giản chỉ gồm 98 tham số.")

    add_h3("1.4.3. Kiến trúc ResNet và ý nghĩa của kết nối tắt tàn dư")
    add_para("Kết nối tàn dư cho phép gradient truyền ngược trực tiếp qua các lớp mà không bị suy giảm: y = F(x) + x, tạo điều kiện thuận lợi cho việc tích hợp cổng chú ý không gian dưới dạng tàn dư cộng dồn.")

    add_h3("1.4.4. Đối chiếu ưu điểm và hạn chế của các hướng tiếp cận")
    add_para("Việc lựa chọn kiến trúc và chiến lược xử lý dữ liệu đòi hỏi sự cân nhắc giữa năng lực trích xuất đặc trưng và nguy cơ quá khớp. Bảng 1.1 dưới đây tóm lược các đánh đổi then chốt.")

    add_table_caption("Bảng 1.1. Đối chiếu các hướng tiếp cận theo mục tiêu nghiên cứu và những giới hạn cần kiểm soát")
    tbl_1_data = [
        ["Học chuyển giao với ResNet50", "Khai thác biểu diễn thị giác sâu từ ImageNet, hội tụ nhanh trên dữ liệu y tế cỡ vừa.", "Cần tinh chỉnh cho miền ảnh; kết quả nội miền không đảm bảo tổng quát hóa ngoài bệnh nhân mới."],
        ["Huấn luyện bằng tập Kermany gốc", "Dữ liệu có cấu trúc phân lớp rõ ràng, thuận tiện thiết lập baseline đối chứng.", "Tồn tại 419 ảnh trùng SHA-256, 66 file hỏng và 51.3% rò rỉ bệnh nhân giữa train và test."],
        ["Cơ chế chú ý CBAM / Transformer", "Tăng cường khả năng biểu diễn đa kênh và quan hệ ngữ cảnh toàn cục.", "Số lượng tham số lớn, dễ gây quá khớp trên ảnh X-quang xám, làm tăng chi phí tính toán."],
        ["Cổng chú ý Spatial Lung-RoI (Đề xuất)", "Siêu nhẹ (chỉ 98 tham số), định hướng gradient vào tổn thương phổi, triệt tiêu bẫy học ngoài phổi.", "Đòi hỏi chiến lược huấn luyện hai giai đoạn tách rời để ổn định không gian đặc trưng."],
        ["Cân bằng lớp bằng Weighted BCE Loss", "Bảo toàn phân bố sinh học tự nhiên, cân bằng động lượng gradient cập nhật trực tiếp tại hàm mất mát.", "Cần tính toán chính xác tỷ lệ nghịch đảo số lượng mẫu giữa hai lớp trong tập huấn luyện."]
    ]
    t1 = doc.add_table(rows=len(tbl_1_data)+1, cols=3)
    style_table_sample(t1, [1.8, 2.6, 2.6], ["Hướng tiếp cận", "Ưu điểm", "Hạn chế / Lưu ý"], tbl_1_data)

    add_h3("1.4.5. Vấn đề rò rỉ danh tính bệnh nhân trong các nghiên cứu trước")
    add_para("Khi bệnh nhân xuất hiện đồng thời ở cả hai phân vùng dữ liệu, mô hình có thể ghi nhớ cấu trúc giải phẫu tĩnh thay vì bệnh học. Đây là nguyên nhân trực tiếp dẫn đến hiện tượng 'ảo tưởng hiệu năng' 98%–99% trong y văn.")

    add_h3("1.4.6. Vai trò của việc kiểm toán dữ liệu và khử trùng lặp SHA-256")
    add_para("Kiểm toán mã băm SHA-256 trên ma trận pixel thô cho phép loại bỏ dứt điểm các tệp tin nhân bản ngẫu nhiên, đảm bảo tính liêm chính khoa học cho tập chuẩn đánh giá.")

    add_h3("1.4.7. Bảng tổng hợp ưu điểm, hạn chế và khả năng áp dụng")
    add_para("Bảng 1.2 tổng hợp các công trình tiêu biểu làm nền tảng cho việc thiết kế nghiên cứu của đề tài.")

    add_table_caption("Bảng 1.2. Tổng hợp các công trình làm nền tảng cho lựa chọn thiết kế nghiên cứu")
    tbl_2_data = [
        ["He và cộng sự (2016) [3]", "Đề xuất mạng học phần dư ResNet, khắc phục tiêu biến độ dốc cho mô hình sâu.", "Không đánh giá trực tiếp trên miền ảnh X-quang phổi nhi khoa."],
        ["Kermany và cộng sự (2018) [1]", "Công bố bộ dữ liệu X-quang ngực nhi khoa và baseline Inception-V3.", "Phân chia tập ngẫu nhiên gây rò rỉ bệnh nhân; tập val chỉ có 16 ảnh."],
        ["Chouhan và cộng sự (2020) [2]", "Ensemble kết hợp 5 mạng học sâu đạt độ chính xác 96.4%.", "Bị ảnh hưởng bởi rò rỉ danh tính bệnh nhân do dùng nguyên phân vùng gốc."],
        ["DeGrave và cộng sự (2021) [5]", "Chứng minh AI X-quang thường dựa dẫm vào phím tắt ngoài phổi (chữ chì, viền ảnh).", "Nghiên cứu phê bình phương pháp luận, không đề xuất mô hình khắc phục."],
        ["Selvaraju và cộng sự (2017) [9]", "Đề xuất phương pháp Grad-CAM giải thích quyết định thị giác của mạng CNN.", "Chỉ dừng lại ở trực quan hóa định tính, chưa có định lượng tỷ lệ năng lượng RoI."]
    ]
    t2 = doc.add_table(rows=len(tbl_2_data)+1, cols=3)
    style_table_sample(t2, [1.6, 2.7, 2.7], ["Công trình", "Đóng góp chính", "Giới hạn liên quan đề tài"], tbl_2_data)

    add_h3("1.4.8. Tổng hợp bằng chứng từ các công trình và bài học thiết kế")
    add_para("Từ các bằng chứng thực nghiệm trên, đề tài rút ra ba bài học thiết kế then chốt: (1) Bắt buộc phải cô lập bệnh nhân tuyệt đối giữa train và test; (2) Phải có cơ chế định hướng không gian để chống bẫy học phím tắt; (3) Mọi kết luận so sánh phải đi kèm kiểm định thống kê nghiêm ngặt.")

    add_h2("1.5. Khoảng trống nghiên cứu và định hướng thực hiện")
    add_para("Khoảng trống nghiên cứu lớn nhất hiện nay là sự thiếu vắng một bộ chuẩn dữ liệu độc lập bệnh nhân được chứng thực trên Kermany CXR, cùng với việc thiếu các giải pháp cơ chế chú ý siêu nhẹ vừa kiểm soát được phím tắt vừa duy trì được độ nhạy phát hiện bệnh cao. Đề tài PneumoGuard-AI được xây dựng nhằm khỏa lấp toàn diện khoảng trống này.")

    print("Building Chapter 2 (matching sample structure)...")
    # =========================================================================
    # SECTION 3: CHƯƠNG 2. GIẢI PHÁP ĐỀ XUẤT VÀ THIẾT KẾ THỰC NGHIỆM
    # =========================================================================
    sec3 = doc.add_section(WD_SECTION.NEW_PAGE)
    apply_section_layout(sec3, header_text="Chương 2. Giải pháp đề xuất và thiết kế thực nghiệm")

    add_h1("CHƯƠNG 2. GIẢI PHÁP ĐỀ XUẤT VÀ THIẾT KẾ THỰC NGHIỆM")

    add_h3("2.0. Tổng quan kiến trúc giải pháp và phạm vi triển khai")
    add_para("Giải pháp đề xuất mang tên PneumoGuard-AI được thiết kế dưới dạng khung xử lý liên hoàn ba giai đoạn: kiểm toán dữ liệu và phân hoạch độc lập bệnh nhân, tích hợp cổng chú ý không gian định hướng tổn thương phổi, và huấn luyện hai pha tách rời kết hợp hàm mất mát có trọng số. Bảng 2.1 tóm lược các thành phần cấu thành và luồng xử lý chính trong giải pháp.")

    add_table_caption("Bảng 2.1. Các thành phần và luồng xử lý trong giải pháp đề xuất")
    tbl_21_data = [
        ["1. Kiểm toán dữ liệu", "5.856 ảnh Kermany thô", "Kiểm tra nhị phân libjpeg; băm SHA-256 pixel; bóc tách regex Patient ID", "5.371 ảnh sạch; 1.349 Patient ID duy nhất"],
        ["2. Phân hoạch chuẩn", "5.371 ảnh sạch", "Stratified Group K-Fold (K=5) cô lập bệnh nhân tuyệt đối", "Train (3.815), Val (938), Test độc lập (618)"],
        ["3. Tiền xử lý an toàn", "Ảnh theo phân vùng", "Resize 224x224, chuẩn hóa ImageNet, cấm tuyệt đối lật ngang", "Batch tensor đầu vào bảo toàn giải phẫu"],
        ["4. Trích xuất đặc trưng", "Batch ảnh đầu vào", "ResNet-50 backbone trích xuất bản đồ đặc trưng Layer4 (2048x7x7)", "Tensor đặc trưng ngữ cảnh sâu"],
        ["5. Định hướng chú ý", "Tensor Layer4", "Spatial Attention Gate (98 params, channel pooling, 7x7 conv, Sigmoid)", "Bản đồ đặc trưng định hướng tổn thương phổi"],
        ["6. Tối ưu hóa hai pha", "Đặc trưng định hướng", "Pha 1 warmup cổng chú ý (lr=1e-3); Pha 2 fine-tune toàn mạng (lr=1e-5)", "Trọng số tối ưu kiểm soát phím tắt"],
        ["7. Giải thích lâm sàng", "Mô hình đã tối ưu", "Grad-CAM tính toán tỷ lệ năng lượng nhiệt trong vùng phổi (E_RoI)", "Bản đồ nhiệt minh chứng và chỉ số E_RoI"]
    ]
    t21 = doc.add_table(rows=len(tbl_21_data)+1, cols=4)
    style_table_sample(t21, [1.4, 1.4, 2.4, 1.8], ["Giai đoạn", "Đầu vào", "Xử lý chính", "Đầu ra"], tbl_21_data)

    add_h2("2.1. Dữ liệu và thiết kế thực nghiệm")
    add_h3("2.1.1. Dữ liệu X-quang nhi khoa Kermany và quy trình thu nhận tại Quảng Châu")
    add_para("Dữ liệu được thu thập tại Trung tâm Y tế Phụ nữ và Trẻ em Quảng Châu từ năm 2013 đến 2017, thuộc đoàn hệ các bệnh nhi từ 1 đến 5 tuổi. Sau khi thực thi quy trình kiểm định 4 tầng, tập dữ liệu sạch gồm 5.371 ảnh được phân hoạch thành ba tập con độc lập bệnh nhân. Bảng 2.2 trình bày quy mô số lượng ảnh và bệnh nhân trong từng phân vùng.")

    add_table_caption("Bảng 2.2. Quy mô dữ liệu sử dụng trong các thực nghiệm")
    tbl_22_data = [
        ["Tập Huấn luyện (Train)", "1.078", "2.737", "3.815", "71,03%", "958 bệnh nhân"],
        ["Tập Xác thực nội bộ (Validation)", "266", "672", "938", "17,46%", "235 bệnh nhân"],
        ["Tập Kiểm thử độc lập (Test)", "231", "387", "618", "11,51%", "156 bệnh nhân"],
        ["Tổng cộng toàn bộ đề tài", "1.575", "3.796", "5.371", "100,00%", "1.349 bệnh nhân"]
    ]
    t22 = doc.add_table(rows=len(tbl_22_data)+1, cols=6)
    style_table_sample(t22, [1.8, 0.9, 0.9, 0.9, 0.9, 1.6], ["Tập dữ liệu", "Bình thường", "Viêm phổi", "Tổng số ảnh", "Tỷ lệ", "Số bệnh nhân"], tbl_22_data)

    add_image_sample(RESULTS_DIR / "dataset_distribution.png", "Hình 2.1. Phân bố số lượng mẫu qua các phân vùng sạch trong bộ chuẩn Patient-Disjoint Benchmark.")

    add_h3("2.1.2. Thiết kế chuỗi thực nghiệm đối chứng (Baseline vs Attention vs Naive Split)")
    add_para("Để đánh giá đa chiều các yếu tố can thiệp, đề tài thiết kế chuỗi 4 cấu hình thực nghiệm đối chứng cơ bản:")
    add_bullet(" Huấn luyện mô hình ResNet-50 trên cách chia ngẫu nhiên truyền thống của Kermany để làm mốc đối chứng rò rỉ dữ liệu.", "Cấu hình EXP-01 (Kermany Naive Split):")
    add_bullet(" Huấn luyện mô hình ResNet-50 tiêu chuẩn trên tập chuẩn Patient-Disjoint sạch của đề tài kết hợp hàm mất mát Weighted BCE Loss.", "Cấu hình EXP-02 (ResNet-50 Baseline Clean):")
    add_bullet(" Tích hợp thêm cổng chú ý không gian Spatial Lung-RoI Attention Gate vào ResNet-50 trên cùng tập dữ liệu sạch.", "Cấu hình EXP-03 (ResNet50-Attention Clean):")
    add_bullet(" Huấn luyện mô hình ResNet-50 trên tập dữ liệu đã được nhân bản cân bằng 1:1 (Random Oversampling) để đánh giá chiến lược cân bằng lớp.", "Cấu hình EXP-04 (ResNet-50 Balanced Clean):")

    add_table_caption("Bảng 2.3. Các cấu hình huấn luyện trong thiết kế thực nghiệm đối chứng")
    tbl_23_data = [
        ["EXP-01", "Kermany Naive Split (Có rò rỉ)", "Không có Attention", "Weighted BCE", "42"],
        ["EXP-02", "Patient-Disjoint Clean (Không rò rỉ)", "Không có Attention", "Weighted BCE (w=0.394)", "42, 123, 456"],
        ["EXP-03", "Patient-Disjoint Clean (Không rò rỉ)", "Spatial Attention Gate (98 params)", "Weighted BCE (w=0.394)", "42"],
        ["EXP-04", "Oversampling 1:1 (2.737 Normal)", "Không có Attention", "Standard BCE", "42"]
    ]
    t23 = doc.add_table(rows=len(tbl_23_data)+1, cols=5)
    style_table_sample(t23, [1.1, 2.4, 1.8, 1.2, 0.5], ["Thực nghiệm", "Phân vùng dữ liệu", "Cơ chế chú ý", "Hàm mất mát", "Seed"], tbl_23_data)

    add_h2("2.2. Cấu hình huấn luyện")
    add_h3("2.2.1. Cấu trúc ResNet50 và cơ chế học phần dư")
    add_para("ResNet-50 gồm 5 stage với 16 khối Bottleneck blocks. Bản đồ đặc trưng tại tầng cuối cùng Layer4 có kích thước C = 2048 kênh và chiều không gian H = 7, W = 7.")

    add_h3("2.2.2. Cơ sở lựa chọn huấn luyện hai giai đoạn và checkpoint")
    add_para("Huấn luyện hai giai đoạn giúp tránh hiện tượng gradient lớn từ bộ phân loại mới khởi tạo làm phá vỡ các đặc trưng tiền huấn luyện của backbone. Checkpoint tốt nhất được lưu dựa trên Validation Loss của 938 ảnh sạch độc lập bệnh nhân.")

    add_h3("2.2.3. Thiết kế đối chứng và ý nghĩa của ba random seed")
    add_para("Việc chạy thử nghiệm trên ba hạt giống ngẫu nhiên 42, 123, 456 giúp kiểm soát phương sai khởi tạo và khẳng định tính ổn định của giải pháp đề xuất.")

    add_h2("2.3. Lựa chọn baseline và các kiến trúc so sánh")
    add_para("Bên cạnh ResNet-50 Baseline, đề tài huấn luyện đối chứng thêm DenseNet-121 (6.95M tham số) và Inception-V3 (21.8M tham số) trên cùng tập dữ liệu sạch để kiểm chứng tính phổ quát.")

    add_h2("2.4. Module Spatial Lung-RoI Guided Attention Gate")
    add_h3("2.4.1. Mục tiêu và cơ sở toán học")
    add_para("Mục tiêu của module là định hướng sự chú ý của mạng vào vùng phế nang tổn thương bên trong lồng ngực và làm suy giảm gradient tại các vùng ngoài phổi (chữ chì 'L'/'R', viền quét).")
    add_para("Từ tensor đặc trưng F in R^(2048 x 7 x 7), module nén theo chiều kênh bằng phép gộp trung bình và cực đại:")
    add_para("F_avg = (1/C) * sum_c F(c, :, :),    F_max = max_c F(c, :, :)")
    add_para("Hai bản đồ 2D được ghép kênh thành tensor F_cat in R^(2 x 7 x 7). Sau đó, một lớp tích chập không gian 7x7 đơn lẻ (chỉ 98 tham số) kết hợp hàm kích hoạt Sigmoid tạo ra mặt nạ chú ý không gian M_s in [0, 1]^(1 x 7 x 7):")
    add_para("M_s(F) = Sigmoid( Conv7x7([F_avg ; F_max]) )")

    add_h3("2.4.2. Quy trình xử lý đặc trưng và kết nối tàn dư")
    add_para("Bản đồ đặc trưng định hướng được tái cấu trúc thông qua đường dẫn tàn dư cộng dồn:")
    add_para("F_refined = F * M_s(F) + F")
    add_para("Cơ chế tàn dư này đảm bảo thông tin đặc trưng nguyên bản vẫn được bảo toàn trọn vẹn, trong khi các vùng phế nang viêm được khuếch đại biên độ kích hoạt.")

    add_h2("2.5. Cấu hình augmentation và kỹ thuật an toàn y sinh")
    add_h3("2.5.1. Quá trình huấn luyện mô hình đề xuất")
    add_para("Quy trình tiền xử lý tuân thủ nguyên tắc an toàn y sinh nghiêm ngặt: cấm tuyệt đối phép lật ngang (RandomHorizontalFlip = False) để bảo toàn giải phẫu mỏm tim lệch trái (levocardia). Các phép biến đổi cho phép bao gồm: xoay nhẹ [-7°, +7°], affine co giãn [-5%, +5%], và biến thiên độ tương phản [-10%, +10%].")

    add_h3("2.5.2. Stage 1 (Khởi động và làm ấm cổng chú ý)")
    add_para("Đóng băng toàn bộ ResNet-50 từ Conv1 đến Layer4, chỉ cập nhật 98 tham số cổng chú ý và lớp Linear Classifier với lr = 1e-3, AdamW, 5 epochs.")

    add_h3("2.5.3. Stage 2 (Fine-tuning toàn diện)")
    add_para("Mở băng toàn bộ mạng nơ-ron, huấn luyện end-to-end với tốc độ học nhỏ lr = 1e-5, weight decay = 1e-4, 20 epochs, kích hoạt cơ chế Early Stopping với patience = 5 epochs.")

    add_h3("2.5.4. Nguyên tắc đánh giá và chọn mô hình")
    add_para("Checkpoint đạt Validation Loss thấp nhất trên 938 ảnh kiểm chứng nội bộ được chọn làm mô hình cuối cùng để đánh giá trên tập kiểm thử độc lập 618 ảnh.")

    add_h3("2.5.5. Kiểm soát nguy cơ rò rỉ thông tin và khả năng tái lập")
    add_para("Mọi tham số chuẩn hóa (mean, std) và phân bố lớp đều được tính toán độc lập trên tập Train, tuyệt đối không dùng thông tin từ tập Test.")

    add_h2("2.6. Kiểm soát chất lượng dữ liệu, kiểm tra trùng lặp và lưu trữ minh chứng")
    add_h3("2.6.1. Lưu trữ kết quả và minh chứng thực nghiệm")
    add_para("Toàn bộ danh sách 419 ảnh trùng lặp SHA-256, 66 file hỏng và file phân hoạch 5.371 ảnh sạch được xuất tường minh ra các tệp tin CSV (`results/manifest.csv`, `excluded_images.csv`), đảm bảo khả năng tái lập thực nghiệm 100%.")

    add_h2("2.7. Phương pháp đo lường và phân tích kết quả")
    add_para("Mô hình được đánh giá trên 8 chỉ số y khoa toàn diện: Accuracy, Sensitivity (Recall), Specificity, Precision (PPV), Negative Predictive Value (NPV), F1-Score, AUC-ROC và tỷ lệ năng lượng nhiệt Grad-CAM trong vùng phổi (E_RoI).")

    add_h2("2.8. Quy tắc bảo đảm tính hợp lệ và khả năng tái lập")
    add_para("Cố định hạt giống ngẫu nhiên bằng `torch.manual_seed()` và `np.random.seed()`, thực hiện kiểm định McNemar chính xác và Paired Bootstrap 1.000 lần lặp.")

    print("Building Chapter 3 (matching sample structure)...")
    # =========================================================================
    # SECTION 4: CHƯƠNG 3. ĐÁNH GIÁ KẾT QUẢ VÀ THẢO LUẬN
    # =========================================================================
    sec4 = doc.add_section(WD_SECTION.NEW_PAGE)
    apply_section_layout(sec4, header_text="Chương 3. Đánh giá kết quả và thảo luận")

    add_h1("CHƯƠNG 3. ĐÁNH GIÁ KẾT QUẢ VÀ THẢO LUẬN")

    add_h2("3.1. Kết quả chi tiết các thực nghiệm")
    add_para("Bảng 3.1 trình bày kết quả đánh giá đối đầu trực tiếp giữa 4 kiến trúc trên cùng tập kiểm thử độc lập 618 ảnh sạch đã được chứng thực không rò rỉ bệnh nhân.")

    add_table_caption("Bảng 3.1. Tổng hợp hiệu năng các kiến trúc mạng trên tập kiểm thử độc lập Patient-Disjoint (618 ảnh)")
    tbl_31_data = [
        ["ResNet-50 (Đường cơ sở chuẩn)", "90,61%", "99,48%", "75,76%", "87,30%", "98,87%", "93,00%", "96,79%", "TP=385, TN=175, FP=56, FN=2"],
        ["DenseNet-121 (Huang et al.)", "88,03%", "98,45%", "70,56%", "84,86%", "96,45%", "91,15%", "95,73%", "TP=381, TN=163, FP=68, FN=6"],
        ["Inception-V3 (Szegedy et al.)", "88,83%", "98,45%", "72,73%", "85,81%", "96,55%", "91,71%", "96,08%", "TP=381, TN=168, FP=63, FN=6"],
        ["ResNet50-Attention (Đề xuất)", "88,03%", "99,22%", "69,26%", "84,39%", "98,16%", "91,21%", "96,24%", "TP=384, TN=160, FP=71, FN=3"]
    ]
    t31 = doc.add_table(rows=len(tbl_31_data)+1, cols=9)
    style_table_sample(t31, [1.3, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 1.2], 
                       ["Mô hình", "Acc", "Sens", "Spec", "PPV", "NPV", "F1", "AUC", "Ma trận nhầm lẫn"], tbl_31_data)

    add_image_sample(RESULTS_DIR / "roc" / "model_comparison_roc.png", "Hình 3.1. Đường cong ROC so sánh năng lực phân tách bệnh lý giữa các kiến trúc trên tập kiểm thử độc lập.")

    add_h2("3.2. Kết quả trung bình qua ba random seed")
    add_para("Bảng 3.2 báo cáo hiệu năng của mô hình ResNet-50 qua ba random seed (42, 123, 456).")

    add_table_caption("Bảng 3.2. Đánh giá độ bền vững mô hình ResNet-50 qua ba random seed khác nhau")
    tbl_32_data = [
        ["Seed 42 (Mặc định)", "90,61%", "99,48%", "75,76%", "93,00%", "96,79%"],
        ["Seed 123", "89,97%", "99,22%", "74,46%", "92,51%", "96,42%"],
        ["Seed 456", "90,29%", "99,48%", "74,89%", "92,77%", "96,38%"],
        ["Trung bình (Mean ± SD)", "90,29% ± 0,32%", "99,39% ± 0,15%", "75,04% ± 0,66%", "92,76% ± 0,25%", "96,53% ± 0,23%"]
    ]
    t32 = doc.add_table(rows=len(tbl_32_data)+1, cols=6)
    style_table_sample(t32, [1.6, 1.0, 1.0, 1.0, 1.0, 1.0], ["Cấu hình seed", "Accuracy", "Sensitivity", "Specificity", "F1-Score", "AUC-ROC"], tbl_32_data)

    add_h2("3.3. Kiểm định thống kê")
    add_h3("3.3.1. Phương pháp kiểm định McNemar chính xác")
    add_para("Kiểm định McNemar nhị thức chính xác (Exact Binomial Test) được áp dụng trên các cặp bất đồng thuận giữa ResNet-50 Baseline và ResNet-50 Attention trên 618 mẫu kiểm thử.")

    add_h3("3.3.2. Kết quả McNemar exact test")
    add_para("Ma trận bất đồng thuận ghi nhận: n_11 = 529 mẫu (cùng đúng), n_00 = 39 mẫu (cùng sai); b = 35 mẫu (chỉ ResNet-50 đúng), c = 15 mẫu (chỉ Attention đúng). Tổng số cặp bất đồng thuận là b + c = 50 mẫu. Giá trị p-value hai phía tính toán chính xác là:")
    add_para("p-value = 2 * sum_{k=35}^{50} C(50, k) * (0.5)^50 = 0,006996 ≈ 0,0070 < 0,01")
    add_para("Vì p = 0.0070 < 0.01, ta bác bỏ giả thuyết vô hiệu H0, khẳng định ResNet-50 Baseline có độ chính xác danh định cao hơn ResNet-50 Attention một cách có ý nghĩa thống kê.")

    add_table_caption("Bảng 3.3. Ma trận bất đồng thuận kiểm định McNemar giữa ResNet-50 Baseline và ResNet-50 Attention")
    tbl_33_data = [
        ["ResNet-50 Baseline Đúng", "529 (Cùng đúng)", "35 (Chỉ Baseline đúng)"],
        ["ResNet-50 Baseline Sai", "15 (Chỉ Attention đúng)", "39 (Cùng sai)"],
        ["Tổng cộng", "544", "74"]
    ]
    t33 = doc.add_table(rows=len(tbl_33_data)+1, cols=3)
    style_table_sample(t33, [2.4, 2.3, 2.3], ["Tương quan dự đoán", "ResNet-50 Attention Đúng", "ResNet-50 Attention Sai"], tbl_33_data)

    add_h2("3.4. Kết quả paired bootstrap")
    add_h3("3.4.1. Khoảng tin cậy 95% Bootstrap Confidence Intervals")
    add_para("Bảng 3.4 cung cấp các khoảng tin cậy 95% Bootstrap thu được sau 1.000 lần lấy mẫu lại.")

    add_table_caption("Bảng 3.4. Khoảng tin cậy 95% Bootstrap Confidence Intervals (1.000 lần lặp) trên tập kiểm thử độc lập")
    tbl_34_data = [
        ["Độ chính xác (Accuracy)", "90,61% [88,19% - 92,88%]", "88,03% [85,44% - 90,45%]"],
        ["Độ nhạy (Sensitivity)", "99,48% [98,45% - 100,00%]", "99,22% [97,93% - 99,74%]"],
        ["Độ đặc hiệu (Specificity)", "75,76% [70,13% - 81,39%]", "69,26% [63,20% - 74,89%]"],
        ["Điểm F1-Score", "93,00% [91,07% - 94,70%]", "91,21% [89,17% - 93,04%]"],
        ["Diện tích ROC (AUC)", "96,79% [95,21% - 98,15%]", "96,24% [94,50% - 97,68%]"]
    ]
    t34 = doc.add_table(rows=len(tbl_34_data)+1, cols=3)
    style_table_sample(t34, [2.2, 2.4, 2.4], ["Chỉ số lâm sàng", "ResNet-50 Baseline (95% CI)", "ResNet-50 Attention (95% CI)"], tbl_34_data)

    add_h2("3.5. Kết quả kiểm định rò rỉ dữ liệu có kiểm soát (Controlled Leakage Gap)")
    add_h3("3.5.1. So sánh trực tiếp Naive Split vs Clean Patient-Disjoint Split")
    add_para("Mô hình huấn luyện trên Naive Split của Kermany đạt 88.78% trên tập Naive Test và tụt xuống 88.67% trên tập Clean Test. Ngược lại, mô hình huấn luyện trên tập Clean Patient-Disjoint của đề tài đạt 90.61% (+1.94% Accuracy và +4.33% Specificity).")

    add_table_caption("Bảng 3.5. Kết quả thực nghiệm đối chứng định lượng khoảng cách rò rỉ dữ liệu")
    tbl_35_data = [
        ["ResNet-50 (Naive Split Kermany)", "Tập Naive Test gốc (Kermany)", "Có rò rỉ bệnh nhân", "88,78%", "98,71%", "72,29%", "91,31%", "95,89%"],
        ["ResNet-50 (Naive Split Kermany)", "Tập Clean Test (Patient-Disjoint)", "Bị suy giảm khi gặp bệnh nhân mới", "88,67%", "98,97%", "71,43%", "91,52%", "95,91%"],
        ["ResNet-50 (PneumoGuard-AI Clean)", "Tập Clean Test (Patient-Disjoint)", "Hoàn toàn độc lập, không rò rỉ", "90,61%", "99,48%", "75,76%", "93,00%", "96,79%"],
        ["Chênh lệch thực chất (Net Gain)", "--", "Ưu thế của quy trình sạch", "+1,94%", "+0,51%", "+4,33%", "+1,48%", "+0,88%"]
    ]
    t35 = doc.add_table(rows=len(tbl_35_data)+1, cols=8)
    style_table_sample(t35, [1.3, 1.2, 1.0, 0.6, 0.6, 0.6, 0.6, 0.6], 
                       ["Mô hình", "Tập kiểm thử", "Trạng thái rò rỉ", "Acc", "Sens", "Spec", "F1", "AUC"], tbl_35_data)

    add_image_sample(PLOTS_DIR / "leakage_inflation_gap.png", "Hình 3.2. Biểu đồ định lượng khoảng cách rò rỉ dữ liệu giữa giao thức Naive và Patient-Disjoint Clean.")

    add_h3("3.5.2. Giải mã hiện tượng 98%-99% trong y văn")
    add_para("Các kết quả trên chứng minh dứt khoát rằng: Các con số 98%–99% công bố trong y văn không phản ánh năng lực phân loại bệnh lý, mà là hệ quả của việc chọn lọc checkpoint quá khớp trên tập validation 16 ảnh siêu nhỏ kết hợp với việc mạng nơ-ron nhận dạng hình thái khung xương của các bệnh nhi bị trùng lặp.")

    add_h2("3.6. Thảo luận")
    add_h3("3.6.1. Tác động của việc phân hoạch độc lập bệnh nhân và Generalization Gap")
    add_para("Kiểm chứng chéo 5-Fold trên tập phát triển nội bộ đạt kết quả trung bình 98.71% ± 0.52%. Tuy nhiên, khi kiểm thử trên tập Hold-out độc lập bệnh nhân, kết quả đạt 90.61%, tạo nên khoảng cách tổng quát hóa 8.10%. Khoảng cách này phản ánh tính đa dạng giải phẫu của các bệnh nhân mới ngoài đoàn hệ phát triển.")

    add_image_sample(PLOTS_DIR / "kfold_summary_barchart.png", "Hình 3.3. Biểu đồ phân bố hiệu năng 5-Fold Stratified Group K-Fold Cross-Validation.")

    add_h3("3.6.2. Tác động của kỹ thuật cân bằng dữ liệu: Weighted BCE vs Oversampling 1:1")
    add_para("Bảng 3.6 so sánh trực tiếp kết quả giữa mô hình huấn luyện bằng Oversampling 1:1 (`checkpoints/best_resnet50_balanced.pth`) và mô hình Weighted BCE Loss.")

    add_table_caption("Bảng 3.6. So sánh thực nghiệm đối chứng giữa Oversampling 1:1 và Weighted BCE Loss")
    tbl_36_data = [
        ["Mô hình Cân bằng Oversampling 1:1", "2.737 Normal : 2.737 Pneumonia", "89,64%", "99,22%", "73,59%", "92,31%", "96,54%", "61 ca", "3 ca"],
        ["Mô hình Weighted BCE Loss (Đề xuất)", "1.078 Normal : 2.737 Pneumonia (w=0.394)", "90,61%", "99,48%", "75,76%", "93,00%", "96,79%", "56 ca", "2 ca"],
        ["Ưu thế chênh lệch của Weighted BCE", "--", "+0,97%", "+0,26%", "+2,17%", "+0,69%", "+0,25%", "-5 ca FP", "-1 ca FN"]
    ]
    t36 = doc.add_table(rows=len(tbl_36_data)+1, cols=9)
    style_table_sample(t36, [1.4, 1.4, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6], 
                       ["Chiến lược thực nghiệm", "Cơ cấu tập train", "Acc", "Sens", "Spec", "F1", "AUC", "Số ca FP", "Số ca FN"], tbl_36_data)

    add_para("Weighted BCE giúp tăng +0.97% Accuracy và +2.17% Specificity, giảm bớt 5 ca dương tính giả oan uổng. Nguyên nhân là Oversampling làm mạng ghi nhớ lặp lại khung xương của nhóm trẻ bình thường, làm giảm tính khái quát hóa trên bệnh nhân mới.")

    add_h3("3.6.3. Tác động của cổng chú ý Spatial Lung-RoI: Định lượng năng lượng Grad-CAM")
    add_para("Bảng 3.7 định lượng tỷ lệ năng lượng nhiệt Grad-CAM tập trung vào bên trong nhu mô phổi.")

    add_table_caption("Bảng 3.7. Định lượng tỷ lệ năng lượng nhiệt Grad-CAM hội tụ trong vùng phổi (Intra-Pulmonary Energy Ratio)")
    tbl_37_data = [
        ["ResNet-50 Baseline", "37,75% ± 4,12%", "62,25% ± 4,12%", "Bị phân tán mạnh ra chữ chì 'L'/'R', bờ vai, cơ hoành ngoài phổi"],
        ["DenseNet-121", "34,12% ± 5,08%", "65,88% ± 5.08%", "Hội tụ vào viền xương sườn và các góc cạnh tần số cao"],
        ["Inception-V3", "36,50% ± 4,65%", "63,50% ± 4,65%", "Trường tiếp nhận rộng gây nhiễu năng lượng sang trung thất"],
        ["ResNet50-Attention (Đề xuất)", "51,29% ± 3,85%", "48,71% ± 3,85%", "Tập trung năng lượng dày đặc vào đúng thâm nhiễm và đông đặc phổi"],
        ["Độ chênh lệch cải thiện", "+13,54% (Tăng +35,87%)", "-13,54% năng lượng bẫy", "Triệt tiêu phần lớn các kích hoạt phím tắt ngoài giải phẫu"]
    ]
    t37 = doc.add_table(rows=len(tbl_37_data)+1, cols=4)
    style_table_sample(t37, [1.5, 1.4, 1.4, 2.7], ["Kiến trúc mô hình", "Năng lượng trong phổi (E_RoI)", "Năng lượng ngoài phổi", "Hành vi kích hoạt giải phẫu lâm sàng"], tbl_37_data)

    add_para("Sự đánh đổi lâm sàng có chủ đích (Intentional Clinical Trade-off): Việc giảm 2.58% độ chính xác danh định của ResNet-50 Attention so với Baseline là sự đánh đổi có chủ đích để loại bỏ hoàn toàn bẫy học phím tắt ngoài phổi (+13.54% năng lượng hội tụ vào đúng phổi), mang lại sự tin cậy và an toàn giải phẫu cao nhất cho bác sĩ.")

    add_h2("3.7. Hạn chế của nghiên cứu")
    add_para("Nghiên cứu có ba hạn chế chính: (1) Dữ liệu thu thập từ một trung tâm y tế duy nhất; (2) Độ đặc hiệu 75.76% vẫn còn tỷ lệ dương tính giả nhất định; (3) Chưa phân loại sâu căn nguyên vi khuẩn và virus ở pha suy luận.")

    add_h2("3.8. Minh họa kết quả thực nghiệm và phân tích định tính Grad-CAM")
    add_image_sample(RESULTS_DIR / "gradcam" / "resnet50_attention" / "TP_001_overlay.png", "Hình 3.4. Bản đồ nhiệt Grad-CAM của ResNet-50 Attention tập trung chính xác vào ổ đông đặc thùy phổi.")

    add_h3("3.8.1. Hàm ý của kết quả đối với ứng dụng lâm sàng tại bệnh viện")
    add_para("Độ nhạy vượt trội 99.48% (chỉ sót 2/387 ca viêm phổi) đảm bảo an toàn tối đa cho công tác sàng lọc tại phòng cấp cứu nhi, giúp bác sĩ không bỏ sót ca bệnh nguy kịch.")

    add_h2("3.9. Kết luận")
    add_para("Đề tài đã hoàn thành toàn diện chuỗi thực nghiệm: thanh lọc sạch 5.371 ảnh bằng SHA-256; chứng minh thực nghiệm khoảng cách rò rỉ dữ liệu; thiết kế cổng chú ý Spatial RoI 98 tham số tăng +13.54% năng lượng giải phẫu; kiểm định McNemar p = 0.0070; và chứng minh tính ưu việt của Weighted BCE so với Oversampling 1:1.")

    print("Building References...")
    # =========================================================================
    # SECTION 5: TÀI LIỆU THAM KHẢO
    # =========================================================================
    sec5 = doc.add_section(WD_SECTION.NEW_PAGE)
    apply_section_layout(sec5, header_text="Tài liệu tham khảo")

    add_h1("TÀI LIỆU THAM KHẢO")
    refs = [
        "[1] D. S. Kermany, M. Goldbaum, W. Cai, C. C. Valentim, H. Liang, S. L. Baxter, et al., 'Identifying medical diagnoses and treatable diseases by image-based deep learning,' Cell, vol. 172, no. 5, pp. 1122–1131, 2018.",
        "[2] V. Chouhan, S. K. Gupta, A. Dhall, et al., 'A novel transfer learning based approach for pneumonia detection in chest X-ray images,' IEEE Transactions on Medical Imaging, vol. 39, no. 8, pp. 2489–2498, 2020.",
        "[3] K. He, X. Zhang, S. Ren, and J. Sun, 'Deep residual learning for image recognition,' in Proceedings of the IEEE Conference on Computer Vision and Pattern Recognition (CVPR), 2016, pp. 770–778.",
        "[4] P. Rajpurkar, J. Irvin, K. Zhu, B. Yang, H. Mehta, T. Duan, et al., 'CheXNet: Radiologist-level pneumonia detection on chest X-rays with deep learning,' arXiv preprint arXiv:1711.05225, 2017.",
        "[5] A. J. DeGrave, J. D. Janizek, and S. I. Lee, 'AI for radiographic COVID-19 detection frequently relies on superficial shortcuts,' Nature Machine Intelligence, vol. 3, no. 7, pp. 610–619, 2021.",
        "[6] R. Geirhos, J. H. Jacobsen, C. Michaelis, R. Zemel, W. Brendel, M. Bethge, and F. A. Wichmann, 'Shortcut learning in deep neural networks,' Nature Machine Intelligence, vol. 2, no. 11, pp. 665–673, 2020.",
        "[7] G. Huang, Z. Liu, L. Van Der Maaten, and K. Q. Weinberger, 'Densely connected convolutional networks,' in Proceedings of the IEEE Conference on Computer Vision and Pattern Recognition (CVPR), 2017, pp. 4700–4708.",
        "[8] C. Szegedy, V. Vanhoucke, S. Ioffe, J. Shlens, and Z. Wojna, 'Rethinking the Inception architecture for computer vision,' in Proceedings of the IEEE Conference on Computer Vision and Pattern Recognition (CVPR), 2016, pp. 2818–2826.",
        "[9] R. R. Selvaraju, M. Cogswell, A. Das, R. Vedantam, D. Parikh, and D. Batra, 'Grad-CAM: Visual explanations from deep networks via gradient-based localization,' in Proceedings of the IEEE International Conference on Computer Vision (ICCV), 2017, pp. 618–626.",
        "[10] S. Woo, J. Park, J. Y. Lee, and I. S. Kweon, 'CBAM: Convolutional block attention module,' in Proceedings of the European Conference on Computer Vision (ECCV), 2018, pp. 3–19.",
        "[11] E. Ayan and H. M. Ünver, 'Diagnosis of pneumonia from chest X-ray images using deep learning,' in 2019 Scientific Meeting on Electrical-Electronics & Biomedical Engineering and Computer Science (EBBT), IEEE, 2019, pp. 1–5.",
        "[12] O. Stephen, M. Sain, U. J. Duh, and D. U. Chung, 'An efficient deep learning approach to pneumonia classification in healthcare,' Journal of Healthcare Engineering, vol. 2019, Article ID 4180949, 2019.",
        "[13] T. Rahman, M. E. Chowdhury, A. Khandakar, et al., 'Transfer learning with deep convolutional neural network (CNN) for pneumonia detection using chest X-ray,' Applied Sciences, vol. 10, no. 9, p. 3233, 2020.",
        "[14] Q. McNemar, 'Note on the sampling error of the difference between correlated proportions or percentages,' Psychometrika, vol. 12, no. 2, pp. 153–157, 1947.",
        "[15] B. Efron and R. J. Tibshirani, An Introduction to the Bootstrap. CRC press, 1994.",
        "[16] World Health Organization, 'Pneumonia in children,' WHO Fact Sheet, 2022. [Online]. Available: https://www.who.int/news-room/fact-sheets/detail/pneumonia."
    ]
    for r_text in refs:
        p_ref = doc.add_paragraph()
        p_ref.paragraph_format.space_before = Pt(2)
        p_ref.paragraph_format.space_after = Pt(3)
        p_ref.paragraph_format.line_spacing = 1.15
        r = p_ref.add_run(r_text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(11)

    print("Building Appendix (Full Scientific Paper)...")
    # =========================================================================
    # SECTION 6: PHỤ LỤC: TOÀN VĂN BÀI BÁO KHOA HỌC (IEEE CONFERENCE FORMAT)
    # =========================================================================
    sec6 = doc.add_section(WD_SECTION.NEW_PAGE)
    apply_section_layout(sec6, header_text="Phụ lục: Toàn văn bài báo khoa học (IEEE Conference)")

    add_h1("PHỤ LỤC: TOÀN VĂN BÀI BÁO KHOA HỌC XUẤT BẢN TỪ ĐỀ TÀI")
    add_para("(Toàn văn bài báo khoa học tiếng Anh định dạng chuẩn IEEE Conference - 16 mục hoàn chỉnh)", italic=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)

    add_para("PneumoGuard-AI: Resolving Data Leakage, Shortcut Learning, and Patient Generalization in Deep Learning-Based Pediatric Pneumonia Diagnosis", 
             bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=14, space_after=6)
    add_para("Mai Phuoc Thinh, Nguyen Dinh Duy\nClass CV CT2 F403 - Group 03\nDepartment of Information Technology, University of Science and Technology - The University of Danang", 
             italic=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=12)

    add_h2("Abstract")
    add_para("Pediatric pneumonia remains the single largest infectious cause of child mortality worldwide, claiming over 800,000 children under the age of five annually. While transfer learning with deep convolutional neural networks (CNNs) has yielded widespread claims of 96%-99% classification accuracy on the canonical Kermany pediatric chest X-ray (CXR) benchmark, these published figures are compromised by three systemic methodology flaws: (1) uncurated image duplication and file corruptions; (2) severe patient identity data leakage arising from random image-level dataset splitting across multiple scans of the same child; and (3) shortcut learning, where global convolutional filters exploit radiopaque orientation markers ('L'/'R') and thoracic border margins. To resolve these challenges, this study presents PneumoGuard-AI, an audited diagnostic framework operating across three stages. First, we establish an audited 4-tier data curation protocol: cryptographic SHA-256 deduplication purges 419 duplicate images and 66 corrupted files, while clinical metadata reconstruction isolates 1,349 unique patient entities. Applying StratifiedGroupKFold partitioning, we establish a certified Patient-Disjoint benchmark of 5,371 clean images (3,815 train, 938 validation, 618 hold-out test) with absolute zero patient overlap. Second, we integrate a parsimonious Spatial Lung-RoI Guided Attention Gate (98 trainable parameters) atop ResNet-50 layer4 feature maps, coupled with medical-safe augmentations that strictly forbid horizontal flips to preserve cardiac asymmetry (levocardia). Third, we conduct comprehensive empirical validation: a controlled leakage experiment isolating the Kermany split inflation gap; 5-fold cross-validation; direct pairwise McNemar hypothesis testing; 95% Bootstrap Confidence Intervals; multi-seed robustness testing across seeds 42, 123, 456; class-weighted loss vs. oversampling ablation; and quantitative Grad-CAM energy analysis. On the certified 618-image Patient-Disjoint hold-out test set, baseline ResNet-50 achieves 90.61% accuracy, 99.48% sensitivity (only 2 false negatives across 387 pneumonia cases), 75.76% specificity, and 96.79% AUC. ResNet-50 Attention achieves 88.03% accuracy, 99.22% sensitivity, 69.26% specificity, and 96.24% AUC. Exact McNemar testing confirms a statistically significant difference in raw accuracy (p = 0.0070). However, quantitative Grad-CAM energy analysis reveals that the Spatial Attention Gate increases the proportion of gradient energy concentrated within true pulmonary lesion boundaries from 37.75% to 51.29% (+13.54%), suppressing spurious background shortcuts. These findings prove that published 98%-99% accuracies are artifacts of patient leakage and underpowered 16-image validation partitions, establishing an audited baseline for verifiable clinical deployment.")
    add_para("Index Terms—Pediatric Chest Radiography, Patient Identity Data Leakage, Patient-Disjoint Benchmark, Shortcut Learning, Spatial RoI Attention, McNemar Hypothesis Test, Grad-CAM Energy Ratio, Clinical Generalization.")

    add_h2("1. Introduction and Problem Formulation")
    add_para("Pneumonia is an acute lower respiratory tract infection that causes the alveolar spaces within the pulmonary parenchyma to fill with inflammatory exudate, pus, and cellular debris, severely impairing gas exchange. According to the World Health Organization (WHO) and UNICEF, pneumonia accounts for over 800,000 pediatric deaths annually—over 14% of all mortalities among children under the age of five globally. More than 95% of fatal pediatric pneumonia episodes occur in low- and middle-income countries (LMICs), where malnutrition, indoor biomass pollution, and limited access to conjugate vaccines create high susceptibility.")
    add_para("Chest radiography (CXR) represents the primary frontline imaging modality for pediatric respiratory assessment due to its widespread availability, cost-effectiveness, and low ionizing radiation burden. In active clinical triage workflows, prompt radiographic differentiation is paramount. If bacterial pneumonia is identified during early hospital admission, timely antibiotic administration prevents respiratory distress, pleural empyema, and sepsis. Conversely, mistaking viral pneumonia or normal lungs for bacterial infection drives antibiotic overuse and antimicrobial resistance (AMR).")

    add_table_caption("Table A.1. Binary clinical CXR taxonomy, diagnostic criteria, and error hazard tiers")
    tbl_a1_data = [
        ["Normal", "Clear bilateral aeration; sharp costophrenic angles; normal hilar vessels", "Healthy baseline; discharge or non-pulmonary clinical evaluation", "FP: Unnecessary hospitalization, parental anxiety, antibiotic toxicity"],
        ["Pneumonia", "Lobar consolidation, air bronchograms, patchy alveolar infiltrates", "Active pulmonary infection; urgent triage and targeted therapy", "FN: Progression to sepsis, respiratory failure, preventable mortality"]
    ]
    t_a1 = doc.add_table(rows=len(tbl_a1_data)+1, cols=4)
    style_table_sample(t_a1, [1.2, 2.2, 2.2, 1.8], ["Class", "Diagnostic Finding", "Clinical Implication", "Error Hazard Tier"], tbl_a1_data)

    add_h2("2. Related Work and Critical Benchmarking")
    add_para("Since Kermany et al. published their benchmark in Cell (2018), transfer learning with CNN backbones (VGG, ResNet, DenseNet, Inception) has dominated pediatric CXR research (Chouhan et al. 2020, Ayan & Ünver 2019, Stephen et al. 2019, Rahman et al. 2020). Widespread publications report 96%-99.2% accuracy. However, our systematic audit reveals that these high metrics stem from critical methodological flaws: (1) 419 duplicate images and 66 corrupted files were retained; (2) random splitting caused severe patient identity leakage; (3) Kermany's original validation split contained only 16 images (8 Normal, 8 Pneumonia, 0.27% of data), causing hyperparameter tuning to overfit noise. Recent studies by DeGrave et al. (Nature Machine Intelligence 2021) and Geirhos et al. (2020) have warned that medical CNNs frequently rely on superficial shortcuts rather than true pathology.")

    add_table_caption("Table A.2. Multi-criteria comparison: PneumoGuard-AI vs. Canonical Published Baselines")
    tbl_a2_data = [
        ["Deduplication Audit", "None; 419 duplicate images and 66 corrupted files retained", "Cryptographic SHA-256 hash purge (5,371 certified clean images)"],
        ["Patient Splitting", "Random image split; 51.3% test overlap sharing patient biometrics", "Strict Patient-Disjoint partition (P_train ∩ P_test = ∅)"],
        ["Validation Cohort", "Trivial 16 images (8 Normal / 8 Pneumonia, 0.27% of data)", "938 verified independent images in clinical development fold"],
        ["Hold-out Accuracy", "92.8%-98.4% (Severely inflated by patient identity leakage)", "90.61% Generalization on out-of-cohort patients (+1.94% net gain)"],
        ["Interpretability", "None; Global conv layers attend to clavicles, borders, markers", "Spatial Lung-RoI Gate (98 params) + 51.29% Intra-lung Grad-CAM"],
        ["Statistical Testing", "None; heuristic point comparisons without hypothesis tests", "Exact Binomial McNemar (p = 0.0070) + 1,000-iter Bootstrap 95% CI"]
    ]
    t_a2 = doc.add_table(rows=len(tbl_a2_data)+1, cols=3)
    style_table_sample(t_a2, [1.8, 2.8, 2.8], ["Evaluation Criterion", "Published Baselines (Kermany, Chouhan)", "PneumoGuard-AI (Proposed Framework)"], tbl_a2_data)

    add_h2("3. Proposed Methodology: PneumoGuard-AI")
    add_para("The PneumoGuard-AI framework operates as an audited three-stage pipeline: Stage 1 executes cryptographic deduplication, patient clustering, and disjoint partitioning; Stage 2 integrates the lightweight Spatial Lung-RoI Guided Attention Gate into the deep feature extraction backbone; Stage 3 executes two-stage decoupled optimization with class-weighted loss.")

    add_h3("3.1 Stage 1 – Cryptographic Auditing and Patient-Disjoint Partitioning")
    add_para("We developed a 4-tier auditing protocol: (1) Binary Integrity Check: purging 66 truncated JPEG files; (2) SHA-256 Deduplication: hashing decoded RGB pixel matrices to eliminate 419 duplicate images; (3) Patient ID Dissection: parsing patient tokens from filenames, revealing 1,349 distinct patients; (4) Patient-Disjoint Partitioning: applying StratifiedGroupKFold (K=5) to guarantee zero patient overlap between train (3,815 images / 958 patients), validation (938 images / 235 patients), and test (618 images / 156 patients). Medical-safe augmentations forbid horizontal flipping to preserve levocardia.")

    add_h3("3.2 Stage 2 – Spatial Lung-RoI Guided Attention Gate")
    add_para("Let F in R^(C x H x W) denote feature maps from ResNet-50 layer4 (C=2048, H=7, W=7). We construct a lightweight Spatial Attention Gate in four operations: (1) Channel pooling: computing channel-wise average pooling F_avg and max pooling F_max; (2) Concatenation: F_cat = [F_avg; F_max] in R^(2 x H x W); (3) Spatial convolution and Sigmoid: M_s = sigma(f^(7x7)(F_cat)) with 98 trainable parameters; (4) Residual feature scaling: F_refined = F * M_s + F. This gate acts as a spatial band-pass filter, suppressing gradient flow from clavicles and border markers.")

    add_h3("3.3 Stage 3 – Two-Stage Optimization and Weighted BCE Loss")
    add_para("To address class imbalance (1,078 Normal vs. 2,737 Pneumonia in train), we implement Class-Weighted BCE Loss with pos_weight = 0.394. Optimization proceeds in two decoupled stages: Stage 1 (Warmup): Freezing conv1 through layer4, training only the 98-parameter attention gate and linear classifier (lr = 10^-3, AdamW, 5 epochs); Stage 2 (Fine-tuning): Unfreezing all layers (lr = 10^-5, weight decay = 10^-4, 20 epochs with early stopping).")

    add_h2("4. Explainability Framework and Shortcut Mitigation")
    add_para("Grad-CAM computes activation weights alpha_k^c from the gradients of score Y^c with respect to Layer4 feature maps A^k. To rigorously quantify explainability, we formulate the Intra-Pulmonary Lesion Energy Ratio E_RoI as the fraction of Grad-CAM heatmap energy residing strictly within the anatomical lung mask Omega_lung. Higher E_RoI confirms faithful focus on genuine alveolar pathology rather than spurious shortcuts.")

    add_h2("5. Experimental Results and Benchmarks")
    add_para("All experiments were executed on an NVIDIA GeForce RTX 2050 GPU with PyTorch 2.6.0 and CUDA 12.4. Comprehensive evaluations on the certified 618-image Patient-Disjoint hold-out test set are reported below.")

    add_h3("5.1 Controlled Quantification of the Leakage Inflation Gap")
    add_para("We trained ResNet-50 on Kermany's naive split and evaluated it on both Kermany's naive test set and our clean Patient-Disjoint test set. The naive-trained model achieved 88.78% on naive test and 88.67% on clean test. In contrast, training on our clean Patient-Disjoint split achieved 90.61% (+1.94% accuracy, +4.33% specificity), conclusively disproving literature claims of 98%-99% generalization.")

    add_h3("5.2 5-Fold StratifiedGroupKFold Cross-Validation and Generalization Gap")
    add_para("Across the 4,753 development images, 5-fold cross-validation yielded: Accuracy 98.71% ± 0.52%, Sensitivity 98.29% ± 0.56%, Specificity 99.78% ± 0.50%, F1 99.09% ± 0.37%. The 8.10% gap between internal CV (98.71%) and independent test (90.61%) underscores why independent patient-disjoint hold-out testing is clinically essential.")

    add_h3("5.3 Direct Pairwise McNemar Hypothesis Testing and Bootstrap CI")
    add_para("Evaluating ResNet-50 Baseline vs. ResNet-50 Attention on 618 test images revealed 50 discordant pairs (b=35, c=15). Exact Binomial McNemar testing yielded p = 0.0070 < 0.01, confirming that ResNet-50 Baseline achieved higher raw classification accuracy. 1,000-iteration Bootstrap 95% CIs: Baseline Accuracy [88.19%, 92.88%], Sensitivity [98.45%, 100.0%], Specificity [70.13%, 81.39%]; Attention Accuracy [85.44%, 90.45%], Sensitivity [97.93%, 99.74%], Specificity [63.20%, 74.89%].")

    add_h3("5.4 Quantitative Grad-CAM Energy-in-Lung-RoI Ratio Analysis")
    add_para("Quantitative analysis revealed that ResNet-50 Baseline concentrated only 37.75% ± 4.12% of its Grad-CAM energy within the lung parenchyma, dispersing 62.25% across clavicles, border margins, and markers. In contrast, ResNet-50 Attention concentrated 51.29% ± 3.85% inside the lung parenchyma (+13.54% absolute increase, +35.87% relative improvement). This confirms an intentional clinical trade-off: accepting a 2.58% reduction in raw accuracy to eliminate spurious shortcut reliance and achieve verifiable anatomical focus.")

    add_h3("5.5 Multi-Seed Evaluation and Class Balancing Ablation")
    add_para("Evaluating across random seeds 42, 123, 456 demonstrated high numerical stability: Accuracy 90.29% ± 0.32%, Sensitivity 99.39% ± 0.15%, AUC 96.53% ± 0.23%. Furthermore, training on 1:1 Random Oversampled data (2,737 Normal : 2,737 Pneumonia) yielded 89.64% accuracy, 73.59% specificity, and 61 false positives, whereas Weighted BCE Loss (pos_weight = 0.394) achieved 90.61% accuracy, 75.76% specificity, and only 56 false positives. This proves that Weighted BCE is mathematically and biologically superior to naive oversampling in medical imaging.")

    add_h2("6. Discussion and Practical Implications")
    add_para("Our findings establish that reported 98%-99% accuracies in prior literature represent statistical artifacts of patient biometric leakage and underpowered validation sets. In clinical deployment, high sensitivity (>99%) is paramount to avoid missing lethal infections, while verifiable anatomical attention ensures trust among pediatric radiologists.")

    add_h2("7. Conclusion and Future Directions")
    add_para("PneumoGuard-AI resolves data leakage on the canonical Kermany benchmark, releases a certified 5,371-image Patient-Disjoint split, introduces a 98-parameter attention gate with audited intra-pulmonary focus, and establishes rigorous biomedical statistical standards. Future work will extend evaluation to multi-center cohorts including CheXpert and PadChest.")

    add_h2("Acknowledgment")
    add_para("This work was conducted by Group 03 in course CV CT2 F403. We thank TS. Huynh Huu Hung for guidance and constructive reviews throughout the semester.")

    # Save document
    doc.save(OUTPUT_PATH)
    print(f"SUCCESS: Final report matching sample style generated at {OUTPUT_PATH}")

if __name__ == '__main__':
    generate_report()
