import docx
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def test_page_number():
    doc = Document()
    sec = doc.sections[0]
    ftr = sec.footer
    p = ftr.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    # Correct OpenXML structure for PAGE field:
    r1 = p.add_run()
    fld1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    r1._r.append(fld1)
    
    r2 = p.add_run()
    instr = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w'))
    r2._r.append(instr)
    
    r3 = p.add_run()
    fld2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    r3._r.append(fld2)
    
    r4 = p.add_run()
    fld3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    r4._r.append(fld3)
    
    doc.add_paragraph("Hello world")
    doc.save("d:/CODE/tgmt/scripts/test_out.docx")

test_page_number()
print("Saved test_out.docx")

