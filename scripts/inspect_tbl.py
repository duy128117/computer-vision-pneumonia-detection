# -*- coding: utf-8 -*-
import docx
import xml.etree.ElementTree as ET

doc = docx.Document(r'd:\CODE\tgmt\Báo cáo cuối kì Thị Giác máy tính.docx')
t = doc.tables[0]
c0 = t.rows[0].cells[0]
print('=== Header Cell 0 tcPr ===')
print(ET.tostring(c0._tc.get_or_add_tcPr(), encoding='unicode'))
print('=== Header Cell 0 text runs ===')
for r in c0.paragraphs[0].runs:
    print(r.text, r.font.name, r.font.size.pt if r.font.size else None, r.bold, r.italic)
