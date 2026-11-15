# -*- coding: utf-8 -*-
import docx

doc = docx.Document(r'd:\CODE\tgmt\Báo cáo cuối kì Thị Giác máy tính.docx')

print('=== HEADER SEC 2 XML ===')
print(doc.sections[2].header.paragraphs[0]._p.xml)

print('\n=== FOOTER SEC 2 XML ===')
print(doc.sections[2].footer.paragraphs[0]._p.xml)

print('\n=== TABLE 0 BORDERS XML ===')
print(doc.tables[0]._tbl.tblPr.xml)

print('\n=== TABLE 0 CELL (0,0) XML ===')
print(doc.tables[0].rows[0].cells[0]._tc.xml)

print('\n=== TABLE 0 CELL (1,0) XML ===')
print(doc.tables[0].rows[1].cells[0]._tc.xml)
