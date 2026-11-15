# -*- coding: utf-8 -*-
import docx

doc = docx.Document(r'd:\CODE\tgmt\Báo cáo cuối kì Thị Giác máy tính.docx')

print('=== ALL HEADINGS WITH TEXT ===')
for i, p in enumerate(doc.paragraphs):
    if p.style.name.startswith('Heading') or p.text.startswith('CHƯƠNG') or p.text == 'MỤC LỤC' or p.text == 'TÀI LIỆU THAM KHẢO':
        print(f'{p.style.name:10s} | {p.text}')

