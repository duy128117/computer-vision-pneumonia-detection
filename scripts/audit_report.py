# -*- coding: utf-8 -*-
import docx

doc = docx.Document('d:/CODE/tgmt/baocao_final_viemphoi.docx')
print(f'Total paragraphs: {len(doc.paragraphs)}')
print(f'Total tables: {len(doc.tables)}')

# Section boundaries based on actual heading indices
# 0..89 is front matter (Title, Acknowledgements, Abstract, TOC, List of Tables/Figures)
sections = [
    ('FRONT_MATTER', 0, 89),
    ('CHƯƠNG 1', 90, 138),
    ('CHƯƠNG 2', 139, 209),
    ('CHƯƠNG 3', 210, 257),
    ('CONCLUSION', 258, 270),
    ('REFERENCES', 271, 287),
    ('APPENDIX', 288, len(doc.paragraphs) - 1)
]

for sec_name, start_idx, end_idx in sections:
    words = 0
    paras = 0
    for i in range(start_idx, min(end_idx + 1, len(doc.paragraphs))):
        p = doc.paragraphs[i]
        t = p.text.strip()
        words += len(t.split())
        paras += 1
    print(f"{sec_name:15}: {words:5} words in {paras:3} paragraphs")

# Also count all words across all paragraphs & tables
all_words = sum(len(p.text.strip().split()) for p in doc.paragraphs)
table_words = 0
for tbl in doc.tables:
    for row in tbl.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                table_words += len(p.text.strip().split())

print(f"\nAll Paragraph Words: {all_words}")
print(f"All Table Words: {table_words}")
print(f"Grand Total Words: {all_words + table_words}")
