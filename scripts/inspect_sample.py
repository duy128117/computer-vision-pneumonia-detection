# -*- coding: utf-8 -*-
import docx

doc = docx.Document(r'd:\CODE\tgmt\Báo cáo cuối kì Thị Giác máy tính.docx')

print('=== COVER PAGE (paragraphs 0 to 30) ===')
for i in range(min(32, len(doc.paragraphs))):
    p = doc.paragraphs[i]
    print(f'P{i:02d} [{p.style.name} | align={p.alignment} | before={p.paragraph_format.space_before} | after={p.paragraph_format.space_after} | ls={p.paragraph_format.line_spacing}]: {repr(p.text)}')

print('\n=== SECTIONS & HEADERS/FOOTERS ===')
for i, s in enumerate(doc.sections):
    hdr = ' | '.join(p.text.strip() for p in s.header.paragraphs if p.text.strip())
    ftr = ' | '.join(p.text.strip() for p in s.footer.paragraphs if p.text.strip())
    print(f'Sec {i}: diff_first_page={s.different_first_page_header_footer}, header="{hdr}", footer="{ftr}"')

print('\n=== STYLES DEFINITION ===')
for s in ['Normal', 'Heading 1', 'Heading 2', 'Heading 3']:
    if s in doc.styles:
        st = doc.styles[s]
        f = st.font
        print(f'{s}: font={f.name}, size={f.size.pt if f.size else None}, bold={f.bold}, italic={f.italic}, color={f.color.rgb if f.color else None}')

print('\n=== TABLES INSPECTION ===')
print(f'Total tables: {len(doc.tables)}')
for t_idx, t in enumerate(doc.tables[:5]):
    print(f'\n--- Table {t_idx} (rows={len(t.rows)}, cols={len(t.columns)}) ---')
    for r_idx, row in enumerate(t.rows[:3]):
        row_txt = [c.text.strip().replace('\n', ' ') for c in row.cells]
        print(f'  Row {r_idx}: {row_txt[:4]}')

