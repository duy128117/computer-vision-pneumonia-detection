import docx

doc = docx.Document("d:/CODE/tgmt/Paper_Pneumonia_Classification.docx")
print("Paper sections:", len(doc.sections))
for i, sec in enumerate(doc.sections):
    cols = sec._sectPr.xpath('w:cols')
    col_str = cols[0].attrib if cols else "None"
    print(f"Sec {i}: orientation={sec.orientation}, cols={col_str}")
print("Paper tables:", len(doc.tables))
for i, t in enumerate(doc.tables):
    col_w = [c.width for c in t.columns]
    print(f"Table {i}: rows={len(t.rows)}, cols={len(t.columns)}, width={col_w}")

