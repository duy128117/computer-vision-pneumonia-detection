import docx

doc = docx.Document("d:/CODE/tgmt/Báo cáo cuối kì Thị Giác máy tính.docx")
print("Sample sections:", len(doc.sections))
for i, sec in enumerate(doc.sections):
    print(f"Sec {i}: orientation={sec.orientation}, start_type={sec.start_type}, cols={sec._sectPr.xpath('w:cols')}")
    # check cols
    cols = sec._sectPr.xpath('w:cols')
    if cols:
        print(f"   cols attr: {cols[0].attrib}")
print("Sample tables:", len(doc.tables))
for i, t in enumerate(doc.tables):
    print(f"Table {i}: rows={len(t.rows)}, cols={len(t.columns)}, width={[c.width for c in t.columns]}")

