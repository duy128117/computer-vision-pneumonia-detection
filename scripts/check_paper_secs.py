import docx

doc = docx.Document("d:/CODE/tgmt/Paper_Pneumonia_Classification.docx")
print("Total sections:", len(doc.sections))
for i, s in enumerate(doc.sections):
    cols_elem = s._sectPr.xpath('w:cols')
    col_info = cols_elem[0].attrib if cols_elem else "None"
    print(f"Sec {i}: start={s.start_type}, cols={col_info}")

