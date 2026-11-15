import docx

doc = docx.Document("d:/CODE/tgmt/baocao_final_viemphoi.docx")
print("Total paragraphs:", len(doc.paragraphs))
print("Total tables:", len(doc.tables))
print("Total sections:", len(doc.sections))

for i, p in enumerate(doc.paragraphs):
    if p.text.startswith("CHƯƠNG") or p.text.startswith("TÀI LIỆU") or p.text.startswith("PHỤ LỤC") or p.text.startswith("MỤC LỤC"):
        print(f"Para {i}: {p.text}")

