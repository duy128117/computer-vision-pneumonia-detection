import docx

doc = docx.Document("d:/CODE/tgmt/Báo cáo cuối kì Thị Giác máy tính.docx")
print("Total paras:", len(doc.paragraphs))
for p in doc.paragraphs:
    txt = p.text.strip()
    if any(txt.startswith(x) for x in ["CHƯƠNG", "1.", "2.", "3.", "TÀI LIỆU"]):
        if len(txt) < 80:
            print(txt)
