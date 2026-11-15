import docx

doc = docx.Document("d:/CODE/tgmt/Báo cáo cuối kì Thị Giác máy tính.docx")
curr_chap = "Header"
chap_counts = {}
for p in doc.paragraphs:
    txt = p.text.strip()
    if txt.startswith("CHƯƠNG 1"):
        curr_chap = "Chương 1"
    elif txt.startswith("CHƯƠNG 2"):
        curr_chap = "Chương 2"
    elif txt.startswith("CHƯƠNG 3"):
        curr_chap = "Chương 3"
    elif txt.startswith("TÀI LIỆU"):
        curr_chap = "Tài liệu"
    
    chap_counts[curr_chap] = chap_counts.get(curr_chap, 0) + len(txt.split())

for k, v in chap_counts.items():
    print(f"{k}: {v} words")

