import docx

doc = docx.Document("d:/CODE/tgmt/Báo cáo cuối kì Thị Giác máy tính.docx")
fig_count = 0
for p in doc.paragraphs:
    if "Hình " in p.text:
        fig_count += 1
        print(p.text[:70])
print("Total figures mentioned in captions:", fig_count)

