import docx

doc = docx.Document("d:/CODE/tgmt/baocao_final_viemphoi.docx")
print("Total tables:", len(doc.tables))
for i, t in enumerate(doc.tables):
    widths = [c.width.inches if c.width else 0 for c in t.columns]
    total_w = sum(widths)
    print(f"Table {i}: cols={len(t.columns)}, total_w={total_w:.2f} in, col_widths={[round(w, 2) for w in widths]}")

