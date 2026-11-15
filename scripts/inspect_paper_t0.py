import docx

doc = docx.Document("d:/CODE/tgmt/Paper_Pneumonia_Classification.docx")
t0 = doc.tables[0]
print("Table 0 rows:", len(t0.rows), "cols:", len(t0.columns))
for r_idx, r in enumerate(t0.rows):
    cell_widths = [c.width.pt if c.width else 0 for c in r.cells]
    cell_texts = [c.text[:20].replace('\n', ' ') for c in r.cells]
    print(f"Row {r_idx}: widths={cell_widths}, texts={cell_texts}")

