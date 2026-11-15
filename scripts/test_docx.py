import docx

try:
    doc = docx.Document("d:/CODE/tgmt/baocao_final_viemphoi.docx")
    print(f"Paragraphs: {len(doc.paragraphs)}, Tables: {len(doc.tables)}, Sections: {len(doc.sections)}")
except Exception as e:
    print(f"Error opening docx: {e}")

