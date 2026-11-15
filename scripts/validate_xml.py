import zipfile
import xml.etree.ElementTree as ET

with zipfile.ZipFile("d:/CODE/tgmt/baocao_final_viemphoi.docx") as z:
    for name in z.namelist():
        if name.endswith(".xml"):
            try:
                data = z.read(name)
                ET.fromstring(data)
            except Exception as e:
                print(f"XML Error in {name}: {e}")
print("All XML files well-formed.")

