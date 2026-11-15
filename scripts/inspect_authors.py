# -*- coding: utf-8 -*-
import docx

doc = docx.Document('d:/CODE/tgmt/baocao_giuaki.docx')
for t in doc.tables[:3]:
    for r in t.rows:
        print([c.text.strip().replace('\n', ' ') for c in r.cells])
