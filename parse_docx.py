# https://github.com/python-openxml/python-docx/issues/74

import docx

doc_style = docx.Document('style.docx')
paragraphs = doc_style.paragraphs
tables = doc_style.tables

for i, par in enumerate(paragraphs):
    if par.text:
        print(i)
        print(par.text)
        print(par._p.xml)
        print(docx.oxml.text.paragraph.CT_P.__dict__)

# for tbl in tables:
#     rows = tbl.rows
#     for i in range(len(rows)):
#         cells = tbl.row_cells(i)
#         for cell in cells:
#             print(cell.text)

doc_demo = docx.Document('demo.docx')
paragraphs = doc_demo.paragraphs
tables = doc_demo.tables

for i, par in enumerate(paragraphs):
    if par.text:
        print(i)
        print(par.text)
        rPr = docx.oxml.shared.OxmlElement('w:rPr')
        c = docx.oxml.shared.OxmlElement('w:color')
        c.set(docx.oxml.shared.qn('w:val'), 'FF8822')
        rPr.append(c)
        par._p.r_lst[0].insert(0, rPr)
        print(par._p.xml)

# for tbl in tables:
#     rows = tbl.rows
#     for i in range(len(rows)):
#         cells = tbl.row_cells(i)
#         for cell in cells:
#             print(cell.text)

doc_demo.save('demo_style.docx')
