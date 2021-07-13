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

for tbl in tables:
    rows = tbl.rows
    for i in range(len(rows)):
        cells = tbl.row_cells(i)
        for cell in cells:
            print(cell.text)
