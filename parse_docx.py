import docx

doc_style = docx.Document('style.docx')

paragraphs = doc_style.paragraphs
tables = doc_style.tables

for tbl in tables:
    rows = tbl.rows
    for i in range(len(rows)):
        cells = tbl.row_cells(i)
        for cell in cells:
            print(cell.text)
