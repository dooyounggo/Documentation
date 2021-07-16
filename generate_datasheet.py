import docx
import numpy as np

PREFIX = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
TBL_WIDTHS = np.array([583, 583, 584, 583, 583, 585, 589, 585, 584, 584, 585, 584, 584, 585, 584, 585])

doc_style = docx.Document('style.docx')
doc_regbank = docx.Document('regbank.docx')

par_style = doc_style.paragraphs
tbl_style = doc_style.tables

# Clear IDs
for par in par_style:
    par._p.attrib.clear()
    for r in par._p.r_lst:
        r.attrib.clear()
for tbl in tbl_style:
    tbl._tbl.attrib.clear()
    for tr in tbl._tbl.tr_lst:
        tr.attrib.clear()
        for tc in tr.tc_lst:
            tc.attrib.clear()
            for p in tc.p_lst:
                p.attrib.clear()
                for r in p.r_lst:
                    r.attrib.clear()

par_name = par_style[2]._p
par_blank = par_style[3]._p

tbl_overview = tbl_style[0]._tbl
tbl_fields = tbl_style[1]._tbl
tbl_details = tbl_style[2]._tbl

cell_field = tbl_fields.tr_lst[1].tc_lst[0]
cell_field.p_lst[0].remove(cell_field.p_lst[0].r_lst[1])
cell_field.p_lst[0].r_lst[0].text = ''

body_regbank = doc_regbank._body._body
registers = list()
reg_found = False
tbl_cnt = 0
for i, obj in enumerate(body_regbank):
    if isinstance(obj, docx.oxml.text.paragraph.CT_P):
        if obj.r_lst:
            if obj.r_lst[0].text.startswith('Register: '):
                name = ''
                for r in obj.r_lst:
                    name += r.text
                name = name.replace('Register: ', '').lstrip()
                registers.append({'name': name, 'reset': None, 'overview': None, 'fields': None, 'details': None})
                reg_found = True
    if isinstance(obj, docx.oxml.table.CT_Tbl):
        if reg_found and tbl_cnt == 0:
            registers[-1]['overview'] = obj
            tbl_cnt += 1
        elif reg_found and tbl_cnt == 1:
            tbl_cnt += 1
            registers[-1]['fields'] = obj
        elif reg_found and tbl_cnt == 2:
            registers[-1]['details'] = obj
            reg_found = False
            tbl_cnt = 0
reg_idx = 0
for tbl in doc_regbank.tables:
    if tbl._tbl.tr_lst[0].tc_lst[0].p_lst[0].r_lst[0].text.startswith('Address'):
        for tr in tbl._tbl.tr_lst[1:]:
            reset = ''
            for r in tr.tc_lst[3].p_lst[0].r_lst:
                reset += r.text
            registers[reg_idx]['reset'] = reset
            reg_idx += 1

for reg in registers:
    new_row = tbl_overview.tr_lst[1].__copy__()
    new_row.tc_lst[0].p_lst[0].r_lst[0].text = reg['name'].upper()

    addr = ''
    for r in reg['overview'].tr_lst[2].tc_lst[1].p_lst[0].r_lst:
        addr += r.text
    new_row.tc_lst[1].p_lst[0].r_lst[0].text = addr.lstrip(': ')

    access = ''
    for tr in reg['details'].tr_lst[1:]:
        for p in tr.tc_lst[3].p_lst:
            access += p.r_lst[0].text
    if 'RW' in access:
        new_row.tc_lst[2].p_lst[0].r_lst[0].text = 'RW'
    elif 'RO' in access:
        new_row.tc_lst[2].p_lst[0].r_lst[0].text = 'RO'
    elif 'WO' in access:
        new_row.tc_lst[2].p_lst[0].r_lst[0].text = 'WO'

    new_row.tc_lst[3].p_lst[0].r_lst[0].text = reg['reset'].replace(' ', '_')

    desc = ''
    for r in reg['overview'].tr_lst[0].tc_lst[1].p_lst[0].r_lst:
        desc += r.text
    new_row.tc_lst[4].p_lst[0].r_lst[0].text = desc.lstrip(': ')

    tbl_overview.append(new_row)
tbl_overview.remove(tbl_overview.tr_lst[1])

doc_style.save('datasheet.docx')

