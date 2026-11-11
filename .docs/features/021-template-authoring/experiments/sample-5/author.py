"""Author Sample 5 using only its source package and the XLSX contract."""
import json
import re
import zipfile
from pathlib import Path
from xml.dom import minidom

BASE = Path(__file__).resolve().parent
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
SOURCE = BASE / 'input/skill-sample-5.xlsx'

def write_json(path, value):
    path.write_text(json.dumps(value, indent=2, ensure_ascii=False) + '\n')

def text(cell):
    return ''.join(t.firstChild.data for t in cell.getElementsByTagNameNS(NS, 't') if t.firstChild)

def set_text(doc, cell, value):
    for child in list(cell.childNodes):
        cell.removeChild(child)
    cell.setAttribute('t', 'inlineStr')
    inline = doc.createElementNS(NS, 'is')
    node = doc.createElementNS(NS, 't')
    node.setAttribute('xml:space', 'preserve')
    node.appendChild(doc.createTextNode(value))
    inline.appendChild(node)
    cell.appendChild(inline)

def records(suppliers, groups):
    result = []
    for key, name, items in groups:
        for _ in suppliers or [None]:
            result.append({'$level': 1, '$path': key, 'Name': name})
        for item_key, item_name, quantity, offers in items:
            for index, _ in enumerate(suppliers or [None]):
                price, lead = offers[index] if suppliers else (None, None)
                result.append({'$level': 2, '$path': item_key, 'Name': item_name,
                               'Quantity': quantity, 'Price': price, 'LeadDays': lead})
    return result

with zipfile.ZipFile(SOURCE) as source:
    doc = minidom.parseString(source.read('xl/worksheets/sheet1.xml'))
    cells = {c.getAttribute('r'): c for c in doc.getElementsByTagNameNS(NS, 'c')}
    suppliers = [{'Company': text(cells[c])} for c in ['C3', 'F3']]
    source_items = []
    for row, key in [(6, 'laptop'), (7, 'dock')]:
        source_items.append((key, text(cells[f'C{row}']), text(cells[f'B{row}']),
                             [(text(cells[f'{c}{row}']), text(cells[f'{d}{row}']))
                              for c, d in [('D', 'E'), ('G', 'H')]]))
    original = {'Title': text(cells['A1']).removeprefix('Tender: '), 'Suppliers': suppliers,
                'Items': records(suppliers, [('it', text(cells['A5']), source_items)])}
    write_json(BASE / 'output/skill-sample-5.json', original)
    sheet_data = doc.getElementsByTagNameNS(NS, 'sheetData')[0]
    column_map = dict(zip('ABCDEI', 'BCDEFG'))
    row_map = {1: 2, 3: 4, 4: 5, 5: 6, 6: 7, 8: 8, 9: 9, 10: 10}
    for row in list(sheet_data.childNodes):
        old_row = int(row.getAttribute('r'))
        if old_row not in row_map:
            sheet_data.removeChild(row)
            continue
        row.setAttribute('r', str(row_map[old_row]))
        for cell in list(row.childNodes):
            old_col = re.match('[A-Z]+', cell.getAttribute('r'))[0]
            if old_col not in column_map:
                row.removeChild(cell)
                continue
            cell.setAttribute('r', column_map[old_col] + str(row_map[old_row]))
    replacements = {'B2': 'Tender: {Title}', 'D4': '{Company}', 'B6': '{Name}',
                    'B7': '{Name} / {../Name}', 'C7': '{Quantity}', 'D7': '{Name}',
                    'E7': '{Price}', 'F7': '{LeadDays}', 'B8': 'Lot total: {Name}'}
    for cell in doc.getElementsByTagNameNS(NS, 'c'):
        address = cell.getAttribute('r')
        if address in replacements:
            set_text(doc, cell, replacements[address])
        if address in ['E8', 'F8']:
            formula = cell.getElementsByTagNameNS(NS, 'f')[0]
            formula.firstChild.data = f'SUM({address[0]}7:{address[0]}7)'
    controls = {'D1': '$h(1)', 'E1': '$h(1)', 'F1': '$h(1)', 'A4': '$header(1).Suppliers',
                'A6': '$level(1).Items', 'A7': '$level(2).Items', 'A8': '$level(1).Items'}
    for address, value in controls.items():
        row_number = re.search('[0-9]+', address)[0]
        rows = {r.getAttribute('r'): r for r in sheet_data.childNodes}
        row = rows.get(row_number)
        if row is None:
            row = doc.createElementNS(NS, 'row')
            row.setAttribute('r', row_number)
            sheet_data.appendChild(row)
        cell = doc.createElementNS(NS, 'c')
        cell.setAttribute('r', address)
        set_text(doc, cell, value)
        row.appendChild(cell)
    for row in sorted(list(sheet_data.childNodes), key=lambda r: int(r.getAttribute('r'))):
        sheet_data.appendChild(row)
        for cell in sorted(list(row.childNodes), key=lambda c: (len(re.match('[A-Z]+', c.getAttribute('r'))[0]), re.match('[A-Z]+', c.getAttribute('r'))[0])):
            row.appendChild(cell)
    cols = doc.getElementsByTagNameNS(NS, 'cols')[0]
    for col in list(cols.childNodes):
        old = chr(64 + int(col.getAttribute('min')))
        if old not in column_map:
            cols.removeChild(col)
        else:
            index = str(ord(column_map[old]) - 64)
            col.setAttribute('min', index)
            col.setAttribute('max', index)
    merges = doc.getElementsByTagNameNS(NS, 'mergeCells')[0]
    merges.removeChild(merges.lastChild)
    merges.firstChild.setAttribute('ref', 'D4:F4')
    merges.setAttribute('count', '1')
    doc.getElementsByTagNameNS(NS, 'dimension')[0].setAttribute('ref', 'A1:G10')
    workbook = minidom.parseString(source.read('xl/workbook.xml'))
    workbook.getElementsByTagNameNS(NS, 'definedName')[0].firstChild.data = "'Tender'!$B$2:$G$10"
    with zipfile.ZipFile(BASE / 'output/skill-sample-5.template.xlsx', 'w', zipfile.ZIP_DEFLATED) as target:
        for entry in source.infolist():
            content = (doc.toxml(encoding='utf-8') if entry.filename == 'xl/worksheets/sheet1.xml'
                       else workbook.toxml(encoding='utf-8') if entry.filename == 'xl/workbook.xml'
                       else source.read(entry.filename))
            target.writestr(entry, content)

modified_suppliers = [{'Company': name} for name in ['Gamma', 'Delta', 'Epsilon']]
modified = {'Title': 'Workplace refresh', 'Suppliers': modified_suppliers,
            'Items': records(modified_suppliers, [
                ('devices', 'Devices', [
                    ('monitor', 'Monitor', '4.00', [(260, 7), (245.5, 9), (275, 3)]),
                    ('keyboard', 'Keyboard', '8.00', [(45, 2), (42, 4), (50, 1)]),
                    ('mouse', 'Mouse', '8.00', [(25, 3), (22.5, 5), (29, 2)])]),
                ('furniture', 'Furniture', [('desk', 'Desk', '2.00', [(400, 14), (380, 18), (425, 10)])])])}
write_json(BASE / 'evidence/modified-data.json', modified)
write_json(BASE / 'evidence/empty-items-data.json', {**original, 'Items': []})
write_json(BASE / 'evidence/empty-lot-data.json', {**original, 'Items': records(suppliers, [('empty', 'Empty lot', [])])})
write_json(BASE / 'evidence/empty-suppliers-data.json', {**original, 'Suppliers': [], 'Items': records([], [('it', 'IT equipment', source_items)])})
write_json(BASE / 'evidence/single-supplier-data.json', {**original, 'Suppliers': suppliers[:1], 'Items': records(suppliers[:1], [('it', 'IT equipment', [(key, name, qty, offers[:1]) for key, name, qty, offers in source_items])])})
write_json(BASE / 'evidence/initial-attempt.json', {'model': 'fixed price/lead-day criteria; repeat complete three-column supplier blocks, including hidden item names; hierarchical lots/items', 'previous_solutions_inspected': False, 'source_numeric_text_preserved': True, 'template': 'output/skill-sample-5.template.xlsx'})
print('Authored initial template, extracted JSON, and five modified/edge datasets.')
