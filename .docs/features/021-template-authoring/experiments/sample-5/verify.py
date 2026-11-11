"""Independently check Sample 5 package, cell, geometry, and formula semantics."""
import hashlib
import json
import re
import time
import zipfile
from pathlib import Path
import xml.etree.ElementTree as E

BASE = Path(__file__).resolve().parent
EV = BASE / 'evidence'
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
TAG = '{' + NS['s'] + '}'
NOTE = 'Compare unit price and lead days in separate columns.'

def package(path):
    with zipfile.ZipFile(path) as z:
        assert z.testzip() is None
        parts = {name: z.read(name) for name in z.namelist()}
    for name, content in parts.items():
        if name.endswith('.xml') or name.endswith('.rels'):
            E.fromstring(content)
    return parts

def sheet(parts):
    return E.fromstring(parts['xl/worksheets/sheet1.xml'])

def value(cell):
    if cell.find('s:f', NS) is not None:
        return ('formula', cell.find('s:f', NS).text)
    if cell.get('t') == 'inlineStr':
        return ('text', ''.join(cell.find('s:is', NS).itertext()))
    v = cell.find('s:v', NS)
    return ('empty', None) if v is None else ('number', v.text)

def cells(root):
    return {c.get('r'): c for c in root.findall('s:sheetData/s:row/s:c', NS)}

def normalized(element):
    return element.tag, sorted(element.attrib.items()), element.text or '', [normalized(c) for c in element]

def column_properties(root):
    result = {}
    for col in root.findall('s:cols/s:col', NS):
        for index in range(int(col.get('min')), int(col.get('max')) + 1):
            result[index] = {k: v for k, v in col.attrib.items() if k not in ['min', 'max']}
    return result

def col_name(index):
    result = ''
    while index:
        index, digit = divmod(index - 1, 26)
        result = chr(65 + digit) + result
    return result

def assert_value(actual, address, expected):
    assert address in actual, address
    if isinstance(expected, (int, float)):
        assert value(actual[address])[0] == 'number', (address, value(actual[address]))
        assert float(value(actual[address])[1]) == expected
    else:
        assert value(actual[address]) == ('text', expected), (address, value(actual[address]), expected)

source = package(BASE / 'input/skill-sample-5.xlsx')
template = package(BASE / 'output/skill-sample-5.template.xlsx')
original = package(EV / 'original.result.xlsx')
source_sheet = sheet(source)
source_cells = cells(source_sheet)
original_sheet = sheet(original)
original_cells = cells(original_sheet)
assert set(source_cells) == set(original_cells), (set(source_cells) ^ set(original_cells))
for address, cell in source_cells.items():
    assert value(cell) == value(original_cells[address]), address
    assert cell.get('s') == original_cells[address].get('s'), address
assert column_properties(source_sheet) == column_properties(original_sheet)
for local in ['mergeCells', 'pageMargins', 'pageSetup', 'sheetFormatPr', 'sheetViews']:
    assert normalized(source_sheet.find('s:' + local, NS)) == normalized(original_sheet.find('s:' + local, NS)), local
source_rows = {r.get('r'): r.attrib for r in source_sheet.findall('s:sheetData/s:row', NS)}
original_rows = {r.get('r'): r.attrib for r in original_sheet.findall('s:sheetData/s:row', NS)}
assert source_rows == original_rows, (source_rows, original_rows)
source_area = E.fromstring(source['xl/workbook.xml']).find('s:definedNames/s:definedName', NS).text
assert E.fromstring(original['xl/workbook.xml']).find('s:definedNames/s:definedName', NS).text == source_area
first = {'unix_seconds': time.time(), 'source_cells_compared': len(source_cells), 'result': 'all original cells, values/types, formulas, style indices, rows, columns, merges, views, print settings and print area match source'}
(EV / 'first-verified-output.json').write_text(json.dumps(first, indent=2) + '\n')
print('Original reconstruction verified:', len(source_cells), 'cells.')

api = json.loads((EV / 'api-cases.json').read_text())
checks = {'original': first}
preservation = {}
source_cols = column_properties(source_sheet)
source_row_elements = {int(r.get('r')): r for r in source_sheet.findall('s:sheetData/s:row', NS)}
for name in ['original', 'modified', 'single-supplier', 'empty-items', 'empty-suppliers']:
    data_file = BASE / 'output/skill-sample-5.json' if name == 'original' else EV / f'{name}-data.json'
    data = json.loads(data_file.read_text())
    path = BASE / 'output/skill-sample-5.result.xlsx' if name == 'modified' else EV / f'{name}.result.xlsx'
    parts = package(path)
    root = sheet(parts)
    actual = cells(root)
    assert set(parts) == set(source)
    unchanged = [part for part in source if source[part] == parts[part]]
    for part in set(source) - {'xl/worksheets/sheet1.xml', 'xl/workbook.xml'}:
        assert parts[part] == source[part], (name, part)
    preservation[name] = {'unchanged_parts': unchanged, 'changed_parts': sorted(set(source) - set(unchanged))}
    assert api[name]['http_status'] == 200 and api[name]['warnings'] == []
    assert_value(actual, 'A1', 'Tender: ' + data['Title'])
    assert_value(actual, 'A4', 'Item')
    assert_value(actual, 'B4', 'Quantity')
    count = len(data['Suppliers'])
    expected_merges = []
    expected_cols = {1: source_cols[1], 2: source_cols[2], 3 + count * 3: source_cols[9]}
    for index, supplier in enumerate(data['Suppliers']):
        first_col = 3 + index * 3
        hidden, price, lead = [col_name(first_col + offset) for offset in range(3)]
        assert_value(actual, hidden + '3', supplier['Company'])
        assert_value(actual, price + '4', 'Unit price')
        assert_value(actual, lead + '4', 'Lead days')
        expected_merges.append(f'{hidden}3:{lead}3')
        for offset in range(3):
            expected_cols[first_col + offset] = source_cols[3 + offset]
    assert column_properties(root) == expected_cols, name
    assert [x.get('ref') for x in root.findall('s:mergeCells/s:mergeCell', NS)] == expected_merges
    logical = []
    for record in data['Items']:
        if not logical or logical[-1][0]['$path'] != record['$path']:
            logical.append([])
        logical[-1].append(record)
    row = 5
    group_name = None
    start = None
    formulas = {}
    expected_styles = {}
    row_mapping = {1: 1, 3: 3, 4: 4}
    def close_group():
        global row
        assert_value(actual, f'A{row}', 'Lot total: ' + group_name)
        row_mapping[row] = 8
        for index in range(count):
            for offset in [1, 2]:
                col = col_name(3 + index * 3 + offset)
                formulas[f'{col}{row}'] = f'SUM({col}{start}:{col}{row - 1})'
        row += 1
    for block in logical:
        record = block[0]
        assert len(block) == max(count, 1)
        if record['$level'] == 1:
            if group_name is not None:
                close_group()
            group_name = record['Name']
            assert_value(actual, f'A{row}', group_name)
            row_mapping[row] = 5
            row += 1
            start = row
        else:
            assert_value(actual, f'A{row}', record['Name'] + ' / ' + group_name)
            assert_value(actual, f'B{row}', record['Quantity'])
            row_mapping[row] = 6
            for index in range(count):
                for offset, field in enumerate(['Name', 'Price', 'LeadDays']):
                    assert_value(actual, f'{col_name(3 + index * 3 + offset)}{row}', block[index][field])
            row += 1
    if group_name is not None:
        close_group()
    row_mapping[row] = 9
    row_mapping[row + 1] = 10
    assert_value(actual, f'A{row + 1}', NOTE)
    expected_area = f"'Tender'!$A$1:${col_name(3 + count * 3)}${row + 1}"
    assert E.fromstring(parts['xl/workbook.xml']).find('s:definedNames/s:definedName', NS).text == expected_area
    assert root.find('s:dimension', NS).get('ref') == f'A1:{col_name(3 + count * 3)}{row + 1}'
    for output_row, source_row in row_mapping.items():
        source_row_cells = {re.match('[A-Z]+', c.get('r'))[0]: c for c in source_row_elements[source_row]}
        col_mapping = {1: 'A', 2: 'B', 3 + count * 3: 'I'}
        for index in range(count):
            col_mapping.update({3 + index * 3 + offset: chr(67 + offset) for offset in range(3)})
        for output_col, source_col in col_mapping.items():
            if source_col in source_row_cells:
                address = col_name(output_col) + str(output_row)
                expected_styles[address] = source_row_cells[source_col].get('s')
        out_row = root.find(f"s:sheetData/s:row[@r='{output_row}']", NS)
        assert out_row is not None
        assert {k: v for k, v in out_row.attrib.items() if k != 'r'} == {k: v for k, v in source_row_elements[source_row].attrib.items() if k != 'r'}
    assert set(actual) == set(expected_styles), (name, set(actual) ^ set(expected_styles))
    assert all(actual[a].get('s') == style for a, style in expected_styles.items())
    assert {a: value(c)[1] for a, c in actual.items() if value(c)[0] == 'formula'} == formulas
    for address, formula in formulas.items():
        assert actual[address].find('s:v', NS) is None
        operands = re.fullmatch(r'SUM\(([A-Z]+)(\d+):\1(\d+)\)', formula)
        col, begin, end = operands.groups()
        assert all(f'{col}{r}' in actual for r in range(int(begin), int(end) + 1))
    for cell in actual.values():
        assert not any(marker in ''.join(cell.itertext()) for marker in ['$level(', '$h(', '$header(', '{Name}', '{Price}', '{Title}'])
    for local in ['pageMargins', 'pageSetup', 'sheetFormatPr', 'sheetViews']:
        assert normalized(source_sheet.find('s:' + local, NS)) == normalized(root.find('s:' + local, NS))
    calc = E.fromstring(parts['xl/workbook.xml']).find('s:calcPr', NS)
    assert calc.get('fullCalcOnLoad') == '1' and calc.get('forceFullCalc') == '1'
    totals = {}
    for address, formula in formulas.items():
        col, begin, end = re.fullmatch(r'SUM\(([A-Z]+)(\d+):\1(\d+)\)', formula).groups()
        totals[address] = sum(float(value(actual[f'{col}{r}'])[1]) for r in range(int(begin), int(end) + 1) if value(actual[f'{col}{r}'])[0] == 'number')
    checks[name] = {'cells_checked': len(actual), 'supplier_count': count, 'logical_records': len(logical), 'print_area': expected_area, 'formulas': formulas, 'independent_SUM_expectations_not_Excel_calculation': totals, 'styles_and_geometry': 'match mapped source', 'warnings': []}
assert api['empty-lot']['http_status'] == 422
assert api['empty-lot']['error']['code'] == 'template.xlsx_unsafe_formula'
checks['empty-lot'] = {'expected_rejection': 'template.xlsx_unsafe_formula', 'reason': 'Subtotal range has no retained item cells.'}
assert set(template) == set(source)
assert [name for name in source if source[name] != template[name]] == ['xl/workbook.xml', 'xl/worksheets/sheet1.xml']
preservation['template'] = {'unchanged_parts': [n for n in source if source[n] == template[n]], 'changed_parts': ['xl/workbook.xml', 'xl/worksheets/sheet1.xml']}
(EV / 'package-preservation.json').write_text(json.dumps(preservation, indent=2) + '\n')
(EV / 'verification.json').write_text(json.dumps(checks, indent=2) + '\n')
hashes = {str(p.relative_to(BASE)): hashlib.sha256(p.read_bytes()).hexdigest() for p in [BASE / 'input/skill-sample-5.xlsx', *sorted((BASE / 'output').glob('*')), *sorted(EV.glob('*.result.xlsx'))]}
(EV / 'artifact-hashes.json').write_text(json.dumps(hashes, indent=2) + '\n')
print('Verified five generated workbooks and one expected rejection.')
