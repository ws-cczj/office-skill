"""Import a CSV and create an auditable cross-sheet group summary."""
import csv
from pathlib import Path

from core.office_batch import run_plan
from core.recipes import input_path
from skills.excel.scripts import excel_native


def column_letter(index):
    if type(index) is not int or not 1 <= index <= 16384:
        raise ValueError('Use a valid Excel column number')
    letters = ''
    while index:
        index, digit = divmod(index - 1, 26)
        letters = chr(ord('A') + digit) + letters
    return letters


def plan(data, folder, params):
    result = {'schema_version': 1, 'target': str(folder / 'excel' / 'workbook.xlsx'),
              'exports': [{'path': str(folder / 'renders' / 'excel' / 'summary.png'),
                           'sheet': 'Summary', 'chart': 'Totals'}]}
    if data.get('update'):
        result.update(update=True, operations=data['operations'])
        return result
    source = input_path(data['csv'], params)
    encoding, delimiter = data.get('encoding', 'utf-8-sig'), data.get('delimiter', ',')
    with Path(source).open(encoding=encoding, newline='') as stream:
        reader = (row for row in csv.reader(stream, delimiter=delimiter) if row)
        headers = next(reader, [])
        last_row = 1 + sum(1 for _ in reader)
    if not headers or last_row < 2:
        raise ValueError('The CSV must have a header and at least one data row')
    group_col, value_col = data['group_column'], data['value_column']
    group, value = column_letter(group_col), column_letter(value_col)
    if group_col == value_col or max(group_col, value_col) > len(headers):
        raise ValueError('Group and value columns must be distinct CSV columns')
    labels = data['groups']
    if (not isinstance(labels, list) or not labels or
            any(not isinstance(label, str) or not label for label in labels) or
            len(set(label.casefold() for label in labels)) != len(labels)):
        raise ValueError('Provide nonempty unique group names')
    operations = [
        {'op': 'sheet', 'name': 'Data', 'rename_from': 1},
        {'op': 'import_csv', 'sheet': 'Data', 'source': source, 'numeric_columns': [value_col],
         'encoding': encoding, 'delimiter': delimiter},
        {'op': 'clean', 'sheet': 'Data', 'range': f'A2:{column_letter(len(headers))}{last_row}',
         'trim': True, 'remove_blank_rows': True, 'unique_by': data.get('unique_by', [])},
        {'op': 'sheet', 'name': 'Summary'},
        {'op': 'write', 'sheet': 'Summary', 'rows': [[headers[group_col-1], headers[value_col-1]],
                                                     *[[label, None] for label in labels]]},
    ]
    groups = 'Data!$%s$2:$%s$%d' % (group, group, last_row)
    values = 'Data!$%s$2:$%s$%d' % (value, value, last_row)
    formulas = [[f'=SUMPRODUCT(--({groups}=A{row}),{values})'] for row in range(2, len(labels)+2)]
    operations.extend([
        {'op': 'formulas', 'sheet': 'Summary', 'start': 'B2', 'rows': formulas},
        {'op': 'format', 'sheet': 'Summary', 'range': 'A1:B1', 'bold': True, 'autofit': True},
        {'op': 'chart', 'sheet': 'Summary', 'name': 'Totals', 'range': f'A1:B{len(labels)+1}',
         'anchor': f'A{len(labels)+3}', 'kind': 'column', 'width_pt': 720,
         'title': data.get('title', f'Total {headers[value_col-1]} by {headers[group_col-1]}')},
    ])
    result['operations'] = operations
    result['snapshot'] = {'ranges': [{'sheet': 'Summary', 'range': f'A1:B{len(labels)+1}'}]}
    return result


def run(plan, data, folder, params):
    return run_plan(excel_native, plan, params)
