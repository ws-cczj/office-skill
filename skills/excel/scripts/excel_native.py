"""Native Excel workbook editing through the desktop object model."""
import csv
import math
import re

from core.office_batch import absolute, existing


PROG_ID = 'Excel.Application'
EXTENSION = '.xlsx'
OPERATIONS = {'sheet', 'write', 'import_csv', 'clean', 'formulas', 'format', 'table', 'chart'}
CHART_TYPES = {'column': 51, 'bar': 57, 'line': 4, 'scatter': -4169, 'pie': 5}
CELL_ADDRESS = r'\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}'


def template_path(value):
    return existing(value, {'.xltx'})


def capabilities(app):
    return {'office_version': str(app.Version)}


def quit(app):
    app.Quit()


def close(book):
    book.Close(False)


def create(app, template):
    return app.Workbooks.Add(str(template) if template else -4167)


def open_document(app, path, readonly):
    return app.Workbooks.Open(Filename=str(path), UpdateLinks=0, ReadOnly=readonly,
                              IgnoreReadOnlyRecommended=True, AddToMru=False)


def save(book, path, existing_copy):
    if existing_copy:
        book.Save()
    else:
        book.SaveAs(Filename=str(path), FileFormat=51, AddToMru=False)


def export_path(spec):
    return absolute(spec['path'], {'.png', '.pdf'})


def export(book, spec, path):
    if path.suffix.lower() == '.png':
        ws = book.Worksheets.Item(spec['sheet'])
        ws.Activate()
        chart = ws.ChartObjects(spec['chart'])
        chart.Activate()
        if not chart.Chart.Export(str(path), 'PNG'):
            raise RuntimeError(f"Excel failed to export chart {spec['chart']}")
    else:
        source = book.Worksheets.Item(spec['sheet']) if 'sheet' in spec else book
        source.ExportAsFixedFormat(0, str(path))


def validate(operations):
    if not isinstance(operations, list):
        raise ValueError('operations must be an array')
    for op in operations:
        if op.get('op') not in OPERATIONS:
            raise ValueError(f"Unknown Excel operation: {op.get('op')}")
        if op['op'] == 'import_csv':
            existing(op['source'], {'.csv'})
        if op['op'] in {'write', 'import_csv', 'formulas'}:
            if not re.fullmatch(CELL_ADDRESS, op.get('start', 'A1'), re.I):
                raise ValueError('The destination start address must be a single A1 cell')
        if 'range' in op:
            validate_address(op['range'])


def validate_address(address):
    if not re.fullmatch(f'{CELL_ADDRESS}(?::{CELL_ADDRESS})?', address, re.I):
        raise ValueError(f'Use a contiguous A1 range: {address}')


def rectangular(rows):
    if not isinstance(rows, list) or not rows or not isinstance(rows[0], list) or not rows[0]:
        raise ValueError('Provide a nonempty matrix')
    cols = len(rows[0])
    if any(not isinstance(row, list) or len(row) != cols for row in rows):
        raise ValueError('All matrix rows must have the same width')
    return len(rows), cols


def literal(value):
    if isinstance(value, str):
        # Excel will otherwise parse strings such as =1+2, 0012 and dates as typed inputs.
        return "'" + value
    if value is None or isinstance(value, (bool, int)) or isinstance(value, float) and math.isfinite(value):
        return value
    raise ValueError('Cells accept strings, numbers, booleans or null')


def write(ws, start, rows, formula_mode=None):
    height, width = rectangular(rows)
    first = ws.Range(start)
    last = ws.Cells.Item(first.Row + height - 1, first.Column + width - 1)
    rng = ws.Range(first, last)
    if formula_mode:
        if formula_mode not in {'a1', 'r1c1'}:
            raise ValueError('Formula notation must be a1 or r1c1')
        if any(not isinstance(item, str) or not item.startswith('=') for row in rows for item in row):
            raise ValueError('Each formula must begin with =')
        property_name = 'FormulaR1C1' if formula_mode == 'r1c1' else 'Formula'
        setattr(rng, property_name, tuple(tuple(row) for row in rows))
    else:
        rng.Value2 = tuple(tuple(literal(value) for value in row) for row in rows)


def import_csv(ws, op):
    with existing(op['source'], {'.csv'}).open(encoding=op.get('encoding', 'utf-8-sig'), newline='') as stream:
        rows = list(csv.reader(stream, delimiter=op.get('delimiter', ',')))
    _, width = rectangular(rows)
    number_columns = op.get('numeric_columns', [])
    if any(type(col) is not int or not 1 <= col <= width for col in number_columns):
        raise ValueError('numeric_columns must use column numbers within the CSV')
    for row in rows[int(op.get('header', True)):]:
        for col in number_columns:
            value = row[col - 1].strip()
            row[col - 1] = float(value) if value else None
    write(ws, op.get('start', 'A1'), rows)


def matrix(rng, prop):
    value = getattr(rng, prop)
    return [list(row) for row in value] if isinstance(value, tuple) else [[value]]


def clean(ws, op):
    rng = ws.Range(op['range'])
    if rng.HasFormula is not False and rng.HasFormula != 0:
        raise ValueError('Clean requires a range containing literal values only')
    rows = matrix(rng, 'Value2')
    width = len(rows[0])
    keys = op.get('unique_by', [])
    if any(type(col) is not int or not 1 <= col <= width for col in keys):
        raise ValueError('unique_by must use column numbers within the selected range')
    cleaned, seen = [], set()
    for row in rows:
        if op.get('trim'):
            row = [value.strip() if isinstance(value, str) else value for value in row]
        if op.get('remove_blank_rows') and all(value in (None, '') for value in row):
            continue
        if keys:
            key = tuple((type(row[col - 1]), row[col - 1]) for col in keys)
            if key in seen:
                continue
            seen.add(key)
        cleaned.append(row)
    rng.ClearContents()
    if cleaned:
        write(ws, rng.Cells.Item(1, 1).Address, cleaned)


def excel_color(color):
    if not re.fullmatch(r'#[0-9A-Fa-f]{6}', color):
        raise ValueError(f'Use #RRGGBB: {color}')
    red, green, blue = (int(color[i:i + 2], 16) for i in (1, 3, 5))
    return red | (green << 8) | (blue << 16)


def format_range(ws, op):
    rng = ws.Range(op['range'])
    for key, prop in {'number_format': 'NumberFormat', 'wrap': 'WrapText',
                      'row_height': 'RowHeight', 'column_width': 'ColumnWidth'}.items():
        if key in op:
            setattr(rng, prop, op[key])
    for key, prop in {'font': 'Name', 'size': 'Size', 'bold': 'Bold', 'italic': 'Italic'}.items():
        if key in op:
            setattr(rng.Font, prop, op[key])
    if 'fill' in op:
        rng.Interior.Color = excel_color(op['fill'])
    if 'font_color' in op:
        rng.Font.Color = excel_color(op['font_color'])
    if op.get('autofit'):
        rng.Columns.AutoFit()


def sheet(book, op):
    name = op['name']
    current = next((ws for ws in book.Worksheets if ws.Name.casefold() == name.casefold()), None)
    if 'rename_from' in op:
        source = book.Worksheets.Item(op['rename_from'])
        if current is not None and source.Name != current.Name:
            raise ValueError(f'The destination worksheet name already exists: {name}')
        source.Name = name
    elif current is None:
        from pythoncom import Missing

        book.Worksheets.Add(Missing, book.Worksheets.Item(book.Worksheets.Count)).Name = name


def table(ws, op):
    rng = ws.Range(op['range'])
    item = ws.ListObjects.Add(SourceType=1, Source=rng, XlListObjectHasHeaders=1 if op.get('header', True) else 2)
    item.Name = op['name']
    if 'style' in op:
        item.TableStyle = op['style']


def chart(ws, op):
    name = op['name']
    if any(item.Name.casefold() == name.casefold() for item in ws.ChartObjects()):
        raise ValueError(f'A chart with that name already exists: {name}')
    anchor = op.get('anchor', 'A12')
    if not re.fullmatch(CELL_ADDRESS, anchor, re.I):
        raise ValueError('The chart anchor must be a single A1 cell')
    position = ws.Range(anchor)
    item = ws.ChartObjects().Add(position.Left, position.Top, op.get('width_pt', 480), op.get('height_pt', 300))
    item.Name = name
    graph = item.Chart
    graph.ChartType = CHART_TYPES[op.get('kind', 'column')]
    graph.SetSourceData(ws.Range(op['range']), 2)
    if 'title' in op:
        graph.HasTitle = True
        graph.ChartTitle.Text = op['title']
    for key, axis_type in (('x_title', 1), ('y_title', 2)):
        if key in op:
            axis = graph.Axes(axis_type)
            axis.HasTitle = True
            axis.AxisTitle.Text = op[key]


def apply(app, book, operations):
    app.EnableEvents = False
    app.ScreenUpdating = False
    try:
        for op in operations:
            kind = op['op']
            if kind == 'sheet':
                sheet(book, op)
                continue
            ws = book.Worksheets.Item(op['sheet'])
            if kind == 'write':
                write(ws, op.get('start', 'A1'), op['rows'])
            elif kind == 'formulas':
                write(ws, op.get('start', 'A1'), op['rows'], op.get('notation', 'a1'))
            elif kind == 'import_csv':
                import_csv(ws, op)
            elif kind == 'clean':
                clean(ws, op)
            elif kind == 'format':
                format_range(ws, op)
            elif kind == 'table':
                table(ws, op)
            elif kind == 'chart':
                chart(ws, op)
        app.Calculate()
    finally:
        app.ScreenUpdating = True


def snapshot(book, options):
    state = {
        'worksheets': [{'name': ws.Name, 'used_range': ws.UsedRange.Address,
                        'tables': [item.Name for item in ws.ListObjects],
                        'charts': [item.Name for item in ws.ChartObjects()]}
                       for ws in book.Worksheets],
        'quality_verdict': 'not_assessed',
    }
    if 'ranges' in options:
        details = []
        for item in options['ranges']:
            validate_address(item['range'])
            rng = book.Worksheets.Item(item['sheet']).Range(item['range'])
            details.append({**item, 'values': matrix(rng, 'Value2'), 'formulas': matrix(rng, 'Formula')})
        state['ranges'] = details
    return state
