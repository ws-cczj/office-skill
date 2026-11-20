"""Native Word manuscript editing through the desktop object model."""
from contextlib import contextmanager
import re

from core.office_batch import absolute, existing


PROG_ID = 'Word.Application'
EXTENSION = '.docx'
MATHTYPE_PROG_ID = 'equation.dsmt4'
MATHTYPE_TOGGLE = 'MTCommand_TeXToggle'
STYLES = {
    'normal': -1, 'title': -63, 'subtitle': -75, 'caption': -35,
    **{f'heading{n}': -1 - n for n in range(1, 10)},
    'abstract': 'Office Abstract', 'reference': 'Office Reference', 'equation': 'Office Equation',
}
SEQUENCES = {'reference': 'Bibliography', 'equation': 'Equation', 'figure': 'Figure', 'table': 'Table'}
OPERATIONS = {
    'configure', 'paragraph', 'format_paragraph', 'equation', 'replace_mathtype',
    'reference', 'figure', 'table', 'caption', 'toc', 'refresh_toc', 'header_footer', 'page_numbers',
}


def template_path(value):
    return existing(value, {'.dotx'})


def capabilities(app):
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, 'Equation.DSMT4\\CLSID'):
            registered = True
    except FileNotFoundError:
        registered = False
    commands_loaded = any(
        addin.Installed and re.fullmatch(r'MathType Commands.*\.dotm', addin.Name, re.I)
        for addin in app.AddIns
    )
    return {'office_version': str(app.Version), 'mathtype_class_registered': registered,
            'mathtype_commands_loaded': commands_loaded}


def quit(app):
    app.Quit(0)


def close(doc):
    doc.Close(0)


def create(app, template):
    return app.Documents.Add(Template=str(template), Visible=False) if template else app.Documents.Add(Visible=False)


def open_document(app, path, readonly):
    return app.Documents.Open(FileName=str(path), ReadOnly=readonly, AddToRecentFiles=False,
                              Visible=False, ConfirmConversions=False, NoEncodingDialog=True)


def save(doc, path, existing_copy):
    if existing_copy:
        doc.Save()
    else:
        doc.SaveAs2(FileName=str(path), FileFormat=12, AddToRecentFiles=False)


def export_path(spec):
    return absolute(spec['path'], {'.pdf'})


def export(doc, spec, path):
    first, last = spec.get('first_page'), spec.get('last_page')
    if first is None and last is None:
        doc.ExportAsFixedFormat(str(path), 17)
        return
    pages = doc.ComputeStatistics(2)
    if type(first) is not int or type(last) is not int or not 1 <= first <= last <= pages:
        raise ValueError(f'PDF export requires a valid first_page and last_page within 1–{pages}')
    doc.ExportAsFixedFormat(OutputFileName=str(path), ExportFormat=17, Range=3, From=first, To=last)


def bookmark(kind, key):
    if not isinstance(key, str) or not re.fullmatch(r'[A-Za-z][A-Za-z0-9_]{0,22}', key):
        raise ValueError('Use a key of 1–23 ASCII letters, digits or underscores, starting with a letter')
    if kind not in SEQUENCES and kind != 'math':
        raise ValueError(f'Unknown reference kind: {kind}')
    return f'office_{kind}_{key}'


def validate(operations):
    if not isinstance(operations, list):
        raise ValueError('operations must be an array')
    for op in operations:
        if op.get('op') not in OPERATIONS:
            raise ValueError(f"Unknown Word operation: {op.get('op')}")
        kind = op['op']
        if kind == 'equation' and op.get('engine', 'word') not in {'word', 'mathtype'}:
            raise ValueError('Equation engine must be word or mathtype')
        if kind == 'caption' and op.get('kind') not in {'figure', 'table'}:
            raise ValueError('Caption kind must be figure or table')
        if kind in {'equation', 'reference', 'caption'}:
            bookmark(op['kind'] if kind == 'caption' else kind, op['id'])
        if kind == 'figure':
            existing(op['image'], {'.png', '.jpg', '.jpeg', '.svg', '.emf', '.tif', '.tiff'})


def style(doc, name):
    key = STYLES.get(name, name)
    if name not in STYLES or not isinstance(key, str):
        return doc.Styles.Item(key)
    item = next((item for item in doc.Styles if item.NameLocal == key), None)
    if item is not None:
        return item
    item = doc.Styles.Add(key, 1)
    normal = doc.Styles.Item(STYLES['normal'])
    item.BaseStyle = normal.NameLocal
    for prop in ('Name', 'NameFarEast', 'Size', 'Bold', 'Italic'):
        setattr(item.Font, prop, getattr(normal.Font, prop))
    return item


def font_settings(font, values):
    for key, prop in {'font': 'Name', 'font_east_asia': 'NameFarEast', 'size': 'Size',
                      'bold': 'Bold', 'italic': 'Italic'}.items():
        if key in values:
            setattr(font, prop, values[key])


def configure(doc, op):
    layout = op.get('page', {})
    for section in doc.Sections:
        page = section.PageSetup
        if 'paper' in layout:
            page.PaperSize = {'a4': 7, 'letter': 2}[layout['paper']]
        if 'orientation' in layout:
            page.Orientation = {'portrait': 0, 'landscape': 1}[layout['orientation']]
        for side in ('top', 'bottom', 'left', 'right'):
            name = f'{side}_pt'
            if name in layout:
                setattr(page, f'{side.title()}Margin', layout[name])
        if 'columns' in layout:
            page.TextColumns.SetCount(layout['columns'])
        if 'column_gap_pt' in layout:
            page.TextColumns.Spacing = layout['column_gap_pt']
    settings = op.get('styles', {})
    if op.get('profile') == 'sci':
        defaults = {'normal': {'font': 'Times New Roman', 'size': 12, 'line_spacing': 'double'},
                    'title': {'font': 'Times New Roman', 'size': 16, 'bold': True},
                    'heading1': {'font': 'Times New Roman', 'size': 12, 'bold': True},
                    'heading2': {'font': 'Times New Roman', 'size': 12, 'italic': True},
                    'reference': {'font': 'Times New Roman', 'size': 12}}
        settings = {name: {**defaults.get(name, {}), **settings.get(name, {})}
                    for name in dict.fromkeys((*defaults, *settings))}
    elif 'profile' in op:
        raise ValueError('Unknown profile; choose sci or use explicit settings')
    for name, values in settings.items():
        item = style(doc, name)
        font_settings(item.Font, values)
        fmt = item.ParagraphFormat
        for key, prop in {'space_before_pt': 'SpaceBefore', 'space_after_pt': 'SpaceAfter',
                          'first_line_pt': 'FirstLineIndent', 'left_indent_pt': 'LeftIndent',
                          'keep_with_next': 'KeepWithNext'}.items():
            if key in values:
                setattr(fmt, prop, values[key])
        if 'align' in values:
            fmt.Alignment = {'left': 0, 'center': 1, 'right': 2, 'justify': 3}[values['align']]
        if 'line_spacing' in values:
            fmt.LineSpacingRule = {'single': 0, '1.5': 1, 'double': 2}[values['line_spacing']]


def insert(doc, position, text):
    rng = doc.Range(position, position)
    rng.InsertAfter(text)
    return rng


def append_position(doc):
    end = doc.Content.End - 1
    if doc.Paragraphs.Last.Range.Start < end:
        # Split off a new paragraph if the source has text in its final paragraph.
        insert(doc, end, '\r')
        end = doc.Content.End - 1
    return end


def paragraph(doc, name):
    start = append_position(doc)
    insert(doc, start, '\r')
    rng = doc.Range(start, start)
    rng.Style = style(doc, name)
    return rng


def set_script(font, script):
    if script not in {'super', 'sub', 'normal'}:
        raise ValueError('script must be super, sub or normal')
    if script == 'super':
        font.Superscript = True
    elif script == 'sub':
        font.Subscript = True
    else:
        font.Superscript = False
        font.Subscript = False


def text_run(doc, start, run):
    text = run['text']
    if not isinstance(text, str) or any(char in text for char in '\r\n'):
        raise ValueError('Pass each paragraph as its own operation')
    rng = insert(doc, start, text)
    set_script(rng.Font, run.get('script', 'normal'))
    font_settings(rng.Font, run)
    return rng.End


def field(doc, pos, code):
    return doc.Fields.Add(doc.Range(pos, pos), -1, code, False)


def numbered_field(doc, pos, kind, key):
    label = bookmark(kind, key)
    if doc.Bookmarks.Exists(label):
        raise ValueError(f'Duplicate {kind} key: {key}')
    number = field(doc, pos, f'SEQ {SEQUENCES[kind]} \\* ARABIC')
    doc.Bookmarks.Add(label, doc.Range(number.Code.Start - 1, number.Result.End + 1))
    return number


def is_mathtype(shape):
    # wdInlineShapeEmbeddedOLEObject = 1. Reading OLEFormat does not activate its server.
    return shape.Type == 1 and shape.OLEFormat.ProgID.casefold() == MATHTYPE_PROG_ID


@contextmanager
def mathtype_writer(app):
    scratch = None

    def write(target, text):
        nonlocal scratch
        if not isinstance(text, str) or not text.strip() or any(c in text for c in '\r\n$'):
            raise ValueError('Use one nonempty MathType TeX expression without delimiters')
        if scratch is None:
            supported = capabilities(app)
            if not (supported['mathtype_class_registered'] and supported['mathtype_commands_loaded']):
                raise RuntimeError('Install MathType 7 desktop and load its Word Commands add-in')
            from win32com.client.dynamic import CDispatch
            if isinstance(app, CDispatch):
                app._FlagAsMethod('Run')
            # Reuse one hidden, unsaved conversion document for the entire batch.
            scratch = app.Documents.Add(Visible=False)

        scratch.Content.Text = '$' + text + '$'
        scratch.Activate()
        scratch.Range(0, scratch.Content.End - 1).Select()
        # Wiris documents this macro as its Toggle TeX command.
        app.Run(MATHTYPE_TOGGLE)
        if scratch.InlineShapes.Count != 1 or not is_mathtype(scratch.InlineShapes.Item(1)):
            raise RuntimeError('MathType did not convert the TeX expression to an editable MathType object')
        start, doc = target.Start, target.Document
        target.FormattedText = scratch.InlineShapes.Item(1).Range.FormattedText
        # Word ranges count the full OLE field, even when Text displays one placeholder.
        shape = next((item for item in doc.InlineShapes if item.Range.Start == start), None)
        if shape is None or not is_mathtype(shape):
            raise RuntimeError('Word did not copy the editable MathType object')
        return shape.Range

    try:
        yield write
    finally:
        if scratch is not None:
            scratch.Close(0)

def write_equation(write_mathtype, target, engine, text):
    if engine == 'mathtype':
        return write_mathtype(target, text)
    if engine != 'word':
        raise ValueError('Equation engine must be word or mathtype')
    if not isinstance(text, str) or not text.strip() or any(c in text for c in '\r\n\t'):
        raise ValueError('Use one nonempty UnicodeMath expression')
    start, doc = target.Start, target.Document
    target.Text = text
    units = len(text.encode('utf-16-le')) // 2
    result = doc.OMaths.Add(doc.Range(start, start + units)).OMaths.Item(1)
    result.BuildUp()
    result.Type = 1  # Let the surrounding paragraph position the inline equation.
    return result.Range


def add_equation(write_mathtype, doc, op):
    rng = paragraph(doc, op.get('style', 'equation'))
    fmt = rng.ParagraphFormat
    columns = rng.Sections.Item(1).PageSetup.TextColumns
    if columns.Count > 1 and not columns.EvenlySpaced:
        raise ValueError('Equation layout requires equally sized text columns')
    width = columns.Item(1).Width
    if width <= 0:
        page = rng.Sections.Item(1).PageSetup
        width = page.PageWidth - page.LeftMargin - page.RightMargin - page.Gutter
    fmt.Alignment = 0
    fmt.LeftIndent = 0
    fmt.RightIndent = 0
    fmt.FirstLineIndent = 0
    fmt.KeepTogether = True
    fmt.TabStops.ClearAll()
    fmt.TabStops.Add(width / 2, 1)
    fmt.TabStops.Add(width, 2)
    start = rng.Start
    # Install both tabs and the sequence before creating the math zone.
    # Appending Word runs directly to an OMath end can put them inside the math zone.
    layout = insert(doc, start, '\t \t()')
    numbered_field(doc, layout.End - 1, 'equation', op['id'])
    formula = write_equation(write_mathtype, doc.Range(start + 1, start + 2), op.get('engine', 'word'), op['text'])
    doc.Bookmarks.Add(bookmark('math', op['id']), formula)

def replace_mathtype(write_mathtype, doc, op):
    label = bookmark('math', op['id']) if 'id' in op else None
    if label:
        rng = doc.Bookmarks.Item(label).Range
        shapes = rng.InlineShapes
        if shapes.Count != 1:
            raise ValueError('The bookmark must contain exactly one MathType object')
        shape = shapes.Item(1)
    else:
        shape = doc.InlineShapes.Item(op['inline_shape'])
    if not is_mathtype(shape):
        raise ValueError('The selected inline shape is not an editable MathType object')
    result = write_mathtype(shape.Range, op['text'])
    if label:
        doc.Bookmarks.Add(label, result)


def add_labelled_paragraph(doc, op):
    kind = op.get('kind', 'reference')
    rng = paragraph(doc, op.get('style', 'reference' if kind == 'reference' else 'caption'))
    if kind == 'reference':
        position = insert(doc, rng.Start, op.get('prefix', '[')).End
        insert(doc, position, op.get('suffix', '] ') + op['text'])
    else:
        label = op.get('label', {'figure': 'Fig.', 'table': 'Table'}[kind])
        position = insert(doc, rng.Start, label + ' ').End
        insert(doc, position, '. ' + op['text'])
    numbered_field(doc, position, kind, op['id'])


def add_paragraph(write_mathtype, doc, op, pending):
    rng = paragraph(doc, op.get('style', 'normal'))
    position = rng.Start
    math_runs = []
    for run in op.get('runs', [{'text': op.get('text', '')}]):
        if 'ref' in run:
            kind = run.get('kind', 'reference')
            label = bookmark(kind, run['ref'])
            pending.add(label)
            left, right = {'reference': ('[', ']'), 'equation': ('(', ')')}.get(kind, ('', ''))
            start = position
            position = insert(doc, start, run.get('prefix', left)).End
            script = run.get('script')
            switch = ' \\* CHARFORMAT' if script else ''
            reference = field(doc, position, f'REF {label} \\h{switch}')
            position = insert(doc, reference.Result.End + 1, run.get('suffix', right)).End
            if script:
                set_script(doc.Range(start, position).Font, script)
        elif 'math' in run:
            math_runs.append((position, run))
            position = insert(doc, position, ' ').End
        else:
            position = text_run(doc, position, run)
    for start, run in reversed(math_runs):
        write_equation(write_mathtype, doc.Range(start, start + 1), run.get('engine', 'word'), run['math'])


def format_paragraph(doc, op):
    rng = doc.Paragraphs.Item(op['paragraph']).Range
    if 'style' in op:
        rng.Style = style(doc, op['style'])
    for run in op.get('runs', []):
        start, end = run['start'], run['end']
        if type(start) is not int or type(end) is not int or not 0 <= start < end < rng.End - rng.Start:
            raise ValueError('Formatting spans must fit within the selected paragraph text')
        target = doc.Range(rng.Start + start, rng.Start + end)
        if 'expect' in run and target.Text != run['expect']:
            raise ValueError(f"Word paragraph {op['paragraph']} text does not match the requested span")
        if 'script' in run:
            set_script(target.Font, run['script'])
        font_settings(target.Font, run)


def add_figure(doc, op):
    rng = paragraph(doc, op.get('style', 'normal'))
    rng.ParagraphFormat.Alignment = 1
    image = doc.InlineShapes.AddPicture(FileName=str(existing(op['image'])), LinkToFile=False,
                                        SaveWithDocument=True, Range=rng)
    if 'width_pt' in op:
        image.LockAspectRatio = True
        image.Width = op['width_pt']


def add_table(doc, op):
    rows = op['rows']
    if not rows or not rows[0] or any(len(row) != len(rows[0]) for row in rows):
        raise ValueError('Table rows must form a nonempty rectangle')
    start = append_position(doc)
    table = doc.Tables.Add(doc.Range(start, start), len(rows), len(rows[0]))
    for i, row in enumerate(rows, 1):
        for j, value in enumerate(row, 1):
            table.Cell(i, j).Range.Text = str(value)
    if 'style' in op:
        table.Style = op['style']
    if op.get('header', True):
        table.Rows.Item(1).HeadingFormat = True


def refresh_fields(doc, toc):
    fields = list(doc.Fields)
    # Populate the number fields before resolving their bookmark references.
    for field_type in (12, 3, 37):
        for item in fields:
            if item.Type == field_type and not item.Locked:
                item.Update()
    if toc:
        for item in doc.TablesOfContents:
            item.Update()
    doc.Repaginate()


def apply(app, doc, operations):
    pending = set()
    refresh_toc = False
    with mathtype_writer(app) as write_mathtype:
        for op in operations:
            kind = op['op']
            if kind == 'configure':
                configure(doc, op)
            elif kind == 'paragraph':
                add_paragraph(write_mathtype, doc, op, pending)
            elif kind == 'format_paragraph':
                format_paragraph(doc, op)
            elif kind == 'equation':
                add_equation(write_mathtype, doc, op)
            elif kind == 'replace_mathtype':
                replace_mathtype(write_mathtype, doc, op)
            elif kind in {'reference', 'caption'}:
                add_labelled_paragraph(doc, op)
            elif kind == 'figure':
                add_figure(doc, op)
            elif kind == 'table':
                add_table(doc, op)
            elif kind == 'toc':
                doc.TablesOfContents.Add(paragraph(doc, 'normal'))
                refresh_toc = True
            elif kind == 'refresh_toc':
                refresh_toc = True
            elif kind == 'header_footer':
                section = doc.Sections.Item(op.get('section', 1))
                location = op['kind']
                if location not in {'header', 'footer'}:
                    raise ValueError('Header/footer kind must be header or footer')
                area = (section.Headers if location == 'header' else section.Footers).Item(1)
                if section.Index > 1:
                    area.LinkToPrevious = False
                area.Range.Text = op['text']
                if 'align' in op:
                    area.Range.ParagraphFormat.Alignment = {'left': 0, 'center': 1, 'right': 2}[op['align']]
            elif kind == 'page_numbers':
                section = doc.Sections.Item(op.get('section', 1))
                section.Footers.Item(1).PageNumbers.Add(2)
    missing = sorted(name for name in pending if not doc.Bookmarks.Exists(name))
    if missing:
        raise ValueError('Unresolved Word reference targets: ' + ', '.join(missing))
    refresh_fields(doc, refresh_toc)


def snapshot(doc, options):
    state = {
        'pages': int(doc.ComputeStatistics(2)), 'paragraphs': int(doc.Paragraphs.Count),
        'sections': int(doc.Sections.Count), 'tables': int(doc.Tables.Count),
        'word_equations': int(doc.OMaths.Count),
        'mathtype_equations': sum(1 for shape in doc.InlineShapes if is_mathtype(shape)),
        'bookmarks': [item.Name for item in doc.Bookmarks],
        'fields': [{'code': item.Code.Text.strip(), 'result': item.Result.Text}
                   for item in doc.Fields if item.Type in {3, 12, 37}],
        'quality_verdict': 'not_assessed',
    }
    if options.get('text'):
        state['text'] = doc.Content.Text
    return state
