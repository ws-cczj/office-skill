"""PowerPoint COM operations. No MCP, file-authoring library, or UI click dependency."""

from hashlib import sha256
import json
import math
import os
from pathlib import Path
import shutil
import time
import traceback

try:
    from .desktop_lock import DesktopLock
except ImportError:
    from desktop_lock import DesktopLock

VERSION = '1.0.0'
OPS = {'add_slide', 'duplicate_slide', 'move_slide', 'delete_slide', 'text', 'shape',
       'line', 'connector', 'picture', 'freeform', 'update', 'delete', 'group', 'ungroup',
       'notes', 'table', 'table_cell', 'chart', 'chart_data', 'animation', 'snapshot', 'export', 'pdf', 'save', 'script'}
SHAPES = {'rectangle': 1, 'rounded_rectangle': 5, 'oval': 9, 'triangle': 7,
          'right_arrow': 33, 'down_arrow': 36, 'chevron': 52}
TYPES = {1: 'auto_shape', 3: 'chart', 5: 'freeform', 6: 'group', 7: 'embedded_ole',
         9: 'line', 11: 'linked_picture', 13: 'picture', 14: 'placeholder', 17: 'text_box',
         19: 'table', 24: 'smart_art', 28: 'graphic'}


def absolute(value):
    path = Path(value)
    if not path.is_absolute():
        raise ValueError(f'Absolute path required: {value}')
    return path.resolve()


def rgb(value):
    text = value.lstrip('#')
    if len(text) != 6 or any(c not in '0123456789abcdefABCDEF' for c in text):
        raise ValueError(f'Invalid RGB color: {value}')
    r, g, b = (int(text[i:i+2], 16) for i in (0, 2, 4))
    return r | g << 8 | b << 16


def utf16_length(text):
    return len(text.encode('utf-16-le')) // 2


def bounds(value):
    if not isinstance(value, list) or len(value) != 4:
        raise ValueError('box must be [left, top, width, height] in points')
    if any(isinstance(n, bool) or not isinstance(n, (float, int)) or not math.isfinite(n) for n in value):
        raise ValueError('box values must be finite numbers')
    if value[2] <= 0 or value[3] <= 0:
        raise ValueError('box width and height must be positive')
    return value


def normalize_text(text):
    return str(text).replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r')


def collection(items):
    return [items.Item(i) for i in range(1, items.Count + 1)]


def walk_shapes(items):
    for shape in collection(items):
        yield shape
        if int(shape.Type) == 6:
            yield from walk_shapes(shape.GroupItems)


def selector_match(obj, selector, *, slide=False):
    if isinstance(selector, str):
        selector = {'name': selector}
    if not isinstance(selector, dict) or len(selector) != 1:
        raise ValueError('Use one selector: id, name, or tag (slides also allow index)')
    key, value = next(iter(selector.items()))
    if key == 'id':
        return (obj.SlideID if slide else obj.Id) == value
    if key == 'name':
        return obj.Name == value
    if key == 'tag':
        return obj.Tags.Item('PPTSKILL').casefold() == str(value).casefold()
    if slide and key == 'index':
        return obj.SlideIndex == value
    raise ValueError(f'Unknown selector: {key}')


def find_unique(items, selector, *, slide=False):
    matches = [item for item in items if selector_match(item, selector, slide=slide)]
    if len(matches) != 1:
        raise ValueError(f'Selector {selector!r} matched {len(matches)} objects; expected exactly one')
    return matches[0]


def validate_plan(plan):
    if plan.get('schema_version') != 1:
        raise ValueError('schema_version must be 1')
    target = absolute(plan['target'])
    if target.suffix.lower() not in {'.pptx', '.pptm', '.potx', '.potm'}:
        raise ValueError('Unsupported target extension')
    if not isinstance(plan.get('operations'), list):
        raise ValueError('operations must be an array')
    create = plan.get('create')
    if create is not None:
        if not isinstance(create, dict):
            raise ValueError('create must be an object')
        if create.get('template'):
            source = absolute(create['template'])
            if not source.is_file() or source == target:
                raise ValueError('Template must exist and differ from target')
            if source.suffix.lower() != target.suffix.lower():
                raise ValueError('Template copy must keep its extension; convert explicitly using PowerPoint')
        selected = create.get('slides')
        if selected is not None and (not isinstance(selected, list) or not selected or
                                    any(type(p) is not int or p < 1 for p in selected) or
                                    len(set(selected)) != len(selected) or not create.get('template')):
            raise ValueError('create.slides must be unique positive source indices with a template')
    protected = {target}
    if create and create.get('template'):
        protected.add(absolute(create['template']))
    for op in plan['operations']:
        if isinstance(op, dict) and op.get('op') in {'picture', 'script'}:
            protected.add(absolute(op['path']))
    if 'verify_export' in plan and absolute(plan['verify_export']) in protected:
        raise ValueError('verify_export must differ from input files')
    if 'verify_snapshot' in plan:
        output = absolute(plan['verify_snapshot'])
        if output in protected or output.suffix.lower() != '.json':
            raise ValueError('verify_snapshot needs a separate .json path')
    for op in plan['operations']:
        if not isinstance(op, dict) or op.get('op') not in OPS:
            raise ValueError(f'Unknown operation: {op}')
        kind = op['op']
        required = {'text': ('box',), 'shape': ('box',), 'picture': ('box', 'path'),
                    'table': ('box', 'data'), 'chart': ('box', 'categories', 'series'),
                    'line': ('start', 'end'), 'connector': ('start', 'end'),
                    'freeform': ('points',), 'group': ('shapes',),
                    'update': ('shape',), 'delete': ('shape',), 'ungroup': ('shape',),
                    'animation': ('shape',), 'notes': ('text',), 'table_cell': ('shape', 'row', 'column', 'text'),
                    'chart_data': ('shape', 'categories', 'series'),
                    'snapshot': ('path',), 'export': ('path',), 'pdf': ('path',), 'script': ('path',)}
        for field in required.get(kind, ()):
            if field not in op:
                raise ValueError(f'{kind} requires {field}')
        if kind not in {'add_slide', 'snapshot', 'export', 'pdf', 'save', 'script'} and 'slide' not in op:
            raise ValueError(f'{kind} requires a slide selector')
        if kind in {'line', 'connector', 'freeform'}:
            points = op['points'] if kind == 'freeform' else [op['start'], op['end']]
            if len(points) < 2 or any(not isinstance(p, list) or len(p) != 2 or
                                      any(type(n) not in (float, int) or not math.isfinite(n) for n in p)
                                      for p in points):
                raise ValueError('Points must be pairs of finite numbers')
        def colors(style):
            if 'color' in style:
                rgb(style['color'])
            fill = style.get('fill')
            if isinstance(fill, str):
                rgb(fill)
            elif isinstance(fill, dict):
                rgb(fill['start'])
                rgb(fill['end'])
            if style.get('line'):
                rgb(style['line'].get('color', '14386C'))
            for run in style.get('runs', []):
                colors(run)
        colors(op)
        for field in ('header_style', 'cell_style'):
            if field in op:
                colors(op[field])
        if 'box' in op:
            bounds(op['box'])
        if op['op'] in {'text', 'shape', 'picture', 'table', 'chart', 'line', 'connector', 'freeform', 'group'}:
            if not isinstance(op.get('name'), str) or not op['name'].strip():
                raise ValueError('New objects require a unique nonempty name')
        if op['op'] in {'snapshot', 'export', 'pdf'}:
            output = absolute(op['path'])
            if output in protected:
                raise ValueError('Output must differ from presentation/template/input files')
            if op['op'] == 'snapshot' and output.suffix.lower() != '.json':
                raise ValueError('Snapshot path must end in .json')
            if op['op'] == 'pdf' and output.suffix.lower() != '.pdf':
                raise ValueError('PDF path must end in .pdf')
        if op['op'] in {'picture', 'script'}:
            if not absolute(op['path']).is_file():
                raise ValueError(f'Input does not exist: {op["path"]}')
        if op['op'] == 'picture' and op.get('fit', 'contain') not in {'contain', 'cover', 'stretch'}:
            raise ValueError('Unknown picture fit')
        if op['op'] == 'shape' and isinstance(op.get('type', 'rectangle'), str) and op.get('type', 'rectangle') not in SHAPES:
            raise ValueError('Unknown shape type')
        if op['op'] == 'script':
            compile(absolute(op['path']).read_text(encoding='utf-8-sig'), op['path'], 'exec')
        if op['op'] == 'table':
            rows = op.get('data', [])
            if not rows or not rows[0] or any(len(row) != len(rows[0]) for row in rows):
                raise ValueError('Table data must be nonempty and rectangular')
        if op['op'] in {'chart', 'chart_data'}:
            if not op.get('categories') or not op.get('series'):
                raise ValueError('Chart needs categories and series')
            if any(len(s['values']) != len(op['categories']) for s in op['series']):
                raise ValueError('Each series must match category count')
    return target


def set_font(font, style):
    if 'font' in style:
        font.Name = style['font']
        font.NameFarEast = style.get('font_east_asia', style['font'])
    elif 'font_east_asia' in style:
        font.NameFarEast = style['font_east_asia']
    for prop, attr in [('size', 'Size'), ('bold', 'Bold'), ('italic', 'Italic')]:
        if prop in style:
            setattr(font, attr, (-1 if style[prop] else 0) if prop in {'bold', 'italic'} else style[prop])
    if 'color' in style:
        font.Fill.ForeColor.RGB = rgb(style['color'])


def style_shape(shape, spec, *, cell=False):
    fixed_box = [shape.Left, shape.Top, shape.Width, shape.Height] if shape.HasTextFrame and not cell else None
    if fixed_box is not None:
        original_auto = shape.TextFrame2.AutoSize
        shape.TextFrame2.AutoSize = 0
    if 'box' in spec:
        shape.LockAspectRatio = 0
        shape.Left, shape.Top, shape.Width, shape.Height = bounds(spec['box'])
    if 'rotation' in spec:
        shape.Rotation = spec['rotation']
    if 'alt_text' in spec:
        shape.AlternativeText = spec['alt_text']
    if 'fill' in spec:
        fill = spec['fill']
        if fill is None:
            shape.Fill.Visible = 0
        elif isinstance(fill, str):
            shape.Fill.Visible = -1
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = rgb(fill)
        else:
            shape.Fill.Visible = -1
            shape.Fill.TwoColorGradient(fill.get('style', 1), fill.get('variant', 1))
            shape.Fill.ForeColor.RGB = rgb(fill['start'])
            shape.Fill.BackColor.RGB = rgb(fill['end'])
            shape.Fill.GradientAngle = fill.get('angle', 0)
    if 'line' in spec:
        line = spec['line']
        shape.Line.Visible = 0 if line is None else -1
        if line is not None:
            shape.Line.ForeColor.RGB = rgb(line.get('color', '14386C'))
            shape.Line.Weight = line.get('width', 1)
            if 'dash' in line:
                shape.Line.DashStyle = line['dash']
            if 'arrow_end' in line:
                shape.Line.EndArrowheadStyle = line['arrow_end']
    if 'shadow' in spec:
        shape.Shadow.Visible = -1 if spec['shadow'] else 0
    if 'text' in spec or 'runs' in spec:
        frame = shape.TextFrame2
        # Preserve template geometry and autosizing after assignment.
        box = [shape.Left, shape.Top, shape.Width, shape.Height] if not cell else None
        runs = spec.get('runs')
        text = ''.join(r['text'] for r in runs) if runs is not None else spec['text']
        frame.TextRange.Text = normalize_text(text)
        if box is not None:
            shape.Left, shape.Top, shape.Width, shape.Height = box
    if any(k in spec for k in ('font', 'font_east_asia', 'size', 'bold', 'italic', 'color', 'align', 'margins', 'vertical', 'wrap')):
        frame = shape.TextFrame2
        set_font(frame.TextRange.Font, spec)
        if 'align' in spec:
            frame.TextRange.ParagraphFormat.Alignment = {'left': 1, 'center': 2, 'right': 3}[spec['align']]
        if 'vertical' in spec:
            frame.VerticalAnchor = {'top': 1, 'middle': 3, 'bottom': 4}[spec['vertical']]
        if 'wrap' in spec:
            frame.WordWrap = -1 if spec['wrap'] else 0
        if 'margins' in spec:
            for side, value in zip(('Left', 'Top', 'Right', 'Bottom'), spec['margins']):
                setattr(frame, 'Margin' + side, value)
    if spec.get('runs'):
        offset = 1
        for run in spec['runs']:
            length = utf16_length(normalize_text(run['text']))
            if length:
                set_font(shape.TextFrame2.TextRange.GetCharacters(offset, length).Font, run)
            offset += length
    if 'zorder' in spec:
        shape.ZOrder({'front': 0, 'back': 1, 'forward': 2, 'backward': 3}[spec['zorder']])
    if fixed_box is not None:
        # Do this after font/runs/paragraph changes as well as text assignment.
        shape.TextFrame2.AutoSize = original_auto
        shape.Left, shape.Top, shape.Width, shape.Height = spec.get('box', fixed_box)


def shape_snapshot(shape):
    record = {'id': int(shape.Id), 'name': shape.Name, 'raw_type': int(shape.Type),
              'kind': TYPES.get(int(shape.Type), 'other'), 'tag': shape.Tags.Item('PPTSKILL'),
              'box': [round(float(getattr(shape, p)), 3) for p in ('Left', 'Top', 'Width', 'Height')]}
    if shape.Connector:
        record['kind'] = 'connector'
    if shape.HasTextFrame:
        frame = shape.TextFrame2
        record['text'] = frame.TextRange.Text.replace('\r', '\n')
        try:
            text = frame.TextRange
            record['text_bounds'] = [round(float(getattr(text, p)), 3)
                                     for p in ('BoundLeft', 'BoundTop', 'BoundWidth', 'BoundHeight')]
            record['font'] = {'latin': text.Font.Name, 'east_asia': text.Font.NameFarEast,
                              'size': float(text.Font.Size)}
        except Exception as exc:
            record['text_metrics_unavailable'] = str(exc)
    if int(shape.Type) == 6:
        record['children'] = [shape_snapshot(s) for s in collection(shape.GroupItems)]
    if shape.HasChart:
        record['chart_type'] = int(shape.Chart.ChartType)
        record['series'] = [{'name': s.Name, 'values': list(s.Values), 'categories': list(s.XValues)}
                            for s in collection(shape.Chart.SeriesCollection())]
    if shape.HasTable:
        table = shape.Table
        record['table'] = [[table.Cell(r, c).Shape.TextFrame.TextRange.Text for c in range(1, table.Columns.Count+1)]
                           for r in range(1, table.Rows.Count+1)]
    return record


def set_chart_data(chart, spec):
    if chart.ChartData.IsLinked:
        raise ValueError('Linked chart data is external; copy/embed it explicitly before editing')
    chart.ChartData.Activate()
    book = chart.ChartData.Workbook
    try:
        sheet = book.Worksheets.Item(1)
        sheet.Cells.Clear()
        sheet.Cells(1, 1).Value = spec.get('category_label', '')
        for row, category in enumerate(spec['categories'], 2):
            sheet.Cells(row, 1).Value = category
        for col, series in enumerate(spec['series'], 2):
            sheet.Cells(1, col).Value = series['name']
            for row, value in enumerate(series['values'], 2):
                sheet.Cells(row, col).Value = value
        area = sheet.Range(sheet.Cells(1, 1), sheet.Cells(len(spec['categories'])+1, len(spec['series'])+1))
        sheet_name = sheet.Name.replace("'", "''")
        chart.SetSourceData(f"'{sheet_name}'!{area.Address}", 2)
    finally:
        book.Close(True)
    if 'title' in spec:
        chart.HasTitle = -1
        chart.ChartTitle.Text = spec['title']


class PowerPoint:
    def __init__(self, app):
        self.app = app
        self.presentation = None
        self.target = None
        self.opened_here = False

    def bind(self, target, create=None):
        self.target = absolute(target)
        if create is not None and self.target.exists():
            raise FileExistsError(f'Will not overwrite an existing target: {self.target}')
        matches = [p for p in collection(self.app.Presentations)
                   if os.path.normcase(os.path.abspath(p.FullName)) == os.path.normcase(str(self.target))]
        if matches:
            if len(matches) != 1 or create is not None:
                raise ValueError('Target must identify exactly one existing presentation')
            self.presentation = matches[0]
        else:
            self.target.parent.mkdir(parents=True, exist_ok=True)
            if create is not None and create.get('template'):
                shutil.copy2(absolute(create['template']), self.target)
            if self.target.is_file():
                self.presentation = self.app.Presentations.Open(str(self.target), 0, 0, -1)
            elif create is not None:
                if self.target.suffix.lower() != '.pptx':
                    raise ValueError('Blank creation supports .pptx; other formats require a source')
                self.presentation = self.app.Presentations.Add(-1)
                self.presentation.PageSetup.SlideWidth = create.get('width', 960)
                self.presentation.PageSetup.SlideHeight = create.get('height', 540)
                self.presentation.SaveAs(str(self.target), 24)
            else:
                raise FileNotFoundError(self.target)
            self.opened_here = True
        if self.presentation.ReadOnly:
            raise ValueError('Target is read-only')
        self.verify()
        if create is not None and 'slides' in create:
            slides = collection(self.presentation.Slides)
            selected = create['slides']
            if max(selected) > len(slides):
                raise ValueError('Requested source slide is out of range')
            ids = [slides[i-1].SlideID for i in selected]
            for item in reversed(slides):
                if item.SlideID not in ids:
                    item.Delete()
            for index, slide_id in enumerate(ids, 1):
                self.presentation.Slides.FindBySlideID(slide_id).MoveTo(index)

    def verify(self):
        if os.path.normcase(os.path.abspath(self.presentation.FullName)) != os.path.normcase(str(self.target)):
            raise RuntimeError('Bound presentation path changed; operation stopped')

    def slide(self, selector):
        if type(selector) is int:
            selector = {'index': selector}
        return find_unique(collection(self.presentation.Slides), selector, slide=True)

    def shape(self, slide, selector):
        return find_unique(walk_shapes(slide.Shapes), selector)

    def register(self, slide, shape, name):
        if name:
            shape.Name = name
            shape.Tags.Add('PPTSKILL', name)
        return shape

    def check_slide_name(self, name):
        if name and any(s.Name == name or s.Tags.Item('PPTSKILL').casefold() == str(name).casefold()
                        for s in collection(self.presentation.Slides)):
            raise ValueError(f'Slide name/tag already exists: {name}')

    def name_slide(self, slide, name=None):
        if name:
            slide.Name = name
            slide.Tags.Add('PPTSKILL', name)
        elif slide.Tags.Item('PPTSKILL'):
            slide.Tags.Delete('PPTSKILL')

    def snapshot(self, pages=None):
        selected = [self.slide(s) for s in pages] if pages is not None else collection(self.presentation.Slides)
        result = []
        for slide in selected:
            objects = [shape_snapshot(s) for s in collection(slide.Shapes)]
            notes = '\n'.join(s.TextFrame.TextRange.Text.replace('\r', '\n')
                              for s in collection(slide.NotesPage.Shapes)
                              if int(s.Type) == 14 and int(s.PlaceholderFormat.Type) == 2)
            result.append({'index': int(slide.SlideIndex), 'id': int(slide.SlideID), 'name': slide.Name,
                           'tag': slide.Tags.Item('PPTSKILL'), 'notes': notes, 'shapes': objects})
        return {'path': str(self.target), 'width': float(self.presentation.PageSetup.SlideWidth),
                'height': float(self.presentation.PageSetup.SlideHeight), 'slides': result}

    def execute(self, op):
        self.verify()
        kind = op['op']
        pres = self.presentation
        if kind == 'add_slide':
            self.check_slide_name(op.get('name'))
            slide = pres.Slides.Add(op.get('index', pres.Slides.Count+1), op.get('layout', 12))
            self.name_slide(slide, op.get('name'))
            return {'slide_id': int(slide.SlideID), 'index': int(slide.SlideIndex)}
        if kind == 'save':
            pres.Save()
            return {'saved': str(self.target)}
        if kind == 'snapshot':
            path = absolute(op['path'])
            if path == self.target or path.suffix.lower() != '.json':
                raise ValueError('Snapshot needs a separate .json path')
            path.parent.mkdir(parents=True, exist_ok=True)
            data = self.snapshot(op.get('pages'))
            path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
            return {'snapshot': str(path), 'pages': len(data['slides'])}
        if kind in {'export', 'pdf'}:
            path = absolute(op['path'])
            if kind == 'pdf':
                path.parent.mkdir(parents=True, exist_ok=True)
                # Generated COM wrappers may default this dispatch argument to 0.
                pres.ExportAsFixedFormat(str(path), 2, PrintRange=None)
                return {'pdf': str(path)}
            path.mkdir(parents=True, exist_ok=True)
            slides = [self.slide(s) for s in op['pages']] if 'pages' in op else collection(pres.Slides)
            width = op.get('width', 1600)
            height = op.get('height', round(width * pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth))
            paths = []
            for slide in slides:
                output = path / f'{slide.SlideIndex:02d}.png'
                slide.Export(str(output), 'PNG', width, height)
                paths.append(str(output))
            return {'images': paths}
        if kind == 'script':
            context = {'app': self.app, 'presentation': pres, 'api': self, '__file__': str(absolute(op['path']))}
            exec(compile(absolute(op['path']).read_text(encoding='utf-8-sig'), op['path'], 'exec'), context)
            self.verify()
            # Extension scripts should put a concise JSON-serializable result here.
            result = context.get('result')
            json.dumps(result)
            return {'result': result}
        slide = self.slide(op['slide'])
        if kind in {'duplicate_slide', 'move_slide', 'delete_slide'}:
            if kind == 'delete_slide':
                slide_id = int(slide.SlideID)
                slide.Delete()
                return {'deleted_slide_id': slide_id}
            if kind == 'duplicate_slide':
                self.check_slide_name(op.get('name'))
                slide = slide.Duplicate().Item(1)
                self.name_slide(slide, op.get('name'))
            if 'index' in op:
                slide.MoveTo(op['index'])
            return {'slide_id': int(slide.SlideID), 'index': int(slide.SlideIndex)}
        if kind == 'notes':
            matches = [s for s in collection(slide.NotesPage.Shapes)
                       if int(s.Type) == 14 and int(s.PlaceholderFormat.Type) == 2]
            if len(matches) != 1:
                raise ValueError('Expected one notes body placeholder')
            matches[0].TextFrame.TextRange.Text = normalize_text(op['text'])
            return {'notes_written': True}
        if kind == 'chart_data':
            shape = self.shape(slide, op['shape'])
            if not shape.HasChart:
                raise ValueError('Selected object is not a native chart')
            set_chart_data(shape.Chart, op)
            return {'id': int(shape.Id), 'data_updated': True}
        if kind == 'table_cell':
            shape = self.shape(slide, op['shape'])
            if not shape.HasTable:
                raise ValueError('Selected object is not a native table')
            cell = shape.Table.Cell(op['row'], op['column']).Shape
            style_shape(cell, op, cell=True)
            return {'id': int(shape.Id), 'cell_updated': [op['row'], op['column']]}
        if kind in {'update', 'delete', 'ungroup', 'animation'}:
            shape = self.shape(slide, op['shape'])
            if kind == 'delete':
                sid = int(shape.Id)
                shape.Delete()
                return {'deleted_id': sid}
            if kind == 'ungroup':
                return {'ungrouped': int(shape.Ungroup().Count)}
            if kind == 'animation':
                effect = slide.TimeLine.MainSequence.AddEffect(shape, op.get('effect', 10), 0, op.get('trigger', 1))
                effect.Timing.Duration = op.get('duration', .5)
                effect.Timing.TriggerDelayTime = op.get('delay', 0)
                return {'effect': int(effect.EffectType)}
            style_shape(shape, op)
            return {'id': int(shape.Id), 'name': shape.Name}
        name = op.get('name')
        if not name:
            raise ValueError('New objects require a unique name')
        if any(s.Name == name or s.Tags.Item('PPTSKILL').casefold() == name.casefold() for s in walk_shapes(slide.Shapes)):
            raise ValueError(f'Object name/tag already exists: {name}')
        if kind == 'text':
            shape = slide.Shapes.AddTextbox(1, *bounds(op['box']))
            shape.TextFrame2.AutoSize = 0
            shape.Fill.Visible = 0
            shape.Line.Visible = 0
            shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb('14386C')
        elif kind == 'shape':
            shape_type = SHAPES[op.get('type', 'rectangle')] if isinstance(op.get('type', 'rectangle'), str) else op['type']
            shape = slide.Shapes.AddShape(shape_type, *bounds(op['box']))
            shape.Shadow.Visible = 0
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = rgb('FFFFFF')
            shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb('14386C')
        elif kind in {'line', 'connector'}:
            start, end = op['start'], op['end']
            shape = (slide.Shapes.AddLine(*start, *end) if kind == 'line' else
                     slide.Shapes.AddConnector(op.get('type', 1), *start, *end))
            if kind == 'connector':
                for side in ('begin', 'end'):
                    if 'connect_'+side in op:
                        ref = op['connect_'+side]
                        getattr(shape.ConnectorFormat, side.capitalize()+'Connect')(
                            self.shape(slide, ref['shape']), ref.get('site', 1))
                shape.RerouteConnections()
        elif kind == 'picture':
            left, top, width, height = bounds(op['box'])
            shape = slide.Shapes.AddPicture(str(absolute(op['path'])), 0, -1, left, top, -1, -1)
            self.register(slide, shape, name)
            pw, ph = float(shape.Width), float(shape.Height)
            mode = op.get('fit', 'contain')
            if mode == 'contain':
                factor = min(width/pw, height/ph)
                shape.LockAspectRatio = -1
                shape.Width = pw*factor
                shape.Left = left+(width-shape.Width)/2
                shape.Top = top+(height-shape.Height)/2
            elif mode == 'cover':
                if pw/ph > width/height:
                    crop = (pw-ph*width/height)/2
                    shape.PictureFormat.CropLeft = crop
                    shape.PictureFormat.CropRight = crop
                else:
                    crop = (ph-pw*height/width)/2
                    shape.PictureFormat.CropTop = crop
                    shape.PictureFormat.CropBottom = crop
                shape.LockAspectRatio = 0
                shape.Width, shape.Height = width, height
            elif mode == 'stretch':
                shape.LockAspectRatio = 0
                shape.Width, shape.Height = width, height
            else:
                raise ValueError('fit must be contain, cover, or stretch')
            if mode != 'contain':
                # Cropping can move the shape in PowerPoint; restore the viewport.
                shape.Left, shape.Top = left, top
        elif kind == 'freeform':
            points = op['points']
            builder = slide.Shapes.BuildFreeform(1, *points[0])
            for point in points[1:]:
                builder.AddNodes(0, 0, *point)
            shape = builder.ConvertToShape()
        elif kind == 'table':
            data = op['data']
            shape = slide.Shapes.AddTable(len(data), len(data[0]), *bounds(op['box']))
            self.register(slide, shape, name)
            for r, row in enumerate(data, 1):
                for c, text in enumerate(row, 1):
                    cell = shape.Table.Cell(r, c).Shape
                    style_shape(cell, {'text': str(text), **op.get('cell_style', {})}, cell=True)
                    if r == 1:
                        style_shape(cell, op.get('header_style', {}), cell=True)
        elif kind == 'chart':
            shape = slide.Shapes.AddChart2(-1, op.get('type', 51), *bounds(op['box']), -1)
            self.register(slide, shape, name)
            set_chart_data(shape.Chart, op)
        elif kind == 'group':
            members = [self.shape(slide, s) for s in op['shapes']]
            if any(int(s.Parent.SlideID) != int(slide.SlideID) for s in members):
                raise ValueError('Grouping requires top-level shapes on the same slide')
            shape = slide.Shapes.Range([s.Name for s in members]).Group()
        else:
            raise ValueError(f'Unsupported operation: {kind}')
        self.register(slide, shape, name)
        # Picture fit has already placed and sized the image.
        style_shape(shape, {k: v for k, v in op.items() if not (kind == 'picture' and k == 'box')})
        return {'id': int(shape.Id), 'name': shape.Name, 'raw_type': int(shape.Type)}


def connect():
    if os.name != 'nt':
        raise RuntimeError('Desktop PowerPoint automation requires Windows')
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch('PowerPoint.Application')
    app.Visible = -1
    return app


def run_plan(plan, report_path, app_factory=connect):
    target = validate_plan(plan)
    digest = sha256(json.dumps(plan, sort_keys=True, ensure_ascii=False).encode('utf-8')).hexdigest()
    report_path = absolute(report_path)
    if report_path == target or report_path.suffix.lower() != '.json':
        raise ValueError('Report needs a separate .json path')
    outputs = [absolute(o['path']) for o in plan['operations'] if o['op'] in {'snapshot', 'pdf', 'export'}]
    if 'verify_snapshot' in plan:
        outputs.append(absolute(plan['verify_snapshot']))
    if 'verify_export' in plan:
        outputs.append(absolute(plan['verify_export']))
    if report_path in outputs:
        raise ValueError('Report path conflicts with an operation output')
    if report_path.exists():
        previous = json.loads(report_path.read_text(encoding='utf-8'))
        if previous.get('plan_sha256') == digest and previous.get('status') == 'completed' and previous.get('saved') and target.is_file():
            if previous.get('target_sha256') == sha256(target.read_bytes()).hexdigest():
                return {**previous, 'replay_skipped': True}
        raise RuntimeError('Report already exists; inspect actual state and write a new plan/report before continuing')
    report = {'engine': 'powerpoint-native', 'version': VERSION, 'target': str(target),
              'plan_sha256': digest, 'status': 'started', 'operations': [], 'mcp_loaded': False}
    report_path.parent.mkdir(parents=True, exist_ok=True)

    def persist():
        # Persist before and after each operation. An interrupted operation is never replayed automatically.
        tmp = report_path.with_suffix(report_path.suffix+'.pending')
        tmp.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        # Windows indexing/antivirus may briefly hold a file without delete sharing.
        # Retry only this filesystem replacement, never a PowerPoint operation.
        for delay in (0, .02, .05, .1, .2, .5):
            if delay:
                time.sleep(delay)
            try:
                tmp.replace(report_path)
                break
            except PermissionError:
                if delay == .5:
                    raise

    started = time.perf_counter()
    with DesktopLock():
        persist()
        api = None
        try:
            api = PowerPoint(app_factory())
            api.bind(target, plan.get('create'))
            if (plan.get('reopen') or plan.get('close')) and not api.opened_here and not plan.get('close_existing', False):
                raise RuntimeError('Target is already open; close_existing=true explicitly permits closing only this target')
            report['office_version'] = str(api.app.Version)
            for index, op in enumerate(plan['operations']):
                item = {'index': index, 'op': op['op'], 'status': 'started'}
                report['operations'].append(item)
                persist()
                tick = time.perf_counter()
                result = api.execute(op)
                item.update(status='completed', seconds=round(time.perf_counter()-tick, 4), result=result)
                persist()
            if plan.get('save', True):
                api.execute({'op': 'save'})
            if plan.get('reopen', False):
                if not plan.get('save', True):
                    raise ValueError('Reopen requires save=true')
                api.presentation.Close()
                api.presentation = api.app.Presentations.Open(str(target), 0, 0, -1)
                api.verify()
                report['reopened'] = True
                if 'verify_snapshot' in plan:
                    api.execute({'op': 'snapshot', 'path': plan['verify_snapshot']})
                if 'verify_export' in plan:
                    api.execute({'op': 'export', 'path': plan['verify_export']})
            report['slides'] = int(api.presentation.Slides.Count)
            report['saved'] = bool(api.presentation.Saved)
            if plan.get('close', False):
                if not plan.get('save', True) and not api.presentation.Saved:
                    raise RuntimeError('Refusing to close unsaved changes')
                api.presentation.Close()
            report['status'] = 'completed'
        except Exception as exc:
            report.update(status='failed', error=f'{type(exc).__name__}: {exc}',
                          operation_may_have_applied=True, traceback=traceback.format_exc())
            # The named document remains open for inspection; no implicit save, retry, or rollback.
        report['seconds'] = round(time.perf_counter()-started, 4)
        if target.is_file():
            report['target_sha256'] = sha256(target.read_bytes()).hexdigest()
        import sys
        report['mcp_loaded'] = any(n == 'mcp' or n.startswith(('mcp.', 'powerpoint_mcp')) for n in sys.modules)
        persist()
    return report
