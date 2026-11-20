"""Local JSON -> desktop Visio COM bridge; no MCP server or network listener."""
from __future__ import annotations
import argparse
import json
import math
from pathlib import Path
import re
import sys
import time
from contextlib import contextmanager
if not __package__:
    sys.path.insert(0, str(Path(__file__).resolve().parents[3]))
from core.desktop_lock import DesktopLock
from skills.visio.scripts import visio_extras, visio_geometry, vsdx_scene

_SRC_CACHE = {}
EXPORT_EXTENSIONS = {".png", ".svg", ".emf"}


def set_cells(shape, formulas):
    """Write style cells in one COM call; cache standard ShapeSheet addresses."""
    stream, values = [], []
    rows = shape.RowCount(3) if any(n.startswith('Char.') for n in formulas) else 1
    for name, formula in formulas.items():
        names = [name] + ([f'{name}[{i+1}]' for i in range(1, rows)] if name.startswith('Char.') else [])
        for n in names:
            if n not in _SRC_CACHE:
                c = shape.CellsU(n)
                _SRC_CACHE[n] = (c.Section, c.Row, c.Column)
            stream.extend(_SRC_CACHE[n])
            values.append(str(formula))
    if stream:
        count = shape.SetFormulas(tuple(stream), tuple(values), 10)
        if count != len(values):
            raise RuntimeError(f'Only {count}/{len(values)} formulas were written')


@contextmanager
def performance_scope(app, enabled=True):
    if not enabled:
        yield
        return
    screen, recalc = app.ScreenUpdating, app.DeferRecalc
    try:
        app.ScreenUpdating = False
        app.DeferRecalc = True
        yield
    finally:
        app.DeferRecalc = recalc
        app.ScreenUpdating = screen


@contextmanager
def png_resolution(app, dpi):
    if dpi is None:
        yield
        return
    settings = app.Settings
    resolution = settings.GetRasterExportResolution()
    size = settings.GetRasterExportSize()
    try:
        settings.SetRasterExportSize(2)
        settings.SetRasterExportResolution(3, dpi, dpi, 0)
        yield
    finally:
        try:
            settings.SetRasterExportResolution(*resolution) if resolution[0] == 3 else settings.SetRasterExportResolution(resolution[0])
        finally:
            settings.SetRasterExportSize(*size) if size[0] == 3 else settings.SetRasterExportSize(size[0])


class RequestError(ValueError):
    pass


def path(value, extensions=None):
    p = Path(value).expanduser()
    if not p.is_absolute():
        raise RequestError(f"Absolute path required: {value}")
    p = p.resolve()
    if extensions and p.suffix.lower() not in extensions:
        raise RequestError(f"Expected extension {extensions}: {p}")
    return p


def number(value):
    if isinstance(value, bool) or not isinstance(value, (float, int)) or not math.isfinite(value):
        raise RequestError(f"Expected finite number: {value!r}")
    return value


def rgb(value):
    if not isinstance(value, str) or not re.fullmatch(r"#[0-9A-Fa-f]{6}", value):
        raise RequestError(f"Expected #RRGGBB: {value!r}")
    return "RGB(%d,%d,%d)" % tuple(int(value[i:i+2], 16) for i in (1, 3, 5))


def cell(shape, name, formula):
    shape.CellsU(name).FormulaU = str(formula)
    # Text assignments can create separate runs for CJK/Latin characters.
    # Format every character row instead of only the first run.
    if name.startswith("Char."):
        for row in range(1, shape.RowCount(3)):
            shape.CellsU(f"{name}[{row+1}]").FormulaU = str(formula)


def connect():
    import pywintypes
    import win32com.client

    try:
        return win32com.client.GetActiveObject("Visio.Application"), "attached"
    except pywintypes.com_error:
        app = win32com.client.Dispatch("Visio.Application")
        app.Visible = True
        return app, "started"


def drawings(app):
    return [d for d in app.Documents if d.Type == 1]


def find_document(app, p):
    return next((d for d in drawings(app) if d.Path and Path(d.FullName).resolve() == p), None)


# visOpenHidden | visOpenMacrosDisabled | visOpenNoWorkspace
BACKGROUND_DRAWING_FLAGS = 64 | 128 | 256
OPEN_AS_COPY = 1


def target(app, req):
    p = path(req["document"], {".vsdx"})
    doc = find_document(app, p)
    if doc is None:
        if not p.is_file():
            raise RequestError(f"Document does not exist: {p}")
        doc = app.Documents.OpenEx(str(p), BACKGROUND_DRAWING_FLAGS)
    return doc, select_page(doc, req)


def select_page(doc, req):
    if "page" in req:
        return doc.Pages.ItemU(req["page"])
    if doc.Pages.Count == 1 or req['action'] == 'pages':
        return doc.Pages.Item(1)
    raise RequestError("Specify page for a document with multiple pages")


def style(s, op, doc):
    formulas = {}
    def put(shape, name, formula):
        formulas[name] = str(formula)
    for key, name in {"fill": "FillForegnd", "line": "LineColor", "text_color": "Char.Color"}.items():
        if key in op:
            put(s, name, rgb(op[key]))
    if "fill" in op:
        put(s, "FillPattern", 1)
    for key, name, unit in [("font_size", "Char.Size", "pt"), ("line_weight", "LineWeight", "pt"), ("rounding", "Rounding", "in")]:
        if key in op:
            put(s, name, f"{op[key]} {unit}")
    if "text_margin" in op:
        for side in ("Left", "Right", "Top", "Bottom"):
            put(s, side+"Margin", f"{op['text_margin']} in")
    if "font" in op:
        put(s, "Char.Font", 'FONT(' + json.dumps(op["font"], ensure_ascii=False) + ')')
        put(s, "Char.AsianFont", 'FONT(' + json.dumps(op["font"], ensure_ascii=False) + ')')
    for key, name in [("bold", "Char.Style"), ("align", "Para.HorzAlign"), ("line_pattern", "LinePattern"),
                      ("end_arrow", "EndArrow"), ("begin_arrow", "BeginArrow")]:
        if key in op:
            put(s, name, int(op[key]) if key == "bold" else op[key])
    if op.get("no_fill"):
        put(s, "FillPattern", 0)
    if op.get("no_line"):
        put(s, "LinePattern", 0)
    set_cells(s, formulas)


def inspect(doc, page, detail='full', recursive=False):
    w, h = (page.PageSheet.CellsU(k).ResultIU for k in ("PageWidth", "PageHeight"))
    shapes, warnings = [], []
    top_ids={s.ID for s in page.Shapes} if recursive else None
    if detail=='summary':
        return {'document':doc.FullName,'page':page.NameU,'width':w,'height':h,'shape_count':page.Shapes.Count,
                'detail':'summary; per-shape validation omitted'}
    def walk(collection):
        for s in collection:
            yield s
            if recursive and s.Type==2:
                yield from walk(s.Shapes)
    for s in walk(page.Shapes):
        item = {"id": s.ID, "key": s.NameU, "text": s.Text, "one_d": bool(s.OneD)}
        if recursive:item['coordinate_space']='page' if s.ID in top_ids else 'parent group'
        if s.OneD:
            item["connections"] = [{"from_cell": c.FromCell.Name, "to_id": c.ToSheet.ID} for c in s.Connects]
            item["route_style"] = s.CellsU("ShapeRouteStyle").ResultIU
        else:
            x, y, sw, sh = (s.CellsU(k).ResultIU for k in ("PinX", "PinY", "Width", "Height"))
            item.update(x=x, y=y, width=sw, height=sh)
            if (not recursive or s.ID in top_ids) and (x-sw/2 < -0.01 or y-sh/2 < -0.01 or x+sw/2 > w+0.01 or y+sh/2 > h+0.01):
                warnings.append({"shape_id": s.ID, "issue": "outside_page"})
        shapes.append(item)
        item['is_container']=False
        if s.CellExistsU('User.msvStructureType',0) and s.CellsU('User.msvStructureType').ResultStr('')=='Container':
            item['is_container']=True
            item['members']=list(s.ContainerProperties.GetMemberShapes(0))
        if s.Text:
            item["font_id"] = s.CellsU("Char.Font").ResultIU
            item["asian_font_id"] = s.CellsU("Char.AsianFont").ResultIU
    return {"document": doc.FullName, "page": page.NameU, "width": w, "height": h,
            "shape_count": len(shapes), "shapes": shapes, "warnings": warnings}


def preflight(req):
    action = req.get("action")
    if action not in {"status", "inspect", "render", "apply", "export", "save", "stencils", "masters", "layout", 'pages','page_size','icons','open_stencil','open','snapshot','rebuild'}:
        raise RequestError(f"Unknown action: {action}")
    if action not in {"status", "stencils", "masters",'icons','open_stencil'}:
        p = path(req["document"], {".vsdx"})
        if action == "render" and p.exists() and not req.get("overwrite", False):
            raise RequestError(f"Refusing to overwrite: {p}")
    for k in ("width", "height"):
        if k in req and number(req[k]) <= 0:
            raise RequestError(f"{k} must be positive")
    if req.get("save_as"):
        p = path(req["save_as"], {".vsdx"})
        if p.exists() and not req.get("overwrite", False):
            raise RequestError(f"Refusing to overwrite: {p}")
    for p in req.get("exports", []):
        path(p, EXPORT_EXTENSIONS)
    if action=='rebuild':
        source = path(req['source'], {'.vsdx'})
        if source == p or req.get('save_as'):
            raise RequestError('Rebuild requires one output document distinct from the source')
        if not source.is_file():
            raise RequestError(f'Source document does not exist: {source}')
        if p.exists() and not req.get('overwrite'):
            raise RequestError(f'Rebuild output already exists: {p}')
        if req.get('scene_output'):
            path(req['scene_output'], {'.json'})
    if req.get('detail','full') not in {'full','summary'}:raise RequestError('detail must be full or summary')
    if req.get('show') and req.get('close_after'):raise RequestError('show and close_after cannot both be true')
    if 'png_dpi' in req and not 36<=number(req['png_dpi'])<=600:raise RequestError('png_dpi must be 36–600')
    if action=='layout' and req.get('style','default') not in {'default','flowchart_tb','flowchart_lr','tree_tb','tree_lr','radial','circular'}:
        raise RequestError('Unknown layout preset')
    for op in req.get("operations", []):
        if op.get("op") not in {"add", "connect", "update", "delete",'container','members','image','badge','group'}:
            raise RequestError(f"Unknown operation: {op.get('op')}")
        for k in ("x", "y", "width", "height", "font_size", "line_weight", "rounding", "text_margin", "align", "line_pattern", "end_arrow", "begin_arrow"):
            if k in op:
                number(op[k])
        for k in ("width", "height"):
            if k in op and op[k] <= 0:
                raise RequestError(f"Shape {k} must be positive")
        for k in ("fill", "line", "text_color"):
            if k in op:
                rgb(op[k])
        if op["op"] in {"add", "connect",'container','image','badge','group'} and not re.fullmatch(r"[A-Za-z][A-Za-z0-9_-]*", op.get("key", "")):
            raise RequestError("New shape key must match [A-Za-z][A-Za-z0-9_-]*")
        if op["op"] == "add" and op.get("kind", "rectangle") not in {"rectangle", "ellipse", "diamond", "text", "master", "cylinder", "polyline"}:
            raise RequestError(f"Unknown shape kind: {op['kind']}")
        if op['op'] == 'add' and op.get('kind') == 'polyline':
            points = op.get('points')
            if not isinstance(points, (list, tuple)) or len(points) < 6 or len(points) % 2:
                raise RequestError('Polyline requires at least three [x, y] pairs as a flat array')
            for coordinate in points:
                number(coordinate)
        if op["op"] == "add" and op.get("kind") == "cylinder":
            cylinder_geometry(op)
            if "cap_fill" in op:
                rgb(op["cap_fill"])
        if op.get("route", "right_angle") not in {"right_angle", "straight",'curved'}:
            raise RequestError("route must be right_angle, straight or curved")
        for k in ('padding','size'):
            if k in op and number(op[k])<0:raise RequestError(f'{k} must not be negative')
        if 'corner' in op and op['corner'] not in {'top_left','top_right','bottom_left','bottom_right'}:
            raise RequestError('Unknown badge corner')
        if op['op']=='members' and op.get('mode','add') not in {'add','remove'}:raise RequestError('Unknown membership mode')


def cylinder_geometry(op):
    def point(name):
        xy = op.get(name)
        if not isinstance(xy,(list,tuple)) or len(xy)!=2:
            raise RequestError(f"Cylinder {name} must be [x, y]")
        return tuple(number(value) for value in xy)
    return visio_geometry.cylinder_paths(point("rear"),point("front"),number(op["radius"]),
                                         number(op["end_ratio"]),number(op.get("rear_scale",1)))


def draw_cylinder(page,doc,op):
    body,cap=cylinder_geometry(op)
    key=op["key"]
    pieces=[]
    for suffix,points,fill in (("Side",body,op.get("fill","#146661")),
                               ("Cap",cap,op.get("cap_fill","#092B29"))):
        piece=page.DrawBezier(points,3,0)
        piece.NameU=f"{key}_{suffix}"
        style(piece,{"fill":fill,"no_line":True},doc)
        pieces.append(piece)
    group=visio_extras.selection(page,pieces).Group()
    group.NameU=key
    return group


def batch(app, doc, page, ops):
    lookup, by_id = {}, {}
    stencil_cache,master_cache={},{}
    def index(collection,recursive=False):
        for s in collection:
            lookup[s.NameU]=s;by_id[s.ID]=s
            if recursive and s.Type==2:index(s.Shapes,True)
    index(page.Shapes)
    def resolve(ref):
        if isinstance(ref, int):
            if ref not in by_id:by_id[ref]=page.Shapes.ItemFromID(ref)
            return by_id[ref]
        if ref not in lookup:
            index(page.Shapes,True)
        if ref not in lookup:
            raise RequestError(f"Unknown shape key: {ref}")
        return lookup[ref]
    scope = app.BeginUndoScope("visio-skill batch")
    try:
        for op in ops:
            kind = op["op"]
            if kind in {"add", "connect",'container','image','badge','group'} and op["key"] in lookup:
                raise RequestError(f"Shape key already exists: {op['key']}")
            if kind == "add" and op.get("kind") == "cylinder":
                s=draw_cylinder(page,doc,op)
                lookup[s.NameU]=s
                by_id[s.ID]=s
                continue
            if kind == "add":
                sk = op.get("kind", "rectangle")
                if sk != "polyline":
                    x, y = op["x"], op["y"]
                    w, h = op.get("width", 2), op.get("height", 0.8)
                if sk == "polyline":
                    s = page.DrawPolyline(tuple(op["points"]), 0)
                elif sk == "ellipse":
                    s = page.DrawOval(x-w/2, y-h/2, x+w/2, y+h/2)
                elif sk == "diamond":
                    pts = [x,y+h/2,x+w/2,y,x,y-h/2,x-w/2,y,x,y+h/2]
                    s = page.DrawPolyline(tuple(pts), 0)
                elif sk == "master":
                    pair=(op['stencil'],op['master'])
                    if pair not in master_cache:
                        if op['stencil'] not in stencil_cache:stencil_cache[op['stencil']]=app.Documents.OpenEx(op['stencil'],64)
                        master_cache[pair]=stencil_cache[op['stencil']].Masters.ItemU(op['master'])
                    s = page.Drop(master_cache[pair], x, y)
                    cell(s, "Width", f"{w} in")
                    cell(s, "Height", f"{h} in")
                else:
                    s = page.DrawRectangle(x-w/2, y-h/2, x+w/2, y+h/2)
                if sk == "text":
                    cell(s, "FillPattern", 0)
                    cell(s, "LinePattern", 0)
                s.NameU = op["key"]
                s.Text = op.get("text", "")
                style(s, op, doc)
                lookup[op["key"]] = s
            elif kind == "connect":
                start, end = resolve(op["from"]), resolve(op["to"])
                s = page.Drop(app.ConnectorToolDataObject, 0, 0)
                s.NameU = op["key"]
                s.CellsU("BeginX").GlueTo(start.CellsU("PinX"))
                s.CellsU("EndX").GlueTo(end.CellsU("PinX"))
                cell(s, "ShapeRouteStyle", 1 if op.get("route", "right_angle") == "right_angle" else 2)
                if op.get('route')=='curved':cell(s,'ConLineRouteExt',2)
                cell(s, "EndArrow", 4)
                s.Text = op.get("text", "")
                style(s, op, doc)
                lookup[op["key"]] = s
            elif kind in {'container','members','image','badge','group'}:
                s=visio_extras.operation(app,doc,page,op,resolve,set_cells)
                style(s,op,doc)
                if 'key' in op:lookup[op['key']]=s
            else:
                s = resolve(op["target"])
                if kind == "delete":
                    key,shape_id = s.NameU,s.ID
                    s.Delete()
                    lookup.pop(key, None)
                    by_id.pop(shape_id,None)
                else:
                    moved_container=False
                    if ('x' in op or 'y' in op) and s.CellExistsU('User.msvStructureType',0) and s.CellsU('User.msvStructureType').ResultStr('')=='Container':
                        visio_extras.move_with_members(page,s,op.get('x',s.CellsU('PinX').ResultIU),op.get('y',s.CellsU('PinY').ResultIU),set_cells)
                        moved_container=True
                    for key, name in (("x", "PinX"), ("y", "PinY"), ("width", "Width"), ("height", "Height")):
                        if key in op:
                            if moved_container and key in {'x','y'}:continue
                            cell(s, name, f"{op[key]} in")
                    if "text" in op:
                        s.Text = op["text"]
                    if "route" in op and s.OneD:
                        cell(s, "ShapeRouteStyle", 1 if op["route"] == "right_angle" else 2)
                        cell(s,'ConLineRouteExt',2 if op['route']=='curved' else 0)
                    style(s, op, doc)
            if kind!='delete':by_id[s.ID]=s
    except Exception as exc:
        try:
            app.EndUndoScope(scope, False)
        except Exception as rollback:
            raise RuntimeError(f"Batch failed: {exc}; rollback failed: {rollback}") from exc
        raise RequestError(f"Batch failed at {op.get('key', op.get('target', '?'))}; drawing rollback requested successfully: {exc}") from exc
    else:
        app.EndUndoScope(scope, True)


def show_document(app,doc,page):
    for window in app.Windows:
        if window.Type==1 and window.Document.ID==doc.ID:
            window.Visible=True;window.Page=page;window.Activate();return
    page.OpenDrawWindow()  # Only if there is no existing window to reuse.


def execute(app,req):
    """Own only documents opened by this request; preserve pre-existing user work."""
    before={d.ID for d in drawings(app)}
    visible = req.get('show',False) or (req['action']=='pages' and req.get('mode')=='activate')
    result=None
    try:
        result=_execute(app,req)
        if req.get('show') and req.get('document'):
            doc=find_document(app,path(req.get('save_as',req['document'])))
            if doc is not None:
                page=doc.Pages.ItemU(req['page']) if req.get('page') else doc.Pages.Item(1)
                show_document(app,doc,page)
        return result
    finally:
        closed=[];kept=[]
        for doc in list(drawings(app)):
            if doc.ID in before:continue
            saved=bool(doc.Saved)
            empty_draft=(result is None and not saved and not doc.Path and
                         all(page.Shapes.Count==0 for page in doc.Pages))
            close=(result is None and (saved or empty_draft)) or (
                result is not None and not visible and req.get('close_after',True) and saved)
            if close:
                name=doc.FullName
                if not saved:doc.Saved=True  # Only an empty, request-owned failed drawing.
                doc.Close();closed.append(name)
            else:kept.append(doc.FullName)
        if result is not None:result['lifecycle']={'closed_owned':closed,'kept_owned':kept,'existing_documents_preserved':True}


def _execute(app, req):
    action = req["action"]
    stages={};started=time.perf_counter()
    if action=='icons':return visio_extras.icons()
    if action=='open_stencil':
        stencil=app.Documents.OpenEx(req['stencil'],64)
        return {'stencil':stencil.FullName,'master_count':stencil.Masters.Count}
    if action == "status":
        return {"version": app.Version, "transport": "local COM",'my_shapes_path':app.MyShapesPath,
                'builtin_stencils_path':str(Path(app.Path)/'Visio Content'),
                'open_stencils':[d.FullName for d in app.Documents if d.Type==2],"documents": [
            {"name": d.Name, "path": d.FullName, "pages": [p.NameU for p in d.Pages], "saved": bool(d.Saved)} for d in drawings(app)],
                'visible_drawing_windows':sum(1 for w in app.Windows if w.Type==1 and w.Visible)}
    if action == "stencils":
        roots = [Path(app.Path)/"Visio Content", Path(app.MyShapesPath)]
        return {"roots": [str(p) for p in roots], "files": [str(f) for root in roots if root.exists()
                for f in root.rglob("*") if f.suffix.lower() in {".vssx", ".vss", ".vssm"}]}
    if action == "masters":
        stencil = app.Documents.OpenEx(req["stencil"], 64)
        query = req.get("query", "").lower()
        return {"stencil": stencil.FullName, "masters": [{"name": m.Name, "name_u": m.NameU}
                for m in stencil.Masters if query in m.Name.lower() or query in m.NameU.lower()]}
    if action == 'rebuild':
        output = path(req['document'])
        if find_document(app, output):
            raise RequestError('Rebuild destination is open')
        source = path(req['source'])
        if req.get('scene_output'):
            model = vsdx_scene.extract(source)
            path(req['scene_output'], {'.json'}).write_text(json.dumps(model, ensure_ascii=False), encoding='utf-8')
        doc = app.Documents.OpenEx(str(source), BACKGROUND_DRAWING_FLAGS | OPEN_AS_COPY)
        page = select_page(doc, req)
    elif action == "render":
        if find_document(app, path(req["document"])):
            raise RequestError("Render destination already open; use apply to edit")
        doc = app.Documents.AddEx(req.get("template", ""),0,BACKGROUND_DRAWING_FLAGS)
        page = doc.Pages.Item(1)
        page.NameU = req.get("page", "Page-1")
        cell(page.PageSheet, "PageWidth", f"{req.get('width', 12)} in")
        cell(page.PageSheet, "PageHeight", f"{req.get('height', 8)} in")
    else:
        doc, page = target(app, req)
    stages['target_total']=time.perf_counter()-started
    if action=='pages':
        mode=req.get('mode','list')
        if mode=='add':
            scope=app.BeginUndoScope('visio-skill add page')
            try:
                if req['name'] in [p.NameU for p in doc.Pages]:raise RequestError('Page name already exists')
                page=doc.Pages.Add();page.NameU=req['name']
            except Exception:
                app.EndUndoScope(scope,False);raise
            else:app.EndUndoScope(scope,True)
            doc.Save()
        elif mode=='activate':
            page=doc.Pages.ItemU(req['name'])
            show_document(app,doc,page)
        elif mode!='list':raise RequestError('pages mode must be list, add or activate')
        return {'pages':[p.NameU for p in doc.Pages],'selected':page.NameU}
    if action=='page_size':
        if req.get('fit_to_contents'):page.ResizeToFitContents()
        else:
            for key,name in [('width','PageWidth'),('height','PageHeight')]:
                if key in req:cell(page.PageSheet,name,f'{req[key]} in')
        if req.get('orientation'):
            cell(page.PageSheet,'PrintPageOrientation',2 if req['orientation']=='landscape' else 1)
    t=time.perf_counter()
    if action in {"render", "apply"}:
        with performance_scope(app, req.get('optimize', True)):
            batch(app, doc, page, req.get("operations", []))
    stages['draw']=time.perf_counter()-t
    if action == "layout":
        scope = app.BeginUndoScope("visio-skill layout")
        try:
            visio_extras.layout(page,req.get('style','default'),req.get('spacing',0.5),req.get('resize_page',True))
        except Exception:
            app.EndUndoScope(scope, False)
            raise
        else:
            app.EndUndoScope(scope, True)
    t=time.perf_counter()
    if action in {"render", "rebuild"} or req.get("save_as"):
        output = path(req.get("save_as", req["document"]), {".vsdx"})
        output.parent.mkdir(parents=True, exist_ok=True)
        doc.SaveAs(str(output))
    elif action in {"save", "layout"} or req.get("save") or (action=='page_size' and req.get('save',True)):
        doc.Save()
    stages['save']=time.perf_counter()-t;t=time.perf_counter()
    exported = []
    for value in req.get("exports", []):
        output = path(value, EXPORT_EXTENSIONS)
        output.parent.mkdir(parents=True, exist_ok=True)
        with png_resolution(app,req.get('png_dpi') if output.suffix.lower()=='.png' else None):
            page.Export(str(output))
        if not output.is_file() or output.stat().st_size == 0:
            raise RuntimeError(f"Export did not produce a nonempty file: {output}")
        exported.append(str(output))
    stages['export']=time.perf_counter()-t;t=time.perf_counter()
    state=inspect(doc,page,req.get('detail','full'),req.get('recursive',False))
    stages['inspect']=time.perf_counter()-t
    return {"state":state, "exports": exported,'timings':{k:round(v,6) for k,v in stages.items()}, "saved": bool(
        action in {"render", "save", "layout",'rebuild'} or (action=='page_size' and req.get('save',True)) or req.get("save_as") or req.get("save"))}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--request", type=Path, help="UTF-8 JSON request file")
    source.add_argument("--action", choices=["status", "stencils",'icons'])
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")
    started = time.perf_counter()
    snapshot = False
    try:
        req = json.loads(args.request.read_text(encoding="utf-8-sig")) if args.request else {"action": args.action}
        snapshot = req.get('action') == 'snapshot'
        if snapshot:
            preflight(req)
            result = vsdx_scene.snapshot(path(req['document']), req.get('recursive', True))
            print(json.dumps({'ok': True, 'result': result}, ensure_ascii=False))
            return 0
        with DesktopLock():
            preflight(req)
            import pythoncom

            pythoncom.CoInitialize()
            try:
                app, connection = connect()
                result = execute(app, req)
            finally:
                pythoncom.CoUninitialize()
        print(json.dumps({"ok": True, "connection": connection, "elapsed_seconds": round(time.perf_counter()-started, 3),
                          "result": result}, ensure_ascii=False))
        return 0
    except Exception as exc:
        failure = {"ok": False, "error_type": type(exc).__name__, "error": str(exc)}
        if snapshot:
            failure['quality_verdict'] = 'not_assessed'
        else:
            failure['elapsed_seconds'] = round(time.perf_counter() - started, 3)
        print(json.dumps(failure, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
