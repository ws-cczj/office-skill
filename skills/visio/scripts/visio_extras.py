"""Advanced desktop operations shared by the JSON bridge."""
from pathlib import Path
import math
import os
import tempfile
import urllib.request
from urllib.parse import urlparse


def icons():
    folder=Path(os.environ.get('VISIO_SKILL_ICONS_DIR',str(Path.home()/'.visio-skill'/'icons')))
    return {'folder':str(folder),'icons':[{'label':p.stem,'path':str(p)} for p in sorted(folder.glob('*'))
            if p.suffix.lower() in {'.svg','.png','.jpg','.jpeg','.emf','.wmf'}]}


def selection(page, members):
    selected=page.CreateSelection(0)
    for shape in members:
        selected.Select(shape,2)
    return selected


def move_with_members(page,shape,x,y,write):
    """COM PinX/PinY writes do not move container members automatically."""
    dx=x-shape.CellsU('PinX').ResultIU;dy=y-shape.CellsU('PinY').ResultIU
    pending=[shape];seen=set();positions=[];containers=[]
    while pending:
        s=pending.pop()
        if s.ID in seen:continue
        seen.add(s.ID)
        positions.append((s,s.CellsU('PinX').ResultIU+dx,s.CellsU('PinY').ResultIU+dy))
        if s.CellExistsU('User.msvStructureType',0) and s.CellsU('User.msvStructureType').ResultStr('')=='Container':
            cp=s.ContainerProperties;containers.append((cp,cp.ResizeAsNeeded));cp.ResizeAsNeeded=0
            pending.extend(page.Shapes.ItemFromID(i) for i in cp.GetMemberShapes(0))
    try:
        for s,nx,ny in reversed(positions):write(s,{'PinX':f'{nx} in','PinY':f'{ny} in'})
    finally:
        for cp,value in containers:cp.ResizeAsNeeded=value


def image(page, op, write):
    source=op['source']
    temporary=None
    try:
        if source.startswith(('https://','http://')):
            suffix=Path(urlparse(source).path).suffix
            with tempfile.NamedTemporaryFile(suffix=suffix or '.png',delete=False) as f:
                temporary=Path(f.name)
            urllib.request.urlretrieve(source,str(temporary))
            source=str(temporary)
        p=Path(source)
        if not p.is_absolute() or not p.is_file():
            raise ValueError('Image source must be an existing absolute local path or HTTP(S) URL')
        s=page.Import(str(p))
        data={'PinX':f"{op['x']} in",'PinY':f"{op['y']} in"}
        if 'width' in op:data['Width']=f"{op['width']} in"
        if 'height' in op:data['Height']=f"{op['height']} in"
        write(s,data)
        return s
    finally:
        if temporary:temporary.unlink(missing_ok=True)


def operation(app,doc,page,op,resolve,write):
    kind=op['op']
    if kind=='image':
        s=image(page,op,write)
    elif kind=='container':
        members=[resolve(ref) for ref in op['members']]
        if not members:raise ValueError('Container requires members')
        stencil=app.Documents.OpenEx(app.GetBuiltInStencilFile(2,0),64)
        master=stencil.Masters.ItemU(op['master']) if op.get('master') else stencil.Masters.Item(1)
        s=page.DropContainer(master,selection(page,members))
        s.ContainerProperties.SetMargin(65,op.get('padding',0.25))
        s.ContainerProperties.FitToContents()
        s.Text=op.get('text','')
    elif kind=='members':
        s=resolve(op['target'])
        for ref in op['members']:
            member=resolve(ref)
            if op.get('mode','add')=='remove':s.ContainerProperties.RemoveMember(member)
            else:s.ContainerProperties.AddMember(member,1)
        return s
    elif kind=='group':
        s=selection(page,[resolve(ref) for ref in op['members']]).Group()
    elif kind=='badge':
        zone=resolve(op['target'])
        x,y,w,h=[zone.CellsU(n).ResultIU for n in ('PinX','PinY','Width','Height')]
        size=op.get('size',0.35);corner=op.get('corner','top_left')
        bx=x+(-1 if corner.endswith('left') else 1)*(w/2-size)
        by=y+(1 if corner.startswith('top') else -1)*(h/2-size)
        if 'source' in op:
            s=image(page,dict(op,x=bx,y=by,width=size,height=size),write)
        else:
            stencil=app.Documents.OpenEx(op['stencil'],64)
            s=page.Drop(stencil.Masters.ItemU(op['master']),bx,by)
            write(s,{'Width':f'{size} in','Height':f'{size} in'})
        zone.ContainerProperties.AddMember(s,0)
    else:
        raise ValueError(f'Unknown advanced operation: {kind}')
    s.NameU=op['key']
    return s


def layout(page,style,spacing=0.5,resize=True):
    """Apply Visio presets; place radial/circular diagrams deterministically."""
    if style in {'radial','circular'}:
        nodes=[s for s in page.Shapes if not s.OneD]
        radius=max(2,len(nodes)*0.4)
        for i,s in enumerate(nodes):
            if style=='radial' and i==0:
                x=y=radius+2
            else:
                angle=2*math.pi*(i-(1 if style=='radial' else 0))/max(1,len(nodes)-(1 if style=='radial' else 0))
                x,y=radius+2+radius*math.cos(angle),radius+2+radius*math.sin(angle)
            s.CellsU('PinX').FormulaU=f'{x} in';s.CellsU('PinY').FormulaU=f'{y} in'
    else:
        routes={'flowchart_tb':5,'flowchart_lr':6,'tree_tb':7,'tree_lr':8,'default':0}
        page.PageSheet.CellsU('RouteStyle').FormulaU=str(routes[style])
        placements={'flowchart_tb':1,'flowchart_lr':2,'tree_tb':17,'tree_lr':23,'default':0}
        page.PageSheet.CellsU('PlaceStyle').FormulaU=str(placements[style])
        page.PageSheet.CellsU('AvenueSizeX').FormulaU=f'{spacing} in'
        page.PageSheet.CellsU('AvenueSizeY').FormulaU=f'{spacing} in'
        page.Layout()
    if resize:page.ResizeToFitContents()
