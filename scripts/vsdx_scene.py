"""Native geometry scene extraction/reconstruction for self-contained VSDX pages.

No COM needed here. This is structural reconstruction, not visual tracing.
Reject master/media-dependent drawings instead of silently dropping their assets.
"""
import hashlib
from pathlib import Path
import re
import tempfile
import zipfile
import xml.etree.ElementTree as E

V = 'http://schemas.microsoft.com/office/visio/2012/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'v': V}
E.register_namespace('', V)
E.register_namespace('r', R)


def xml(root):
    return E.tostring(root, encoding='utf-8', xml_declaration=True)


def snapshot(filename, recursive=True):
    with zipfile.ZipFile(filename) as z:
        pages = E.fromstring(z.read('visio/pages/pages.xml'))
        rels = E.fromstring(z.read('visio/pages/_rels/pages.xml.rels'))
        targets = {r.get('Id'):r.get('Target') for r in rels}
        result = []
        for page in pages:
            rel = page.find('v:Rel', NS)
            name = targets[rel.get('{'+R+'}id')]
            part = 'visio/pages/' + name
            root = E.fromstring(z.read(part))
            shapes = root.findall('.//v:Shape', NS) if recursive else root.findall('v:Shapes/v:Shape', NS)
            items = []
            for s in shapes:
                cells = {c.get('N'):c.get('V') for c in s.findall('v:Cell', NS)}
                t = s.find('v:Text', NS)
                items.append({'id':int(s.get('ID')), 'key':s.get('NameU'), 'type':s.get('Type'),
                              'text': ''.join(t.itertext()) if t is not None else '', 'cells':cells})
            sheet = page.find('v:PageSheet', NS)
            result.append({'name':page.get('NameU'), 'part':part, 'shape_count':len(items),
                           'top_level_count':len(root.findall('v:Shapes/v:Shape', NS)),
                           'page_cells':{c.get('N'):c.get('V') for c in sheet.findall('v:Cell', NS)},
                           'shapes':items})
        return {'document':str(Path(filename).resolve()), 'state_source':'saved VSDX, not live unsaved state',
                'recursive':recursive, 'pages':result}


def extract(filename):
    with zipfile.ZipFile(filename) as z:
        if any(n.startswith(('visio/media/', 'visio/masters/')) for n in z.namelist()):
            raise ValueError('Scene reconstruction currently requires standalone geometry without masters/media')
        summary = snapshot(filename)
        document = E.fromstring(z.read('visio/document.xml'))
        # Copy only formatting resources; do not carry source window state or package.
        support = [E.tostring(e,encoding='unicode') for e in document if e.tag.rsplit('}',1)[-1] in
                   {'DocumentSettings','Colors','FaceNames','StyleSheets','DocumentSheet'}]
        pages_root = E.fromstring(z.read('visio/pages/pages.xml'))
        model = {'format':'visio-native-scene-v1', 'source_sha256':hashlib.sha256(Path(filename).read_bytes()).hexdigest(),
                 'support':support, 'pages':[]}
        for entry, page in zip(summary['pages'], pages_root):
            content = E.fromstring(z.read(entry['part']))
            if any('Master' in s.attrib for s in content.findall('.//v:Shape',NS)):
                raise ValueError('Unresolved master references')
            if content.findall('.//v:ForeignData',NS):
                raise ValueError('Foreign data cannot be reconstructed as native geometry')
            model['pages'].append({'name':entry['name'], 'sheet':E.tostring(page.find('v:PageSheet',NS),encoding='unicode'),
                                   'content':E.tostring(content,encoding='unicode')})
        return model


def reconstruct(model, destination, overwrite=False, *, readback=True):
    if model.get('format') != 'visio-native-scene-v1':
        raise ValueError('Unsupported scene version')
    destination=Path(destination)
    if destination.exists() and not overwrite:
        raise FileExistsError(destination)
    document=E.Element('{'+V+'}VisioDocument')
    for item in model['support']:
        document.append(E.fromstring(item))
    pages=E.Element('{'+V+'}Pages')
    relationships=E.Element('Relationships',xmlns=P)
    parts={}
    for index, item in enumerate(model['pages'],1):
        page=E.SubElement(pages,'{'+V+'}Page',ID=str(index-1),NameU=item['name'],Name=item['name'],
                          IsCustomName='1',IsCustomNameU='1')
        page.append(E.fromstring(item['sheet']))
        E.SubElement(page,'{'+V+'}Rel',{'{'+R+'}id':f'rId{index}'})
        E.SubElement(relationships,'Relationship',Id=f'rId{index}',
                     Type='http://schemas.microsoft.com/visio/2010/relationships/page',Target=f'page{index}.xml')
        content=E.fromstring(item['content'])
        # Re-key every object and update formula/connection references. This creates
        # a new native shape tree rather than copying the source ZIP or a picture.
        mapping={s.get('ID'):str(10000+i) for i,s in enumerate(content.findall('.//v:Shape',NS),1)}
        for node in content.iter():
            if node.tag=='{'+V+'}Shape':
                node.set('ID',mapping[node.get('ID')])
            for key in ('FromSheet','ToSheet'):
                if key in node.attrib:
                    node.set(key,mapping.get(node.get(key),node.get(key)))
            if 'F' in node.attrib:
                node.set('F',re.sub(r'Sheet\.(\d+)(?=!)',lambda m:'Sheet.'+mapping.get(m[1],m[1]),node.get('F')))
        parts[f'visio/pages/page{index}.xml']=xml(content)
    parts['visio/document.xml']=xml(document)
    parts['visio/pages/pages.xml']=xml(pages)
    parts['visio/pages/_rels/pages.xml.rels']=xml(relationships)
    parts['_rels/.rels']=f'''<?xml version="1.0"?><Relationships xmlns="{P}"><Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/document" Target="visio/document.xml"/></Relationships>'''.encode()
    parts['visio/_rels/document.xml.rels']=f'''<?xml version="1.0"?><Relationships xmlns="{P}"><Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/pages" Target="pages/pages.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/visio/2010/relationships/windows" Target="windows.xml"/></Relationships>'''.encode()
    parts['visio/windows.xml']=f'''<?xml version="1.0"?><Windows xmlns="{V}" ClientWidth="1600" ClientHeight="1000"><Window ID="0" WindowType="Drawing" WindowState="268435456" ContainerType="Page" Page="0" ViewScale="-1"><ShowGrid>0</ShowGrid><ShowPageBreaks>0</ShowPageBreaks></Window></Windows>'''.encode()
    types=E.Element('Types',xmlns='http://schemas.openxmlformats.org/package/2006/content-types')
    E.SubElement(types,'Default',Extension='rels',ContentType='application/vnd.openxmlformats-package.relationships+xml')
    E.SubElement(types,'Default',Extension='xml',ContentType='application/xml')
    for part in parts:
        kind='page' if re.search(r'/page\d+\.xml$',part) else 'pages' if part.endswith('/pages.xml') else 'drawing.main' if part=='visio/document.xml' else 'windows' if part=='visio/windows.xml' else None
        if kind:
            E.SubElement(types,'Override',PartName='/'+part,ContentType='application/vnd.ms-visio.'+kind+'+xml')
    parts['[Content_Types].xml']=xml(types)
    destination.parent.mkdir(parents=True,exist_ok=True)
    # Write a complete temporary package before replacing the destination.
    with tempfile.NamedTemporaryFile(dir=destination.parent,suffix='.vsdx',delete=False) as f:
        temporary=Path(f.name)
    try:
        with zipfile.ZipFile(temporary,'w',zipfile.ZIP_DEFLATED) as z:
            for name,data in parts.items():
                z.writestr(name,data)
        if destination.exists() and not overwrite:
            raise FileExistsError(destination)
        temporary.replace(destination)
    finally:
        temporary.unlink(missing_ok=True)
    if readback:
        return snapshot(destination)
