"""Turn an explicit flow diagram blueprint into a native Visio render request."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
import unicodedata


FONT = 'Microsoft YaHei'
INK = '#162D43'
ACCENT = '#17665C'
MARGIN = .56
ROLES = {
    'step': dict(kind='rectangle', fill='#F0F4F8', line='#CCD9E4', text_color=INK, rounding=.12),
    'decision': dict(kind='diamond', fill='#FFF4D9', line='#CC9A3B', text_color='#5A391B'),
    'start': dict(kind='rectangle', fill='#E1F2EC', line='#83B4A7', text_color=INK, rounding=.18, bold=True),
    'end': dict(kind='rectangle', fill=ACCENT, no_line=True, text_color='#FFFFFF', rounding=.18, bold=True),
}


def text_width(text):
    return sum(.35 if c.isspace() else 1 if unicodedata.east_asian_width(c) in 'WF'
               else .56 for c in text)


def wrap(text, limit):
    """Prefer word breaks; permit character breaks in CJK."""
    lines = []
    for paragraph in text.strip().splitlines():
        line = ''
        for token in re.findall(r'[\x21-\x7E]+|\s+|[^\x00-\x7F]', paragraph):
            if token.isspace() and not line:
                continue
            if line and text_width(line + token) > limit and not token.isspace():
                lines.append(line.rstrip())
                line = ''
            line += token
        lines.append(line.rstrip())
    return lines


def node_shape(node):
    role = node.get('role', 'step')
    if role not in ROLES:
        raise ValueError(f"Unknown visual role: {role!r}")
    label = node['text']
    if not isinstance(label, str) or not label.strip():
        raise ValueError('Node text must be nonempty')
    decision = role == 'decision'
    lines = wrap(label, 8.5 if decision else 13)
    if len(lines) > 4:
        raise ValueError(f"Shorten the label for {node['key']!r} or choose a custom layout")
    text_in = max(text_width(s) for s in lines) * (13/72) * 1.16
    if decision:
        # Text occupies the narrower inner region of a diamond.
        width, height = max(2.52, text_in*1.7+.36), max(1.08, .61+.26*len(lines))
    else:
        width, height = max(2.02, text_in+.38), max(.68, .25+.23*len(lines))
    return dict(ROLES[role], op='add', key=node['key'], text='\n'.join(lines),
                width=width, height=height, font=FONT, font_size=13, text_margin=.09, align=1)


def cell_centers(sizes, gap):
    cursor = 0
    centers = {}
    for index,size in sorted(sizes.items()):
        centers[index] = cursor + size/2
        cursor += size + gap
    return centers, cursor - gap


def render_request(blueprint, document, preview=None):
    """Place ranked nodes and keep all arrows as glueable Visio connectors."""
    direction = blueprint.get('direction', 'lr')
    if direction not in {'lr', 'tb'}:
        raise ValueError('direction must be lr or tb')
    title = blueprint.get('title')
    if not isinstance(title,str) or not title.strip() or len(title.splitlines()) != 1 or text_width(title) > 48:
        raise ValueError('Provide a short single-line title')
    nodes = blueprint['nodes']
    if not nodes:
        raise ValueError('At least one node is required')
    keys,slots = {},set()
    columns,rows = {},{}
    for node in nodes:
        key,stage,track = node['key'],node['stage'],node.get('track',0)
        if (not isinstance(key,str) or not re.fullmatch(r'[A-Za-z][A-Za-z0-9_-]*',key)
                or key in keys or key.lower().startswith('flow_')):
            raise ValueError(f'Invalid or duplicate node key: {key!r}')
        if type(stage) is not int or stage < 0 or type(track) is not int:
            raise ValueError('stage must be a nonnegative integer and track an integer')
        col,row = (stage,track) if direction=='lr' else (track,stage)
        if (col,row) in slots:
            raise ValueError(f'More than one node in slot {(col,row)}')
        slots.add((col,row))
        shape = node_shape(node)
        keys[key] = col,row,shape
        columns[col] = max(columns.get(col,0),shape['width'])
        rows[row] = max(rows.get(row,0),shape['height'])

    xs, content_width = cell_centers(columns, .78)
    ys, content_height = cell_centers(rows, .75)
    width = max(content_width, text_width(title)*(20/72)*1.18+.18, 4.2) + 2*MARGIN
    height = content_height + 2*MARGIN + .85
    ops = [
        dict(op='add',key='Flow_Canvas',x=width/2,y=height/2,width=width,height=height,
             fill='#FFFFFF',no_line=True),
        dict(op='add',kind='text',key='Flow_Title',x=width/2,y=height-MARGIN-.15,
             width=width-2*MARGIN,height=.44,text=title,font=FONT,font_size=20,
             text_color=INK,bold=True,align=0,text_margin=0),
        dict(op='add',key='Flow_Rule',x=MARGIN+.28,y=height-MARGIN-.52,width=.56,height=.035,
             fill=ACCENT,no_line=True),
    ]
    left = (width - content_width)/2
    top = height-MARGIN-.85
    for col,row,shape in keys.values():
        ops.append(dict(shape,x=left+xs[col],y=top-ys[row]))

    col_order = {k:i for i,k in enumerate(sorted(columns))}
    row_order = {k:i for i,k in enumerate(sorted(rows))}
    for index,edge in enumerate(blueprint.get('edges',[]),1):
        start,end = edge['from'],edge['to']
        if start not in keys or end not in keys or start==end:
            raise ValueError(f'Invalid flow edge: {start!r} -> {end!r}')
        a,b,_ = keys[start]
        c,d,_ = keys[end]
        distance = abs(col_order[a]-col_order[c])+abs(row_order[b]-row_order[d])
        ops.append(dict(op='connect',key=f'Flow_Edge_{index}',text=edge.get('label',''),
                        font=FONT,font_size=10,text_color=INK,line='#55728A',line_weight=1.25,
                        route='straight' if distance==1 else 'right_angle',
                        **{'from':start,'to':end}))

    request = dict(action='render', document=str(Path(document).resolve()), width=width, height=height,
                   operations=ops, page='Diagram')
    if preview is not None:
        request.update(exports=[str(Path(preview).resolve())], png_dpi=300)
    return request


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--blueprint',type=Path,required=True)
    parser.add_argument('--document',type=Path,required=True)
    parser.add_argument('--preview',type=Path)
    args=parser.parse_args()
    data=json.loads(args.blueprint.read_text(encoding='utf-8-sig'))
    print(json.dumps(render_request(data,args.document,args.preview), ensure_ascii=True))


if __name__=='__main__':
    main()
