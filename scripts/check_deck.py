"""Read-only checks for delivered PPTX files. Visual review remains separate."""

import argparse
from dataclasses import dataclass
from hashlib import sha256
import json
from pathlib import Path, PurePosixPath
import posixpath
import re
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile


NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
REL_ID = f'{{{NS["r"]}}}id'
IMAGE_LABEL = re.compile(
    r'^\[(?:图表|图片|图表\s*/\s*图片|研究示意图|核心研究示意图|示意图|高图框|'
    r'结果图\s*\d+|图\s*\d+|共用证据\s*/\s*图表)\]$'
)
MONTH = re.compile(r'(?<!\d)(20\d{2})\s*[./年–-]\s*(0?[1-9]|1[0-2])(?!\d)')
SCRIPT_HEADING = re.compile(r'^##\s+(?:第\s*)?(\d+)(?:页)?\s*[｜|]\s*.+$', re.M)
INTERNAL_COPY = re.compile(
    r'^(?:依据|来源|出处|引用来源)\s*[：:]\s*(?:据\s*)?'
    r'(?:(?:原始|本次|用户提供的)\s*)?(?:(?:开题\s*)?报告|Word)(?![a-z])'
    r'|^(?:原始|本次|开题)?\s*报告\s*第?\s*(?:[§图表]\s*\d|\d+(?:\.\d+)+)'
    r'|^(?:制作说明|绘制说明|素材路径|验收记录|导图检查|执行\s*Prompt)\s*[：:]'
    r'|(?<![\w])(?:[A-Z]:[\\/]|/(?:home|mnt/data)/)[^\s]+',
    re.I | re.M,
)


def paragraphs(element):
    return '\n'.join(''.join(t.text or '' for t in p.findall('.//a:t', NS))
                     for p in element.findall('.//a:p', NS))


def relationships(package, part):
    path = PurePosixPath(part)
    rel = str(path.parent / '_rels' / (path.name + '.rels'))
    if rel not in package.namelist():
        return {}
    return {
        r.attrib['Id']: (r.attrib['Type'].rsplit('/', 1)[-1],
                         posixpath.normpath(posixpath.join(str(path.parent), r.attrib['Target'])).lstrip('/'))
        for r in ET.fromstring(package.read(rel)) if r.attrib.get('TargetMode') != 'External'
    }


@dataclass(frozen=True)
class Shape:
    name: str
    kind: str
    box: tuple | None
    fill: str | None
    text: str


def shapes(slide):
    for obj in slide.find('p:cSld/p:spTree', NS):
        props = obj.find('./*/p:cNvPr', NS)
        if props is None or props.get('hidden') in {'true', '1'}:
            continue
        xfrm = next((x for path in ('p:spPr/a:xfrm', 'p:grpSpPr/a:xfrm', 'p:xfrm')
                     if (x := obj.find(path, NS)) is not None), None)
        box = None
        if xfrm is not None:
            offset, extent = xfrm.find('a:off', NS), xfrm.find('a:ext', NS)
            if offset is not None and extent is not None:
                box = tuple(int(num) for num in (offset.get('x'), offset.get('y'), extent.get('cx'), extent.get('cy')))
        fill = obj.find('p:spPr/a:solidFill/a:srgbClr', NS)
        yield Shape(props.get('name', ''), obj.tag.rsplit('}', 1)[-1],
                    box, fill.get('val') if fill is not None else None, paragraphs(obj))


def read_deck(path):
    slides = []
    part = 'ppt/presentation.xml'
    with ZipFile(path) as package:
        slide_ids = ET.fromstring(package.read(part)).findall('./p:sldIdLst/p:sldId', NS)
        if not slide_ids:
            raise ValueError('PPTX 没有幻灯片')
        rels = relationships(package, part)
        for slide_id in slide_ids:
            role, slide_path = rels[slide_id.attrib[REL_ID]]
            if role != 'slide':
                raise ValueError('演示文稿的页面关系无效')
            slide = ET.fromstring(package.read(slide_path))
            notes = ''
            for kind, target in relationships(package, slide_path).values():
                if kind == 'notesSlide':
                    node = ET.fromstring(package.read(target))
                    notes = '\n'.join(paragraphs(s) for s in node.findall('./p:cSld/p:spTree/p:sp', NS)
                                      if s.find('./p:nvSpPr/p:nvPr/p:ph[@type="body"]', NS) is not None)
                    break
            slides.append((list(shapes(slide)), notes))
    return slides


def internal_copy_findings(slides):
    for page, (objects, notes) in enumerate(slides, 1):
        for shape in objects:
            if INTERNAL_COPY.search(shape.text.strip()):
                yield {'code': 'internal_copy_on_slide', 'page': page, 'object': shape.name,
                       'message': '正文对象含报告定位、制作说明或本机路径，核对其是否应留在工作稿。'}
        if INTERNAL_COPY.search(notes.strip()):
            yield {'code': 'internal_copy_in_notes', 'page': page,
                   'message': '备注含报告定位、制作说明或本机路径，核对其是否属于现场口述。'}


def inside(outer, inner):
    x, y, w, h = outer
    ix, iy, iw, ih = inner
    return x <= ix and y <= iy and ix + iw <= x + w and iy + ih <= y + h


def overlap_fraction(outer, inner):
    x, y, w, h = outer
    ix, iy, iw, ih = inner
    width = max(0, min(x+w, ix+iw) - max(x, ix))
    height = max(0, min(y+h, iy+ih) - max(y, iy))
    return width * height / (w * h)


def signature(shape):
    return shape.name, shape.kind, shape.box, shape.fill


def original_slots(template):
    result = set()
    for objects, _ in template:
        for label in objects:
            if not label.box or not IMAGE_LABEL.fullmatch(label.text.strip()):
                continue
            frames = [s for s in objects if s.kind == 'sp' and s.fill and s.box and not s.text.strip()
                      and s.box[2] >= 45 * 12700 and s.box[3] >= 40 * 12700 and inside(s.box, label.box)]
            if frames:
                result.add(signature(min(frames, key=lambda s: s.box[2] * s.box[3])))
    return result


def empty_figure_slots(slides, template):
    slots = original_slots(template)
    baseline = {signature(s) for objects, _ in template for s in objects}
    for page, (objects, _) in enumerate(slides, 1):
        for i, frame in enumerate(objects):
            if signature(frame) not in slots:
                continue
            # Inspect items in front of the retained source frame. A native diagram
            # may have no media, so accept new primitives/groups as visual candidates.
            foreground = [s for s in objects[i+1:] if s.box and signature(s) not in baseline]
            image = any(s.kind in {'pic', 'grpSp', 'graphicFrame'} and overlap_fraction(frame.box, s.box) >= .05
                        for s in foreground)
            primitives = [s for s in foreground if inside(frame.box, s.box)
                          and (s.kind == 'cxnSp' or (s.kind == 'sp' and not s.text.strip() and s.box != frame.box))]
            diagram = len(primitives) >= 2
            if not (image or diagram):
                yield {'code': 'retained_figure_slot', 'page': page, 'object': frame.name,
                       'message': '保留了源模板图位，前景没有检测到图片、图表或新绘示意。看实际导图，填图或移除图框后重排。'}


def timeline_end(value):
    parsed = re.fullmatch(r'([1-9]\d*):(20\d{2})-(0[1-9]|1[0-2])', value)
    if not parsed:
        raise argparse.ArgumentTypeError('时间轴参数应为 页码:YYYY-MM，例如 12:2028-04')
    return tuple(map(int, parsed.groups()))


def check_timeline(slides, expected):
    page, year, month = expected
    if page > len(slides):
        return {'code': 'timeline_page_missing', 'page': page, 'message': '时间轴页码不存在。'}
    found = max((tuple(map(int, match.groups())) for s in slides[page-1][0]
                 for match in MONTH.finditer(s.text)), default=None)
    if found != (year, month):
        return {'code': 'timeline_end_mismatch', 'page': page, 'expected': f'{year}-{month:02d}',
                'actual_latest_label': f'{found[0]}-{found[1]:02d}' if found else None,
                'message': '本页最晚的月份标签与需求不符。请核对截止月与条带位置。'}
    return None


def check_script(slides, path):
    if not path.is_file():
        yield {'code': 'speaker_script_missing', 'message': '独立发言稿文件不存在。'}
        return
    text = path.read_text(encoding='utf-8-sig')
    headings = list(SCRIPT_HEADING.finditer(text))
    if [int(m[1]) for m in headings] != list(range(1, len(slides)+1)):
        yield {'code': 'speaker_script_page_order', 'message': '按实际页序使用 ## 01｜标题，并为每页写一段口述。'}
        return
    for page, (match, (_, notes)) in enumerate(zip(headings, slides), 1):
        end = headings[page].start() if page < len(headings) else len(text)
        script = text[match.end():end].split()
        if not script or script != notes.split():
            yield {'code': 'speaker_notes_mismatch', 'page': page,
                   'message': '本页备注正文与独立发言稿不同或为空。'}


def page_selection(value):
    """Explicit content pages: omit covers, agendas and references."""
    pages = set()
    for item in value.split(','):
        match = re.fullmatch(r'([1-9]\d*)(?:-([1-9]\d*))?', item.strip())
        if not match:
            raise argparse.ArgumentTypeError('正文页应为 3-12 或 3,4,9-11')
        first, last = int(match[1]), int(match[2] or match[1])
        if last < first or last - first > 10000:
            raise argparse.ArgumentTypeError('正文页范围无效')
        pages.update(range(first, last+1))
    return sorted(pages)


def content_layout_review(slides, pages, size):
    """Report facts and repeated container geometry, never semantic visual quality.

    Header/footer objects are excluded by a documented 15%/95% body band.
    Native technical diagrams can also repeat; every finding requires visual review.
    """
    if pages and max(pages) > len(slides):
        raise ValueError('指定的正文页不存在')
    width, height = size
    profiles, families = [], {}
    for page in pages:
        objects = [s for s in slides[page-1][0] if s.box and
                   s.box[1] >= height*.15 and s.box[1]+s.box[3] <= height*.95]
        media = [s for s in objects if s.kind in {'pic', 'graphicFrame', 'grpSp'}]
        # Text is deliberately ignored: replacing all labels leaves this skeleton unchanged.
        containers = [s for s in objects if s.kind == 'sp' and not s.text.strip()
                      and s.fill and s.box[2]*s.box[3] >= width*height*.01]
        key = tuple(sorted(tuple(round(v/12700/2)*2 for v in s.box) + (s.fill,)
                           for s in containers))
        profiles.append({'page': page, 'body_media_or_groups': len(media),
                         'large_containers': len(containers)})
        if len(containers) >= 2 and not media:
            families.setdefault(key, []).append(page)
    findings = [{'code': 'repeated_text_container_layout', 'pages': group,
                 'message': '至少三页的大容器几何相同，且正文未检测到图片、图表/表格或分组。逐页看原生图形是否表达具体对象与关系；重复布局本身不等于错误。'}
                for group in families.values() if len(group) >= 3]
    return profiles, findings


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pptx', type=Path)
    parser.add_argument('--source-template', type=Path, help='复用的源模板。检查保留原位置和尺寸的示意图框。')
    parser.add_argument('--timeline-end', type=timeline_end, action='append', default=[])
    parser.add_argument('--speaker-script', type=Path)
    parser.add_argument('--content-pages', type=page_selection,
                        help='如 3-12。输出正文对象概况，并提示三页以上重复的无媒体大容器；需要人工看图裁定。')
    parser.add_argument('--output', type=Path, help='将同样的 JSON 写入 UTF-8 文件')
    args = parser.parse_args()
    report = {'file': str(args.pptx.resolve()), 'checks': [], 'issues': [],
              'quality_verdict': 'not_assessed',
              'quality_scope': '只核对已列文件条件与可选布局风险，不验收论证、技术图有效性、风格或整体美感。'}
    try:
        report['sha256'] = sha256(args.pptx.read_bytes()).hexdigest()
        slides = read_deck(args.pptx)
        report['slides'] = len(slides)
        report['checks'].extend(['pptx_page_structure', 'internal_copy_literals'])
        report['issues'].extend(internal_copy_findings(slides))
        if args.source_template:
            report['checks'].append('retained_figure_slots')
            report['issues'].extend(empty_figure_slots(slides, read_deck(args.source_template)))
        for end in args.timeline_end:
            report['checks'].append(f'timeline_end:{end[0]}')
            if issue := check_timeline(slides, end):
                report['issues'].append(issue)
        if args.speaker_script:
            report['checks'].append('speaker_script_and_notes')
            report['issues'].extend(check_script(slides, args.speaker_script))
        if args.content_pages is not None:
            with ZipFile(args.pptx) as package:
                size = ET.fromstring(package.read('ppt/presentation.xml')).find('p:sldSz', NS)
                if size is None:
                    raise ValueError('未找到画布尺寸，无法检查正文布局')
                dimensions = int(size.get('cx')), int(size.get('cy'))
                if min(dimensions) <= 0:
                    raise ValueError('画布尺寸无效')
            report['checks'].append('content_layout_review')
            profiles, findings = content_layout_review(slides, args.content_pages, dimensions)
            report['content_profiles'] = profiles
            report['issues'].extend(findings)
        report['status'] = 'needs_review' if report['issues'] else 'selected_checks_passed'
        code = 1 if report['issues'] else 0
    except (OSError, UnicodeError, BadZipFile, ET.ParseError, ValueError, KeyError, TypeError) as exc:
        report.update(status='input_error', detail=str(exc))
        code = 2
    payload = json.dumps(report, ensure_ascii=False, indent=2) + '\n'
    if args.output:
        args.output.write_text(payload, encoding='utf-8')
    print(payload, end='')
    return code


if __name__ == '__main__':
    raise SystemExit(main())
