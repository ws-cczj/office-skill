"""Read saved DOCX equation, numbering, and reference structures without Word."""
import argparse
from collections import Counter
import json
from pathlib import Path
import re
import sys
from xml.etree import ElementTree as ET
from zipfile import ZipFile


W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
O = '{urn:schemas-microsoft-com:office:office}'
REFERENCE = re.compile(r'^(?:REF|PAGEREF)\s+"?([^\s"]+)', re.I)
SEQUENCE = re.compile(r'^SEQ\s+(\S+)', re.I)


def fields(document):
    stack = []
    for element in document.iter():
        tag = element.tag
        if tag == W + 'fldSimple':
            yield element.get(W + 'instr', '').strip(), ''.join(
                item.text or '' for item in element.iter(W + 't')
            )
        elif tag == W + 'fldChar':
            kind = element.get(W + 'fldCharType')
            if kind == 'begin':
                stack.append({'code': [], 'result': [], 'separated': False})
            elif kind == 'separate' and stack:
                stack[-1]['separated'] = True
            elif kind == 'end' and stack:
                data = stack.pop()
                yield ''.join(data['code']).strip(), ''.join(data['result'])
        elif stack and tag == W + 'instrText' and not stack[-1]['separated']:
            stack[-1]['code'].append(element.text or '')
        elif stack and tag == W + 't':
            for data in stack:
                if data['separated']:
                    data['result'].append(element.text or '')


def check(path):
    with ZipFile(path) as archive:
        document = ET.fromstring(archive.read('word/document.xml'))
    bookmarks = {item.get(W + 'name', '').casefold() for item in document.iter(W + 'bookmarkStart')}
    scripts = Counter(item.get(W + 'val') for item in document.iter(W + 'vertAlign'))
    sequences, references = [], []
    for code, result in fields(document):
        if match := REFERENCE.match(code):
            target = match[1]
            references.append({'target': target, 'exists': target.casefold() in bookmarks, 'cached': result})
        elif match := SEQUENCE.match(code):
            sequences.append({'sequence': match[1], 'cached': result})
    displays = []
    for index, paragraph in enumerate(document.iter(W + 'p'), 1):
        ids = [item.get(W + 'name', '') for item in paragraph.iter(W + 'bookmarkStart')]
        equations = [name for name in ids if name.startswith('office_equation_')]
        if equations:
            tabs = [{'alignment': tab.get(W + 'val'), 'position_twips': int(tab.get(W + 'pos', '0'))}
                    for tab in paragraph.findall(f'{W}pPr/{W}tabs/{W}tab')]
            displays.append({'paragraph': index, 'equations': equations, 'tabs': tabs})
    return {
        'scope': 'Saved main document XML and cached field results',
        'word_equations': sum(1 for _ in document.iter(M + 'oMath')),
        'mathtype_equations': sum(
            1 for item in document.iter(O + 'OLEObject')
            if item.get('ProgID', '').casefold() == 'equation.dsmt4'
        ),
        'subscript_runs': scripts['subscript'],
        'superscript_runs': scripts['superscript'],
        'sequences': sequences, 'references': references, 'equation_layouts': displays,
        'quality_verdict': 'not_assessed',
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('docx', type=Path)
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding='utf-8')
    report = check(args.docx)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return int(any(not ref['exists'] for ref in report['references']))


if __name__ == '__main__':
    raise SystemExit(main())
