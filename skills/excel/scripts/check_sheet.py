"""Audit saved XLSX formula definitions and cached errors without Excel."""
import argparse
import json
from pathlib import Path
from posixpath import join, normpath
import sys
from xml.etree import ElementTree as ET
from zipfile import ZipFile


S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def check(path, names=()):
    worksheets = []
    with ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        rels = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        targets = {
            item.get('Id'): item.get('Target')
            for item in rels
            if item.get('Type', '').endswith('/worksheet') and item.get('TargetMode') != 'External'
        }
        for sheet in workbook.find(S + 'sheets'):
            name = sheet.get('name')
            target = targets.get(sheet.get(REL + 'id'))
            if not target or names and name not in names:
                continue
            part = target.lstrip('/') if target.startswith('/') else normpath(join('xl', target))
            root = ET.fromstring(archive.read(part))
            formulas, errors = [], []
            for cell in root.iter(S + 'c'):
                address = cell.get('r')
                cached = cell.findtext(S + 'v')
                formula = cell.find(S + 'f')
                if formula is not None:
                    formulas.append({'cell': address, 'formula': formula.text,
                                     'attributes': dict(sorted(formula.attrib.items())), 'cached': cached})
                if cell.get('t') == 'e':
                    errors.append({'cell': address, 'error': cached, 'has_formula': formula is not None})
            worksheets.append({'name': name, 'formulas': formulas, 'errors': errors})
    selected = {sheet['name'] for sheet in worksheets}
    if missing := set(names) - selected:
        raise ValueError('Unknown workbook worksheets: ' + ', '.join(sorted(missing)))
    return {'scope': 'Saved worksheet formulas and cached results; no calculation',
            'worksheets': worksheets, 'quality_verdict': 'not_assessed'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('xlsx', type=Path)
    parser.add_argument('--sheet', action='append', default=[])
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding='utf-8')
    report = check(args.xlsx, args.sheet)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return int(any(sheet['errors'] for sheet in report['worksheets']))


if __name__ == '__main__':
    raise SystemExit(main())
