"""Regression test for exact matrix writes and chart export in desktop Excel."""
from importlib.util import find_spec
import os
from pathlib import Path
from struct import unpack
from tempfile import TemporaryDirectory
import unittest

from core.office_batch import run_plan
from skills.excel.scripts import excel_native
from skills.excel.scripts.check_sheet import check


@unittest.skipUnless(os.name == 'nt' and find_spec('win32com'), 'Requires Windows desktop Excel and pywin32')
class ExcelComTests(unittest.TestCase):
    def test_clean_matrix_cross_sheet_formulas_and_chart(self):
        with TemporaryDirectory() as folder:
            root = Path(folder)
            target, chart = root / 'workbook.xlsx', root / 'chart.png'
            data = root / 'data.csv'
            data.write_text('Group,Value\n A ,2\nA,2\nB,4\nA,6\n', encoding='utf-8')
            result = run_plan(excel_native, {
                'schema_version': 1, 'target': str(target),
                'snapshot': {'ranges': [
                    {'sheet': 'Data', 'range': 'A1:B4'},
                    {'sheet': 'Summary', 'range': 'A1:C3'},
                ]},
                'exports': [{'path': str(chart), 'sheet': 'Summary', 'chart': 'Totals'}],
                'operations': [
                    {'op': 'sheet', 'name': 'Data', 'rename_from': 1},
                    {'op': 'import_csv', 'sheet': 'Data', 'source': str(data), 'numeric_columns': [2]},
                    {'op': 'clean', 'sheet': 'Data', 'range': 'A2:B5', 'trim': True, 'unique_by': [1, 2]},
                    {'op': 'table', 'sheet': 'Data', 'range': 'A1:B4', 'name': 'MeasurementData'},
                    {'op': 'sheet', 'name': 'Summary'},
                    {'op': 'write', 'sheet': 'Summary', 'rows': [
                        ['Group', 'Total', 'Text'], ['A', None, '=1+2'], ['B', None, '0012'],
                    ]},
                    {'op': 'formulas', 'sheet': 'Summary', 'start': 'B2', 'rows': [
                        ['=SUMIF(Data!A2:A4,A2,Data!B2:B4)'], ['=SUMIF(Data!A2:A4,A3,Data!B2:B4)'],
                    ]},
                    {'op': 'chart', 'sheet': 'Summary', 'name': 'Totals', 'range': 'A1:B3', 'width_pt': 720},
                ],
            }, root / 'plan.json')
            self.assertTrue(result['reopened'])
            self.assertEqual(result['state']['ranges'][0]['values'], [['Group', 'Value'], ['A', 2], ['B', 4], ['A', 6]])
            self.assertEqual(result['state']['ranges'][1]['values'][1:], [['A', 8, '=1+2'], ['B', 4, '0012']])
            saved, = check(target, ['Summary'])['worksheets']
            self.assertEqual([(item['cell'], item['cached']) for item in saved['formulas']], [('B2', '8'), ('B3', '4')])
            self.assertEqual(saved['errors'], [])
            image = chart.read_bytes()
            self.assertEqual(image[:8], b'\x89PNG\r\n\x1a\n')
            width, height = unpack('>II', image[16:24])
            self.assertGreaterEqual(width, 720)
            self.assertGreater(height, 0)


if __name__ == '__main__':
    unittest.main()
