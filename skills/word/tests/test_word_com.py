"""Regression tests for editable formulas and SEQ/REF placement in desktop Word."""
from importlib.util import find_spec
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from core.office_batch import office, run_plan
from skills.word.scripts import word_native
from skills.word.scripts.check_doc import check


@unittest.skipUnless(os.name == 'nt' and find_spec('win32com'), 'Requires Windows desktop Word and pywin32')
class WordComTests(unittest.TestCase):
    def test_numbered_equation_and_forward_references_survive_reopen(self):
        with TemporaryDirectory() as folder:
            root = Path(folder)
            source, target, pdf = root / 'source.docx', root / 'paper.docx', root / 'paper.pdf'
            with office(word_native) as app:
                doc = word_native.create(app, None)
                try:
                    doc.Content.Text = 'Existing final paragraph'
                    word_native.save(doc, source, False)
                finally:
                    word_native.close(doc)
            result = run_plan(word_native, {
                'schema_version': 1, 'target': str(target), 'source': str(source),
                'snapshot': {'text': True},
                'exports': [{'path': str(pdf), 'first_page': 1, 'last_page': 1}],
                'operations': [
                    {'op': 'configure', 'profile': 'sci', 'page': {'paper': 'a4', 'columns': 2}},
                    {'op': 'header_footer', 'kind': 'header', 'text': 'Manuscript'},
                    {'op': 'paragraph', 'style': 'title', 'text': 'Sample Manuscript'},
                    {'op': 'paragraph', 'runs': [
                        {'text': 'H'}, {'text': '2', 'script': 'sub'}, {'text': 'O, x'},
                        {'text': '2', 'script': 'super'}, {'text': ' and '}, {'math': 'a_1/b'},
                        {'text': ' follow '}, {'ref': 'source_a'},
                        {'text': '. See Eq. '}, {'ref': 'balance', 'kind': 'equation'},
                    ]},
                    {'op': 'equation', 'id': 'balance', 'text': 'x_1^2/y'},
                    {'op': 'table', 'rows': [['Quantity', 'Value'], ['A', 1.5]]},
                    {'op': 'caption', 'kind': 'table', 'id': 'conditions', 'text': 'Conditions.'},
                    {'op': 'reference', 'id': 'source_a', 'text': 'Author. Article. Journal.'},
                ],
            }, root / 'plan.json')
            self.assertTrue(result['reopened'])
            first_lines = result['state']['text'].split('\r', 2)[:2]
            self.assertEqual(first_lines, ['Existing final paragraph', 'Sample Manuscript'])
            self.assertEqual(pdf.read_bytes()[:5], b'%PDF-')
            self.assertEqual(result['state']['tables'], 1)
            report = check(target)
            self.assertEqual(report['word_equations'], 2)
            self.assertEqual(report['subscript_runs'], 1)
            self.assertEqual(report['superscript_runs'], 1)
            self.assertEqual([item['cached'] for item in report['references']], ['1', '1'])
            self.assertTrue(all(item['exists'] for item in report['references']))
            self.assertEqual(report['sequences'][0], {'sequence': 'Equation', 'cached': '1'})
            center, right = report['equation_layouts'][0]['tabs']
            self.assertEqual((center['alignment'], right['alignment']), ('center', 'right'))
            self.assertAlmostEqual(center['position_twips'] * 2, right['position_twips'], delta=1)

    def test_replace_mathtype_preserves_number_and_original(self):
        with office(word_native) as app:
            info = word_native.capabilities(app)
        if not (info['mathtype_class_registered'] and info['mathtype_commands_loaded']):
            self.skipTest('Requires MathType 7 desktop and the Word Commands add-in')
        with TemporaryDirectory() as folder:
            root = Path(folder)
            source, target = root / 'original.docx', root / 'edited.docx'
            run_plan(word_native, {
                'schema_version': 1, 'target': str(source),
                'operations': [
                    {'op': 'equation', 'id': 'balance', 'engine': 'mathtype', 'text': r'\frac{x_1^2}{y}'},
                    {'op': 'paragraph', 'runs': [{'ref': 'balance', 'kind': 'equation'}]},
                ],
            }, root / 'create.json')
            original = source.read_bytes()
            result = run_plan(word_native, {
                'schema_version': 1, 'source': str(source), 'target': str(target),
                'operations': [{'op': 'replace_mathtype', 'id': 'balance', 'text': r'\frac{x_1^3}{z}'}],
            }, root / 'edit.json')
            self.assertEqual(source.read_bytes(), original)
            self.assertEqual(result['state']['mathtype_equations'], 1)
            before, after = check(source), check(target)
            self.assertEqual(before['sequences'], after['sequences'])
            self.assertEqual(before['references'], after['references'])
            self.assertNotEqual(self.math_payload(source), self.math_payload(target))

    def math_payload(self, path):
        with ZipFile(path) as archive:
            names = [name for name in archive.namelist() if name.startswith('word/embeddings/')]
            self.assertEqual(len(names), 1)
            return archive.read(names[0])


if __name__ == '__main__':
    unittest.main()
