import json
from pathlib import Path
import subprocess
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch, MagicMock

from core.desktop_lock import DesktopBusyError, DesktopLock
from skills.ppt.scripts.ppt_native import (PowerPoint, bounds, find_unique, rgb, run_plan,
                                selector_match, set_chart_data, style_shape, utf16_length, validate_plan)


class NativeInputTests(unittest.TestCase):
    def test_invalid_batch_is_rejected_before_connecting(self):
        with tempfile.TemporaryDirectory() as directory:
            target = str(Path(directory)/'deck.pptx')
            factory = MagicMock()
            for operation in ({'op': 'typo'}, {'op': 'text', 'name': 'x', 'box': [0, 0, -1, 3]},
                              {'op': 'snapshot', 'path': target}):
                with self.assertRaises(ValueError):
                    run_plan({'schema_version': 1, 'target': target, 'operations': [operation]},
                             str(Path(directory)/'report.json'), factory)
            factory.assert_not_called()

    def test_report_cannot_overwrite_the_presentation(self):
        with tempfile.TemporaryDirectory() as directory:
            target = str(Path(directory)/'deck.pptx')
            with self.assertRaises(ValueError):
                run_plan({'schema_version': 1, 'target': target, 'operations': []}, target, MagicMock())

    def test_source_selection_rejects_duplicates_and_zero(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory)/'source.pptx'
            source.write_bytes(b'fixture')
            for indices in ([1, 1], [0], []):
                with self.assertRaises(ValueError):
                    validate_plan({'schema_version': 1, 'target': str(Path(directory)/'output.pptx'),
                                   'create': {'template': str(source), 'slides': indices}, 'operations': []})

    def test_color_and_geometry_validation(self):
        self.assertEqual(rgb('#14386C'), 0x6C3814)
        for value in ([0, 0, float('nan'), 2], [0, 0, True, 2], [0, 0, 0, 2]):
            with self.assertRaises(ValueError):
                bounds(value)
        with self.assertRaises(ValueError):
            rgb('ZZZZZZ')

    def test_ambiguous_names_are_not_silently_selected(self):
        objects = [SimpleNamespace(Name='same', Id=i) for i in (1, 2)]
        with self.assertRaises(ValueError):
            find_unique(objects, 'same')
        self.assertIs(find_unique(objects, {'id': 2}), objects[1])

    def test_unicode_rich_text_uses_powerpoint_utf16_offsets(self):
        self.assertEqual(utf16_length('中😀A'), 4)
        shape = MagicMock()
        shape.Left, shape.Top, shape.Width, shape.Height = 10, 20, 200, 60
        style_shape(shape, {'runs': [{'text': '中😀', 'bold': True}, {'text': 'A', 'color': 'FF0000'}],
                            'size': 20, 'font': 'Arial'})
        calls = shape.TextFrame2.TextRange.GetCharacters.call_args_list
        self.assertEqual([tuple(c.args) for c in calls], [(1, 3), (4, 1)])
        self.assertEqual((shape.Left, shape.Top, shape.Width, shape.Height), (10, 20, 200, 60))

    def test_table_cell_does_not_change_unsupported_autosize_or_geometry(self):
        shape = MagicMock()
        style_shape(shape, {'text': '单元格', 'size': 18}, cell=True)
        self.assertNotIn('AutoSize', shape.TextFrame2.__dict__)
        self.assertNotIn('Width', shape.__dict__)

    def test_autosize_and_template_geometry_survive_all_format_changes(self):
        shape = MagicMock()
        shape.Left, shape.Top, shape.Width, shape.Height = 1, 2, 300, 44
        shape.TextFrame2.AutoSize = 1
        style_shape(shape, {'text': '长标题', 'size': 30, 'font': 'Microsoft YaHei', 'margins': [0, 0, 0, 0]})
        self.assertEqual(shape.TextFrame2.AutoSize, 1)
        self.assertEqual((shape.Left, shape.Top, shape.Width, shape.Height), (1, 2, 300, 44))

    def test_tags_are_case_insensitive_across_com_versions(self):
        shape = SimpleNamespace(Tags=SimpleNamespace(Item=lambda _: 'my-key'))
        self.assertTrue(selector_match(shape, {'tag': 'MY-KEY'}))

    def test_linked_chart_does_not_edit_an_external_workbook(self):
        chart = MagicMock()
        chart.ChartData.IsLinked = True
        with self.assertRaises(ValueError):
            set_chart_data(chart, {'categories': ['A'], 'series': [{'name': 'x', 'values': [1]}]})
        chart.ChartData.Activate.assert_not_called()

    def test_reopen_snapshot_cannot_replace_target(self):
        with tempfile.TemporaryDirectory() as directory:
            target = str(Path(directory)/'deck.pptx')
            with self.assertRaises(ValueError):
                validate_plan({'schema_version': 1, 'target': target, 'operations': [], 'verify_snapshot': target})

    def test_invalid_style_and_picture_fit_are_preflight_errors(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            picture = root/'image.png'
            picture.write_bytes(b'fixture')
            bad = [{'op': 'shape', 'slide': 1, 'name': 'node', 'box': [0, 0, 1, 1], 'fill': 'bad'},
                   {'op': 'picture', 'slide': 1, 'name': 'img', 'box': [0, 0, 1, 1],
                    'path': str(picture), 'fit': 'bad'}, {'op': 'text', 'slide': 1, 'name': 'missing-box'}]
            factory = MagicMock()
            for op in bad:
                with self.assertRaises(ValueError):
                    run_plan({'schema_version': 1, 'target': str(root/'deck.pptx'), 'operations': [op]},
                             root/'report.json', factory)
            factory.assert_not_called()


class NativeLifecycleTests(unittest.TestCase):
    def test_connection_failure_is_recorded_without_replaying(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            def unavailable():
                raise RuntimeError('PowerPoint unavailable')
            result = run_plan({'schema_version': 1, 'target': str(root/'deck.pptx'), 'operations': []},
                              root/'report.json', unavailable)
            self.assertEqual(result['status'], 'failed')
            self.assertIn('PowerPoint unavailable', result['error'])
            self.assertEqual(result['operations'], [])
    def test_named_slide_is_findable_and_unnamed_copy_drops_inherited_tag(self):
        api = PowerPoint(None)
        slide = MagicMock()
        api.name_slide(slide, 'background')
        self.assertEqual(slide.Name, 'background')
        slide.Tags.Add.assert_called_with('PPTSKILL', 'background')
        api.name_slide(slide)
        slide.Tags.Delete.assert_called_with('PPTSKILL')

    def test_completed_unchanged_batch_is_not_replayed(self):
        from hashlib import sha256
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            target = root/'deck.pptx'
            target.write_bytes(b'saved')
            plan = {'schema_version': 1, 'target': str(target), 'operations': []}
            digest = sha256(json.dumps(plan, sort_keys=True, ensure_ascii=False).encode()).hexdigest()
            report = {'plan_sha256': digest, 'status': 'completed', 'saved': True, 'target_sha256': sha256(b'saved').hexdigest()}
            path = root/'report.json'
            path.write_text(json.dumps(report), encoding='utf-8')
            factory = MagicMock()
            self.assertTrue(run_plan(plan, path, factory)['replay_skipped'])
            factory.assert_not_called()
            target.write_bytes(b'changed')
            with self.assertRaises(RuntimeError):
                run_plan(plan, path, factory)
    def test_explicit_target_does_not_use_active_presentation(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory)/'deck.pptx'
            pres = SimpleNamespace(FullName=str(target), ReadOnly=False)
            presentations = SimpleNamespace(Count=1, Item=lambda _: pres)
            app = SimpleNamespace(Presentations=presentations)
            api = PowerPoint(app)
            api.bind(target)
            self.assertIs(api.presentation, pres)
            self.assertFalse(api.opened_here)

    def test_changed_bound_path_stops_execution(self):
        api = PowerPoint(None)
        api.target = Path('deck.pptx').resolve()
        api.presentation = SimpleNamespace(FullName=str(Path('other.pptx').resolve()))
        with self.assertRaises(RuntimeError):
            api.execute({'op': 'save'})

    def test_existing_target_is_not_overwritten(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory)/'deck.pptx'
            path.write_bytes(b'original')
            with self.assertRaises(FileExistsError):
                PowerPoint(None).bind(path, {})
            self.assertEqual(path.read_bytes(), b'original')

    def test_partial_failure_is_journaled_and_never_replayed(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory)
            plan = {'schema_version': 1, 'target': str(path/'deck.pptx'),
                    'operations': [{'op': 'add_slide'}, {'op': 'add_slide'}]}
            api = MagicMock()
            api.app.Version = 'test'
            api.execute.side_effect = [{'slide_id': 1}, RuntimeError('interrupted')]
            with patch('skills.ppt.scripts.ppt_native.PowerPoint', return_value=api):
                result = run_plan(plan, path/'report.json', lambda: None)
            self.assertEqual(result['status'], 'failed')
            self.assertEqual([i['status'] for i in result['operations']], ['completed', 'started'])
            self.assertEqual(json.loads((path/'report.json').read_text(encoding='utf-8'))['status'], 'failed')
            with self.assertRaises(RuntimeError):
                run_plan(plan, path/'report.json', lambda: self.fail('should not reconnect'))

    def test_shared_lock_excludes_a_second_process(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory)/'test.lock'
            source = 'from core.desktop_lock import DesktopLock;\nwith DesktopLock('+repr(str(path))+'): pass'
            with DesktopLock(path):
                result = subprocess.run([sys.executable, '-c', source], capture_output=True)
                self.assertNotEqual(result.returncode, 0)
            with DesktopLock(path):
                pass
