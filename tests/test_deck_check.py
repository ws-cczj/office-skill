import tempfile
import unittest
from pathlib import Path

from scripts.check_deck import (Shape, check_script, check_timeline, empty_figure_slots,
                               content_layout_review, internal_copy_findings, page_selection)


def pt(x, y, w, h):
    return tuple(value * 12700 for value in (x, y, w, h))


class DeckCheckTests(unittest.TestCase):
    def setUp(self):
        self.frame = Shape('source-frame', 'sp', pt(50, 100, 200, 150), 'F8F8F8', '')
        label = Shape('source-label', 'sp', pt(60, 130, 180, 30), None, '[研究示意图]')
        self.template = [([self.frame, label], '')]

    def test_text_in_original_figure_slot_is_flagged(self):
        words = Shape('source-label', 'sp', pt(60, 130, 180, 70), None, '输入\n↓\n输出')
        issues = list(empty_figure_slots([([self.frame, words], '')], self.template))
        self.assertEqual([(i['page'], i['object']) for i in issues], [(1, self.frame.name)])

    def test_foreground_picture_or_diagram_is_accepted(self):
        picture = Shape('figure', 'pic', pt(55, 110, 185, 115), None, '')
        node = Shape('node', 'sp', pt(60, 130, 60, 35), '14386C', '')
        arrow = Shape('arrow', 'cxnSp', pt(120, 145, 55, 0), None, '')
        for visual in ([picture], [node, arrow]):
            self.assertEqual(list(empty_figure_slots([([self.frame, *visual], '')], self.template)), [])
        self.assertTrue(list(empty_figure_slots([([picture, self.frame], '')], self.template)))

    def test_timeline_checks_latest_label_not_only_schedule_text(self):
        labels = [Shape('task', 'sp', None, None, '2028.03—2028.04'),
                  Shape('axis', 'sp', None, None, '2027.11—2028.06')]
        issue = check_timeline([(labels, '')], (1, 2028, 4))
        self.assertEqual(issue['actual_latest_label'], '2028-06')
        self.assertIsNone(check_timeline([(labels[:1], '')], (1, 2028, 4)))

    def test_note_must_match_the_corresponding_spoken_body(self):
        with tempfile.TemporaryDirectory() as root:
            path = Path(root) / 'speech.md'
            path.write_text('# 发言稿\n\n## 01｜说明\n\n先看右图。\n\n## 02｜验证\n\n再看对照。\n', encoding='utf-8')
            issues = list(check_script([([], '先看右图。'), ([], '来源：报告第二节。')], path))
            self.assertEqual([(i['code'], i['page']) for i in issues], [('speaker_notes_mismatch', 2)])

    def test_flags_internal_source_text_without_flagging_academic_attribution(self):
        objects = [
            Shape('footnote', 'sp', None, None, '依据：开题报告 §1.1、§2.4。'),
            Shape('caption', 'sp', None, None, '来源：Zhang 等（2025），图 4，CC BY 4.0'),
            Shape('title', 'sp', None, None, '硕士学位论文开题报告'),
            Shape('number', 'sp', None, None, '03'),
        ]
        items = list(internal_copy_findings([(objects, '先看右图。\n素材路径：E:/work/figure.png')]))
        self.assertEqual([(item['code'], item['page'], item.get('object')) for item in items],
                         [('internal_copy_on_slide', 1, 'footnote'),
                          ('internal_copy_in_notes', 1, None)])

    def test_new_names_and_text_cannot_hide_repeated_body_containers(self):
        slides = []
        for n in range(3):
            objects = [Shape(f'new-{n}-{x}', 'sp', pt(x, 180, 400, 240), 'F0F4F8', '')
                       for x in (40, 510)]
            objects.append(Shape(f'label-{n}', 'sp', pt(60, 240, 300, 80), None, str(n)))
            slides.append((objects, ''))
        size = (960*12700, 540*12700)
        profiles, findings = content_layout_review(slides, [1, 2, 3], size)
        self.assertEqual(findings[0]['pages'], [1, 2, 3])
        # A header logo does not count as body evidence. One actual body image breaks the trio.
        slides[0][0].append(Shape('logo', 'pic', pt(10, 5, 40, 40), None, ''))
        self.assertEqual(len(content_layout_review(slides, [1, 2, 3], size)[1]), 1)
        slides[0][0].append(Shape('model', 'pic', pt(80, 220, 200, 160), None, ''))
        self.assertEqual(content_layout_review(slides, [1, 2, 3], size)[1], [])
        self.assertEqual(profiles[0]['body_media_or_groups'], 0)

    def test_content_page_ranges_and_missing_pages(self):
        self.assertEqual(page_selection('3,4,9-11,3'), [3, 4, 9, 10, 11])
        with self.assertRaises(ValueError):
            content_layout_review([([], '')], [2], (960, 540))
