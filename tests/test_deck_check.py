import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile

from scripts.check_deck import (NS, Shape, check_script, check_timeline, empty_figure_slots,
                               content_layout_review, internal_copy_findings, page_selection,
                               third_person_pov_findings)


def pt(x, y, w, h):
    return tuple(value * 12700 for value in (x, y, w, h))


def write_text_deck(path, visible, hidden, notes):
    namespaces = ' '.join(f'xmlns:{key}="{uri}"' for key, uri in NS.items())

    def text_shape(name, text, hidden=False, note=False):
        visibility = ' hidden="1"' if hidden else ''
        placeholder = '<p:ph type="body"/>' if note else ''
        return (f'<p:sp><p:nvSpPr><p:cNvPr name="{name}"{visibility}/><p:nvPr>{placeholder}</p:nvPr>'
                f'</p:nvSpPr><p:txBody><a:p><a:r><a:t>{escape(text)}</a:t></a:r></a:p></p:txBody></p:sp>')

    def shape_tree(tag, objects):
        return f'<p:{tag} {namespaces}><p:cSld><p:spTree>{objects}</p:spTree></p:cSld></p:{tag}>'

    def relationship(role, target):
        rel_type = f'{NS["r"]}/{role}'
        return f'<Relationships><Relationship Id="rId1" Type="{rel_type}" Target="{target}"/></Relationships>'

    group = ('<p:grpSp><p:nvGrpSpPr><p:cNvPr name="分组"/></p:nvGrpSpPr>'
             f'{text_shape("可见图注", visible)}{text_shape("隐藏说明", hidden, hidden=True)}</p:grpSp>')
    with ZipFile(path, 'w') as package:
        package.writestr('ppt/presentation.xml', f'<p:presentation {namespaces}>'
                         '<p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>')
        package.writestr('ppt/_rels/presentation.xml.rels', relationship('slide', 'slides/slide1.xml'))
        package.writestr('ppt/slides/slide1.xml', shape_tree('sld', group))
        package.writestr('ppt/slides/_rels/slide1.xml.rels', relationship('notesSlide', '../notesSlides/notesSlide1.xml'))
        package.writestr('ppt/notesSlides/notesSlide1.xml', shape_tree('notes', text_shape('备注', notes, note=True)))


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
        self.assertEqual([(i['code'], i['page'], i.get('object'), i['matched']) for i in items], [
            ('internal_copy_on_slide', 1, 'footnote', '依据：开题报告 §1.1'),
            ('internal_copy_in_notes', 1, None, '素材路径：'),
            ('internal_copy_in_notes', 1, None, 'E:/work/figure.png'),
        ])

    def test_internal_copy_includes_exact_inline_fragments(self):
        examples = [
            ('证据对照（依据：参见本次 Word）', '依据：参见本次 Word'),
            ('原图（来源：用户提供的Word文档）', '来源：用户提供的Word文档'),
            ('测点分布，来源：用户提供的原始报告', '来源：用户提供的原始报告'),
            ('工作边界，出处：随附的报告。', '出处：随附的报告'),
            ('图示取自报告图17 · 设备实物', '报告图17'),
            ('参数见原报告第表 3。', '原报告第表 3'),
            ('控制模型见报告式（8）。', '报告式（8）'),
            ('参见开题报告 §1.1。', '开题报告 §1.1'),
            ('对照原报告第二章。', '原报告第二章'),
            ('对照原报告第 2 章。', '原报告第 2 章'),
            ('参数对照，依据：§2.4。', '依据：§2.4'),
            ('边界按报告 2.4.1 设置。', '报告 2.4.1'),
            ('条目编号沿用原报告。', '编号沿用原报告'),
            ('流程清单，路线按本次汇报重绘。', '按本次汇报重绘'),
            ('流程按本次汇报重新绘制。', '按本次汇报重新绘制'),
            ('装置图按原要求重绘。', '按原要求重绘'),
            ('输入与输出（流向为符号与核算示意）。', '流向为符号与核算示意'),
            ('箭头仅为方向示意。', '箭头仅为方向示意'),
            ('配色说明，颜色仅为符号示意。', '颜色仅为符号示意'),
            ('图中比例与流向仅作示意', '流向仅作示意'),
            ('试件分组，示意用深蓝色表示材料分段。', '示意用深蓝色表示'),
            ('进度按报告修订版。', '报告修订版'),
            ('验证方法（修订版）', '验证方法（修订版）'),
            ('研究方案修订版', '研究方案修订版'),
            ('工作记录 · 修订版', '· 修订版'),
            ('工作记录 · 修订版，待核对', '· 修订版'),
        ]
        for text, fragment in examples:
            with self.subTest(text=text):
                shape = Shape('图注', 'sp', None, None, text)
                hits = list(internal_copy_findings([([shape], '')]))
                self.assertEqual([hit['matched'] for hit in hits], [fragment])

    def test_third_person_pov_includes_exact_self_report_fragments(self):
        examples = [
            ('报告记载已有对照实验。', '报告记载'),
            ('报告中列出了参数。', '报告中列出'),
            ('开题报告里给出了参数。', '开题报告里给出'),
            ('以原报告所述为准。', '原报告所述'),
            ('该报告写明了边界条件。', '该报告写明'),
            ('原始报告提到相关条件。', '原始报告提到'),
            ('右图是原报告提供的装置。', '右图是原报告提供的'),
            ('进度按开题报告调整。', '进度按开题报告'),
            ('进度按报告修订版安排。', '进度按报告修订版'),
        ]
        for text, fragment in examples:
            with self.subTest(text=text):
                hits = list(third_person_pov_findings([([], text)]))
                self.assertEqual([hit['matched'] for hit in hits], [fragment])

    def test_ordinary_academic_terminology_and_attribution_are_allowed(self):
        valid = [
            '采用中文版修订版量表。',
            '试件按测试规范修订版制作。',
            '参见实验规范（修订版）。',
            '使用软件修订版和最新求解器版本。',
            '来源：Li 等（2024），原图 4，CC BY 4.0。',
            '来源：机构技术报告，DOI: 10.1000/182。',
            '依据：实验报告与研究报告。',
            '依照测试规范 §3.2 的方法加载。',
            '图 4 中的比例尺为 100 μm。',
            '右图展示了结构示意图与机理示意图。',
            '图中箭头表示流向，蓝色对应试件。',
            '图中颜色为应力分布示意，红色为高应力区。',
            '流向为热量传递示意。',
            '箭头为载荷方向示意。',
            '线条为边界示意。',
            '文献[3]指出该差异，前人研究报道了相同趋势。',
            '文献报告了试验结果，作者发现了该关系。',
            '公开技术报告指出这一差异，实验报告提供了结果。',
            '第三方检测报告提供的数据显示屈服强度为 350 MPa。',
            '检验报告给出了误差范围。',
            '调查报告列出了常见的失效模式。',
            '临床报告记载了 120 例数据。',
            '监测报告指出温度出现周期变化。',
            '既有研究报告列出了可比条件。',
            '《公开研究报告》指出这一差异。',
            '我们按报告规范记录不确定度。',
            '进度按报告格式统一呈现。',
            '本课题拟使用该装置。',
        ]
        for text in valid:
            with self.subTest(text=text):
                shape = Shape('正文', 'sp', None, None, text)
                slides = [([shape], text)]
                self.assertEqual(list(internal_copy_findings(slides)), [])
                self.assertEqual(list(third_person_pov_findings(slides)), [])

    def test_script_copy_in_preamble_and_each_page(self):
        notes = '先介绍方案，再按报告修订版说明进度。依据：原报告表2。'
        slides = [([], notes), ([], '本课题拟验证。')]
        with tempfile.TemporaryDirectory() as root:
            path = Path(root) / 'speech.md'
            path.write_text('# 发言稿\n\n报告中列出了设备。按原报告重新绘制。\n\n'
                            f'## 01｜方案\n\n{notes}\n\n## 02｜验证\n\n本课题拟验证。\n',
                            encoding='utf-8')
            hits = list(check_script(slides, path))
        self.assertEqual([(hit['code'], hit.get('page'), hit.get('line'), hit['matched']) for hit in hits], [
            ('internal_copy_in_script', None, 3, '按原报告重新绘制'),
            ('third_person_pov_in_script', None, 3, '报告中列出'),
            ('internal_copy_in_script', 1, None, '报告修订版'),
            ('internal_copy_in_script', 1, None, '依据：原报告表2'),
            ('third_person_pov_in_script', 1, None, '按报告修订版'),
        ])

    def test_cli_checks_both_patterns_on_all_surfaces_and_skips_hidden_shapes(self):
        dirty = '条目编号沿用原报告。右图是报告提供的装置。'
        clean = '右图展示了机理示意图，文献[3]指出该差异。'
        checker = Path(__file__).resolve().parents[1] / 'scripts' / 'check_deck.py'
        with tempfile.TemporaryDirectory() as root:
            pptx = Path(root) / 'sample.pptx'
            script = Path(root) / 'speech.md'
            for text, exit_code in ((dirty, 1), (clean, 0)):
                with self.subTest(exit_code=exit_code):
                    write_text_deck(pptx, text, '依据：开题报告 §9.9。报告中提到。', text)
                    script.write_text(f'# 发言稿\n\n## 01｜方案\n\n{text}\n', encoding='utf-8')
                    result = subprocess.run([sys.executable, '-X', 'utf8', str(checker), str(pptx),
                                             '--speaker-script', str(script)], capture_output=True, encoding='utf-8')
                    self.assertEqual(result.returncode, exit_code, result.stderr or result.stdout)
                    report = json.loads(result.stdout)
                    self.assertEqual(report['quality_verdict'], 'not_assessed')
                    self.assertTrue({'internal_copy_literals', 'third_person_pov',
                                     'internal_copy_in_script', 'third_person_pov_in_script'} <= set(report['checks']))
                    found = [(i['code'], i['page'], i.get('object'), i['matched']) for i in report['issues']]
                    expected = [
                        (f'{code}_{location}', 1, name, fragment)
                        for code, fragment in [('internal_copy', '编号沿用原报告'),
                                               ('third_person_pov', '右图是报告提供的')]
                        for location, name in [('on_slide', '可见图注'), ('in_notes', None), ('in_script', None)]
                    ] if exit_code else []
                    self.assertCountEqual(found, expected)

    def test_cli_cannot_write_report_over_speaker_script(self):
        checker = Path(__file__).resolve().parents[1] / 'scripts' / 'check_deck.py'
        with tempfile.TemporaryDirectory() as root:
            script = Path(root) / 'speech.md'
            script.write_text('本课题拟验证。', encoding='utf-8')
            result = subprocess.run([sys.executable, '-X', 'utf8', str(checker), str(script.with_suffix('.pptx')),
                                     '--speaker-script', str(script), '--output', str(script)],
                                    capture_output=True, encoding='utf-8')
            self.assertEqual(result.returncode, 2)
            self.assertIn('--output 必须', result.stderr)
            self.assertEqual(script.read_text(encoding='utf-8'), '本课题拟验证。')

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
