"""Apply a manuscript plan with optional journal settings and PDF export."""
from core.office_batch import run_plan
from core.recipes import input_path
from skills.word.scripts import word_native


def plan(data, folder, params):
    result = {
        'schema_version': 1, 'target': str(folder / 'word' / 'manuscript.docx'),
        'operations': list(data.get('operations', [])),
        'exports': [{'path': str(folder / 'renders' / 'word' / 'manuscript.pdf')}],
    }
    for key in ('source', 'template'):
        if key in data:
            result[key] = input_path(data[key], params)
    if 'update' in data:
        result['update'] = data['update']
    if 'layout' in data:
        result['operations'].insert(0, {**data['layout'], 'op': 'configure'})
    elif not (data.get('source') or data.get('template') or data.get('update')):
        result['operations'].insert(0, {'op': 'configure', 'profile': 'sci'})
    return result


def run(plan, data, folder, params):
    return run_plan(word_native, plan, params)
