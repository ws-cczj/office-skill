"""Fill existing template objects and notes, then reopen and verify the same deck."""
from pathlib import Path
import re

from core.recipes import input_path
from skills.ppt.scripts.ppt_native import run_plan


def plan(data, folder, params):
    update = data.get('update', False)
    if update and 'template' in data:
        raise ValueError('Use template for the first run and update=true for the working deck')
    pages = data['pages']
    operations = []
    for page in pages:
        selector = page['slide']
        for shape in page.get('shapes', []):
            operations.append({**shape, 'op': 'update', 'slide': selector})
        if 'notes' in page:
            operations.append({'op': 'notes', 'slide': selector, 'text': page['notes']})
    if not operations:
        raise ValueError('Provide a shape update or notes')
    result = {
        'schema_version': 1, 'target': str(folder / 'ppt' / 'deck.pptx'),
        'operations': operations, 'reopen': True, 'close': True,
        'verify_snapshot': str(folder / 'qa' / 'ppt-objects.json'),
        'verify_export': str(folder / 'renders' / 'ppt'),
    }
    if update:
        result['verify_pages'] = [page['slide'] for page in pages]
    else:
        result['create'] = {'template': input_path(data['template'], params)}
        if 'template_slides' in data:
            result['create']['slides'] = data['template_slides']
    return result


def run(plan, data, folder, params):
    batch = data.get('batch', 'initial')
    if not isinstance(batch, str) or not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9_-]*', batch):
        raise ValueError('Use a simple batch name for the execution journal')
    if data.get('update') and batch == 'initial':
        raise ValueError('Specify a new batch name for an update')
    report = folder / 'qa' / f'ppt-{batch}.json'
    protected = [report, Path(plan['target']), Path(plan['verify_snapshot'])]
    if params in protected:
        raise ValueError('Recipe parameters must be kept separate from outputs')
    result = run_plan(plan, report)
    keys = ('status', 'target', 'seconds', 'slides', 'reopened', 'foreground_operations', 'error', 'replay_skipped')
    return {'report': str(report), **{key: result[key] for key in keys if key in result}}
