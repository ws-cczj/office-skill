"""Run the bundled editable flowchart blueprint or revise the same drawing."""
from skills.visio.scripts.visio_bridge import run_request
from skills.visio.scripts.visio_flow import render_request


def plan(data, folder, params):
    document = folder / 'visio' / 'diagram.vsdx'
    preview = folder / 'renders' / 'visio' / 'diagram.png'
    if data.get('update'):
        return {'action': 'apply', 'document': str(document), 'page': data.get('page', 'Diagram'),
                'operations': data['operations'], 'save': True, 'reopen': True,
                'exports': [str(preview)], 'png_dpi': 300}
    result = render_request(data, document, preview)
    result['reopen'] = True
    return result


def run(plan, data, folder, params):
    result = run_request(plan)
    state = result['result']
    return {'ok': True, 'target': plan['document'], 'connection': result['connection'],
            'seconds': result['elapsed_seconds'], 'reopened': state['reopened'],
            'shapes': state['state']['shape_count'], 'exports': state['exports']}
