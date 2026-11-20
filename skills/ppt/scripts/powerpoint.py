#!/usr/bin/env python3
"""Skill-local PowerPoint desktop control. Native COM, no MCP server or SDK."""

import argparse
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import venv

ROOT = Path(__file__).resolve().parents[3]
RUNTIME = ROOT / '.runtime' / 'powerpoint-native'
if not __package__:  # Also support direct invocation with the isolated interpreter (-I).
    sys.path.insert(0, str(ROOT))
from core.desktop_lock import DesktopLock
from skills.ppt.scripts.ppt_native import VERSION, OPS, PowerPoint, connect, run_plan, absolute


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    sub.add_parser('inspect', help='Inspect runtime dependencies without starting PowerPoint')
    sub.add_parser('setup', help='Install only pywin32 into a skill-local virtual environment')
    sub.add_parser('probe', help='Connect to the real PowerPoint application without editing documents')
    run = sub.add_parser('run', help='Execute a JSON plan as one serialized desktop batch')
    run.add_argument('--plan', type=Path, required=True)
    run.add_argument('--report', type=Path)
    snap = sub.add_parser('snapshot', help='Read a named document and save its object state')
    snap.add_argument('--target', required=True)
    snap.add_argument('--output', required=True)
    snap.add_argument('--pages', type=int, nargs='+')
    args = parser.parse_args()
    if args.command == 'inspect':
        return {'engine': 'powerpoint-native', 'version': VERSION, 'windows': os.name == 'nt',
                'python': sys.executable, 'pywin32': importlib.util.find_spec('win32com') is not None,
                'isolated_python': str(RUNTIME/'Scripts/python.exe'), 'operations': sorted(OPS),
                'requires': ['Windows', 'Microsoft PowerPoint', 'Python 3.10+', 'pywin32'],
                'office_connection': 'not_probed', 'requires_mcp': False}
    if os.name != 'nt':
        raise RuntimeError('Windows desktop PowerPoint is required')
    if args.command == 'setup':
        if not (RUNTIME/'Scripts/python.exe').is_file():
            venv.EnvBuilder(with_pip=True).create(RUNTIME)
        subprocess.run([str(RUNTIME/'Scripts/python.exe'), '-m', 'pip', 'install', 'pywin32>=306,<400'],
                       check=True, stdout=sys.stderr)
        return {'runtime': str(RUNTIME), 'installed': ['pywin32'], 'requires_mcp': False}
    if importlib.util.find_spec('win32com') is None:
        python = RUNTIME/'Scripts/python.exe'
        if python.is_file() and Path(sys.executable).resolve() != python.resolve():
            raise SystemExit(subprocess.call([str(python), '-I', str(Path(__file__).resolve()), *sys.argv[1:]]))
        raise RuntimeError('pywin32 is unavailable; run python skills/ppt/scripts/powerpoint.py setup from the repository root')
    if args.command == 'snapshot':
        with DesktopLock():
            api = PowerPoint(connect())
            try:
                api.bind(args.target)
                return api.execute({'op': 'snapshot', 'path': args.output,
                                    **({'pages': args.pages} if args.pages else {})})
            finally:
                if api.opened_here and api.presentation is not None:
                    api.presentation.Close()
    if args.command == 'probe':
        with DesktopLock():
            app = connect()
            return {'engine': 'powerpoint-native', 'office_connection': 'verified',
                    'office_version': str(app.Version), 'presentations': int(app.Presentations.Count),
                    'requires_mcp': False}
    plan = json.loads(args.plan.read_text(encoding='utf-8-sig'))
    report = (args.report or args.plan.with_suffix('.result.json')).resolve()
    if report == args.plan.resolve() or any(absolute(o['path']) == args.plan.resolve()
                                         for o in plan['operations'] if o.get('op') in {'snapshot', 'pdf'}):
        raise ValueError('Outputs must not overwrite the input plan')
    if plan.get('verify_snapshot') and absolute(plan['verify_snapshot']) == args.plan.resolve():
        raise ValueError('verify_snapshot must not overwrite the input plan')
    result = run_plan(plan, report)
    return {'status': result['status'], 'report': str(report), 'target': result['target'],
            'slides': result.get('slides'), 'seconds': result.get('seconds'),
            'completed_operations': sum(i['status'] == 'completed' for i in result['operations']),
            'mcp_loaded': result['mcp_loaded'], 'replay_skipped': result.get('replay_skipped', False),
            **({'error': result['error']} if 'error' in result else {})}


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    try:
        result = main()
    except Exception as exc:
        result = {'status': 'input_or_connection_error', 'error': f'{type(exc).__name__}: {exc}'}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(0 if result.get('status') not in {'failed', 'input_or_connection_error'} else 1)
