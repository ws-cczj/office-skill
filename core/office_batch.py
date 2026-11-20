"""Run a Word or Excel plan in a private desktop COM instance."""
import argparse
from contextlib import ExitStack, contextmanager
from importlib.util import find_spec
import json
import os
from pathlib import Path
import shutil
import sys
from tempfile import TemporaryDirectory

from core.desktop_lock import DesktopLock


def absolute(value, suffixes=()):
    path = Path(value).expanduser()
    if not path.is_absolute():
        raise ValueError(f"Use an absolute path: {path}")
    path = path.resolve()
    if suffixes and path.suffix.lower() not in suffixes:
        raise ValueError(f"Expected {', '.join(sorted(suffixes))}: {path}")
    return path


def existing(value, suffixes=()):
    path = absolute(value, suffixes)
    if not path.is_file():
        raise FileNotFoundError(path)
    return path


@contextmanager
def office(adapter):
    import pythoncom
    from win32com.client import DispatchEx
    from win32com.client.dynamic import CDispatch

    with DesktopLock():
        pythoncom.CoInitialize()
        app = None
        try:
            app = DispatchEx(adapter.PROG_ID)
            if isinstance(app, CDispatch):
                app._FlagAsMethod('Quit')
            app.DisplayAlerts = 0
            app.AutomationSecurity = 3  # Disable macros in files opened by automation.
            yield app
        finally:
            try:
                if app is not None:
                    adapter.quit(app)
            finally:
                app = None
                pythoncom.CoUninitialize()


@contextmanager
def document(adapter, app, path, readonly):
    doc = adapter.open_document(app, path, readonly)
    try:
        yield doc
    finally:
        adapter.close(doc)


def staging_path(stack, destination):
    destination.parent.mkdir(parents=True, exist_ok=True)
    work = Path(stack.enter_context(TemporaryDirectory(prefix='.office-', dir=destination.parent)))
    return work / destination.name


def publish(pairs):
    published = []
    try:
        for temporary, destination in pairs:
            # On Windows rename within the volume is atomic and refuses an existing destination.
            temporary.rename(destination)
            published.append(destination)
    except OSError:
        for destination in reversed(published):
            destination.unlink()
        raise


def run_plan(adapter, plan, plan_path):
    if plan.get('schema_version') != 1:
        raise ValueError('schema_version must be 1')
    target = absolute(plan['target'], {adapter.EXTENSION})
    source = existing(plan['source'], {adapter.EXTENSION}) if 'source' in plan else None
    template = adapter.template_path(plan['template']) if 'template' in plan else None
    if source and template:
        raise ValueError('Choose source or template')
    operations = plan.get('operations', [])
    adapter.validate(operations)
    exports = [(adapter.export_path(spec), spec) for spec in plan.get('exports', [])]
    destinations = [target, *(path for path, _ in exports)]
    if len(destinations) != len(set(destinations)) or set(destinations) & {source, template, plan_path}:
        raise ValueError('Output paths must be unique and distinct from inputs')

    with ExitStack() as stack:
        with office(adapter) as app:
            for path in destinations:
                if path.exists():
                    raise FileExistsError(f'Refusing to overwrite: {path}')
            working = staging_path(stack, target)
            if source:
                shutil.copyfile(source, working)
                doc = adapter.open_document(app, working, False)
            else:
                doc = adapter.create(app, template)
            try:
                adapter.apply(app, doc, operations)
                adapter.save(doc, working, bool(source))
            finally:
                adapter.close(doc)

            pairs = [(working, target)]
            with document(adapter, app, working, True) as reopened:
                state = adapter.snapshot(reopened, plan.get('snapshot', {}))
                for path, spec in exports:
                    temporary = staging_path(stack, path)
                    adapter.export(reopened, spec, temporary)
                    if not temporary.is_file() or not temporary.stat().st_size:
                        raise RuntimeError(f'Office did not export a nonempty file: {path}')
                    pairs.append((temporary, path))
        publish(pairs)
    return {'target': str(target), 'reopened': True, 'state': state,
            'exports': [str(path) for path, _ in exports]}


def cli(adapter):
    parser = argparse.ArgumentParser(description=f'Native desktop COM for {adapter.PROG_ID}')
    sub = parser.add_subparsers(dest='command', required=True)
    sub.add_parser('inspect', help='Check Python and pywin32 without launching Office')
    sub.add_parser('probe', help='Connect to the desktop application and report capabilities')
    run = sub.add_parser('run', help='Apply a JSON plan to a new output or source copy')
    run.add_argument('--plan', type=Path, required=True)
    snapshot = sub.add_parser('snapshot', help='Read saved state through desktop COM')
    snapshot.add_argument('--target', type=Path, required=True)
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding='utf-8')
    try:
        dependencies = {
            'application': adapter.PROG_ID,
            'windows': os.name == 'nt',
            'python': sys.executable,
            'pywin32': find_spec('win32com') is not None,
            'operations': sorted(adapter.OPERATIONS),
        }
        if args.command == 'inspect':
            result = {**dependencies, 'office_connection': 'not_probed'}
        else:
            if not dependencies['windows'] or not dependencies['pywin32']:
                raise RuntimeError('Windows desktop Office and pywin32 are required')
            if args.command == 'run':
                plan_path = args.plan.resolve(strict=True)
                plan = json.loads(plan_path.read_text(encoding='utf-8-sig'))
                result = run_plan(adapter, plan, plan_path)
            elif args.command == 'probe':
                with office(adapter) as app:
                    result = {**dependencies, 'office_connection': 'verified', **adapter.capabilities(app)}
            else:
                target = existing(args.target, {adapter.EXTENSION})
                with office(adapter) as app, document(adapter, app, target, True) as doc:
                    result = {'target': str(target), 'state': adapter.snapshot(doc, {})}
        print(json.dumps({'ok': True, **result}, ensure_ascii=False))
        return 0
    except Exception as error:
        print(json.dumps({'ok': False, 'error_type': type(error).__name__, 'error': str(error)},
                         ensure_ascii=False))
        return 1
