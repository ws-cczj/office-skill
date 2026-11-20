"""Run a Word or Excel plan in a private desktop COM instance."""
import argparse
from contextlib import ExitStack, contextmanager, nullcontext
from importlib.util import find_spec
import json
import os
from pathlib import Path
import shutil
import sys
from tempfile import TemporaryDirectory
import time

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
def office(adapter, *, acquire_lock=True):
    import pythoncom
    from win32com.client import DispatchEx
    from win32com.client.dynamic import CDispatch

    with (DesktopLock() if acquire_lock else nullcontext()):
        pythoncom.CoInitialize()
        app = None
        try:
            app = DispatchEx(adapter.PROG_ID)
            if isinstance(app, CDispatch):
                app._FlagAsMethod('Quit')
            app.Visible = False
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


def stage(stack, workspaces, destination):
    destination.parent.mkdir(parents=True, exist_ok=True)
    device = destination.parent.stat().st_dev
    if device not in workspaces:
        workspaces[device] = Path(stack.enter_context(
            TemporaryDirectory(prefix='.office-', dir=destination.parent)
        ))
    workspace = workspaces[device]
    folder = workspace / str(sum(1 for _ in workspace.iterdir()))
    folder.mkdir()
    return folder / destination.name


def file_state(path):
    try:
        info = path.stat()
    except FileNotFoundError:
        return None
    if not path.is_file():
        raise ValueError(f'Expected a file path: {path}')
    return info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns


def publish(pairs, expected):
    published = []
    try:
        for index, (temporary, destination) in enumerate(pairs):
            before = expected[destination]
            if file_state(destination) != before:
                raise RuntimeError(f'Output changed while Office was working: {destination}')
            backup = None
            if before is not None:
                # The document is published last. Only replaced exports need rollback data.
                if index < len(pairs) - 1:
                    backup = temporary.parent / '.previous'
                    shutil.copy2(destination, backup)
                temporary.replace(destination)
            else:
                temporary.rename(destination)  # On Windows this cannot overwrite another file.
            published.append((destination, backup))
    except (OSError, RuntimeError, ValueError):
        for destination, backup in reversed(published):
            if backup is None:
                destination.unlink()
            else:
                backup.replace(destination)
        raise


def run_plan(adapter, plan, plan_path):
    started = time.perf_counter()
    if plan.get('schema_version') != 1:
        raise ValueError('schema_version must be 1')
    target = absolute(plan['target'], {adapter.EXTENSION})
    update = plan.get('update', False)
    if type(update) is not bool:
        raise ValueError('update must be a boolean')
    source = existing(plan['source'], {adapter.EXTENSION}) if 'source' in plan else None
    template = adapter.template_path(plan['template']) if 'template' in plan else None
    if (source and template) or (update and (source or template)):
        raise ValueError('Choose source, template or update')
    operations = plan.get('operations', [])
    adapter.validate(operations)
    exports = [(adapter.export_path(spec), spec) for spec in plan.get('exports', [])]
    destinations = [target, *(path for path, _ in exports)]
    if len(destinations) != len(set(destinations)) or set(destinations) & {source, template, plan_path}:
        raise ValueError('Output paths must be unique and distinct from inputs')

    with DesktopLock(), ExitStack() as stack:
        expected = {path: file_state(path) for path in destinations}
        if update:
            if expected[target] is None:
                raise FileNotFoundError(f'Working draft does not exist: {target}')
            source = target
        else:
            for path, state in expected.items():
                if state is not None:
                    raise FileExistsError(f'Refusing to overwrite: {path}')
        workspaces = {}
        working = stage(stack, workspaces, target)
        with office(adapter, acquire_lock=False) as app:
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

            pairs = []
            with document(adapter, app, working, True) as reopened:
                state = adapter.snapshot(reopened, plan.get('snapshot', {}))
                for path, spec in exports:
                    temporary = stage(stack, workspaces, path)
                    adapter.export(reopened, spec, temporary)
                    if not temporary.is_file() or not temporary.stat().st_size:
                        raise RuntimeError(f'Office did not export a nonempty file: {path}')
                    pairs.append((temporary, path))
        publish([*pairs, (working, target)], expected)
    return {'target': str(target), 'reopened': True, 'updated': update, 'state': state,
            'exports': [str(path) for path, _ in exports], 'seconds': round(time.perf_counter() - started, 3)}


def cli(adapter):
    parser = argparse.ArgumentParser(description=f'Native desktop COM for {adapter.PROG_ID}')
    sub = parser.add_subparsers(dest='command', required=True)
    sub.add_parser('inspect', help='Check Python and pywin32 without launching Office')
    sub.add_parser('probe', help='Connect to the desktop application and report capabilities')
    run = sub.add_parser('run', help='Apply a JSON plan to a new output or the same working draft')
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
