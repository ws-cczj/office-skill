"""Run a bundled Office recipe from a small parameter file."""
import argparse
from importlib import import_module
import json
from pathlib import Path
import sys


def input_path(value, params_path):
    return str((params_path.parent / value).resolve())


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('app', choices=('ppt', 'visio', 'word', 'excel'))
    parser.add_argument('--params', type=Path, required=True)
    parser.add_argument('--run-dir', type=Path, required=True)
    parser.add_argument('--plan-only', action='store_true', help='Print the expanded plan without starting Office')
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding='utf-8')
    try:
        params = args.params.resolve(strict=True)
        data = json.loads(params.read_text(encoding='utf-8-sig'))
        if not isinstance(data, dict):
            raise ValueError('Recipe parameters must be a JSON object')
        if type(data.get('update', False)) is not bool:
            raise ValueError('update must be a boolean')
        folder = args.run_dir.resolve()
        recipe = import_module(f'skills.{args.app}.scripts.recipe')
        native = recipe.plan(data, folder, params)
        result = native if args.plan_only else recipe.run(native, data, folder, params)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return int(result.get('ok') is False or result.get('status') == 'failed')
    except Exception as exc:
        print(json.dumps({'ok': False, 'error_type': type(exc).__name__, 'error': str(exc)}, ensure_ascii=False))
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
