"""Repository entry point for Microsoft Excel."""
from pathlib import Path
import sys

if not __package__:
    sys.path.insert(0, str(Path(__file__).resolve().parents[3]))

from core.office_batch import cli
from skills.excel.scripts import excel_native


if __name__ == '__main__':
    raise SystemExit(cli(excel_native))
