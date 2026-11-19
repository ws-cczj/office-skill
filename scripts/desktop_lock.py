"""Cooperative Windows desktop lock shared by all copies of this skill."""

from pathlib import Path
import os
import tempfile


class DesktopBusyError(RuntimeError):
    pass


class DesktopLock:
    def __init__(self, path=None):
        self.path = Path(path) if path else Path(tempfile.gettempdir()) / 'ppt-skill-desktop.lock'
        self.stream = None

    def __enter__(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.stream = self.path.open('a+b')
        if not self.path.stat().st_size:
            self.stream.write(b'\0')
            self.stream.flush()
        self.stream.seek(0)
        try:
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(self.stream.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            self.stream.close()
            raise DesktopBusyError('Another cooperating PowerPoint task holds the desktop lock') from exc
        return self

    def __exit__(self, *args):
        self.stream.close()
