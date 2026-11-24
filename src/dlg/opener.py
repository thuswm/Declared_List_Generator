"""Open a file or its folder with the operating system's default program.

Used for the GUI's "Open document" / "Open folder" buttons after an export and
the "Open" buttons in Settings (database, template, layout, output folder).
The server runs on the user's own computer, so the program opens on their
desktop (the endpoint is local-only).
"""

from __future__ import annotations

import os
import subprocess
import sys
from collections.abc import Callable
from pathlib import Path

#: Signature of an opener: (file, reveal_in_folder) -> None.
Opener = Callable[[Path, bool], None]


class OpenError(Exception):
    """Raised when the operating system cannot open the file or folder."""


def open_path(path: Path, folder: bool, platform: str | None = None) -> None:
    """Open a file in its default program, or show it in its folder.

    A directory with ``folder=False`` is opened itself (Explorer, Finder or
    the file manager).

    Windows: ``os.startfile`` (the program registered for ``.docx``, normally
    Word), and Explorer with the file selected for the folder. macOS: ``open``
    and ``open -R``. Elsewhere: ``xdg-open`` (the folder is opened; file
    managers have no common "select this file" option).

    Args:
        path: The file (or a directory, see above).
        folder: True to show the containing folder instead of opening the file.
        platform: Value of ``sys.platform`` (injectable for tests).

    Raises:
        OpenError: If the program cannot be started.
    """
    platform = platform if platform is not None else sys.platform
    try:
        if platform.startswith("win"):
            if folder:
                # One argument "/select,<path>": Explorer parses it itself. Its exit code is
                # meaningless (1 even on success), so it is not checked.
                subprocess.Popen(["explorer", f"/select,{path}"])  # noqa: S603,S607
            else:
                os.startfile(path)  # type: ignore[attr-defined]  # Windows only
        elif platform == "darwin":
            subprocess.Popen(["open", "-R", str(path)] if folder else ["open", str(path)])  # noqa: S603,S607
        else:
            target = path.parent if folder else path
            # Detached and silent: xdg-open may print warnings and should not tie up the server.
            subprocess.Popen(  # noqa: S603,S607
                ["xdg-open", str(target)],
                stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                start_new_session=True,
            )
    except OSError as exc:
        what = "the folder" if folder else path.name
        raise OpenError(f"Cannot open {what}: {exc.strerror or exc}") from exc
