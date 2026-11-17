"""Command-line launcher: ``python -m dlg``.

Starts the local web server and opens the GUI in the default browser.
"""

from __future__ import annotations

import argparse
import threading
import webbrowser
from collections.abc import Sequence

import uvicorn


def main(argv: Sequence[str] | None = None) -> None:
    """Start the Declared List Generator.

    The server binds to localhost by default because the app has no
    authentication; only change ``--host`` on a trusted network.

    Args:
        argv: Command-line arguments (defaults to ``sys.argv[1:]``).
    """
    parser = argparse.ArgumentParser(prog="dlg", description="Declared List Generator")
    parser.add_argument("--host", default="127.0.0.1", help="interface to bind (default: 127.0.0.1)")
    parser.add_argument("--port", type=int, default=8000, help="port to listen on (default: 8000)")
    parser.add_argument("--no-browser", action="store_true", help="do not open the browser")
    args = parser.parse_args(argv)

    if not args.no_browser:
        # Delay slightly so the server is accepting connections when the tab opens.
        url = f"http://{'localhost' if args.host == '127.0.0.1' else args.host}:{args.port}/"
        threading.Timer(1.0, webbrowser.open, args=(url,)).start()
    uvicorn.run("dlg.app:app", host=args.host, port=args.port)


if __name__ == "__main__":
    main()
