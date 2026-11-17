"""Command-line launcher: ``python -m dlg``.

Starts the local web server and opens the GUI in the default browser.
"""

from __future__ import annotations

import argparse
import sys
import threading
import webbrowser
from collections.abc import Sequence
from pathlib import Path

import uvicorn

from dlg.app import create_app
from dlg.config import ConfigError, load_config


def main(argv: Sequence[str] | None = None) -> None:
    """Start the Declared List Generator.

    The server binds to localhost by default because the app has no
    authentication; only change ``--host`` on a trusted network.

    Args:
        argv: Command-line arguments (defaults to ``sys.argv[1:]``).
    """
    parser = argparse.ArgumentParser(prog="dlg", description="Declared List Generator")
    parser.add_argument("--config", type=Path, help="configuration file (default: per-user config.toml)")
    parser.add_argument("--host", help="interface to bind (default: from config, 127.0.0.1)")
    parser.add_argument("--port", type=int, help="port to listen on (default: from config, 8000)")
    parser.add_argument("--no-browser", action="store_true", help="do not open the browser")
    args = parser.parse_args(argv)

    try:
        config = load_config(args.config)  # creates the file with defaults on first run
    except ConfigError as exc:
        print(f"dlg: {exc}", file=sys.stderr)
        raise SystemExit(2) from exc
    host = args.host or config.host
    port = args.port or config.port

    if not args.no_browser:
        # Delay slightly so the server is accepting connections when the tab opens.
        url = f"http://{'localhost' if host == '127.0.0.1' else host}:{port}/"
        threading.Timer(1.0, webbrowser.open, args=(url,)).start()
    uvicorn.run(create_app(config), host=host, port=port)


if __name__ == "__main__":
    main()
