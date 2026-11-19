"""FastAPI application: JSON API plus the static browser GUI.

Normally started through ``python -m dlg`` (see README). File locations come from
the per-user :class:`dlg.config.Config` (see ``dlg/config.py``).
"""

from __future__ import annotations

import re
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

from fastapi import BackgroundTasks, Depends, FastAPI, HTTPException, Request, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from dlg import get_version
from dlg.config import Config, ConfigError, save_config
from dlg.docgen import DocgenError, OutputLockedError, generate_declared_list
from dlg.filebrowser import BrowseError, BrowseNotFoundError, Listing, list_folder
from dlg.models import Column, Item, ItemTable, Session
from dlg.sessions import WINDOWS_RESERVED_NAMES, SessionError, SessionNotFoundError, SessionStore
from dlg.sources.excel import ExcelItemSource, ItemSourceError

STATIC_DIR = Path(__file__).resolve().parent / "static"
#: Characters Windows does not allow in file names (checked on every platform so a
#: name that works on Linux also works on Windows).
_BAD_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

#: Client addresses treated as "this computer" for the local-only endpoints.
LOCAL_CLIENTS = frozenset({"127.0.0.1", "::1", "localhost"})


def require_local(request: Request) -> None:
    """Refuse requests that do not come from this computer.

    The folder browser and settings reveal and change local paths, so they
    stay local even if the server is started with ``--host`` on a network.

    Args:
        request: The incoming request.

    Raises:
        HTTPException: 403 for any other client.
    """
    host = request.client.host if request.client else ""
    if host not in LOCAL_CLIENTS:
        raise HTTPException(status_code=403, detail="Only available on the computer running the tool.")


class SettingsResponse(BaseModel):
    """Settings editable in the GUI (``GET``/``PUT /api/settings``).

    Attributes:
        items_file: Excel item database.
        id_column: Header of the database column holding the unique item IDs.
        output_dir: Folder for generated Word documents.
        config_file: Configuration file the settings are saved to ('' if none).
    """

    items_file: str
    id_column: str
    output_dir: str
    config_file: str


class SettingsUpdate(BaseModel):
    """Body of ``PUT /api/settings``; absent fields are left unchanged.

    Attributes:
        items_file: New Excel item database (absolute path to an existing .xlsx).
        id_column: New ID column header (non-empty).
        output_dir: New output folder (absolute path to an existing folder).
    """

    items_file: str | None = None
    id_column: str | None = Field(default=None, max_length=200)
    output_dir: str | None = None


class InfoResponse(BaseModel):
    """Response of ``GET /api/info``.

    Attributes:
        version: Application version from ``pyproject.toml``.
    """

    version: str


class ItemsResponse(BaseModel):
    """Response of ``GET /api/items``.

    Attributes:
        source: Name of the item database.
        id_column: Name of the column holding the unique item IDs.
        columns: Shown columns, in database order.
        items: All selectable items.
    """

    source: str
    id_column: str
    columns: list[Column]
    items: list[Item]


class SessionSummary(BaseModel):
    """Short description of a saved session for listings.

    Attributes:
        name: Session name.
        configuration_name: Configuration the selection belongs to.
        modified: Last modification time.
        source: Item database the selection was made from.
        count: Number of selected items.
    """

    name: str
    configuration_name: str
    modified: datetime
    source: str
    count: int


class SaveSessionRequest(BaseModel):
    """Body of ``POST /api/sessions``.

    Attributes:
        name: Session name (an existing session of that name is overwritten).
        configuration_name: Configuration the selection belongs to.
        selected_ids: IDs of the selected items.
        user_values: Values typed by the user, per item ID and column name.
    """

    name: str = Field(min_length=1, max_length=200)
    configuration_name: str = Field(default="", max_length=500)
    selected_ids: list[str]
    user_values: dict[str, dict[str, str]] = Field(default_factory=dict)


class ExportRequest(BaseModel):
    """Body of ``POST /api/export``.

    Attributes:
        selected_ids: IDs of the items to put in the declared list.
        user_values: Values typed by the user for the user-input columns,
            per item ID and column name.
        file_name: Name of the Word file in the output folder (``.docx`` optional).
        configuration_name: Text for ``{{CONFIGURATION_NAME}}`` in the template.
        overwrite: Replace an existing file of that name (otherwise 409).
    """

    selected_ids: list[str]
    user_values: dict[str, dict[str, str]] = Field(default_factory=dict)
    file_name: str = Field(default="declared_list", max_length=200)
    configuration_name: str = Field(default="", max_length=500)
    overwrite: bool = False


class ExportResponse(BaseModel):
    """Result of ``POST /api/export``.

    Attributes:
        path: Full path of the written file.
        file_name: Its file name.
        count: Number of items in the declared list.
    """

    path: str
    file_name: str
    count: int


def output_file_name(name: str) -> str:
    """Validate a user-given output file name and add ``.docx``.

    Invalid names are rejected rather than silently changed, so the user
    gets exactly the file name they see confirmed.

    Args:
        name: File name as typed (with or without ``.docx``).

    Returns:
        The file name ending in ``.docx``.

    Raises:
        HTTPException: 400 if the name is empty, contains characters not
            allowed on Windows, or is a reserved Windows device name.
    """
    stem = name.strip()
    if stem.lower().endswith(".docx"):
        stem = stem[: -len(".docx")]
    stem = stem.rstrip(" .")  # Windows silently drops trailing dots/spaces
    if not stem:
        raise HTTPException(status_code=400, detail="Enter a file name for the Word document.")
    if _BAD_FILENAME_CHARS.search(stem):
        raise HTTPException(
            status_code=400, detail='A file name cannot contain any of these characters: < > : " / \\ | ? *'
        )
    if stem.split(".")[0].upper() in WINDOWS_RESERVED_NAMES:
        raise HTTPException(status_code=400, detail=f"'{stem}' is reserved by Windows; choose another file name.")
    return f"{stem}.docx"


def _checked_path(value: str, label: str) -> Path:
    """Turn a user-supplied path into an absolute Path.

    Args:
        value: Path text from the GUI.
        label: What the path is for (used in error messages).

    Returns:
        The path.

    Raises:
        HTTPException: 400 if the path is empty or relative (it would be
            resolved against the server's working folder, which is surprising).
    """
    path = Path(value.strip()).expanduser()
    if not value.strip() or not path.is_absolute():
        raise HTTPException(status_code=400, detail=f"{label}: enter a full path, e.g. chosen with Browse.")
    return path


def create_app(
    settings: Config | None = None,
    config_path: Path | None = None,
    on_shutdown: Callable[[], None] | None = None,
) -> FastAPI:
    """Build the FastAPI application.

    Args:
        settings: Configuration; defaults to the built-in :class:`Config`
            defaults (the per-user file is loaded by the launcher).
        config_path: File that settings changed in the GUI are saved to;
            None keeps changes in memory only (used by tests).
        on_shutdown: Called to stop the server when the user clicks Exit;
            None disables the Exit endpoint (no server to stop).

    Returns:
        The configured application.
    """
    cfg = settings or Config()
    store = SessionStore(cfg.sessions_dir)
    app = FastAPI(title="Declared List Generator", version=get_version())

    @app.middleware("http")
    async def no_cache(request: Request, call_next: Callable) -> Response:
        """Make the browser revalidate every response before using a cached copy.

        Without a ``Cache-Control`` header browsers cache the GUI files
        heuristically, so after an update the old ``app.js`` kept running until
        the user pressed F5 (seen on Windows). ``no-cache`` still allows the
        cache, but only after a cheap conditional request (304 when unchanged,
        thanks to the ETag/Last-Modified that ``StaticFiles`` sends).
        """
        response = await call_next(request)
        response.headers.setdefault("Cache-Control", "no-cache")
        return response

    def source() -> ExcelItemSource:
        """Return the item source for the *current* settings (they can change at runtime)."""
        return ExcelItemSource(cfg.items_file, id_column=cfg.id_column)

    def load_table() -> ItemTable:
        """Read the item database (fresh each call so file edits show up).

        Returns:
            Columns and items.

        Raises:
            HTTPException: 500 if the database is missing or malformed.
        """
        try:
            return source().load()
        except ItemSourceError as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc

    @app.get("/api/info", response_model=InfoResponse)
    def get_info() -> InfoResponse:
        """Return facts about the running application (shown in the GUI header)."""
        return InfoResponse(version=app.version)

    @app.get("/api/items", response_model=ItemsResponse)
    def get_items() -> ItemsResponse:
        """Return all selectable items."""
        table = load_table()
        return ItemsResponse(
            source=source().name, id_column=table.id_column, columns=table.columns, items=table.items
        )

    @app.get("/api/sessions", response_model=list[SessionSummary])
    def list_sessions() -> list[SessionSummary]:
        """List saved sessions, newest first."""
        return [
            SessionSummary(
                name=s.name,
                configuration_name=s.configuration_name,
                modified=s.modified,
                source=s.source,
                count=len(s.selected_ids),
            )
            for s in store.list()
        ]

    # ":path" so names containing "/" (sent URL-encoded as %2F) still match the route.
    @app.get("/api/sessions/{name:path}", response_model=Session)
    def get_session(name: str) -> Session:
        """Return one saved session."""
        try:
            return store.load(name)
        except SessionNotFoundError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from exc
        except SessionError as exc:
            raise HTTPException(status_code=422, detail=str(exc)) from exc

    @app.post("/api/sessions", response_model=Session)
    def save_session(req: SaveSessionRequest) -> Session:
        """Create or overwrite a session.

        IDs are stored as given (not checked against the database) so a session
        can be saved even if the database is temporarily unavailable; they are
        validated when exporting.
        """
        session = Session(
            name=req.name,
            configuration_name=req.configuration_name.strip(),
            source=source().name,
            selected_ids=req.selected_ids,
            user_values=req.user_values,
        )
        try:
            store.save(session)
            return store.load(req.name)  # reflects de-duplication and timestamps
        except SessionNotFoundError as exc:  # pragma: no cover - just written
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        except SessionError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.delete("/api/sessions/{name:path}", status_code=204)
    def delete_session(name: str) -> Response:
        """Delete a saved session."""
        try:
            store.delete(name)
        except SessionNotFoundError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from exc
        except SessionError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return Response(status_code=204)

    @app.post("/api/export", response_model=ExportResponse)
    def export(req: ExportRequest) -> ExportResponse:
        """Write the declared list .docx for the selected items to the output folder.

        Rows follow database order (not click order) so the output is stable.
        An existing file is only replaced when ``overwrite`` is set (409 otherwise),
        so the GUI can ask first.
        """
        if not req.selected_ids:
            raise HTTPException(status_code=400, detail="No items selected.")
        name = output_file_name(req.file_name)
        target = cfg.output_dir / name
        table = load_table()
        wanted = set(req.selected_ids)
        unknown = sorted(wanted - {i.id for i in table.items})
        if unknown:
            raise HTTPException(status_code=422, detail=f"Unknown item IDs: {', '.join(unknown)}")
        if target.exists() and not req.overwrite:
            raise HTTPException(status_code=409, detail=f"{name} already exists in {cfg.output_dir}.")
        items = [i for i in table.with_user_values(req.user_values) if i.id in wanted]
        try:
            generate_declared_list(
                cfg.template_file, table.column_names, items, target, req.configuration_name.strip()
            )
        except OutputLockedError as exc:
            raise HTTPException(status_code=423, detail=str(exc)) from exc
        except DocgenError as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        return ExportResponse(path=str(target), file_name=name, count=len(items))

    def current_settings() -> SettingsResponse:
        """Return the GUI-editable settings."""
        return SettingsResponse(
            items_file=str(cfg.items_file),
            id_column=cfg.id_column,
            output_dir=str(cfg.output_dir),
            config_file=str(config_path) if config_path else "",
        )

    @app.get("/api/settings", response_model=SettingsResponse, dependencies=[Depends(require_local)])
    def get_settings() -> SettingsResponse:
        """Return the item database, ID column and output folder."""
        return current_settings()

    @app.put("/api/settings", response_model=SettingsResponse, dependencies=[Depends(require_local)])
    def put_settings(req: SettingsUpdate) -> SettingsResponse:
        """Change the item database, ID column and/or output folder and save them to the config file.

        Everything is validated before anything changes, so a bad value never
        leaves the settings half-updated.
        """
        items_file, id_column, output_dir = cfg.items_file, cfg.id_column, cfg.output_dir
        if req.items_file is not None:
            items_file = _checked_path(req.items_file, "Item database")
            if items_file.suffix.lower() != ".xlsx":
                raise HTTPException(status_code=400, detail=f"Item database must be an .xlsx file: {items_file}")
            if not items_file.is_file():
                raise HTTPException(status_code=400, detail=f"Item database not found: {items_file}")
        if req.id_column is not None:
            id_column = req.id_column.strip()
            if not id_column:
                raise HTTPException(status_code=400, detail="ID column: enter the header of the column with the item IDs.")
        # Only a *changed* folder must exist: the default output folder is created on first export,
        # and re-sending an unchanged value must not block saving the other settings.
        if req.output_dir is not None and req.output_dir.strip() != str(cfg.output_dir):
            output_dir = _checked_path(req.output_dir, "Output folder")
            if not output_dir.is_dir():
                raise HTTPException(status_code=400, detail=f"Output folder not found: {output_dir}")
        old = (cfg.items_file, cfg.id_column, cfg.output_dir)
        cfg.items_file, cfg.id_column, cfg.output_dir = items_file, id_column, output_dir
        if config_path is not None:
            try:
                save_config(cfg, config_path)
            except ConfigError as exc:
                cfg.items_file, cfg.id_column, cfg.output_dir = old  # keep memory and file in step
                raise HTTPException(status_code=500, detail=str(exc)) from exc
        return current_settings()

    @app.get("/api/fs", dependencies=[Depends(require_local)])
    def browse(path: str = "", files: bool = True) -> Listing:
        """List a folder for the browse dialog.

        Args:
            path: Folder (or a file inside it); empty for the home folder.
            files: Include .xlsx files (False when choosing a folder).
        """
        try:
            return list_folder(path or None, (".xlsx",) if files else ())
        except BrowseNotFoundError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from exc
        except BrowseError as exc:
            raise HTTPException(status_code=403, detail=str(exc)) from exc

    @app.post("/api/shutdown", status_code=202, dependencies=[Depends(require_local)])
    def shutdown(background: BackgroundTasks) -> dict[str, str]:
        """Stop the server (the GUI's Exit button).

        The stop is requested in a background task, which runs after the
        response has been sent, so the browser still gets its answer.
        """
        if on_shutdown is None:
            raise HTTPException(status_code=503, detail="Exit is not available here; stop the server with Ctrl+C.")
        background.add_task(on_shutdown)
        return {"status": "stopping"}

    # Mounted last so it never shadows the /api routes.
    if STATIC_DIR.is_dir():
        app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
    return app

