"""FastAPI application: JSON API plus the static browser GUI.

Normally started through ``python -m dlg`` (see README). File locations come from
the per-user :class:`dlg.config.Config` (see ``dlg/config.py``).
"""

from __future__ import annotations

import re
from typing import Literal
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

from fastapi import BackgroundTasks, Depends, FastAPI, HTTPException, Request, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from dlg import get_version
from dlg.config import Config, ConfigError, empty_input_text_problem, font_size_problem, save_config
from dlg.docgen import (
    DocgenError, OutputLockedError, check_template, column_weights, fill_empty_inputs, generate_declared_list,
)
from dlg.filebrowser import BrowseError, BrowseNotFoundError, Listing, list_folder
from dlg.layout import HeaderCell, LayoutError, TableLayout, apply_layout, load_layout
from dlg.models import Column, Item, ItemTable, Session
from dlg.opener import OpenError, Opener, open_path
from dlg.sessions import WINDOWS_RESERVED_NAMES, SessionError, SessionNotFoundError, SessionStore
from dlg.sources.excel import ExcelItemSource, ItemSourceError

STATIC_DIR = Path(__file__).resolve().parent / "static"
#: Characters Windows does not allow in file names (checked on every platform so a
#: name that works on Linux also works on Windows).
_BAD_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

#: File types the browse dialog can list (item database, Word template).
BROWSE_EXTENSIONS = frozenset({".xlsx", ".docx"})

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


class HeadersResponse(BaseModel):
    """Headers of an item database that can be chosen as the ID column.

    Attributes:
        headers: Header texts in sheet order (no ``#`` or ``*`` columns).
    """

    headers: list[str]


class SettingsResponse(BaseModel):
    """Settings editable in the GUI (``GET``/``PUT /api/settings``).

    Attributes:
        items_file: Excel item database.
        id_column: Header of the database column holding the unique item IDs.
        template_file: Word template with the ``{{DECLARED_LIST}}`` placeholder.
        output_dir: Folder for generated Word documents.
        sessions_dir: The user's session library folder.
        global_sessions_dir: Shared session library folder ('' if not set).
        table_layout_file: Workbook with the Word layout sheet ('' = the item database).
        table_layout_sheet: Name of the Word layout sheet.
        table_font_size: Word table font size in pt (0 = template font).
        mark_empty_inputs: Write ``empty_input_text`` into empty input cells in Word.
        empty_input_text: Text for empty input cells (e.g. ``TBD``).
        config_file: Configuration file the settings are saved to ('' if none).
    """

    items_file: str
    id_column: str
    template_file: str
    output_dir: str
    sessions_dir: str
    global_sessions_dir: str
    table_layout_file: str
    table_layout_sheet: str
    table_font_size: float
    mark_empty_inputs: bool
    empty_input_text: str
    config_file: str


class SettingsUpdate(BaseModel):
    """Body of ``PUT /api/settings``; absent fields are left unchanged.

    Attributes:
        items_file: New Excel item database (absolute path to an existing .xlsx).
        id_column: New ID column header (non-empty).
        template_file: New Word template (absolute path to an existing .docx
            containing the placeholder).
        output_dir: New output folder (absolute path to an existing folder).
        sessions_dir: New user session library (absolute path to an existing folder).
        global_sessions_dir: New shared session library (existing folder, or
            '' to unset it).
        table_layout_file: New layout workbook (existing .xlsx, or '' for the
            item database).
        table_layout_sheet: New layout sheet name (non-empty).
        table_font_size: New Word table font size in pt (0 = template font,
            else 4–72 in steps of 0.5).
        mark_empty_inputs: Write ``empty_input_text`` into empty input cells.
        empty_input_text: New text for empty input cells (non-empty).
    """

    items_file: str | None = None
    id_column: str | None = Field(default=None, max_length=200)
    template_file: str | None = None
    output_dir: str | None = None
    sessions_dir: str | None = None
    global_sessions_dir: str | None = None
    table_layout_file: str | None = None
    table_layout_sheet: str | None = Field(default=None, max_length=100)
    table_font_size: float | None = None
    mark_empty_inputs: bool | None = None
    empty_input_text: str | None = None


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
        columns: Shown columns, in database order (then input columns that
            only the layout sheet adds).
        items: All selectable items.
        layout: Where the Word table layout comes from (e.g.
            ``items.xlsx › Word layout``); '' for the automatic table.
        word_columns: Columns of the Word table, in order (blank columns
            marked ``(blank)``).
        warnings: Word table layout problems that do not stop the export.
    """

    source: str
    id_column: str
    columns: list[Column]
    items: list[Item]
    layout: str = ""
    word_columns: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


class LayoutColumnInfo(BaseModel):
    """One Word table column in ``GET /api/layout``.

    Attributes:
        label: Display text (blank columns marked ``(blank)``).
        user_input: True for an input (``*``) column.
        blank: True if the column is exported empty.
    """

    label: str
    user_input: bool = False
    blank: bool = False


class HeaderCellInfo(BaseModel):
    """One (possibly merged) Word header cell in ``GET /api/layout``.

    Attributes:
        row: Zero-based header row.
        col: Zero-based column.
        text: Cell text (may contain line breaks).
        row_span: Header rows covered.
        col_span: Columns covered.
        rotation: '' (horizontal), ``down`` or ``up``.
    """

    row: int
    col: int
    text: str
    row_span: int = 1
    col_span: int = 1
    rotation: str = ""


class LayoutResponse(BaseModel):
    """Response of ``GET /api/layout``: the Word table as it will be exported.

    Attributes:
        source: Origin of the layout ('' for the automatic table).
        header_rows: Number of header rows.
        columns: Word table columns, in order.
        cells: Header cells covering the ``header_rows`` x ``len(columns)`` grid.
        widths: Relative column widths as fractions of the table width (sum 1).
        warnings: Layout problems that do not stop the export.
    """

    source: str
    header_rows: int
    columns: list[LayoutColumnInfo]
    cells: list[HeaderCellInfo]
    widths: list[float]
    warnings: list[str] = Field(default_factory=list)


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


class LibraryEntry(SessionSummary):
    """A session file in a library folder (``GET /api/library``).

    Attributes:
        file: File name, e.g. ``Project A.json``.
        path: Full path, used to load or delete exactly this file.
    """

    file: str
    path: str


class SkippedFile(BaseModel):
    """A ``.json`` file in a library folder that is not a readable session.

    Attributes:
        file: File name.
        error: Why it was skipped.
    """

    file: str
    error: str


class LibraryResponse(BaseModel):
    """Response of ``GET /api/library``.

    Attributes:
        folder: The folder that was listed.
        sessions: Readable sessions, most recently modified first.
        skipped: Files that could not be read as sessions.
    """

    folder: str
    sessions: list[LibraryEntry]
    skipped: list[SkippedFile]


class OpenRequest(BaseModel):
    """Body of ``POST /api/open``.

    Attributes:
        file_name: Name of an exported document in the output folder (as
            returned by the export); not a path.
        folder: True to show the output folder (with the file selected where
            the system supports it) instead of opening the document.
    """

    file_name: str = Field(min_length=1, max_length=260)
    folder: bool = False


class OpenSettingRequest(BaseModel):
    """Body of ``POST /api/open-setting``.

    Attributes:
        setting: Which saved setting to open: the item database, the Word
            template, the table layout workbook or the output folder.
    """

    setting: Literal["items_file", "template_file", "table_layout_file", "output_dir"]


class SaveSessionRequest(BaseModel):
    """Body of ``POST /api/sessions``.

    Attributes:
        name: Session name (an existing session of that name is overwritten).
        configuration_name: Configuration the selection belongs to.
        selected_ids: IDs of the selected items.
        user_values: Values typed by the user, per item ID and column name.
        folder: Folder to save into ("Save to…"); empty for the user's library.
    """

    name: str = Field(min_length=1, max_length=200)
    configuration_name: str = Field(default="", max_length=500)
    selected_ids: list[str]
    user_values: dict[str, dict[str, str]] = Field(default_factory=dict)
    folder: str = ""


class SavedSession(Session):
    """Response of ``POST /api/sessions``: the stored session and where it was written.

    Attributes:
        path: The session file.
    """

    path: str


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
        warnings: Word table layout problems (the document was still written).
        empty_inputs: Number of empty cells per input column of the Word
            table, only columns with at least one, in table order.
        empty_input_text: Text written into those cells; '' when they were
            left empty (setting off).
    """

    path: str
    file_name: str
    count: int
    warnings: list[str] = Field(default_factory=list)
    empty_inputs: dict[str, int] = Field(default_factory=dict)
    empty_input_text: str = ""


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
    opener: Opener = open_path,
) -> FastAPI:
    """Build the FastAPI application.

    Args:
        settings: Configuration; defaults to the built-in :class:`Config`
            defaults (the per-user file is loaded by the launcher).
        config_path: File that settings changed in the GUI are saved to;
            None keeps changes in memory only (used by tests).
        on_shutdown: Called to stop the server when the user clicks Exit;
            None disables the Exit endpoint (no server to stop).
        opener: Opens a document or its folder on the desktop (injectable so
            tests do not start real programs).

    Returns:
        The configured application.
    """
    cfg = settings or Config()
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
        # The layout sheet never holds the data when it lives in the database workbook.
        skip = cfg.table_layout_sheet if cfg.table_layout_file is None else None
        return ExcelItemSource(cfg.items_file, id_column=cfg.id_column, skip_sheet=skip)

    def load_data() -> tuple[ItemTable, TableLayout | None]:
        """Read the item database and the Word table layout (fresh each call).

        Returns:
            The table (with the layout's columns merged in) and the layout,
            or None when there is no layout sheet.

        Raises:
            HTTPException: 500 if the database or the layout sheet cannot be used.
        """
        try:
            table = source().load()
            if cfg.table_layout_file is None:
                layout = load_layout(cfg.items_file, cfg.table_layout_sheet, required=False)
            else:
                layout = load_layout(cfg.table_layout_file, cfg.table_layout_sheet, required=True)
            return apply_layout(table, layout) if layout else (table, None)
        except ItemSourceError as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        except LayoutError as exc:
            raise HTTPException(status_code=500, detail=f"Word table layout: {exc}") from exc

    def store() -> SessionStore:
        """Return the session store for the *current* user library (it can change at runtime)."""
        return SessionStore(cfg.sessions_dir)

    def load_table() -> ItemTable:
        """Read the item database with the layout's columns (fresh each call so file edits show up).

        Returns:
            Columns and items.

        Raises:
            HTTPException: 500 if the database or the layout sheet cannot be used.
        """
        return load_data()[0]

    @app.get("/api/info", response_model=InfoResponse)
    def get_info() -> InfoResponse:
        """Return facts about the running application (shown in the GUI header)."""
        return InfoResponse(version=app.version)

    @app.get("/api/items", response_model=ItemsResponse)
    def get_items() -> ItemsResponse:
        """Return all selectable items."""
        table, layout = load_data()
        return ItemsResponse(
            source=source().name, id_column=table.id_column, columns=table.columns, items=table.items,
            layout=layout.source if layout else "",
            word_columns=layout.column_labels if layout else table.column_names,
            warnings=layout.warnings if layout else [],
        )

    @app.get("/api/layout", response_model=LayoutResponse)
    def get_layout() -> LayoutResponse:
        """Return the Word table header and columns, for checking the layout before exporting.

        Without a layout sheet this describes the automatic table; its widths
        are estimated from all items (the export uses the selected ones).
        """
        table, layout = load_data()
        if layout is None:
            columns = [LayoutColumnInfo(label=c.name, user_input=c.user_input) for c in table.columns]
            cells = [HeaderCell(0, i, c.name) for i, c in enumerate(table.columns)]
            weights = column_weights(table.column_names, table.items)
            header_rows, source, warnings = 1, "", []
        else:
            columns = [LayoutColumnInfo(label=c.label, user_input=c.user_input, blank=c.blank)
                       for c in layout.columns]
            cells, weights = layout.cells, layout.widths
            header_rows, source, warnings = layout.header_rows, layout.source, layout.warnings
        total = sum(weights) or 1.0
        return LayoutResponse(
            source=source, header_rows=header_rows, columns=columns,
            cells=[HeaderCellInfo(row=c.row, col=c.col, text=c.text, row_span=c.row_span, col_span=c.col_span,
                                  rotation=c.rotation)
                   for c in cells],
            widths=[round(w / total, 4) for w in weights], warnings=warnings,
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
            for s in store().list()
        ]

    # ":path" so names containing "/" (sent URL-encoded as %2F) still match the route.
    @app.get("/api/sessions/{name:path}", response_model=Session)
    def get_session(name: str) -> Session:
        """Return one saved session."""
        try:
            return store().load(name)
        except SessionNotFoundError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from exc
        except SessionError as exc:
            raise HTTPException(status_code=422, detail=str(exc)) from exc

    @app.post("/api/sessions", response_model=SavedSession)
    def save_session(req: SaveSessionRequest, request: Request) -> SavedSession:
        """Create or overwrite a session in the user's library or a chosen folder.

        IDs are stored as given (not checked against the database) so a session
        can be saved even if the database is temporarily unavailable; they are
        validated when exporting.
        """
        if req.folder.strip():
            require_local(request)  # writing to a chosen path: like the other path endpoints
            folder = _checked_path(req.folder, "Session folder")
            if not folder.is_dir():
                raise HTTPException(status_code=400, detail=f"Session folder not found: {folder}")
            target = SessionStore(folder)
        else:
            target = store()
        session = Session(
            name=req.name,
            configuration_name=req.configuration_name.strip(),
            source=source().name,
            selected_ids=req.selected_ids,
            user_values=req.user_values,
        )
        try:
            path = target.save(session)
            # Read back: reflects de-duplication and timestamps.
            return SavedSession(**SessionStore.read_file(path).model_dump(), path=str(path))
        except SessionNotFoundError as exc:  # pragma: no cover - just written
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        except SessionError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.delete("/api/sessions/{name:path}", status_code=204)
    def delete_session(name: str) -> Response:
        """Delete a saved session."""
        try:
            store().delete(name)
        except SessionNotFoundError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from exc
        except SessionError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return Response(status_code=204)

    @app.get("/api/library", response_model=LibraryResponse, dependencies=[Depends(require_local)])
    def library(folder: str = "") -> LibraryResponse:
        """List the session files in a library folder.

        Args:
            folder: Folder to list; empty for the user's library (``sessions_dir``),
                which may not exist yet (it is created on the first save).
        """
        if folder.strip():
            path = _checked_path(folder, "Session folder")
            if not path.is_dir():
                raise HTTPException(status_code=404, detail=f"Session folder not found: {path}")
        else:
            path = cfg.sessions_dir
        result = SessionStore(path).scan()
        return LibraryResponse(
            folder=str(path),
            sessions=[
                LibraryEntry(
                    name=s.name, configuration_name=s.configuration_name, modified=s.modified,
                    source=s.source, count=len(s.selected_ids), file=p.name, path=str(p),
                )
                for p, s in result.sessions
            ],
            skipped=[SkippedFile(file=p.name, error=e) for p, e in result.skipped],
        )

    @app.get("/api/library/session", response_model=Session, dependencies=[Depends(require_local)])
    def library_session(path: str) -> Session:
        """Return the session stored in one file (from any folder)."""
        try:
            return SessionStore.read_file(_checked_path(path, "Session file"))
        except SessionNotFoundError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from exc
        except SessionError as exc:
            raise HTTPException(status_code=422, detail=str(exc)) from exc

    @app.delete("/api/library/session", status_code=204, dependencies=[Depends(require_local)])
    def delete_library_session(path: str) -> Response:
        """Delete one session file (from any folder)."""
        try:
            SessionStore.delete_file(_checked_path(path, "Session file"))
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
        table, layout = load_data()
        wanted = set(req.selected_ids)
        unknown = sorted(wanted - {i.id for i in table.items})
        if unknown:
            raise HTTPException(status_code=422, detail=f"Unknown item IDs: {', '.join(unknown)}")
        if target.exists() and not req.overwrite:
            raise HTTPException(status_code=409, detail=f"{name} already exists in {cfg.output_dir}.")
        items = [i for i in table.with_user_values(req.user_values) if i.id in wanted]
        # Counted even with the setting off, so the user hears about missing input either way.
        word_keys = layout.column_keys if layout else table.column_names
        input_names = {c.name for c in table.columns if c.user_input}
        inputs = [k for k in word_keys if k is not None and k in input_names]
        filled, counts = fill_empty_inputs(items, inputs, cfg.empty_input_text)
        if cfg.mark_empty_inputs:
            items = filled
        try:
            generate_declared_list(
                cfg.template_file, table.column_names, items, target, req.configuration_name.strip(), layout,
                cfg.table_font_size,
            )
        except OutputLockedError as exc:
            raise HTTPException(status_code=423, detail=str(exc)) from exc
        except DocgenError as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        return ExportResponse(
            path=str(target), file_name=name, count=len(items), warnings=layout.warnings if layout else [],
            empty_inputs={k: n for k, n in counts.items() if n},
            empty_input_text=cfg.empty_input_text if cfg.mark_empty_inputs else "",
        )

    @app.post("/api/open", status_code=204, dependencies=[Depends(require_local)])
    def open_document(req: OpenRequest) -> Response:
        """Open an exported document in Word, or show it in its folder.

        Only a plain file name is accepted and it must be an existing ``.docx``
        in the output folder, so the browser can never make the server open
        some other file or program.
        """
        name = req.file_name
        if name != Path(name).name or "/" in name or "\\" in name or not name.lower().endswith(".docx"):
            raise HTTPException(status_code=400, detail=f"Not a document name: {name}")
        path = cfg.output_dir / name
        if not path.is_file():
            raise HTTPException(status_code=404, detail=f"{name} is no longer in {cfg.output_dir}.")
        try:
            opener(path, req.folder)
        except OpenError as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        return Response(status_code=204)

    @app.post("/api/open-setting", status_code=204, dependencies=[Depends(require_local)])
    def open_setting(req: OpenSettingRequest) -> Response:
        """Open a file or folder named in the settings in its program (Excel, Word, Explorer).

        Helps while adjusting the layout or template: edit, save, then check
        or export again (both read the files afresh). Only the *saved* path of
        one of four settings can be opened, so the browser can never make the
        server open some other file or program.
        """
        path: Path
        if req.setting == "table_layout_file":
            # Without a separate layout file the layout sheet lives in the item database.
            path = cfg.table_layout_file or cfg.items_file
        else:
            path = getattr(cfg, req.setting)
        exists = path.is_dir() if req.setting == "output_dir" else path.is_file()
        if not exists:
            raise HTTPException(status_code=404, detail=f"Not found: {path}")
        try:
            # folder=False on a directory opens the directory itself on every platform.
            opener(path, False)
        except OpenError as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        return Response(status_code=204)

    def current_settings() -> SettingsResponse:
        """Return the GUI-editable settings."""
        return SettingsResponse(
            items_file=str(cfg.items_file),
            id_column=cfg.id_column,
            template_file=str(cfg.template_file),
            output_dir=str(cfg.output_dir),
            sessions_dir=str(cfg.sessions_dir),
            global_sessions_dir=str(cfg.global_sessions_dir) if cfg.global_sessions_dir else "",
            table_layout_file=str(cfg.table_layout_file) if cfg.table_layout_file else "",
            table_layout_sheet=cfg.table_layout_sheet,
            table_font_size=cfg.table_font_size,
            mark_empty_inputs=cfg.mark_empty_inputs,
            empty_input_text=cfg.empty_input_text,
            config_file=str(config_path) if config_path else "",
        )

    @app.get("/api/settings", response_model=SettingsResponse, dependencies=[Depends(require_local)])
    def get_settings() -> SettingsResponse:
        """Return the GUI-editable settings."""
        return current_settings()

    @app.put("/api/settings", response_model=SettingsResponse, dependencies=[Depends(require_local)])
    def put_settings(req: SettingsUpdate) -> SettingsResponse:
        """Change the GUI-editable settings and save them to the config file.

        Everything is validated before anything changes, so a bad value never
        leaves the settings half-updated.
        """
        items_file, id_column, output_dir = cfg.items_file, cfg.id_column, cfg.output_dir
        template_file = cfg.template_file
        sessions_dir, global_sessions_dir = cfg.sessions_dir, cfg.global_sessions_dir
        table_layout_file, table_layout_sheet = cfg.table_layout_file, cfg.table_layout_sheet
        table_font_size = cfg.table_font_size
        if req.table_font_size is not None:
            problem = font_size_problem(req.table_font_size)
            if problem:
                raise HTTPException(status_code=400, detail=problem)
            table_font_size = float(req.table_font_size)
        mark_empty_inputs, empty_input_text = cfg.mark_empty_inputs, cfg.empty_input_text
        if req.mark_empty_inputs is not None:
            mark_empty_inputs = req.mark_empty_inputs
        if req.empty_input_text is not None:
            empty_input_text = req.empty_input_text.strip()
            problem = empty_input_text_problem(empty_input_text)
            if problem:
                raise HTTPException(status_code=400, detail=problem)
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
        # Only a *changed* template is checked, so an unchanged (e.g. since deleted) one
        # does not block saving the other settings; export reports it then.
        if req.template_file is not None and req.template_file.strip() != str(cfg.template_file):
            template_file = _checked_path(req.template_file, "Word template")
            if template_file.suffix.lower() != ".docx":
                raise HTTPException(status_code=400, detail=f"Word template must be a .docx file: {template_file}")
            try:
                check_template(template_file)
            except DocgenError as exc:
                raise HTTPException(status_code=400, detail=f"Word template: {exc}") from exc
        # Only a *changed* folder must exist: the default output folder is created on first export,
        # and re-sending an unchanged value must not block saving the other settings.
        if req.output_dir is not None and req.output_dir.strip() != str(cfg.output_dir):
            output_dir = _checked_path(req.output_dir, "Output folder")
            if not output_dir.is_dir():
                raise HTTPException(status_code=400, detail=f"Output folder not found: {output_dir}")
        if req.sessions_dir is not None and req.sessions_dir.strip() != str(cfg.sessions_dir):
            sessions_dir = _checked_path(req.sessions_dir, "My sessions folder")
            if not sessions_dir.is_dir():
                raise HTTPException(status_code=400, detail=f"My sessions folder not found: {sessions_dir}")
        current_global = str(cfg.global_sessions_dir) if cfg.global_sessions_dir else ""
        if req.global_sessions_dir is not None and req.global_sessions_dir.strip() != current_global:
            if not req.global_sessions_dir.strip():
                global_sessions_dir = None  # unset
            else:
                global_sessions_dir = _checked_path(req.global_sessions_dir, "Shared library folder")
                if not global_sessions_dir.is_dir():
                    raise HTTPException(
                        status_code=400, detail=f"Shared library folder not found: {global_sessions_dir}")
        current_layout = str(cfg.table_layout_file) if cfg.table_layout_file else ""
        if req.table_layout_file is not None and req.table_layout_file.strip() != current_layout:
            if not req.table_layout_file.strip():
                table_layout_file = None  # back to the sheet inside the item database
            else:
                table_layout_file = _checked_path(req.table_layout_file, "Table layout file")
                if table_layout_file.suffix.lower() != ".xlsx":
                    raise HTTPException(
                        status_code=400, detail=f"Table layout file must be an .xlsx file: {table_layout_file}")
                if not table_layout_file.is_file():
                    raise HTTPException(status_code=400, detail=f"Table layout file not found: {table_layout_file}")
        if req.table_layout_sheet is not None:
            table_layout_sheet = req.table_layout_sheet.strip()
            if not table_layout_sheet:
                raise HTTPException(status_code=400, detail="Layout sheet: enter the name of the sheet.")
        names = ("items_file", "id_column", "template_file", "output_dir", "sessions_dir", "global_sessions_dir",
                 "table_layout_file", "table_layout_sheet", "table_font_size", "mark_empty_inputs",
                 "empty_input_text")
        old = {n: getattr(cfg, n) for n in names}
        new = (items_file, id_column, template_file, output_dir, sessions_dir, global_sessions_dir,
               table_layout_file, table_layout_sheet, table_font_size, mark_empty_inputs, empty_input_text)
        for n, value in zip(names, new):
            setattr(cfg, n, value)
        if config_path is not None:
            try:
                save_config(cfg, config_path)
            except ConfigError as exc:
                for n, value in old.items():  # keep memory and file in step
                    setattr(cfg, n, value)
                raise HTTPException(status_code=500, detail=str(exc)) from exc
        return current_settings()

    @app.get("/api/headers", response_model=HeadersResponse, dependencies=[Depends(require_local)])
    def headers(items_file: str, layout_file: str = "", layout_sheet: str = "") -> HeadersResponse:
        """List the headers of an item database, for the ID column pick list.

        Reads the file named in the Settings field, which may not be saved
        yet, so the list follows the database the user is choosing.

        Args:
            items_file: Path of the item database (.xlsx).
            layout_file: Table layout file field; when set, the layout is not
                in the database, so no sheet is skipped.
            layout_sheet: Layout sheet field (empty: the saved setting); this
                sheet of the database is skipped when finding the data sheet.
        """
        path = Path(items_file.strip()).expanduser()
        if path.suffix.lower() != ".xlsx":
            raise HTTPException(status_code=400, detail="The item database must be an .xlsx file.")
        # Same rule as source(): the layout sheet never holds the data when it is in the database.
        skip = None if layout_file.strip() else (layout_sheet.strip() or cfg.table_layout_sheet)
        source = ExcelItemSource(path, skip_sheet=skip)
        try:
            return HeadersResponse(headers=source.headers())
        except ItemSourceError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from exc

    @app.get("/api/fs", dependencies=[Depends(require_local)])
    def browse(path: str = "", files: bool = True, ext: str = ".xlsx") -> Listing:
        """List a folder for the browse dialog.

        Args:
            path: Folder (or a file inside it); empty for the home folder.
            files: Include files (False when choosing a folder).
            ext: Extension of the files to list (one of ``BROWSE_EXTENSIONS``).
        """
        ext = ext.lower()
        if ext not in BROWSE_EXTENSIONS:
            raise HTTPException(status_code=400, detail=f"Cannot browse for {ext} files.")
        try:
            return list_folder(path or None, (ext,) if files else ())
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

