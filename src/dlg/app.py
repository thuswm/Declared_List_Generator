"""FastAPI application: JSON API plus the static browser GUI.

Normally started through ``python -m dlg`` (see README). File locations come from
the per-user :class:`dlg.config.Config` (see ``dlg/config.py``).
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, HTTPException, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from dlg.config import Config
from dlg.docgen import DocgenError, render_to_bytes
from dlg.models import Item, Session
from dlg.sessions import SessionError, SessionNotFoundError, SessionStore
from dlg.sources.excel import ExcelItemSource, ItemSourceError

STATIC_DIR = Path(__file__).resolve().parent / "static"
DOCX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


class ItemsResponse(BaseModel):
    """Response of ``GET /api/items``.

    Attributes:
        source: Name of the item database.
        items: All selectable items.
    """

    source: str
    items: list[Item]


class SessionSummary(BaseModel):
    """Short description of a saved session for listings.

    Attributes:
        name: Session name.
        modified: Last modification time.
        source: Item database the selection was made from.
        count: Number of selected items.
    """

    name: str
    modified: datetime
    source: str
    count: int


class SaveSessionRequest(BaseModel):
    """Body of ``POST /api/sessions``.

    Attributes:
        name: Session name (an existing session of that name is overwritten).
        selected_ids: IDs of the selected items.
    """

    name: str = Field(min_length=1, max_length=200)
    selected_ids: list[str]


class ExportRequest(BaseModel):
    """Body of ``POST /api/export``.

    Attributes:
        selected_ids: IDs of the items to put in the declared list.
    """

    selected_ids: list[str]


def create_app(settings: Config | None = None) -> FastAPI:
    """Build the FastAPI application.

    Args:
        settings: Configuration; defaults to the built-in :class:`Config`
            defaults (the per-user file is loaded by the launcher).

    Returns:
        The configured application.
    """
    cfg = settings or Config()
    source = ExcelItemSource(cfg.items_file)
    store = SessionStore(cfg.sessions_dir)
    app = FastAPI(title="Declared List Generator")

    def load_items() -> list[Item]:
        """Read the item database (fresh each call so file edits show up).

        Returns:
            All items.

        Raises:
            HTTPException: 500 if the database is missing or malformed.
        """
        try:
            return source.load_items()
        except ItemSourceError as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc

    @app.get("/api/items", response_model=ItemsResponse)
    def get_items() -> ItemsResponse:
        """Return all selectable items."""
        return ItemsResponse(source=source.name, items=load_items())

    @app.get("/api/sessions", response_model=list[SessionSummary])
    def list_sessions() -> list[SessionSummary]:
        """List saved sessions, newest first."""
        return [
            SessionSummary(name=s.name, modified=s.modified, source=s.source, count=len(s.selected_ids))
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
        session = Session(name=req.name, source=source.name, selected_ids=req.selected_ids)
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

    @app.post("/api/export")
    def export(req: ExportRequest) -> Response:
        """Generate the declared list .docx for the selected items.

        Rows follow database order (not click order) so the output is stable.
        """
        if not req.selected_ids:
            raise HTTPException(status_code=400, detail="No items selected.")
        items = load_items()
        wanted = set(req.selected_ids)
        unknown = sorted(wanted - {i.id for i in items})
        if unknown:
            raise HTTPException(status_code=422, detail=f"Unknown item IDs: {', '.join(unknown)}")
        try:
            data = render_to_bytes(cfg.template_file, [i for i in items if i.id in wanted])
        except DocgenError as exc:
            raise HTTPException(status_code=500, detail=str(exc)) from exc
        return Response(
            content=data,
            media_type=DOCX_MEDIA_TYPE,
            headers={"Content-Disposition": 'attachment; filename="declared_list.docx"'},
        )

    # Mounted last so it never shadows the /api routes.
    if STATIC_DIR.is_dir():
        app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
    return app

