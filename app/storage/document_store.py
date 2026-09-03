"""Sessões de documentos abertos, histórico e snapshots para rollback."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from uuid import uuid4

from docx.document import Document as DocxDocument

from app.config import settings
from app.docx_engine import ids as id_mod
from app.docx_engine.reader import ensure_all_ids, get_structure
from app.docx_engine.writer import document_bytes, open_bytes, save_document
from app.exceptions import NotFoundError
from app.models.commands import HistoryEntry
from app.models.document import DocumentInfo


def _now() -> datetime:
    return datetime.now(timezone.utc)


def new_document_id() -> str:
    return f"doc_{uuid4().hex[:8]}"


@dataclass
class DocumentSession:
    id: str
    filename: str
    document: DocxDocument
    original_path: Path | None
    working_path: Path
    registry: id_mod.IdRegistry
    created_at: datetime
    updated_at: datetime
    dirty: bool = False
    history: list[HistoryEntry] = field(default_factory=list)
    snapshots: list[bytes] = field(default_factory=list)
    original_bytes: bytes | None = None

    def info(self) -> DocumentInfo:
        return DocumentInfo(
            id=self.id,
            filename=self.filename,
            title=self.document.core_properties.title or None,
            author=self.document.core_properties.author or None,
            original_filename=self.original_path.name if self.original_path else None,
            dirty=self.dirty,
            created_at=self.created_at,
            updated_at=self.updated_at,
            block_count=len(get_structure(self.document)),
        )


class DocumentStore:
    def __init__(self) -> None:
        self._sessions: dict[str, DocumentSession] = {}

    def add(
        self,
        document: DocxDocument,
        filename: str,
        *,
        original_path: Path | None = None,
        original_bytes: bytes | None = None,
    ) -> DocumentSession:
        doc_id = new_document_id()
        working_dir = settings.workspace_dir / doc_id
        working_dir.mkdir(parents=True, exist_ok=True)
        working_path = working_dir / "working.docx"
        registry = ensure_all_ids(document)
        save_document(document, working_path)
        now = _now()
        session = DocumentSession(
            id=doc_id,
            filename=filename,
            document=document,
            original_path=original_path,
            working_path=working_path,
            registry=registry,
            created_at=now,
            updated_at=now,
            original_bytes=original_bytes,
        )
        self._sessions[doc_id] = session
        return session

    def get(self, document_id: str) -> DocumentSession:
        session = self._sessions.get(document_id)
        if session is None:
            raise NotFoundError("DOCUMENT_NOT_FOUND", f"Document {document_id} was not found.")
        return session

    def list(self) -> list[DocumentSession]:
        return list(self._sessions.values())

    def persist_working(self, session: DocumentSession) -> None:
        save_document(session.document, session.working_path)
        session.updated_at = _now()
        session.dirty = True

    def snapshot(self, session: DocumentSession) -> bytes:
        data = document_bytes(session.document)
        session.snapshots.append(data)
        if len(session.snapshots) > settings.max_history_snapshots:
            session.snapshots.pop(0)
        return data

    def restore(self, session: DocumentSession, data: bytes) -> None:
        session.document = open_bytes(data)
        session.registry = ensure_all_ids(session.document)
        self.persist_working(session)

    def record(
        self,
        session: DocumentSession,
        action: str,
        *,
        target: str | None = None,
        before: str | None = None,
        after: str | None = None,
        success: bool = True,
    ) -> None:
        session.history.append(
            HistoryEntry(
                timestamp=_now().isoformat(),
                action=action,
                target=target,
                before=_clip(before),
                after=_clip(after),
                success=success,
            )
        )


def _clip(value: str | None, limit: int = 240) -> str | None:
    if value is None:
        return None
    if len(value) <= limit:
        return value
    return value[:limit] + "…"
