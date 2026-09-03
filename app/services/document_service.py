"""Fachada principal usada pela API REST e pelo servidor MCP."""

from __future__ import annotations

import logging
import shutil
from pathlib import Path
from typing import Any

from fastapi import UploadFile

from app.config import settings
from app.docx_engine.reader import get_block_detail, get_structure, inspect_document
from app.docx_engine.writer import (
    clone_from_bytes,
    core_title,
    create_document,
    document_bytes,
    is_valid_docx,
    open_document,
    save_document,
)
from app.exceptions import AppError
from app.models.blocks import HeadingContent, ImageContent, ListContent, ParagraphContent, TableContent
from app.models.commands import Command, CommandBatch, FindRequest, HistoryEntry
from app.models.document import CreateDocumentRequest, DocumentInfo, DocumentInspect, DocumentStructure, HeaderFooterRequest, OpenDocumentRequest, SaveDocumentRequest
from app.models.responses import PreviewData, PreviewPage
from app.models.styles import StyleUpsertRequest
from app.services.editing_service import EditingService
from app.services.rendering_service import RenderingService
from app.services.template_service import TemplateService
from app.storage.document_store import DocumentSession, DocumentStore
from app.utils.paths import resolve_existing_file, safe_join, sanitize_filename

logger = logging.getLogger(__name__)


class DocumentService:
    def __init__(
        self,
        store: DocumentStore | None = None,
        editing: EditingService | None = None,
        templates: TemplateService | None = None,
        rendering: RenderingService | None = None,
    ) -> None:
        self.store = store or DocumentStore()
        self.editing = editing or EditingService(self.store)
        self.templates = templates or TemplateService()
        self.rendering = rendering or RenderingService()

    def create(self, payload: CreateDocumentRequest) -> DocumentInfo:
        filename = sanitize_filename(payload.filename)
        if payload.template:
            root = settings.templates_dir if payload.template_source == "templates" else settings.documents_dir
            path = resolve_existing_file(root, payload.template)
            document = open_document(path)
            logger.info("document_create from_template=%s source=%s", path.name, payload.template_source)
        else:
            document = create_document(
                title=payload.title,
                author=payload.author,
                page=payload.page,
                language=payload.language,
            )
            logger.info("document_create filename=%s", filename)
        if payload.title:
            document.core_properties.title = payload.title
        if payload.author:
            document.core_properties.author = payload.author
        if payload.language:
            from app.docx_engine.language import set_document_language

            set_document_language(document, payload.language)
        session = self.store.add(document, filename)
        self.store.record(session, "create_document", after=filename)
        return session.info()

    def open(self, payload: OpenDocumentRequest) -> DocumentInfo:
        root = settings.templates_dir if payload.source == "templates" else settings.documents_dir
        path = resolve_existing_file(root, payload.filename)
        document = open_document(path)
        session = self.store.add(
            document,
            path.name,
            original_path=path,
            original_bytes=path.read_bytes(),
        )
        logger.info("document_open id=%s file=%s", session.id, path.name)
        self.store.record(session, "open_document", after=path.name)
        return session.info()

    def open_upload(self, upload: UploadFile) -> DocumentInfo:
        filename = sanitize_filename(upload.filename or "upload.docx")
        data = upload.file.read()
        document = clone_from_bytes(data)
        dest = safe_join(settings.documents_dir, filename)
        dest.write_bytes(data)
        session = self.store.add(document, filename, original_path=dest, original_bytes=data)
        logger.info("document_open_upload id=%s file=%s", session.id, filename)
        self.store.record(session, "open_document", after=filename)
        return session.info()

    def get(self, document_id: str) -> DocumentInfo:
        return self.store.get(document_id).info()

    def list_open(self) -> list[DocumentInfo]:
        return [session.info() for session in self.store.list()]

    def structure(self, document_id: str) -> DocumentStructure:
        session = self.store.get(document_id)
        return DocumentStructure(
            document_id=session.id,
            filename=session.filename,
            title=core_title(session.document),
            blocks=get_structure(session.document),
        )

    def inspect(self, document_id: str) -> DocumentInspect:
        session = self.store.get(document_id)
        return inspect_document(session.document, document_id=session.id, filename=session.filename)

    def get_block(self, document_id: str, block_id: str):
        session = self.store.get(document_id)
        return get_block_detail(session.document, block_id)

    def save(self, document_id: str, payload: SaveDocumentRequest | None = None) -> dict[str, Any]:
        session = self.store.get(document_id)
        payload = payload or SaveDocumentRequest()
        self.store.persist_working(session)
        target: Path | None = None
        if payload.filename:
            target = safe_join(settings.documents_dir, sanitize_filename(payload.filename))
        elif payload.overwrite_original:
            if session.original_path is None:
                raise AppError("NO_ORIGINAL", "Document has no original path to overwrite.")
            target = session.original_path
        if payload.keep_version:
            self._write_version(session)
        if target is not None:
            save_document(session.document, target)
            if not is_valid_docx(target):
                raise AppError("CORRUPT_DOCX", "Saved file is not a valid DOCX archive.")
            session.filename = target.name
            logger.info("document_save id=%s file=%s", session.id, target.name)
            self.store.record(session, "save", after=str(target.name))
            return {"path": target.name, "overwrite_original": payload.overwrite_original}
        logger.info("document_save_working id=%s", session.id)
        self.store.record(session, "save")
        return {"path": session.working_path.name, "working_copy": True}

    def save_as(self, document_id: str, filename: str) -> dict[str, Any]:
        return self.save(document_id, SaveDocumentRequest(filename=filename, keep_version=True))

    def export_bytes(self, document_id: str) -> tuple[str, bytes]:
        session = self.store.get(document_id)
        self.store.persist_working(session)
        return session.filename, document_bytes(session.document)

    def history(self, document_id: str) -> list[HistoryEntry]:
        return self.store.get(document_id).history

    def add_paragraph(self, document_id: str, payload: ParagraphContent, after_id: str | None = None, before_id: str | None = None) -> str:
        session = self.store.get(document_id)
        block_id = self._mutate(session, "add_paragraph", lambda: self.editing.add_paragraph(session, payload, after_id, before_id), target=after_id, after=payload.text)
        return block_id

    def add_heading(self, document_id: str, payload: HeadingContent, after_id: str | None = None, before_id: str | None = None) -> str:
        session = self.store.get(document_id)
        return self._mutate(session, "add_heading", lambda: self.editing.add_heading(session, payload, after_id, before_id), after=payload.text)

    def update_block(self, document_id: str, block_id: str, payload: dict[str, Any]) -> None:
        session = self.store.get(document_id)
        before = self.get_block(document_id, block_id).text
        self._mutate(session, "update_block", lambda: self.editing.update_block(session, block_id, payload) or True, target=block_id, before=before, after=payload.get("text"))

    def delete_block(self, document_id: str, block_id: str) -> None:
        session = self.store.get(document_id)
        before = self.get_block(document_id, block_id).text
        self._mutate(session, "delete_block", lambda: self.editing.delete_block(session, block_id) or True, target=block_id, before=before)

    def add_table(self, document_id: str, payload: TableContent, after_id: str | None = None) -> str:
        session = self.store.get(document_id)
        return self._mutate(session, "add_table", lambda: self.editing.add_table(session, payload, after_id=after_id))

    def set_cell(self, document_id: str, table_id: str, row: int, column: int, text: str, **kwargs: Any) -> None:
        session = self.store.get(document_id)
        self._mutate(session, "set_cell", lambda: self.editing.set_cell(session, table_id, row, column, text, **kwargs) or True, target=table_id, after=text)

    def table_op(self, document_id: str, table_id: str, action: str, **kwargs: Any) -> None:
        session = self.store.get(document_id)
        self._mutate(session, action, lambda: self.editing.table_op(session, table_id, action, **kwargs) or True, target=table_id)

    def add_image(self, document_id: str, image_bytes: bytes, payload: ImageContent, after_id: str | None = None) -> str:
        session = self.store.get(document_id)
        return self._mutate(session, "add_image", lambda: self.editing.add_image_bytes(session, image_bytes, payload, after_id=after_id))

    def add_list(self, document_id: str, payload: ListContent, after_id: str | None = None) -> list[str]:
        session = self.store.get(document_id)
        return self._mutate(session, "add_list", lambda: self.editing.add_list(session, payload, after_id=after_id))

    def add_page_break(self, document_id: str, after_id: str | None = None) -> str:
        session = self.store.get(document_id)
        return self._mutate(session, "page_break", lambda: self.editing.add_page_break(session, after_id))

    def add_section_break(self, document_id: str, break_type: str = "next_page", after_id: str | None = None) -> str:
        session = self.store.get(document_id)
        return self._mutate(session, "section_break", lambda: self.editing.add_section_break(session, break_type, after_id))

    def set_header(self, document_id: str, payload: HeaderFooterRequest, image_bytes: bytes | None = None) -> None:
        session = self.store.get(document_id)
        self._mutate(session, "set_header", lambda: self.editing.set_header(session, payload, image_bytes) or True)

    def set_footer(self, document_id: str, payload: HeaderFooterRequest, image_bytes: bytes | None = None) -> None:
        session = self.store.get(document_id)
        self._mutate(session, "set_footer", lambda: self.editing.set_footer(session, payload, image_bytes) or True)

    def find(self, document_id: str, payload: FindRequest):
        session = self.store.get(document_id)
        return self.editing.find(session, payload)

    def replace_text(self, document_id: str, search: str, replace: str, block_id: str | None = None) -> int:
        session = self.store.get(document_id)
        return self._mutate(session, "replace_text", lambda: self.editing.replace_text(session, search, replace, block_id), before=search, after=replace)

    def set_variables(self, document_id: str, variables: dict[str, str]) -> int:
        session = self.store.get(document_id)
        return self._mutate(session, "set_variables", lambda: self.templates.set_variables(session.document, variables))

    def list_styles(self, document_id: str):
        return self.editing.list_styles(self.store.get(document_id))

    def get_style(self, document_id: str, name: str):
        return self.editing.get_style(self.store.get(document_id), name)

    def upsert_style(self, document_id: str, payload: StyleUpsertRequest):
        session = self.store.get(document_id)
        return self._mutate(session, "create_style", lambda: self.editing.upsert_style(session, payload))

    def apply_style(self, document_id: str, block_id: str, style_name: str) -> None:
        session = self.store.get(document_id)
        self._mutate(session, "apply_style", lambda: self.editing.apply_style(session, block_id, style_name) or True, target=block_id)

    def execute_command(self, document_id: str, command: Command) -> dict[str, Any]:
        session = self.store.get(document_id)
        logger.info("command document_id=%s action=%s", document_id, command.action)
        return self._mutate(session, command.action, lambda: self.editing.execute_command(session, command), target=command.target_id)

    def execute_commands(self, document_id: str, batch: CommandBatch) -> list[dict[str, Any]]:
        session = self.store.get(document_id)
        snapshot = self.store.snapshot(session)
        results: list[dict[str, Any]] = []
        try:
            for command in batch.commands:
                logger.info("command document_id=%s action=%s", document_id, command.action)
                result = self.editing.execute_command(session, command)
                self.store.record(session, command.action, target=command.target_id)
                results.append(result)
            self.store.persist_working(session)
            return results
        except Exception:
            self.store.restore(session, snapshot)
            self.store.record(session, "batch_rollback", success=False)
            logger.exception("command_batch_failed document_id=%s", document_id)
            raise

    def render(self, document_id: str) -> PreviewData:
        session = self.store.get(document_id)
        self.store.persist_working(session)
        logger.info("render document_id=%s", document_id)
        return self.rendering.render_preview(session.working_path, session.id)

    def preview(self, document_id: str) -> PreviewData:
        out_dir = settings.previews_dir / document_id
        pages = sorted(out_dir.glob("page_*.png"))
        if not pages:
            return self.render(document_id)
        return PreviewData(
            pages=[PreviewPage(page=i, image=f"/previews/{document_id}/{p.name}") for i, p in enumerate(pages, start=1)],
            pdf=f"/previews/{document_id}/preview.pdf" if (out_dir / "preview.pdf").is_file() else None,
        )

    def _mutate(self, session: DocumentSession, action: str, fn, *, target: str | None = None, before: str | None = None, after: str | None = None):
        self.store.snapshot(session)
        try:
            result = fn()
            self.store.persist_working(session)
            self.store.record(session, action, target=target, before=before, after=after)
            return result
        except Exception:
            logger.exception("command_error action=%s document_id=%s", action, session.id)
            raise

    def _write_version(self, session: DocumentSession) -> None:
        history_dir = settings.history_dir / session.id
        history_dir.mkdir(parents=True, exist_ok=True)
        existing = list(history_dir.glob("document_v*.docx"))
        next_n = len(existing) + 1
        dest = history_dir / f"document_v{next_n:03d}.docx"
        shutil.copy2(session.working_path, dest)
