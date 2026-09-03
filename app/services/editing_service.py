"""Operações de edição de alto nível sobre o documento aberto."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

from lxml import etree

from app.docx_engine import ids as id_mod
from app.docx_engine import xml_utils as xu
from app.docx_engine.headers import set_header_footer
from app.docx_engine.images import add_image, load_image_bytes, replace_image
from app.docx_engine.lists import add_list_items
from app.docx_engine.paragraphs import (
    heading_style,
    insert_paragraph,
    require_paragraph,
    update_paragraph,
    wrap_paragraph,
)
from app.docx_engine.reader import classify_element, get_block_detail, resolve_anchor
from app.docx_engine.search import find_matches, paragraph_text, replace_in_document
from app.docx_engine.sections import add_section_break
from app.docx_engine.styles import apply_style, create_or_update_style, get_style, list_styles
from app.docx_engine.fields import insert_caption, insert_index, next_caption_number
from app.docx_engine.language import set_document_language
from app.docx_engine.tables import (
    add_column,
    add_row,
    add_table,
    delete_column,
    delete_row,
    merge_cells,
    normalize_table,
    require_table,
    set_cell_text,
)
from app.exceptions import AppError, NotFoundError
from app.models.blocks import HeadingContent, ImageContent, ListContent, ParagraphContent, TableContent
from app.models.commands import Command, FindRequest
from app.models.document import HeaderFooterRequest
from app.models.styles import StyleUpsertRequest
from app.storage.document_store import DocumentSession, DocumentStore


class EditingService:
    def __init__(self, store: DocumentStore) -> None:
        self.store = store

    def add_paragraph(self, session: DocumentSession, payload: ParagraphContent, after_id: str | None = None, before_id: str | None = None) -> str:
        after_el, before_el = self._anchors(session, after_id, before_id)
        block_id, _ = insert_paragraph(
            session.document,
            session.registry,
            text=payload.text,
            runs=payload.runs,
            properties=payload,
            after_el=after_el,
            before_el=before_el,
            kind="paragraph",
            style=payload.style,
        )
        return block_id

    def add_heading(self, session: DocumentSession, payload: HeadingContent, after_id: str | None = None, before_id: str | None = None) -> str:
        after_el, before_el = self._anchors(session, after_id, before_id)
        style = heading_style(payload.level)
        props = payload.model_dump(exclude_none=True)
        props["style"] = style
        block_id, _ = insert_paragraph(
            session.document,
            session.registry,
            text=payload.text,
            runs=payload.runs,
            properties=props,
            after_el=after_el,
            before_el=before_el,
            kind="heading",
            style=style,
        )
        return block_id

    def update_block(self, session: DocumentSession, block_id: str, payload: dict[str, Any]) -> None:
        detail = get_block_detail(session.document, block_id)
        if detail.type.value in {"paragraph", "heading", "list", "image", "page_break"}:
            paragraph = require_paragraph(session.document, block_id)
            if detail.type.value == "heading" and payload.get("level"):
                payload = {**payload, "style": heading_style(int(payload["level"]))}
            update_paragraph(
                paragraph,
                text=payload.get("text"),
                runs=payload.get("runs"),
                properties=payload,
            )
            return
        if detail.type.value == "table" and "data" in payload:
            table = require_table(session.document, block_id)
            for r, row in enumerate(payload["data"]):
                for c, value in enumerate(row):
                    if r < len(table.rows) and c < len(table.columns):
                        set_cell_text(table.cell(r, c), str(value))
            return
        raise AppError("UNSUPPORTED_UPDATE", f"Block {block_id} cannot be updated with the provided fields.")

    def delete_block(self, session: DocumentSession, block_id: str) -> None:
        element = resolve_anchor(session.document, block_id)
        id_mod.remove_id(session.document, block_id)
        parent = element.getparent()
        if parent is None:
            raise AppError("DELETE_FAILED", f"Could not delete block {block_id}.")
        parent.remove(element)

    def move_block(self, session: DocumentSession, target_id: str, after_id: str | None = None, before_id: str | None = None) -> None:
        target = resolve_anchor(session.document, target_id)
        after_el, before_el = self._anchors(session, after_id, before_id)
        if after_el is None and before_el is None:
            raise AppError("INVALID_MOVE", "after_id or before_id is required.")
        xu.relocate_cluster(target, after=after_el, before=before_el)

    def duplicate_block(self, session: DocumentSession, target_id: str) -> str:
        element = resolve_anchor(session.document, target_id)
        clone = deepcopy(element)
        for bookmark in list(clone.iter(xu.w("bookmarkStart"))) + list(clone.iter(xu.w("bookmarkEnd"))):
            parent = bookmark.getparent()
            if parent is not None:
                parent.remove(bookmark)
        element.addnext(clone)
        from app.docx_engine.reader import classify_element, kind_for_allocate

        kind = kind_for_allocate(classify_element(clone, session.document))
        return id_mod.ensure_block_id(clone, kind, session.registry, session.document.element)

    def add_table(self, session: DocumentSession, payload: TableContent, after_id: str | None = None, before_id: str | None = None) -> str:
        after_el, before_el = self._anchors(session, after_id, before_id)
        block_id, _ = add_table(
            session.document,
            session.registry,
            payload.rows,
            payload.columns,
            data=payload.data,
            style=payload.style,
            after_el=after_el,
            before_el=before_el,
            width_cm=payload.width_cm,
        )
        return block_id

    def set_cell(self, session: DocumentSession, table_id: str, row: int, column: int, text: str, bold: bool | None = None, alignment: str | None = None, shading: str | None = None) -> None:
        table = require_table(session.document, table_id)
        try:
            cell = table.cell(row, column)
        except IndexError as exc:
            raise AppError("CELL_NOT_FOUND", f"Cell ({row}, {column}) was not found.") from exc
        set_cell_text(cell, text, bold=bold, alignment=alignment, shading=shading)

    def table_op(self, session: DocumentSession, table_id: str, action: str, **kwargs: Any) -> None:
        table = require_table(session.document, table_id)
        if action == "add_row":
            add_row(table, kwargs.get("values"))
        elif action == "delete_row":
            delete_row(table, int(kwargs["index"]))
        elif action == "add_column":
            add_column(table, kwargs.get("values"))
        elif action == "delete_column":
            delete_column(table, int(kwargs["index"]))
        elif action == "merge_cells":
            merge_cells(
                table,
                int(kwargs["start_row"]),
                int(kwargs["start_col"]),
                int(kwargs["end_row"]),
                int(kwargs["end_col"]),
            )
        else:
            raise AppError("UNKNOWN_ACTION", f"Unknown table action '{action}'.")

    def add_image_from_path(
        self,
        session: DocumentSession,
        image_path: Path,
        payload: ImageContent,
        after_id: str | None = None,
        before_id: str | None = None,
    ) -> str:
        after_el, before_el = self._anchors(session, after_id, before_id)
        return add_image(
            session.document,
            session.registry,
            load_image_bytes(image_path),
            filename=image_path.name,
            width_cm=payload.width_cm,
            height_cm=payload.height_cm,
            alignment=payload.alignment.value if payload.alignment else None,
            caption=payload.caption,
            after_el=after_el,
            before_el=before_el,
        )

    def add_image_bytes(self, session: DocumentSession, data: bytes, payload: ImageContent, after_id: str | None = None, before_id: str | None = None) -> str:
        after_el, before_el = self._anchors(session, after_id, before_id)
        return add_image(
            session.document,
            session.registry,
            data,
            filename=payload.filename or "image.png",
            width_cm=payload.width_cm,
            height_cm=payload.height_cm,
            alignment=payload.alignment.value if payload.alignment else None,
            caption=payload.caption,
            after_el=after_el,
            before_el=before_el,
        )

    def replace_image_bytes(self, session: DocumentSession, block_id: str, data: bytes, payload: ImageContent) -> None:
        replace_image(session.document, block_id, data, width_cm=payload.width_cm, height_cm=payload.height_cm)

    def add_list(self, session: DocumentSession, payload: ListContent, after_id: str | None = None) -> list[str]:
        after_el, _ = self._anchors(session, after_id, None)
        items = payload.items or ([payload.text] if payload.text else [])
        return add_list_items(
            session.document,
            session.registry,
            items,
            list_type=payload.list_type,
            level=payload.level,
            after_el=after_el,
        )

    def add_page_break(self, session: DocumentSession, after_id: str | None = None) -> str:
        after_el, _ = self._anchors(session, after_id, None)
        from app.docx_engine.paragraphs import add_page_break_paragraph

        return add_page_break_paragraph(session.document, session.registry, after_el=after_el)

    def add_section_break(self, session: DocumentSession, break_type: str = "next_page", after_id: str | None = None) -> str:
        after_el, _ = self._anchors(session, after_id, None)
        return add_section_break(session.document, session.registry, break_type=break_type, after_el=after_el)

    def set_header(self, session: DocumentSession, payload: HeaderFooterRequest, image_bytes: bytes | None = None) -> None:
        set_header_footer(session.document, payload, kind="header", document_name=session.filename, image_bytes=image_bytes)

    def set_footer(self, session: DocumentSession, payload: HeaderFooterRequest, image_bytes: bytes | None = None) -> None:
        set_header_footer(session.document, payload, kind="footer", document_name=session.filename, image_bytes=image_bytes)

    def find(self, session: DocumentSession, payload: FindRequest):
        return find_matches(
            session.document,
            text=payload.text,
            block_type=payload.block_type,
            style=payload.style,
            block_id=payload.block_id,
        )

    def replace_text(self, session: DocumentSession, search: str, replace: str, block_id: str | None = None) -> int:
        return replace_in_document(session.document, search, replace, block_id=block_id)

    def list_styles(self, session: DocumentSession):
        return list_styles(session.document)

    def get_style(self, session: DocumentSession, name: str):
        return get_style(session.document, name)

    def upsert_style(self, session: DocumentSession, payload: StyleUpsertRequest):
        return create_or_update_style(session.document, payload)

    def apply_style(self, session: DocumentSession, block_id: str, style_name: str) -> None:
        element = resolve_anchor(session.document, block_id)
        if element.tag == xu.w("tbl"):
            from app.docx_engine.tables import wrap_table

            apply_style(session.document, wrap_table(element, session.document), style_name)
            return
        apply_style(session.document, wrap_paragraph(element, session.document), style_name)

    def execute_command(self, session: DocumentSession, command: Command) -> dict[str, Any]:
        action = command.action
        extra = command.model_extra or {}
        content = command.content or {}
        merged = {**extra, **content}
        if command.text is not None:
            merged.setdefault("text", command.text)
        if command.level is not None:
            merged.setdefault("level", command.level)
        if command.style is not None:
            merged.setdefault("style", command.style)

        created_id = None
        data: dict[str, Any] = {}

        if action in {"add_paragraph", "insert_paragraph", "insert_paragraph_after"}:
            payload = ParagraphContent.model_validate(merged)
            created_id = self.add_paragraph(session, payload, after_id=command.after_id, before_id=command.before_id)
        elif action == "insert_paragraph_before":
            payload = ParagraphContent.model_validate(merged)
            created_id = self.add_paragraph(session, payload, before_id=command.before_id or command.target_id)
        elif action in {"add_heading", "insert_heading"}:
            payload = HeadingContent.model_validate(merged)
            created_id = self.add_heading(session, payload, after_id=command.after_id, before_id=command.before_id)
        elif action in {"update_paragraph", "update_heading", "update_block"}:
            if not command.target_id:
                raise AppError("INVALID_COMMAND", "target_id is required.")
            self.update_block(session, command.target_id, merged)
        elif action == "format":
            if not command.target_id:
                raise AppError("INVALID_COMMAND", "target_id is required.")
            self.update_block(session, command.target_id, command.properties or merged)
        elif action in {"delete", "delete_paragraph", "delete_table", "delete_image", "delete_block"}:
            if not command.target_id:
                raise AppError("INVALID_COMMAND", "target_id is required.")
            self.delete_block(session, command.target_id)
        elif action == "move":
            if not command.target_id:
                raise AppError("INVALID_COMMAND", "target_id is required.")
            self.move_block(session, command.target_id, after_id=command.after_id, before_id=command.before_id)
        elif action in {"insert_before", "insert_after"}:
            payload = ParagraphContent.model_validate(merged)
            if action == "insert_before":
                created_id = self.add_paragraph(session, payload, before_id=command.target_id or command.before_id)
            else:
                created_id = self.add_paragraph(session, payload, after_id=command.target_id or command.after_id)
        elif action == "duplicate":
            if not command.target_id:
                raise AppError("INVALID_COMMAND", "target_id is required.")
            created_id = self.duplicate_block(session, command.target_id)
        elif action == "add_table":
            payload = TableContent.model_validate(merged)
            created_id = self.add_table(session, payload, after_id=command.after_id, before_id=command.before_id)
        elif action == "set_cell":
            table_id = merged.get("table_id") or command.target_id
            self.set_cell(
                session,
                table_id,
                int(merged.get("row", 0)),
                int(merged.get("column", merged.get("col", 0))),
                str(merged.get("text", "")),
                bold=merged.get("bold"),
                alignment=merged.get("alignment"),
                shading=merged.get("shading"),
            )
        elif action in {"add_row", "delete_row", "add_column", "delete_column", "merge_cells"}:
            table_id = merged.get("table_id") or command.target_id
            self.table_op(session, table_id, action, **merged)
        elif action in {"bullet_list", "numbered_list", "add_list"}:
            list_type = "numbered" if action == "numbered_list" else merged.get("list_type", "bullet")
            payload = ListContent.model_validate({**merged, "list_type": list_type})
            ids = self.add_list(session, payload, after_id=command.after_id)
            created_id = ids[0] if ids else None
            data["created_block_ids"] = ids
        elif action == "page_break":
            created_id = self.add_page_break(session, after_id=command.after_id)
        elif action == "section_break":
            created_id = self.add_section_break(session, merged.get("break_type", "next_page"), after_id=command.after_id)
        elif action == "replace_text":
            search = command.search or merged.get("search")
            replace = command.replace if command.replace is not None else merged.get("replace", "")
            if not search:
                raise AppError("INVALID_COMMAND", "search is required.")
            data["replaced"] = self.replace_text(session, search, replace, block_id=command.target_id)
        elif action == "set_variables":
            variables = command.variables or merged.get("variables") or {}
            from app.services.template_service import TemplateService

            data["replaced"] = TemplateService().set_variables(session.document, variables)
        elif action == "set_header":
            self.set_header(session, HeaderFooterRequest.model_validate(merged))
        elif action == "set_footer":
            self.set_footer(session, HeaderFooterRequest.model_validate(merged))
        elif action == "apply_style":
            if not command.target_id or not (command.style or merged.get("style")):
                raise AppError("INVALID_COMMAND", "target_id and style are required.")
            self.apply_style(session, command.target_id, command.style or merged["style"])
        elif action == "create_style":
            data["style"] = self.upsert_style(session, StyleUpsertRequest.model_validate(merged)).model_dump()
        elif action in {"insert_toc", "insert_index", "insert_summary"}:
            kind = merged.get("kind", "toc")
            if action == "insert_toc":
                kind = "toc"
            result = insert_index(
                session.document,
                session.registry,
                kind=kind,
                after_el=self._anchors(session, command.after_id, command.before_id)[0],
                title=merged.get("title") or command.text,
                levels=str(merged.get("levels", "1-3")),
            )
            created_id = result["field_id"]
            data = result
        elif action in {"insert_list_of_figures", "insert_tof"}:
            result = insert_index(
                session.document,
                session.registry,
                kind="figures",
                after_el=self._anchors(session, command.after_id, command.before_id)[0],
                title=merged.get("title") or command.text,
            )
            created_id = result["field_id"]
            data = result
        elif action in {"insert_list_of_tables", "insert_tot"}:
            result = insert_index(
                session.document,
                session.registry,
                kind="tables",
                after_el=self._anchors(session, command.after_id, command.before_id)[0],
                title=merged.get("title") or command.text,
            )
            created_id = result["field_id"]
            data = result
        elif action in {"insert_list_of_quadros", "insert_toq"}:
            result = insert_index(
                session.document,
                session.registry,
                kind="quadros",
                after_el=self._anchors(session, command.after_id, command.before_id)[0],
                title=merged.get("title") or command.text,
            )
            created_id = result["field_id"]
            data = result
        elif action == "insert_caption":
            kind = str(merged.get("kind") or merged.get("label") or "figura")
            after_el, _ = self._anchors(session, command.after_id, command.before_id)
            created_id = insert_caption(
                session.document,
                session.registry,
                text=command.text or merged.get("text") or "",
                kind=kind,
                after_el=after_el,
                number=int(merged.get("number") or next_caption_number(session.document, kind)),
            )
        elif action in {"set_language", "set_proofing_language"}:
            lang = merged.get("language") or command.text or "pt-BR"
            set_document_language(session.document, lang)
            data["language"] = lang
        elif action in {"normalize_tables", "fix_tables"}:
            table_id = merged.get("table_id") or command.target_id
            width_cm = merged.get("width_cm")
            if table_id:
                normalize_table(require_table(session.document, table_id), width_cm=width_cm)
                data["normalized"] = [table_id]
            else:
                ids = []
                for table in session.document.tables:
                    block_id = id_mod.get_block_id(table._tbl, session.document.element) or ""
                    normalize_table(table, width_cm=width_cm)
                    ids.append(block_id)
                data["normalized"] = ids
        else:
            raise AppError("UNKNOWN_ACTION", f"Unknown action '{action}'.")

        session.registry = id_mod.scan_registry(session.document)
        return {"created_block_id": created_id, "data": data or None}

    def _anchors(self, session: DocumentSession, after_id: str | None, before_id: str | None) -> tuple[etree._Element | None, etree._Element | None]:
        after_el = resolve_anchor(session.document, after_id) if after_id else None
        before_el = resolve_anchor(session.document, before_id) if before_id else None
        return after_el, before_el
