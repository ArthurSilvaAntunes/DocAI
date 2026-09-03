"""Servidor MCP que reutiliza o mesmo DocumentService da API REST."""

from __future__ import annotations

from typing import Any, Optional

from mcp.server.mcpserver import MCPServer

from app.config import ensure_directories, settings
from app.container import document_service
from app.exceptions import AppError
from app.models.blocks import HeadingContent, ImageContent, ListContent, ParagraphContent, TableContent
from app.models.commands import Command, FindRequest
from app.models.document import CreateDocumentRequest, OpenDocumentRequest, PageSetup, SaveDocumentRequest
from app.models.responses import ApiResponse
from app.utils.paths import resolve_existing_file

ensure_directories()
mcp = MCPServer(
    name="ai-docx-editor",
    title="AI DOCX Editor",
    description="Motor profissional de edição de documentos DOCX controlável por IA.",
)


def _ok(**kwargs: Any) -> dict:
    return ApiResponse.ok(**kwargs).model_dump()


def _fail(exc: Exception) -> dict:
    if isinstance(exc, AppError):
        return ApiResponse.fail(exc.code, exc.message).model_dump()
    return ApiResponse.fail("INTERNAL_ERROR", str(exc)).model_dump()


@mcp.tool()
def document_create(
    filename: str = "documento.docx",
    title: Optional[str] = None,
    author: Optional[str] = None,
    template: Optional[str] = None,
    template_source: str = "templates",
    page_size: str = "A4",
    orientation: str = "portrait",
    language: str = "pt-BR",
) -> dict:
    """Cria um novo documento DOCX. Use template + template_source para clonar um arquivo existente."""
    try:
        info = document_service.create(
            CreateDocumentRequest(
                filename=filename,
                title=title,
                author=author,
                template=template,
                template_source=template_source,
                language=language,
                page=PageSetup(size=page_size, orientation=orientation),
            )
        )
        return _ok(document_id=info.id, data=info.model_dump(mode="json"))
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


@mcp.tool()
def document_open(filename: str, source: str = "documents") -> dict:
    """Abre um DOCX existente em /documents ou /templates."""
    try:
        info = document_service.open(OpenDocumentRequest(filename=filename, source=source))
        return _ok(document_id=info.id, data=info.model_dump(mode="json"))
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


@mcp.tool()
def document_get_structure(document_id: str) -> dict:
    """Retorna a estrutura simplificada do documento com IDs de blocos."""
    try:
        structure = document_service.structure(document_id)
        return _ok(document_id=document_id, data=structure.model_dump())
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


@mcp.tool()
def document_inspect(document_id: str) -> dict:
    """Lê um documento já criado: sumário de títulos, placeholders, tabelas, figuras e legendas."""
    try:
        data = document_service.inspect(document_id)
        return _ok(document_id=document_id, data=data.model_dump())
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


@mcp.tool()
def document_get_block(document_id: str, block_id: str) -> dict:
    """Retorna detalhes completos de um único bloco."""
    try:
        detail = document_service.get_block(document_id, block_id)
        return _ok(document_id=document_id, data=detail.model_dump())
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


@mcp.tool()
def document_add_heading(document_id: str, text: str, level: int = 1, after_id: Optional[str] = None) -> dict:
    """Insere um título (Heading 1-4) no documento."""
    try:
        block_id = document_service.add_heading(
            document_id, HeadingContent(text=text, level=level), after_id=after_id
        )
        return _ok(document_id=document_id, created_block_id=block_id)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


@mcp.tool()
def document_add_paragraph(document_id: str, text: str, after_id: Optional[str] = None, style: Optional[str] = None) -> dict:
    """Insere um parágrafo de texto."""
    try:
        block_id = document_service.add_paragraph(
            document_id, ParagraphContent(text=text, style=style), after_id=after_id
        )
        return _ok(document_id=document_id, created_block_id=block_id)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


@mcp.tool()
def document_add_table(
    document_id: str,
    rows: int = 2,
    columns: int = 2,
    data: Optional[list[list[str]]] = None,
    after_id: Optional[str] = None,
) -> dict:
    """Insere uma tabela."""
    try:
        block_id = document_service.add_table(
            document_id, TableContent(rows=rows, columns=columns, data=data), after_id=after_id
        )
        return _ok(document_id=document_id, created_block_id=block_id)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


@mcp.tool()
def document_add_image(
    document_id: str,
    filename: str,
    width_cm: Optional[float] = None,
    height_cm: Optional[float] = None,
    caption: Optional[str] = None,
    after_id: Optional[str] = None,
) -> dict:
    """Insere uma imagem inline a partir de um arquivo em /documents."""
    try:
        path = resolve_existing_file(settings.documents_dir, filename)
        block_id = document_service.add_image(
            document_id,
            path.read_bytes(),
            ImageContent(filename=filename, width_cm=width_cm, height_cm=height_cm, caption=caption),
            after_id=after_id,
        )
        return _ok(document_id=document_id, created_block_id=block_id)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


@mcp.tool()
def document_add_list(
    document_id: str,
    items: list[str],
    list_type: str = "bullet",
    after_id: Optional[str] = None,
) -> dict:
    """Insere uma lista com marcadores ou numerada."""
    try:
        ids = document_service.add_list(
            document_id, ListContent(items=items, list_type=list_type), after_id=after_id
        )
        return _ok(document_id=document_id, created_block_id=ids[0] if ids else None, data={"created_block_ids": ids})
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


@mcp.tool()
def document_update_block(document_id: str, block_id: str, text: Optional[str] = None, style: Optional[str] = None) -> dict:
    """Atualiza um bloco existente."""
    try:
        payload: dict[str, Any] = {}
        if text is not None:
            payload["text"] = text
        if style is not None:
            payload["style"] = style
        document_service.update_block(document_id, block_id, payload)
        return _ok(document_id=document_id)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


@mcp.tool()
def document_delete_block(document_id: str, block_id: str) -> dict:
    """Remove um bloco do documento."""
    try:
        document_service.delete_block(document_id, block_id)
        return _ok(document_id=document_id)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


@mcp.tool()
def document_move_block(document_id: str, target_id: str, after_id: Optional[str] = None, before_id: Optional[str] = None) -> dict:
    """Move um bloco para outra posição."""
    try:
        result = document_service.execute_command(
            document_id,
            Command(action="move", target_id=target_id, after_id=after_id, before_id=before_id),
        )
        return _ok(document_id=document_id, data=result.get("data"))
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


@mcp.tool()
def document_find(document_id: str, text: Optional[str] = None, block_type: Optional[str] = None, style: Optional[str] = None) -> dict:
    """Busca blocos por texto, tipo ou estilo."""
    try:
        matches = document_service.find(document_id, FindRequest(text=text, block_type=block_type, style=style))
        return _ok(document_id=document_id, data={"matches": [m.model_dump() for m in matches]})
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


@mcp.tool()
def document_replace(document_id: str, search: str, replace: str) -> dict:
    """Substitui texto, inclusive quando dividido entre runs."""
    try:
        count = document_service.replace_text(document_id, search, replace)
        return _ok(document_id=document_id, data={"replaced": count})
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


@mcp.tool()
def document_render(document_id: str) -> dict:
    """Renderiza o documento para PDF e imagens de preview."""
    try:
        preview = document_service.render(document_id)
        return _ok(document_id=document_id, data=preview.model_dump())
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


@mcp.tool()
def document_save(document_id: str, filename: Optional[str] = None, overwrite_original: bool = False) -> dict:
    """Salva o documento. Sem filename, grava a cópia de trabalho. Nunca sobrescreve o original sem overwrite_original."""
    try:
        data = document_service.save(
            document_id,
            SaveDocumentRequest(filename=filename, overwrite_original=overwrite_original),
        )
        return _ok(document_id=document_id, data=data)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
