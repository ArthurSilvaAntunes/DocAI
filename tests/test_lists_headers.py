"""Testes extras: listas, cabeçalho/rodapé, runs e IDs estáveis."""

from __future__ import annotations

from app.models.blocks import ListContent, ParagraphContent, RunModel
from app.models.document import CreateDocumentRequest, HeaderFooterRequest
from app.services.document_service import DocumentService
from app.storage.document_store import DocumentStore


def test_bullet_and_numbered_lists():
    service = DocumentService(DocumentStore())
    info = service.create(CreateDocumentRequest(filename="listas.docx"))
    bullets = service.add_list(info.id, ListContent(list_type="bullet", items=["Alpha", "Beta"], level=0))
    numbers = service.add_list(info.id, ListContent(list_type="numbered", items=["Um", "Dois"], level=0))
    assert bullets[0].startswith("list_")
    detail = service.get_block(info.id, numbers[0])
    assert detail.type.value == "list"
    assert detail.list_type == "numbered"


def test_runs_formatting():
    service = DocumentService(DocumentStore())
    info = service.create(CreateDocumentRequest(filename="runs.docx"))
    pid = service.add_paragraph(
        info.id,
        ParagraphContent(
            runs=[
                RunModel(text="Este texto é "),
                RunModel(text="importante", bold=True),
                RunModel(text="."),
            ]
        ),
    )
    detail = service.get_block(info.id, pid)
    assert detail.text == "Este texto é importante."
    assert any(run.bold for run in (detail.runs or []))


def test_header_footer_page_fields():
    service = DocumentService(DocumentStore())
    info = service.create(CreateDocumentRequest(filename="header.docx", title="Doc"))
    service.set_header(info.id, HeaderFooterRequest(text="Relatório", include_document_name=True))
    service.set_footer(info.id, HeaderFooterRequest(include_page_number=True))
    session = service.store.get(info.id)
    footer_xml = session.document.sections[0].footer._element.xml
    assert "PAGE" in footer_xml
    assert "NUMPAGES" in footer_xml
