"""Testes da API REST."""

from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

from app.services.rendering_service import find_soffice


def test_health(client: TestClient):
    response = client.get("/health")
    assert response.status_code == 200
    assert response.json()["success"] is True


def test_create_structure_edit_save_flow(client: TestClient):
    created = client.post("/documents", json={"filename": "relatorio.docx", "title": "Relatório", "author": "IA"})
    assert created.status_code == 200
    body = created.json()
    assert body["success"] is True
    doc_id = body["document_id"]

    heading = client.post(f"/documents/{doc_id}/headings", json={"text": "Introdução", "level": 1})
    heading_id = heading.json()["created_block_id"]

    para = client.post(
        f"/documents/{doc_id}/paragraphs",
        params={"after_id": heading_id},
        json={"text": "Texto introdutório {{CLIENTE}}."},
    )
    para_id = para.json()["created_block_id"]

    client.patch(f"/documents/{doc_id}/blocks/{para_id}", json={"text": "Texto introdutório atualizado {{CLIENTE}}."})

    h2 = client.post(f"/documents/{doc_id}/headings", json={"text": "Resultados", "level": 1})
    results_id = h2.json()["created_block_id"]

    table = client.post(
        f"/documents/{doc_id}/tables",
        params={"after_id": results_id},
        json={"rows": 2, "columns": 2, "data": [["Item", "Valor"], ["A", "10"]]},
    )
    table_id = table.json()["created_block_id"]
    client.post(
        f"/documents/{doc_id}/tables/{table_id}/cells",
        json={"row": 1, "column": 1, "text": "R$ 1.500,00"},
    )

    riscos = client.post(
        "/documents/{}/commands".format(doc_id),
        json={"action": "insert_heading", "after_id": table_id, "text": "Riscos", "level": 1},
    )
    riscos_id = riscos.json()["created_block_id"]
    client.post(
        f"/documents/{doc_id}/commands",
        json={"action": "insert_paragraph", "after_id": riscos_id, "text": "Os principais riscos identificados foram..."},
    )

    client.post(
        f"/documents/{doc_id}/commands",
        json={
            "commands": [
                {"action": "replace_text", "search": "{{CLIENTE}}", "replace": "Empresa ABC"},
                {"action": "add_paragraph", "text": "Documento gerado automaticamente."},
            ]
        },
    )

    structure = client.get(f"/documents/{doc_id}/structure").json()
    assert structure["success"] is True
    blocks = structure["data"]["blocks"]
    assert any(b["type"] == "heading" and b["text"] == "Riscos" for b in blocks)
    assert any(b["id"] == table_id for b in blocks)

    found = client.post(f"/documents/{doc_id}/find", json={"text": "Conclusão"})
    assert found.json()["success"] is True

    saved = client.post(f"/documents/{doc_id}/save", json={"filename": "relatorio_final.docx"})
    assert saved.json()["success"] is True
    assert saved.json()["data"]["path"] == "relatorio_final.docx"

    opened = client.post("/documents/open", json={"filename": "relatorio_final.docx"})
    assert opened.json()["success"] is True
    re_id = opened.json()["document_id"]
    re_struct = client.get(f"/documents/{re_id}/structure").json()["data"]["blocks"]
    texts = " ".join(b.get("text") or "" for b in re_struct)
    assert "Empresa ABC" in texts
    assert "Riscos" in texts


def test_block_not_found(client: TestClient):
    created = client.post("/documents", json={"filename": "x.docx"})
    doc_id = created.json()["document_id"]
    response = client.get(f"/documents/{doc_id}/blocks/paragraph_999")
    assert response.status_code == 404
    body = response.json()
    assert body["success"] is False
    assert body["error"]["code"] == "BLOCK_NOT_FOUND"


def test_path_traversal_rejected(client: TestClient):
    response = client.post("/documents/open", json={"filename": "../secret.docx"})
    assert response.status_code == 400
    assert response.json()["error"]["code"] in {"PATH_TRAVERSAL", "INVALID_FILENAME"}


@pytest.mark.skipif(find_soffice() is None, reason="LibreOffice não está instalado")
def test_render_pdf(client: TestClient):
    created = client.post("/documents", json={"filename": "render.docx", "title": "Preview"})
    doc_id = created.json()["document_id"]
    client.post(f"/documents/{doc_id}/headings", json={"text": "Capa", "level": 1})
    rendered = client.post(f"/documents/{doc_id}/render")
    assert rendered.status_code == 200
    body = rendered.json()
    assert body["success"] is True
    assert body["data"]["pages"]
    preview = client.get(f"/documents/{doc_id}/preview")
    assert preview.json()["data"]["pages"][0]["image"].endswith(".png")
