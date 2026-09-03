"""Endpoints de documentos, comandos, busca e salvamento."""

from __future__ import annotations

from fastapi import APIRouter, File, UploadFile
from fastapi.responses import Response

from app.container import document_service
from app.models.commands import Command, CommandBatch, FindRequest, ReplaceRequest, SetVariablesRequest
from app.models.document import CreateDocumentRequest, HeaderFooterRequest, OpenDocumentRequest, SaveDocumentRequest
from app.models.responses import ApiResponse

router = APIRouter(tags=["documents"])


@router.post("/documents")
def create_document(payload: CreateDocumentRequest) -> ApiResponse:
    info = document_service.create(payload)
    return ApiResponse.ok(document_id=info.id, data=info.model_dump(mode="json"))


@router.post("/documents/open")
def open_document(payload: OpenDocumentRequest) -> ApiResponse:
    info = document_service.open(payload)
    return ApiResponse.ok(document_id=info.id, data=info.model_dump(mode="json"))


@router.post("/documents/upload")
def upload_document(file: UploadFile = File(...)) -> ApiResponse:
    info = document_service.open_upload(file)
    return ApiResponse.ok(document_id=info.id, data=info.model_dump(mode="json"))


@router.get("/documents")
def list_documents() -> ApiResponse:
    items = [item.model_dump(mode="json") for item in document_service.list_open()]
    return ApiResponse.ok(data=items)


@router.get("/documents/{document_id}")
def get_document(document_id: str) -> ApiResponse:
    info = document_service.get(document_id)
    return ApiResponse.ok(document_id=info.id, data=info.model_dump(mode="json"))


@router.get("/documents/{document_id}/inspect")
def inspect_document(document_id: str) -> ApiResponse:
    data = document_service.inspect(document_id)
    return ApiResponse.ok(document_id=document_id, data=data.model_dump())


@router.get("/documents/{document_id}/structure")
def get_structure(document_id: str) -> ApiResponse:
    structure = document_service.structure(document_id)
    return ApiResponse.ok(document_id=document_id, data=structure.model_dump())


@router.post("/documents/{document_id}/find")
def find(document_id: str, payload: FindRequest) -> ApiResponse:
    matches = document_service.find(document_id, payload)
    return ApiResponse.ok(document_id=document_id, data={"matches": [m.model_dump() for m in matches]})


@router.post("/documents/{document_id}/replace")
def replace(document_id: str, payload: ReplaceRequest) -> ApiResponse:
    count = document_service.replace_text(document_id, payload.search, payload.replace, payload.block_id)
    return ApiResponse.ok(document_id=document_id, data={"replaced": count})


@router.post("/documents/{document_id}/variables")
def set_variables(document_id: str, payload: SetVariablesRequest) -> ApiResponse:
    count = document_service.set_variables(document_id, payload.variables)
    return ApiResponse.ok(document_id=document_id, data={"replaced": count})


@router.post("/documents/{document_id}/commands")
def run_commands(document_id: str, payload: dict) -> ApiResponse:
    if isinstance(payload, dict) and "commands" in payload:
        results = document_service.execute_commands(document_id, CommandBatch.model_validate(payload))
        created = [item.get("created_block_id") for item in results if item.get("created_block_id")]
        return ApiResponse.ok(
            document_id=document_id,
            created_block_id=created[-1] if created else None,
            data={"results": results},
        )
    result = document_service.execute_command(document_id, Command.model_validate(payload))
    return ApiResponse.ok(
        document_id=document_id,
        created_block_id=result.get("created_block_id"),
        data=result.get("data"),
    )


@router.post("/documents/{document_id}/save")
def save_document(document_id: str, payload: SaveDocumentRequest | None = None) -> ApiResponse:
    data = document_service.save(document_id, payload)
    return ApiResponse.ok(document_id=document_id, data=data)


@router.post("/documents/{document_id}/save-as")
def save_document_as(document_id: str, payload: SaveDocumentRequest) -> ApiResponse:
    if not payload.filename:
        return ApiResponse.fail("INVALID_FILENAME", "filename is required.", document_id=document_id)
    data = document_service.save_as(document_id, payload.filename)
    return ApiResponse.ok(document_id=document_id, data=data)


@router.get("/documents/{document_id}/download")
def download_document(document_id: str) -> Response:
    filename, data = document_service.export_bytes(document_id)
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.get("/documents/{document_id}/history")
def get_history(document_id: str) -> ApiResponse:
    items = [item.model_dump() for item in document_service.history(document_id)]
    return ApiResponse.ok(document_id=document_id, data=items)


@router.put("/documents/{document_id}/header")
def set_header(document_id: str, payload: HeaderFooterRequest) -> ApiResponse:
    document_service.set_header(document_id, payload)
    return ApiResponse.ok(document_id=document_id)


@router.put("/documents/{document_id}/footer")
def set_footer(document_id: str, payload: HeaderFooterRequest) -> ApiResponse:
    document_service.set_footer(document_id, payload)
    return ApiResponse.ok(document_id=document_id)
