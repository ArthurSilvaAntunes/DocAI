"""Renderização e preview por página."""

from __future__ import annotations

from fastapi import APIRouter

from app.container import document_service
from app.models.responses import ApiResponse

router = APIRouter(tags=["rendering"])


@router.post("/documents/{document_id}/render")
def render_document(document_id: str) -> ApiResponse:
    preview = document_service.render(document_id)
    return ApiResponse.ok(document_id=document_id, data=preview.model_dump())


@router.get("/documents/{document_id}/preview")
def get_preview(document_id: str) -> ApiResponse:
    preview = document_service.preview(document_id)
    return ApiResponse.ok(document_id=document_id, data=preview.model_dump())
