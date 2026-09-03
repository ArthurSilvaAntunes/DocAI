"""Estilos nativos do documento."""

from __future__ import annotations

from fastapi import APIRouter

from app.container import document_service
from app.models.responses import ApiResponse
from app.models.styles import StyleUpsertRequest
from pydantic import BaseModel

router = APIRouter(tags=["styles"])


class ApplyStyleRequest(BaseModel):
    style: str


@router.get("/documents/{document_id}/styles")
def list_styles(document_id: str) -> ApiResponse:
    styles = [item.model_dump() for item in document_service.list_styles(document_id)]
    return ApiResponse.ok(document_id=document_id, data=styles)


@router.get("/documents/{document_id}/styles/{name}")
def get_style(document_id: str, name: str) -> ApiResponse:
    style = document_service.get_style(document_id, name)
    return ApiResponse.ok(document_id=document_id, data=style.model_dump())


@router.post("/documents/{document_id}/styles")
def create_style(document_id: str, payload: StyleUpsertRequest) -> ApiResponse:
    style = document_service.upsert_style(document_id, payload)
    return ApiResponse.ok(document_id=document_id, data=style.model_dump())


@router.patch("/documents/{document_id}/styles")
def update_style(document_id: str, payload: StyleUpsertRequest) -> ApiResponse:
    style = document_service.upsert_style(document_id, payload)
    return ApiResponse.ok(document_id=document_id, data=style.model_dump())


@router.post("/documents/{document_id}/blocks/{block_id}/style")
def apply_style(document_id: str, block_id: str, payload: ApplyStyleRequest) -> ApiResponse:
    document_service.apply_style(document_id, block_id, payload.style)
    return ApiResponse.ok(document_id=document_id)
