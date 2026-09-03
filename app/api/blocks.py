"""CRUD de blocos, parágrafos, títulos, listas e quebras."""

from __future__ import annotations

from typing import Optional

from fastapi import APIRouter, Query

from app.container import document_service
from app.models.blocks import HeadingContent, ListContent, ParagraphContent
from app.models.responses import ApiResponse

router = APIRouter(tags=["blocks"])


@router.get("/documents/{document_id}/blocks/{block_id}")
def get_block(document_id: str, block_id: str) -> ApiResponse:
    detail = document_service.get_block(document_id, block_id)
    return ApiResponse.ok(document_id=document_id, data=detail.model_dump())


@router.patch("/documents/{document_id}/blocks/{block_id}")
def patch_block(document_id: str, block_id: str, payload: dict) -> ApiResponse:
    document_service.update_block(document_id, block_id, payload)
    return ApiResponse.ok(document_id=document_id)


@router.delete("/documents/{document_id}/blocks/{block_id}")
def delete_block(document_id: str, block_id: str) -> ApiResponse:
    document_service.delete_block(document_id, block_id)
    return ApiResponse.ok(document_id=document_id)


@router.post("/documents/{document_id}/paragraphs")
def add_paragraph(
    document_id: str,
    payload: ParagraphContent,
    after_id: Optional[str] = Query(default=None),
    before_id: Optional[str] = Query(default=None),
) -> ApiResponse:
    block_id = document_service.add_paragraph(document_id, payload, after_id=after_id, before_id=before_id)
    return ApiResponse.ok(document_id=document_id, created_block_id=block_id)


@router.post("/documents/{document_id}/headings")
def add_heading(
    document_id: str,
    payload: HeadingContent,
    after_id: Optional[str] = Query(default=None),
    before_id: Optional[str] = Query(default=None),
) -> ApiResponse:
    block_id = document_service.add_heading(document_id, payload, after_id=after_id, before_id=before_id)
    return ApiResponse.ok(document_id=document_id, created_block_id=block_id)


@router.post("/documents/{document_id}/lists")
def add_list(document_id: str, payload: ListContent, after_id: Optional[str] = Query(default=None)) -> ApiResponse:
    ids = document_service.add_list(document_id, payload, after_id=after_id)
    return ApiResponse.ok(document_id=document_id, created_block_id=ids[0] if ids else None, data={"created_block_ids": ids})


@router.post("/documents/{document_id}/page-breaks")
def add_page_break(document_id: str, after_id: Optional[str] = Query(default=None)) -> ApiResponse:
    block_id = document_service.add_page_break(document_id, after_id=after_id)
    return ApiResponse.ok(document_id=document_id, created_block_id=block_id)


@router.post("/documents/{document_id}/section-breaks")
def add_section_break(
    document_id: str,
    break_type: str = Query(default="next_page"),
    after_id: Optional[str] = Query(default=None),
) -> ApiResponse:
    block_id = document_service.add_section_break(document_id, break_type=break_type, after_id=after_id)
    return ApiResponse.ok(document_id=document_id, created_block_id=block_id)
