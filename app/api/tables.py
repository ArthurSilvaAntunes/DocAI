"""Operações de tabelas e imagens."""

from __future__ import annotations

from typing import Optional

from fastapi import APIRouter, File, Query, UploadFile
from pydantic import BaseModel

from app.container import document_service
from app.models.blocks import ImageContent, TableContent
from app.models.responses import ApiResponse

router = APIRouter(tags=["tables-images"])


class SetCellRequest(BaseModel):
    row: int
    column: int
    text: str
    bold: Optional[bool] = None
    alignment: Optional[str] = None
    shading: Optional[str] = None


class RowRequest(BaseModel):
    index: Optional[int] = None
    values: Optional[list[str]] = None


class MergeRequest(BaseModel):
    start_row: int
    start_col: int
    end_row: int
    end_col: int


@router.post("/documents/{document_id}/tables")
def add_table(document_id: str, payload: TableContent, after_id: Optional[str] = Query(default=None)) -> ApiResponse:
    block_id = document_service.add_table(document_id, payload, after_id=after_id)
    return ApiResponse.ok(document_id=document_id, created_block_id=block_id)


@router.post("/documents/{document_id}/tables/{table_id}/cells")
def set_cell(document_id: str, table_id: str, payload: SetCellRequest) -> ApiResponse:
    document_service.set_cell(
        document_id,
        table_id,
        payload.row,
        payload.column,
        payload.text,
        bold=payload.bold,
        alignment=payload.alignment,
        shading=payload.shading,
    )
    return ApiResponse.ok(document_id=document_id)


@router.post("/documents/{document_id}/tables/{table_id}/rows")
def add_row(document_id: str, table_id: str, payload: RowRequest | None = None) -> ApiResponse:
    values = payload.values if payload else None
    document_service.table_op(document_id, table_id, "add_row", values=values)
    return ApiResponse.ok(document_id=document_id)


@router.delete("/documents/{document_id}/tables/{table_id}/rows/{index}")
def delete_row(document_id: str, table_id: str, index: int) -> ApiResponse:
    document_service.table_op(document_id, table_id, "delete_row", index=index)
    return ApiResponse.ok(document_id=document_id)


@router.post("/documents/{document_id}/tables/{table_id}/columns")
def add_column(document_id: str, table_id: str, payload: RowRequest | None = None) -> ApiResponse:
    values = payload.values if payload else None
    document_service.table_op(document_id, table_id, "add_column", values=values)
    return ApiResponse.ok(document_id=document_id)


@router.delete("/documents/{document_id}/tables/{table_id}/columns/{index}")
def delete_column(document_id: str, table_id: str, index: int) -> ApiResponse:
    document_service.table_op(document_id, table_id, "delete_column", index=index)
    return ApiResponse.ok(document_id=document_id)


@router.post("/documents/{document_id}/tables/{table_id}/merge")
def merge(document_id: str, table_id: str, payload: MergeRequest) -> ApiResponse:
    document_service.table_op(document_id, table_id, "merge_cells", **payload.model_dump())
    return ApiResponse.ok(document_id=document_id)


@router.post("/documents/{document_id}/images")
def add_image(
    document_id: str,
    file: UploadFile = File(...),
    after_id: Optional[str] = Query(default=None),
    width_cm: Optional[float] = Query(default=None),
    height_cm: Optional[float] = Query(default=None),
    caption: Optional[str] = Query(default=None),
    alignment: Optional[str] = Query(default=None),
) -> ApiResponse:
    payload = ImageContent(
        filename=file.filename,
        width_cm=width_cm,
        height_cm=height_cm,
        caption=caption,
        alignment=alignment,
    )
    block_id = document_service.add_image(document_id, file.file.read(), payload, after_id=after_id)
    return ApiResponse.ok(document_id=document_id, created_block_id=block_id)
