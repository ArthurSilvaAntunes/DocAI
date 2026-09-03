"""Respostas curtas e previsíveis para consumo por IA."""

from __future__ import annotations

from typing import Any, Optional

from pydantic import BaseModel, Field


class ErrorDetail(BaseModel):
    code: str
    message: str


class ApiResponse(BaseModel):
    success: bool
    document_id: Optional[str] = None
    created_block_id: Optional[str] = None
    data: Optional[Any] = None
    error: Optional[ErrorDetail] = None

    @classmethod
    def ok(
        cls,
        *,
        document_id: Optional[str] = None,
        created_block_id: Optional[str] = None,
        data: Any = None,
    ) -> "ApiResponse":
        return cls(
            success=True,
            document_id=document_id,
            created_block_id=created_block_id,
            data=data,
        )

    @classmethod
    def fail(cls, code: str, message: str, *, document_id: Optional[str] = None) -> "ApiResponse":
        return cls(
            success=False,
            document_id=document_id,
            error=ErrorDetail(code=code, message=message),
        )


class MatchItem(BaseModel):
    block_id: str
    text: str
    type: Optional[str] = None


class FindResponseData(BaseModel):
    matches: list[MatchItem] = Field(default_factory=list)


class PreviewPage(BaseModel):
    page: int
    image: str


class PreviewData(BaseModel):
    pages: list[PreviewPage] = Field(default_factory=list)
    pdf: Optional[str] = None
