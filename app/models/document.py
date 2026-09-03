"""Modelos de documento, página e metadados."""

from __future__ import annotations

from datetime import datetime
from typing import Literal, Optional

from pydantic import BaseModel, ConfigDict, Field

from app.models.blocks import BlockSummary


class PageSetup(BaseModel):
    model_config = ConfigDict(extra="ignore")

    size: Literal["A4", "Letter"] = "A4"
    orientation: Literal["portrait", "landscape"] = "portrait"
    margin_top_cm: float = 2.54
    margin_bottom_cm: float = 2.54
    margin_left_cm: float = 2.54
    margin_right_cm: float = 2.54


class CreateDocumentRequest(BaseModel):
    filename: str = "documento.docx"
    title: Optional[str] = None
    author: Optional[str] = None
    template: Optional[str] = None
    template_source: Literal["documents", "templates"] = "templates"
    language: Optional[str] = "pt-BR"
    page: PageSetup = Field(default_factory=PageSetup)


class OpenDocumentRequest(BaseModel):
    filename: str
    source: Literal["documents", "templates"] = "documents"


class SaveDocumentRequest(BaseModel):
    filename: Optional[str] = None
    overwrite_original: bool = False
    keep_version: bool = True


class DocumentInfo(BaseModel):
    id: str
    filename: str
    title: Optional[str] = None
    author: Optional[str] = None
    original_filename: Optional[str] = None
    dirty: bool = False
    created_at: datetime
    updated_at: datetime
    block_count: int = 0


class DocumentStructure(BaseModel):
    document_id: str
    filename: str
    title: Optional[str] = None
    blocks: list[BlockSummary]


class OutlineEntry(BaseModel):
    id: str
    level: int
    text: str


class TablePreview(BaseModel):
    id: str
    rows: int
    columns: int
    header: list[str] = Field(default_factory=list)


class CaptionPreview(BaseModel):
    id: str
    kind: str
    text: str


class DocumentInspect(BaseModel):
    document_id: str
    filename: str
    title: Optional[str] = None
    language: Optional[str] = None
    placeholders: list[str] = Field(default_factory=list)
    outline: list[OutlineEntry] = Field(default_factory=list)
    tables: list[TablePreview] = Field(default_factory=list)
    images: list[str] = Field(default_factory=list)
    captions: list[CaptionPreview] = Field(default_factory=list)
    styles_in_use: list[str] = Field(default_factory=list)
    block_count: int = 0


class HeaderFooterRequest(BaseModel):
    text: Optional[str] = None
    include_page_number: bool = False
    include_document_name: bool = False
    include_date: bool = False
    number_only: bool = False
    image_filename: Optional[str] = None
