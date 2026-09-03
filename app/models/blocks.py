"""Modelos Pydantic de blocos, runs e formatação."""

from __future__ import annotations

from enum import Enum
from typing import Any, Literal, Optional

from pydantic import BaseModel, ConfigDict, Field


class BlockType(str, Enum):
    heading = "heading"
    paragraph = "paragraph"
    table = "table"
    image = "image"
    list = "list"
    page_break = "page_break"
    section_break = "section_break"


class Alignment(str, Enum):
    left = "left"
    center = "center"
    right = "right"
    justify = "justify"


class RunFormat(BaseModel):
    model_config = ConfigDict(extra="ignore")

    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    strikethrough: Optional[bool] = None
    font: Optional[str] = None
    font_size: Optional[float] = None
    highlight: Optional[str] = None
    color: Optional[str] = None
    superscript: Optional[bool] = None
    subscript: Optional[bool] = None


class RunModel(RunFormat):
    text: str = ""


class ParagraphProperties(BaseModel):
    model_config = ConfigDict(extra="ignore")

    style: Optional[str] = None
    alignment: Optional[Alignment] = None
    indent_left_cm: Optional[float] = None
    first_line_indent_cm: Optional[float] = None
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None
    line_spacing: Optional[float] = None
    keep_with_next: Optional[bool] = None
    page_break_before: Optional[bool] = None


class ParagraphContent(ParagraphProperties):
    text: Optional[str] = None
    runs: Optional[list[RunModel]] = None


class HeadingContent(ParagraphContent):
    level: int = Field(default=1, ge=1, le=9)


class ListContent(ParagraphContent):
    list_type: Literal["bullet", "numbered"] = "bullet"
    level: int = Field(default=0, ge=0, le=8)
    items: Optional[list[str]] = None


class TableCellContent(BaseModel):
    model_config = ConfigDict(extra="ignore")

    text: str = ""
    bold: Optional[bool] = None
    alignment: Optional[Alignment] = None
    shading: Optional[str] = None


class TableContent(BaseModel):
    model_config = ConfigDict(extra="ignore")

    rows: int = Field(default=2, ge=1, le=200)
    columns: int = Field(default=2, ge=1, le=50)
    data: Optional[list[list[str]]] = None
    style: Optional[str] = None
    width_cm: Optional[float] = None


class ImageContent(BaseModel):
    model_config = ConfigDict(extra="ignore")

    filename: Optional[str] = None
    width_cm: Optional[float] = None
    height_cm: Optional[float] = None
    alignment: Optional[Alignment] = None
    caption: Optional[str] = None


class BlockSummary(BaseModel):
    id: str
    type: BlockType
    text: Optional[str] = None
    level: Optional[int] = None
    rows: Optional[int] = None
    columns: Optional[int] = None
    style: Optional[str] = None
    list_type: Optional[str] = None


class BlockDetail(BaseModel):
    id: str
    type: BlockType
    text: Optional[str] = None
    level: Optional[int] = None
    style: Optional[str] = None
    properties: Optional[ParagraphProperties] = None
    runs: Optional[list[RunModel]] = None
    rows: Optional[int] = None
    columns: Optional[int] = None
    cells: Optional[list[list[str]]] = None
    list_type: Optional[str] = None
    extra: dict[str, Any] = Field(default_factory=dict)
