"""Modelos de estilos DOCX."""

from __future__ import annotations

from enum import Enum
from typing import Optional

from pydantic import BaseModel, ConfigDict

from app.models.blocks import Alignment, RunFormat


class StyleKind(str, Enum):
    paragraph = "paragraph"
    character = "character"
    table = "table"


class StyleInfo(BaseModel):
    name: str
    type: StyleKind
    builtin: bool = False
    font: Optional[str] = None
    font_size: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    alignment: Optional[Alignment] = None


class StyleUpsertRequest(RunFormat):
    model_config = ConfigDict(extra="ignore")

    name: str
    type: StyleKind = StyleKind.paragraph
    based_on: Optional[str] = None
    alignment: Optional[Alignment] = None
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None
