"""Comandos genéricos para agentes de IA."""

from __future__ import annotations

from typing import Any, Optional

from pydantic import BaseModel, ConfigDict, Field


class Command(BaseModel):
    """Comando genérico. Campos extras são aceitos para flexibilidade da IA."""

    model_config = ConfigDict(extra="allow")

    action: str
    target_id: Optional[str] = None
    after_id: Optional[str] = None
    before_id: Optional[str] = None
    text: Optional[str] = None
    search: Optional[str] = None
    replace: Optional[str] = None
    content: Optional[dict[str, Any]] = None
    properties: Optional[dict[str, Any]] = None
    level: Optional[int] = None
    style: Optional[str] = None
    filename: Optional[str] = None
    variables: Optional[dict[str, str]] = None


class CommandBatch(BaseModel):
    commands: list[Command] = Field(min_length=1)


class FindRequest(BaseModel):
    text: Optional[str] = None
    block_type: Optional[str] = None
    style: Optional[str] = None
    block_id: Optional[str] = None


class ReplaceRequest(BaseModel):
    search: str
    replace: str
    block_id: Optional[str] = None


class SetVariablesRequest(BaseModel):
    variables: dict[str, str]


class HistoryEntry(BaseModel):
    timestamp: str
    action: str
    target: Optional[str] = None
    before: Optional[str] = None
    after: Optional[str] = None
    success: bool = True
