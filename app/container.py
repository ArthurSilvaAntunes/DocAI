"""Instância compartilhada dos serviços (API REST e MCP)."""

from __future__ import annotations

from app.services.document_service import DocumentService
from app.services.editing_service import EditingService
from app.services.rendering_service import RenderingService
from app.services.template_service import TemplateService
from app.storage.document_store import DocumentStore

store = DocumentStore()
editing_service = EditingService(store)
template_service = TemplateService()
rendering_service = RenderingService()
document_service = DocumentService(store, editing_service, template_service, rendering_service)
