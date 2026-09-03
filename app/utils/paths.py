"""Utilitários de caminho seguro (previne path traversal)."""

from __future__ import annotations

import re
from pathlib import Path

from app.exceptions import AppError

SAFE_FILENAME = re.compile(r"^[A-Za-z0-9._\- ]+$")


def sanitize_filename(name: str, *, default_ext: str = ".docx") -> str:
    raw_input = (name or "").strip()
    if not raw_input:
        raise AppError("INVALID_FILENAME", "Filename is required.")
    normalized = raw_input.replace("\\", "/")
    if normalized.startswith("/") or ":" in Path(raw_input).drive or ".." in normalized.split("/"):
        raise AppError("PATH_TRAVERSAL", "Path traversal is not allowed.")
    if any(part == ".." for part in Path(raw_input).parts):
        raise AppError("PATH_TRAVERSAL", "Path traversal is not allowed.")
    raw = normalized.split("/")[-1]
    if not raw:
        raise AppError("INVALID_FILENAME", "Filename is required.")
    if not SAFE_FILENAME.match(raw):
        raise AppError("INVALID_FILENAME", "Filename contains invalid characters.")
    if default_ext and not raw.lower().endswith(default_ext):
        raw += default_ext
    return raw


def safe_join(root: Path, user_path: str) -> Path:
    root_resolved = root.resolve()
    candidate = Path(user_path)
    if candidate.is_absolute():
        raise AppError("PATH_TRAVERSAL", "Absolute paths are not allowed.")
    joined = (root_resolved / user_path).resolve()
    try:
        joined.relative_to(root_resolved)
    except ValueError as exc:
        raise AppError("PATH_TRAVERSAL", "Path is outside the allowed directory.") from exc
    return joined


def resolve_existing_file(root: Path, user_path: str) -> Path:
    filename = sanitize_filename(user_path, default_ext="")
    path = safe_join(root, filename)
    if not path.is_file():
        raise AppError("FILE_NOT_FOUND", f"File was not found: {path.name}", status=404)
    return path
