"""Configuração e diretórios controlados da aplicação."""

from __future__ import annotations

from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_prefix="AIDOC_", extra="ignore")

    app_name: str = "AI DOCX Editor"
    host: str = "127.0.0.1"
    port: int = 8000
    soffice_path: str = ""
    max_history_snapshots: int = 20

    @property
    def root_dir(self) -> Path:
        return Path(__file__).resolve().parent.parent

    @property
    def documents_dir(self) -> Path:
        return self.root_dir / "documents"

    @property
    def templates_dir(self) -> Path:
        return self.root_dir / "templates"

    @property
    def previews_dir(self) -> Path:
        return self.root_dir / "previews"

    @property
    def history_dir(self) -> Path:
        return self.root_dir / "history"

    @property
    def workspace_dir(self) -> Path:
        return self.documents_dir / ".workspace"


settings = Settings()


def ensure_directories() -> None:
    for path in (
        settings.documents_dir,
        settings.templates_dir,
        settings.previews_dir,
        settings.history_dir,
        settings.workspace_dir,
    ):
        path.mkdir(parents=True, exist_ok=True)
