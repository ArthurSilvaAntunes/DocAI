"""Renderização DOCX → PDF (LibreOffice) → imagens PNG (PyMuPDF)."""

from __future__ import annotations

import logging
import os
import shutil
import subprocess
from pathlib import Path

from app.config import settings
from app.exceptions import AppError
from app.models.responses import PreviewData, PreviewPage
from app.utils.paths import sanitize_filename

logger = logging.getLogger(__name__)


def find_soffice() -> Path | None:
    candidates: list[Path] = []
    if settings.soffice_path:
        candidates.append(Path(settings.soffice_path))
    env = os.environ.get("SOFFICE_PATH")
    if env:
        candidates.append(Path(env))
    which = shutil.which("soffice") or shutil.which("soffice.exe") or shutil.which("libreoffice")
    if which:
        candidates.append(Path(which))
    candidates.extend(
        [
            Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
            Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
            Path("/usr/bin/soffice"),
            Path("/usr/bin/libreoffice"),
            Path("/usr/lib/libreoffice/program/soffice"),
            Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        ]
    )
    for path in candidates:
        if path.is_file():
            return path
    return None


class RenderingService:
    def render_pdf(self, docx_path: Path, document_id: str) -> Path:
        soffice = find_soffice()
        if soffice is None:
            raise AppError(
                "RENDERER_UNAVAILABLE",
                "LibreOffice (soffice) was not found. Install LibreOffice or set AIDOC_SOFFICE_PATH.",
                status=503,
            )
        out_dir = settings.previews_dir / document_id
        out_dir.mkdir(parents=True, exist_ok=True)
        logger.info("render_pdf document_id=%s", document_id)
        cmd = [
            str(soffice),
            "--headless",
            "--nologo",
            "--nofirststartwizard",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out_dir),
            str(docx_path),
        ]
        try:
            completed = subprocess.run(
                cmd,
                check=False,
                capture_output=True,
                text=True,
                timeout=120,
            )
        except subprocess.TimeoutExpired as exc:
            raise AppError("RENDER_TIMEOUT", "PDF conversion timed out.") from exc
        if completed.returncode != 0:
            logger.error("render_pdf failed document_id=%s code=%s", document_id, completed.returncode)
            raise AppError("RENDER_FAILED", "LibreOffice failed to convert the document to PDF.")
        pdf_path = out_dir / (docx_path.stem + ".pdf")
        expected = out_dir / "working.pdf"
        if not pdf_path.is_file():
            generated = list(out_dir.glob("*.pdf"))
            if generated:
                pdf_path = generated[0]
            elif expected.is_file():
                pdf_path = expected
            else:
                raise AppError("RENDER_FAILED", "PDF file was not produced.")
        stable = out_dir / "preview.pdf"
        if pdf_path != stable:
            shutil.copy2(pdf_path, stable)
        return stable

    def render_preview(self, docx_path: Path, document_id: str) -> PreviewData:
        pdf_path = self.render_pdf(docx_path, document_id)
        try:
            import fitz
        except ImportError as exc:
            raise AppError("RENDERER_UNAVAILABLE", "PyMuPDF is not installed.") from exc
        out_dir = settings.previews_dir / document_id
        pages: list[PreviewPage] = []
        with fitz.open(pdf_path) as pdf:
            for index, page in enumerate(pdf, start=1):
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                filename = f"page_{index:03d}.png"
                image_path = out_dir / filename
                pix.save(str(image_path))
                pages.append(PreviewPage(page=index, image=f"/previews/{document_id}/{filename}"))
        logger.info("render_preview document_id=%s pages=%s", document_id, len(pages))
        return PreviewData(pages=pages, pdf=f"/previews/{document_id}/preview.pdf")
