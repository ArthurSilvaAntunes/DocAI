"""Estilos nativos do DOCX."""

from __future__ import annotations

from docx.document import Document as DocxDocument
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE
from docx.shared import Pt, RGBColor
from docx.styles.style import BaseStyle

from app.docx_engine.paragraphs import ALIGN_MAP, _parse_color
from app.exceptions import AppError, NotFoundError
from app.models.blocks import Alignment
from app.models.styles import StyleInfo, StyleKind, StyleUpsertRequest

TYPE_MAP = {
    StyleKind.paragraph: WD_STYLE_TYPE.PARAGRAPH,
    StyleKind.character: WD_STYLE_TYPE.CHARACTER,
    StyleKind.table: WD_STYLE_TYPE.TABLE,
}
REVERSE_TYPE = {
    WD_STYLE_TYPE.PARAGRAPH: StyleKind.paragraph,
    WD_STYLE_TYPE.CHARACTER: StyleKind.character,
    WD_STYLE_TYPE.TABLE: StyleKind.table,
}


def list_styles(document: DocxDocument) -> list[StyleInfo]:
    result: list[StyleInfo] = []
    for style in document.styles:
        kind = REVERSE_TYPE.get(style.type)
        if kind is None:
            continue
        font = getattr(style.font, "name", None)
        size = style.font.size.pt if getattr(style.font, "size", None) else None
        alignment = None
        if kind == StyleKind.paragraph:
            try:
                raw = style.paragraph_format.alignment
                reverse = {
                    WD_ALIGN_PARAGRAPH.LEFT: Alignment.left,
                    WD_ALIGN_PARAGRAPH.CENTER: Alignment.center,
                    WD_ALIGN_PARAGRAPH.RIGHT: Alignment.right,
                    WD_ALIGN_PARAGRAPH.JUSTIFY: Alignment.justify,
                }
                alignment = reverse.get(raw)
            except (AttributeError, ValueError):
                alignment = None
        result.append(
            StyleInfo(
                name=style.name,
                type=kind,
                builtin=bool(getattr(style, "builtin", False)),
                font=font,
                font_size=size,
                bold=style.font.bold,
                italic=style.font.italic,
                alignment=alignment,
            )
        )
    return result


def get_style(document: DocxDocument, name: str) -> StyleInfo:
    try:
        style = document.styles[name]
    except KeyError as exc:
        raise NotFoundError("STYLE_NOT_FOUND", f"Style '{name}' was not found.") from exc
    for item in list_styles(document):
        if item.name == style.name:
            return item
    raise NotFoundError("STYLE_NOT_FOUND", f"Style '{name}' was not found.")


def _apply_style_font(style: BaseStyle, payload: StyleUpsertRequest) -> None:
    font = style.font
    if payload.bold is not None:
        font.bold = payload.bold
    if payload.italic is not None:
        font.italic = payload.italic
    if payload.underline is not None:
        font.underline = WD_UNDERLINE.SINGLE if payload.underline else WD_UNDERLINE.NONE
    if payload.strikethrough is not None:
        font.strike = payload.strikethrough
    if payload.font:
        font.name = payload.font
    if payload.font_size is not None:
        font.size = Pt(payload.font_size)
    if payload.color:
        font.color.rgb = _parse_color(payload.color)
    if payload.superscript is not None:
        font.superscript = payload.superscript
    if payload.subscript is not None:
        font.subscript = payload.subscript


def create_or_update_style(document: DocxDocument, payload: StyleUpsertRequest) -> StyleInfo:
    try:
        style = document.styles[payload.name]
    except KeyError:
        style = document.styles.add_style(payload.name, TYPE_MAP[payload.type])
    if payload.based_on:
        try:
            style.base_style = document.styles[payload.based_on]
        except KeyError as exc:
            raise AppError("STYLE_NOT_FOUND", f"Base style '{payload.based_on}' was not found.") from exc
    _apply_style_font(style, payload)
    if payload.type == StyleKind.paragraph:
        if payload.alignment:
            style.paragraph_format.alignment = ALIGN_MAP.get(payload.alignment)
        if payload.space_before_pt is not None:
            style.paragraph_format.space_before = Pt(payload.space_before_pt)
        if payload.space_after_pt is not None:
            style.paragraph_format.space_after = Pt(payload.space_after_pt)
    return get_style(document, payload.name)


def apply_style(document: DocxDocument, element, style_name: str) -> None:
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    try:
        document.styles[style_name]
    except KeyError as exc:
        raise AppError("STYLE_NOT_FOUND", f"Style '{style_name}' was not found.") from exc
    if isinstance(element, Paragraph):
        element.style = style_name
        return
    if isinstance(element, Table):
        element.style = style_name
        return
    raise AppError("INVALID_STYLE_TARGET", "Style can only be applied to paragraphs or tables.")
