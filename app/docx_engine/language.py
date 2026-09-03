"""Idioma de revisão do Word (evita ondulações vermelhas em texto correto)."""

from __future__ import annotations

from docx.document import Document as DocxDocument
from docx.oxml.ns import qn
from lxml import etree

from app.docx_engine import xml_utils as xu

DEFAULT_LANGUAGE = "pt-BR"


def set_document_language(document: DocxDocument, language: str = DEFAULT_LANGUAGE) -> None:
    lang = language.strip() or DEFAULT_LANGUAGE
    _set_doc_defaults(document, lang)
    _set_settings_language(document, lang)
    for el in document.element.iter(xu.w("rPr")):
        _set_lang_on_rpr(el, lang)


def get_document_language(document: DocxDocument) -> str | None:
    defaults = document.styles.element.find(qn("w:docDefaults"))
    if defaults is None:
        return None
    lang = defaults.find(".//" + xu.w("lang"))
    if lang is None:
        return None
    return xu.get_attr(lang, "val")


def apply_run_language(run_el: etree._Element, language: str = DEFAULT_LANGUAGE) -> None:
    rpr = xu.first_child(run_el, "rPr")
    if rpr is None:
        rpr = xu.make("rPr")
        run_el.insert(0, rpr)
    _set_lang_on_rpr(rpr, language)


def _set_lang_on_rpr(rpr: etree._Element, language: str) -> None:
    xu.remove_children(rpr, "lang")
    rpr.append(xu.make("lang", val=language, eastAsia=language, bidi="ar-SA"))


def _set_doc_defaults(document: DocxDocument, language: str) -> None:
    styles = document.styles.element
    defaults = styles.find(qn("w:docDefaults"))
    if defaults is None:
        defaults = xu.make("docDefaults")
        styles.insert(0, defaults)
    rpr_default = defaults.find(qn("w:rPrDefault"))
    if rpr_default is None:
        rpr_default = xu.make("rPrDefault")
        defaults.insert(0, rpr_default)
    rpr = rpr_default.find(qn("w:rPr"))
    if rpr is None:
        rpr = xu.make("rPr")
        rpr_default.append(rpr)
    _set_lang_on_rpr(rpr, language)


def _set_settings_language(document: DocxDocument, language: str) -> None:
    try:
        settings = document.part.settings.element
    except (AttributeError, KeyError, ValueError):
        return
    theme_font = settings.find(qn("w:themeFontLang"))
    if theme_font is None:
        settings.append(xu.make("themeFontLang", val=language, eastAsia=language, bidi="ar-SA"))
    else:
        xu.set_attr(theme_font, "val", language)
        xu.set_attr(theme_font, "eastAsia", language)
    _ensure_update_fields(settings)


def _ensure_update_fields(settings: etree._Element) -> None:
    existing = settings.find(qn("w:updateFields"))
    if existing is None:
        settings.append(xu.make("updateFields", val="true"))
    else:
        xu.set_attr(existing, "val", "true")


def enable_update_fields_on_open(document: DocxDocument) -> None:
    try:
        settings = document.part.settings.element
    except (AttributeError, KeyError, ValueError):
        return
    _ensure_update_fields(settings)
