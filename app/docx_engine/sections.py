"""Seções, tamanho de página, margens e quebras."""

from __future__ import annotations

from copy import deepcopy

from docx.document import Document as DocxDocument
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.shared import Cm, Mm, Inches, Emu
from lxml import etree

from app.docx_engine import ids as id_mod
from app.docx_engine import xml_utils as xu
from app.docx_engine.paragraphs import insert_paragraph
from app.models.document import PageSetup

PAGE_SIZES = {
    "A4": (Mm(210), Mm(297)),
    "Letter": (Inches(8.5), Inches(11)),
}

SECTION_TYPES = {
    "next_page": WD_SECTION.NEW_PAGE,
    "continuous": WD_SECTION.CONTINUOUS,
    "even_page": WD_SECTION.EVEN_PAGE,
    "odd_page": WD_SECTION.ODD_PAGE,
}

SECTION_TYPE_XML = {
    "next_page": "nextPage",
    "continuous": "continuous",
    "even_page": "evenPage",
    "odd_page": "oddPage",
}


def apply_page_setup(document: DocxDocument, page: PageSetup) -> None:
    width, height = PAGE_SIZES.get(page.size, PAGE_SIZES["A4"])
    for section in document.sections:
        if page.orientation == "landscape":
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width = height
            section.page_height = width
        else:
            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = width
            section.page_height = height
        section.top_margin = Cm(page.margin_top_cm)
        section.bottom_margin = Cm(page.margin_bottom_cm)
        section.left_margin = Cm(page.margin_left_cm)
        section.right_margin = Cm(page.margin_right_cm)


def add_section_break(
    document: DocxDocument,
    registry: id_mod.IdRegistry,
    *,
    break_type: str = "next_page",
    after_el: etree._Element | None = None,
) -> str:
    xml_type = SECTION_TYPE_XML.get(break_type, "nextPage")
    current = document.sections[-1]._sectPr
    new_sect = deepcopy(current)
    type_el = xu.first_child(new_sect, "type")
    if type_el is None:
        new_sect.append(xu.make("type", val=xml_type))
    else:
        xu.set_attr(type_el, "val", xml_type)
    block_id, paragraph = insert_paragraph(
        document,
        registry,
        after_el=after_el,
        kind="section",
    )
    ppr = paragraph._p.get_or_add_pPr()
    existing = xu.first_child(ppr, "sectPr")
    if existing is not None:
        ppr.remove(existing)
    ppr.append(new_sect)
    return block_id
