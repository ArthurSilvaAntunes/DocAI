"""Listas com marcadores e numeração multinível."""

from __future__ import annotations

from docx.document import Document as DocxDocument
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.oxml.ns import qn
from docx.oxml.parser import parse_xml
from docx.parts.numbering import NumberingPart
from lxml import etree

from app.docx_engine import ids as id_mod
from app.docx_engine import xml_utils as xu
from app.docx_engine.paragraphs import insert_paragraph
from app.exceptions import AppError

BULLET_NUM_ID = 10
NUMBER_NUM_ID = 11

NUMBERING_ROOT = (
    f'<w:numbering xmlns:w="{xu.W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
)

ABSTRACT_NUM_BULLET = f"""
<w:abstractNum w:abstractNumId="10" xmlns:w="{xu.W_NS}">
  <w:multiLevelType w:val="hybridMultilevel"/>
  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/>
    <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>
  </w:lvl>
  <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="○"/><w:lvlJc w:val="left"/>
    <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:hint="default"/></w:rPr>
  </w:lvl>
  <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="▪"/><w:lvlJc w:val="left"/>
    <w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/></w:rPr>
  </w:lvl>
</w:abstractNum>
"""

ABSTRACT_NUM_NUMBER = f"""
<w:abstractNum w:abstractNumId="11" xmlns:w="{xu.W_NS}">
  <w:multiLevelType w:val="multilevel"/>
  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>
    <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
  <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/>
    <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
  <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3."/><w:lvlJc w:val="left"/>
    <w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr></w:lvl>
</w:abstractNum>
"""


def get_numbering_part(document: DocxDocument) -> NumberingPart:
    part = document.part
    try:
        return part.part_related_by(RT.NUMBERING)
    except KeyError:
        numbering_part = NumberingPart(
            PackURI("/word/numbering.xml"),
            CT.WML_NUMBERING,
            parse_xml(NUMBERING_ROOT.encode("utf-8")),
            part.package,
        )
        part.relate_to(numbering_part, RT.NUMBERING)
        return numbering_part


def _ensure_numbering(document: DocxDocument) -> None:
    numbering_part = get_numbering_part(document)
    root = numbering_part._element
    existing_ids = {el.get(qn("w:abstractNumId")) for el in root.findall(qn("w:abstractNum"))}
    existing_num_ids = {el.get(qn("w:numId")) for el in root.findall(qn("w:num"))}
    if "10" not in existing_ids:
        root.append(etree.fromstring(ABSTRACT_NUM_BULLET))
    if "11" not in existing_ids:
        root.append(etree.fromstring(ABSTRACT_NUM_NUMBER))
    if "10" not in existing_num_ids:
        num = xu.make("num", numId="10")
        num.append(xu.make("abstractNumId", val="10"))
        root.append(num)
    if "11" not in existing_num_ids:
        num = xu.make("num", numId="11")
        num.append(xu.make("abstractNumId", val="11"))
        root.append(num)


def add_list_items(
    document: DocxDocument,
    registry: id_mod.IdRegistry,
    items: list[str],
    *,
    list_type: str = "bullet",
    level: int = 0,
    after_el: etree._Element | None = None,
) -> list[str]:
    if not items:
        raise AppError("INVALID_LIST", "List items are required.")
    _ensure_numbering(document)
    num_id = BULLET_NUM_ID if list_type == "bullet" else NUMBER_NUM_ID
    created: list[str] = []
    current_after = after_el
    for item in items:
        block_id, paragraph = insert_paragraph(
            document,
            registry,
            text=item,
            after_el=current_after,
            kind="list",
        )
        xu.set_num_pr(paragraph._p, num_id, max(0, min(8, level)))
        created.append(block_id)
        current_after = paragraph._p
    return created


def apply_list_format(p_el: etree._Element, document: DocxDocument, list_type: str, level: int) -> None:
    _ensure_numbering(document)
    num_id = BULLET_NUM_ID if list_type == "bullet" else NUMBER_NUM_ID
    xu.set_num_pr(p_el, num_id, max(0, min(8, level)))


def list_type_of(p_el: etree._Element) -> str | None:
    if not xu.paragraph_has_num_pr(p_el):
        return None
    ppr = xu.first_child(p_el, "pPr")
    num_pr = xu.first_child(ppr, "numPr") if ppr is not None else None
    num_id_el = xu.first_child(num_pr, "numId") if num_pr is not None else None
    raw = xu.get_attr(num_id_el, "val") if num_id_el is not None else None
    if raw == str(NUMBER_NUM_ID):
        return "numbered"
    return "bullet"
