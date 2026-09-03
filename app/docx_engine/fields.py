"""Campos nativos do Word: TOC, listas de figuras/tabelas/quadros e legendas SEQ."""

from __future__ import annotations

from typing import Literal

from docx.document import Document as DocxDocument
from lxml import etree

from app.docx_engine import ids as id_mod
from app.docx_engine import xml_utils as xu
from app.docx_engine.language import DEFAULT_LANGUAGE
from app.docx_engine.paragraphs import heading_level_from_style, insert_paragraph, wrap_paragraph
from app.exceptions import AppError

IndexKind = Literal["toc", "figures", "tables", "quadros"]

CAPTION_LABELS = {
    "figura": "Figura",
    "table": "Tabela",
    "tabela": "Tabela",
    "quadro": "Quadro",
    "figure": "Figura",
}

INDEX_TITLES = {
    "toc": "Sumário",
    "figures": "Lista de figuras",
    "tables": "Lista de tabelas",
    "quadros": "Lista de quadros",
}

TOC_INSTRUCTIONS = {
    "toc": 'TOC \\o "1-3" \\h \\z \\u',
    "figures": 'TOC \\h \\z \\c "Figura"',
    "tables": 'TOC \\h \\z \\c "Tabela"',
    "quadros": 'TOC \\h \\z \\c "Quadro"',
}


def insert_caption(
    document: DocxDocument,
    registry: id_mod.IdRegistry,
    *,
    text: str,
    kind: str = "figura",
    after_el: etree._Element | None = None,
    number: int | None = None,
) -> str:
    label = CAPTION_LABELS.get(kind.lower(), kind.title())
    seq_name = label
    block_id, paragraph = insert_paragraph(
        document,
        registry,
        after_el=after_el,
        kind="paragraph",
        style="Caption" if _has_style(document, "Caption") else None,
        properties={"alignment": "center", "space_after_pt": 12},
    )
    p_el = paragraph._p
    xu.clear_element_text_runs(p_el)
    _append_text(p_el, f"{label} ")
    _append_seq_field(p_el, seq_name, cached=str(number or 1))
    cleaned = _strip_caption_prefix(text, label)
    if cleaned:
        _append_text(p_el, f" — {cleaned}")
    paragraph.alignment = 1  # center
    return block_id


def write_caption_paragraph(p_el: etree._Element, *, kind: str, text: str, number: int) -> None:
    label = CAPTION_LABELS.get(kind.lower(), kind.title())
    xu.clear_element_text_runs(p_el)
    _append_text(p_el, f"{label} ")
    _append_seq_field(p_el, label, cached=str(number))
    cleaned = _strip_caption_prefix(text, label)
    if cleaned:
        _append_text(p_el, f" — {cleaned}")


def _strip_caption_prefix(text: str, label: str) -> str:
    cleaned = text.strip()
    for prefix in (f"{label} ", "Figura ", "Tabela ", "Quadro ", "Figure ", "Table "):
        if cleaned.lower().startswith(prefix.lower()):
            cleaned = cleaned[len(prefix) :]
            break
    return cleaned.lstrip("0123456789.-—–: ").lstrip()


def insert_index(
    document: DocxDocument,
    registry: id_mod.IdRegistry,
    *,
    kind: IndexKind = "toc",
    after_el: etree._Element | None = None,
    title: str | None = None,
    levels: str = "1-3",
) -> dict[str, str]:
    if kind not in TOC_INSTRUCTIONS:
        raise AppError("INVALID_COMMAND", f"Unknown index kind '{kind}'.")
    heading_text = title or INDEX_TITLES[kind]
    heading_id, heading = insert_paragraph(
        document,
        registry,
        text=heading_text,
        after_el=after_el,
        kind="heading",
        style="Heading 1",
    )
    instruction = TOC_INSTRUCTIONS[kind]
    if kind == "toc":
        instruction = f'TOC \\o "{levels}" \\h \\z \\u'
    field_id, field_p = insert_paragraph(
        document,
        registry,
        after_el=heading._p,
        kind="paragraph",
        style="Normal",
    )
    _write_toc_field(field_p._p, instruction, _index_preview(document, kind))
    from app.docx_engine.language import enable_update_fields_on_open

    enable_update_fields_on_open(document)
    return {"heading_id": heading_id, "field_id": field_id}


def _index_preview(document: DocxDocument, kind: IndexKind) -> list[str]:
    lines: list[str] = []
    body = document.element.body
    for element in xu.iter_body_elements(body):
        text = "".join((t.text or "") for t in element.iter(xu.w("t"))).strip()
        if not text:
            continue
        if element.tag == xu.w("tbl"):
            continue
        paragraph = wrap_paragraph(element, document)
        style_name = paragraph.style.name if paragraph.style is not None else ""
        is_heading = heading_level_from_style(style_name) is not None
        if kind == "toc" and is_heading:
            lines.append(text)
        elif kind == "figures" and text.lower().startswith("figura"):
            lines.append(text)
        elif kind == "tables" and text.lower().startswith("tabela"):
            lines.append(text)
        elif kind == "quadros" and text.lower().startswith("quadro"):
            lines.append(text)
    if not lines:
        hint = {
            "toc": "(O Word atualiza este sumário ao abrir o arquivo.)",
            "figures": "(Nenhuma figura legendada ainda.)",
            "tables": "(Nenhuma tabela legendada ainda.)",
            "quadros": "(Nenhum quadro legendado ainda.)",
        }
        return [hint[kind]]
    return lines[:40]


def _write_toc_field(p_el: etree._Element, instruction: str, preview: list[str]) -> None:
    xu.clear_element_text_runs(p_el)
    begin = xu.make("r")
    begin.append(xu.make("fldChar", fldCharType="begin"))
    p_el.append(begin)

    instr_run = xu.make("r")
    instr_el = xu.make("instrText")
    instr_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr_el.text = f" {instruction} "
    instr_run.append(instr_el)
    p_el.append(instr_run)

    sep = xu.make("r")
    sep.append(xu.make("fldChar", fldCharType="separate"))
    p_el.append(sep)

    _append_text(p_el, preview[0] if preview else "")
    for extra in preview[1:]:
        _append_text(p_el, "\n" + extra)

    end = xu.make("r")
    end.append(xu.make("fldChar", fldCharType="end"))
    p_el.append(end)


def _append_text(p_el: etree._Element, text: str, *, bold: bool = False) -> None:
    run = xu.make("r")
    rpr = xu.make("rPr")
    if bold:
        rpr.append(xu.make("b"))
    rpr.append(xu.make("lang", val=DEFAULT_LANGUAGE, eastAsia=DEFAULT_LANGUAGE, bidi="ar-SA"))
    run.append(rpr)
    t = xu.make("t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    run.append(t)
    p_el.append(run)


def _append_seq_field(p_el: etree._Element, seq_name: str, *, cached: str = "1") -> None:
    begin = xu.make("r")
    begin.append(xu.make("fldChar", fldCharType="begin"))
    p_el.append(begin)

    instr_run = xu.make("r")
    instr_el = xu.make("instrText")
    instr_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr_el.text = f" SEQ {seq_name} \\* ARABIC "
    instr_run.append(instr_el)
    p_el.append(instr_run)

    sep = xu.make("r")
    sep.append(xu.make("fldChar", fldCharType="separate"))
    p_el.append(sep)

    result = xu.make("r")
    t = xu.make("t")
    t.text = cached
    result.append(t)
    p_el.append(result)

    end = xu.make("r")
    end.append(xu.make("fldChar", fldCharType="end"))
    p_el.append(end)


def _has_style(document: DocxDocument, name: str) -> bool:
    try:
        document.styles[name]
        return True
    except KeyError:
        return False


def next_caption_number(document: DocxDocument, kind: str) -> int:
    label = CAPTION_LABELS.get(kind.lower(), kind.title())
    count = 0
    for instr in document.element.iter(xu.w("instrText")):
        raw = (instr.text or "").strip()
        if raw.upper().startswith("SEQ ") and label.upper() in raw.upper():
            count += 1
    if count:
        return count + 1
    prefix = label.lower()
    for t_el in document.element.iter(xu.w("t")):
        text = (t_el.text or "").strip().lower()
        if text.startswith(prefix):
            count += 1
    return count + 1
