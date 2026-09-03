"""Helpers centralizados de WordprocessingML. Nenhum XML bruto deve vazar para a API."""

from __future__ import annotations

from copy import deepcopy
from typing import Iterable, Optional

from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NSMAP = {"w": W_NS, "r": R_NS, "a": A_NS, "wp": WP_NS}


def w(tag: str) -> str:
    return qn(f"w:{tag}")


def make(tag: str, **attrs: str) -> etree._Element:
    el = OxmlElement(f"w:{tag}")
    for key, value in attrs.items():
        if value is None:
            continue
        el.set(qn(f"w:{key}"), str(value))
    return el


def set_attr(el: etree._Element, name: str, value: str) -> None:
    el.set(qn(f"w:{name}"), value)


def get_attr(el: etree._Element, name: str) -> Optional[str]:
    return el.get(qn(f"w:{name}"))


def first_child(el: etree._Element, tag: str) -> Optional[etree._Element]:
    return el.find(w(tag))


def children(el: etree._Element, tag: str) -> list[etree._Element]:
    return list(el.findall(w(tag)))


def ensure_child(parent: etree._Element, tag: str, *, index: int = 0) -> etree._Element:
    existing = first_child(parent, tag)
    if existing is not None:
        return existing
    child = make(tag)
    parent.insert(index, child)
    return child


def remove_children(el: etree._Element, tag: str) -> None:
    for child in children(el, tag):
        el.remove(child)


def clone(el: etree._Element) -> etree._Element:
    return deepcopy(el)


def paragraph_has_num_pr(p_el: etree._Element) -> bool:
    ppr = first_child(p_el, "pPr")
    if ppr is None:
        return False
    return first_child(ppr, "numPr") is not None


def paragraph_num_level(p_el: etree._Element) -> int:
    ppr = first_child(p_el, "pPr")
    if ppr is None:
        return 0
    num_pr = first_child(ppr, "numPr")
    if num_pr is None:
        return 0
    ilvl = first_child(num_pr, "ilvl")
    if ilvl is None:
        return 0
    try:
        return int(get_attr(ilvl, "val") or "0")
    except ValueError:
        return 0


def set_num_pr(p_el: etree._Element, num_id: int, level: int) -> None:
    ppr = p_el.get_or_add_pPr() if hasattr(p_el, "get_or_add_pPr") else ensure_child(p_el, "pPr")
    remove_children(ppr, "numPr")
    num_pr = make("numPr")
    num_pr.append(make("ilvl", val=str(level)))
    num_pr.append(make("numId", val=str(num_id)))
    ppr.append(num_pr)


def has_page_break(p_el: etree._Element) -> bool:
    for br in p_el.iter(w("br")):
        if get_attr(br, "type") == "page":
            return True
    return False


def has_drawing(p_el: etree._Element) -> bool:
    return p_el.find(f".//{w('drawing')}") is not None or p_el.find(".//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline") is not None


def has_section_properties(p_el: etree._Element) -> bool:
    ppr = first_child(p_el, "pPr")
    if ppr is None:
        return False
    return first_child(ppr, "sectPr") is not None


def add_page_break_run(p_el: etree._Element) -> None:
    run = make("r")
    run.append(make("br", type="page"))
    p_el.append(run)


def add_page_field(parent_p: etree._Element, label_before: str = "Página ", label_mid: str | None = " de ") -> None:
    """Insere campo PAGE. Com label_mid=None, omite NUMPAGES (padrão ABNT)."""
    if label_before:
        _append_text_run(parent_p, label_before)
    _append_field(parent_p, "PAGE")
    if label_mid:
        _append_text_run(parent_p, label_mid)
        _append_field(parent_p, "NUMPAGES")


def _append_text_run(p_el: etree._Element, text: str) -> None:
    run = make("r")
    t = make("t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    run.append(t)
    p_el.append(run)


def _append_field(p_el: etree._Element, instr: str) -> None:
    begin = make("r")
    begin.append(make("fldChar", fldCharType="begin"))
    p_el.append(begin)

    instr_run = make("r")
    instr_el = make("instrText")
    instr_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr_el.text = f" {instr} "
    instr_run.append(instr_el)
    p_el.append(instr_run)

    sep = make("r")
    sep.append(make("fldChar", fldCharType="separate"))
    p_el.append(sep)

    result = make("r")
    t = make("t")
    t.text = "1"
    result.append(t)
    p_el.append(result)

    end = make("r")
    end.append(make("fldChar", fldCharType="end"))
    p_el.append(end)


def set_cell_shading(tc: etree._Element, fill_hex: str) -> None:
    tc_pr = ensure_child(tc, "tcPr")
    remove_children(tc_pr, "shd")
    color = fill_hex.lstrip("#").upper()
    tc_pr.append(make("shd", val="clear", color="auto", fill=color))


def set_cell_width(tc: etree._Element, dxa: int) -> None:
    tc_pr = ensure_child(tc, "tcPr")
    remove_children(tc_pr, "tcW")
    tc_pr.append(make("tcW", w=str(int(dxa)), type="dxa"))


def cm_to_dxa(cm: float) -> int:
    """Converte centímetros para twips (dxa). Nunca use int(Cm()) — isso devolve EMU."""
    from docx.shared import Cm

    return int(round(float(Cm(cm).twips)))


def set_table_width_dxa(tbl: etree._Element, total_dxa: int, cols: int) -> None:
    tbl_pr = ensure_child(tbl, "tblPr")
    remove_children(tbl_pr, "tblW")
    remove_children(tbl_pr, "tblLayout")
    tbl_pr.append(make("tblW", w=str(total_dxa), type="dxa"))
    tbl_pr.append(make("tblLayout", type="fixed"))
    col_widths = _split_dxa(total_dxa, cols)
    grid = first_child(tbl, "tblGrid")
    if grid is None:
        grid = make("tblGrid")
        tbl_pr.addnext(grid)
    for child in list(grid):
        if child.tag == w("gridCol"):
            grid.remove(child)
    for width in col_widths:
        grid.append(make("gridCol", w=str(width)))
    rows = [child for child in tbl if child.tag == w("tr")]
    for row in rows:
        cells = [child for child in row if child.tag == w("tc")]
        for cell, width in zip(cells, col_widths):
            set_cell_width(cell, width)


def set_table_full_width(tbl: etree._Element) -> None:
    """100% da área útil — evita tabela estourar a margem."""
    tbl_pr = ensure_child(tbl, "tblPr")
    remove_children(tbl_pr, "tblW")
    remove_children(tbl_pr, "tblLayout")
    tbl_pr.append(make("tblW", w="5000", type="pct"))
    tbl_pr.append(make("tblLayout", type="autofit"))


def _split_dxa(total: int, cols: int) -> list[int]:
    if cols <= 0:
        return []
    base = total // cols
    widths = [base] * cols
    widths[-1] += total - base * cols
    return widths


def mark_header_row(tr: etree._Element) -> None:
    tr_pr = ensure_child(tr, "trPr")
    if first_child(tr_pr, "tblHeader") is None:
        tr_pr.append(make("tblHeader", val="on"))


def set_table_borders(tbl: etree._Element, color: str = "000000", sz: int = 4) -> None:
    tbl_pr = ensure_child(tbl, "tblPr")
    remove_children(tbl_pr, "tblBorders")
    borders = make("tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(make(edge, val="single", sz=str(sz), space="0", color=color))
    tbl_pr.append(borders)


def iter_body_elements(body: etree._Element) -> Iterable[etree._Element]:
    for child in body.iterchildren():
        if child.tag in (w("p"), w("tbl")):
            yield child


def next_bookmark_id(root: etree._Element) -> int:
    max_id = 0
    for el in root.iter(w("bookmarkStart")):
        raw = get_attr(el, "id") or "0"
        try:
            max_id = max(max_id, int(raw))
        except ValueError:
            continue
    return max_id + 1


def attach_bookmark(block_el: etree._Element, name: str, bookmark_id: int) -> None:
    """Anexa bookmark estável ao parágrafo ou imediatamente ao redor da tabela."""
    start = make("bookmarkStart", **{"id": str(bookmark_id), "name": name})
    end = make("bookmarkEnd", **{"id": str(bookmark_id)})
    if block_el.tag == w("p"):
        block_el.insert(0, start)
        block_el.append(end)
        return
    parent = block_el.getparent()
    if parent is None:
        return
    index = list(parent).index(block_el)
    parent.insert(index, start)
    parent.insert(index + 2, end)


def find_bookmark(root: etree._Element, name: str) -> Optional[etree._Element]:
    for el in root.iter(w("bookmarkStart")):
        if get_attr(el, "name") == name:
            return el
    return None


def element_for_bookmark(bookmark: etree._Element) -> Optional[etree._Element]:
    parent = bookmark.getparent()
    if parent is None:
        return None
    if parent.tag == w("p"):
        return parent
    sibling = bookmark.getnext()
    while sibling is not None:
        if sibling.tag in (w("p"), w("tbl")):
            return sibling
        if sibling.tag == w("bookmarkEnd"):
            break
        sibling = sibling.getnext()
    return parent if parent.tag in (w("p"), w("tbl")) else None


def block_cluster(element: etree._Element) -> list[etree._Element]:
    """Retorna o elemento e bookmarks imediatamente associados (necessário para tabelas)."""
    nodes: list[etree._Element] = []
    if element.tag == w("tbl"):
        prev = element.getprevious()
        if prev is not None and prev.tag == w("bookmarkStart"):
            nodes.append(prev)
        nodes.append(element)
        nxt = element.getnext()
        if nxt is not None and nxt.tag == w("bookmarkEnd"):
            nodes.append(nxt)
        return nodes
    return [element]


def relocate_cluster(
    element: etree._Element,
    *,
    after: etree._Element | None = None,
    before: etree._Element | None = None,
) -> None:
    nodes = block_cluster(element)
    if before is not None:
        for node in reversed(nodes):
            before.addprevious(node)
        return
    if after is not None:
        for node in reversed(nodes):
            after.addnext(node)


def remove_bookmark_pair(root: etree._Element, name: str) -> None:
    start = find_bookmark(root, name)
    if start is None:
        return
    bid = get_attr(start, "id")
    parent = start.getparent()
    if parent is not None:
        parent.remove(start)
    if bid is None:
        return
    for el in list(root.iter(w("bookmarkEnd"))):
        if get_attr(el, "id") == bid:
            end_parent = el.getparent()
            if end_parent is not None:
                end_parent.remove(el)
            break


def clear_element_text_runs(p_el: etree._Element) -> None:
    for child in list(p_el):
        if child.tag == w("r"):
            p_el.remove(child)
