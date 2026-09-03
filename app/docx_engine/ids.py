"""Identificadores estáveis persistidos como bookmarks nativos do Word."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

from docx.document import Document as DocxDocument
from lxml import etree

from app.docx_engine import xml_utils as xu

# Prefixo com underscore: bookmark oculto no Word (não aparece como marcação cinza/vermelha).
BOOKMARK_PREFIX = "_Aidoc_"
LEGACY_BOOKMARK_PREFIX = "Aidoc_"
ID_RE = re.compile(r"^([a-z]+)_(\d+)$")


@dataclass
class IdRegistry:
    counters: dict[str, int] = field(default_factory=dict)
    next_bookmark: int = 1

    def allocate(self, kind: str) -> str:
        n = self.counters.get(kind, 0) + 1
        self.counters[kind] = n
        return f"{kind}_{n:03d}"

    def observe(self, block_id: str) -> None:
        match = ID_RE.match(block_id)
        if not match:
            return
        kind, raw = match.group(1), int(match.group(2))
        self.counters[kind] = max(self.counters.get(kind, 0), raw)


def bookmark_name(block_id: str) -> str:
    return f"{BOOKMARK_PREFIX}{block_id}"


def bookmark_names(block_id: str) -> tuple[str, ...]:
    return (f"{BOOKMARK_PREFIX}{block_id}", f"{LEGACY_BOOKMARK_PREFIX}{block_id}")


def block_id_from_bookmark(name: str | None) -> Optional[str]:
    if not name:
        return None
    for prefix in (BOOKMARK_PREFIX, LEGACY_BOOKMARK_PREFIX):
        if name.startswith(prefix):
            return name[len(prefix) :]
    return None


def scan_registry(document: DocxDocument) -> IdRegistry:
    registry = IdRegistry()
    root = document.element
    registry.next_bookmark = xu.next_bookmark_id(root)
    for el in root.iter(xu.w("bookmarkStart")):
        bid = block_id_from_bookmark(xu.get_attr(el, "name"))
        if bid:
            registry.observe(bid)
    return registry


def ensure_block_id(element: etree._Element, kind: str, registry: IdRegistry, root: etree._Element) -> str:
    existing = get_block_id(element, root)
    if existing:
        registry.observe(existing)
        return existing
    block_id = registry.allocate(kind)
    xu.attach_bookmark(element, bookmark_name(block_id), registry.next_bookmark)
    registry.next_bookmark += 1
    return block_id


def get_block_id(element: etree._Element, root: etree._Element) -> Optional[str]:
    if element.tag == xu.w("p"):
        for child in element:
            if child.tag == xu.w("bookmarkStart"):
                bid = block_id_from_bookmark(xu.get_attr(child, "name"))
                if bid:
                    return bid
        return None
    prev = element.getprevious()
    while prev is not None:
        if prev.tag == xu.w("bookmarkStart"):
            bid = block_id_from_bookmark(xu.get_attr(prev, "name"))
            if bid:
                return bid
        if prev.tag in (xu.w("p"), xu.w("tbl")):
            break
        prev = prev.getprevious()
    return None


def find_element(document: DocxDocument, block_id: str) -> Optional[etree._Element]:
    for name in bookmark_names(block_id):
        bookmark = xu.find_bookmark(document.element, name)
        if bookmark is not None:
            return xu.element_for_bookmark(bookmark)
    return None


def remove_id(document: DocxDocument, block_id: str) -> None:
    for name in bookmark_names(block_id):
        xu.remove_bookmark_pair(document.element, name)
