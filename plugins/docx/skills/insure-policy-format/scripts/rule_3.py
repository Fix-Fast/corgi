"""Rule 3: page layout + running header.

Per format.md §3:

    Margins:    top 1.0"  bottom 1.0"  left 1.0"  right 1.0"
    Header dist: 0.5"
    Header text: <Title><TAB><Policy Code>
    Header para: left-aligned, right-aligned tab stop at 6.3",
                 Inter, 10pt, gray RGB 128,128,128

If only one of <title>/<policy_code> is supplied, the header still uses
the available value (no tab in that case).
"""
from __future__ import annotations

from lxml import etree

from _docx import (
    Doc,
    NSMAP,
    RunFormat,
    W,
    apply_run_format,
    get_or_create_rPr,
    inches_to_twips,
    make_element,
)
from parts import ResolvedParts


HEADER_PART_NAME = "header_corgi.xml"
HEADER_REL_ID = "rIdCorgiHeader"
HEADER_RELS_TARGET = HEADER_PART_NAME


_HEADER_FMT = RunFormat(font="Inter", size_pt=10, color_hex="808080")


def _set_section_geometry(sectPr: etree._Element) -> None:
    """Set pgMar and remove any existing header/footer references."""
    # Page margins.
    for el in sectPr.findall(W + "pgMar"):
        sectPr.remove(el)
    sectPr.append(make_element("pgMar", {
        "top": str(inches_to_twips(1.0)),
        "bottom": str(inches_to_twips(1.0)),
        "left": str(inches_to_twips(1.0)),
        "right": str(inches_to_twips(1.0)),
        "header": str(inches_to_twips(0.5)),
        "footer": str(inches_to_twips(0.5)),
        "gutter": "0",
    }))


def _replace_header_reference(sectPr: etree._Element, rel_id: str) -> None:
    """Drop any existing headerReference children and add a single 'default' one."""
    for el in sectPr.findall(W + "headerReference"):
        sectPr.remove(el)
    href = make_element("headerReference", {"type": "default"})
    href.set(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
        rel_id,
    )
    # headerReference belongs at the start of sectPr per OOXML; insert at index 0.
    sectPr.insert(0, href)


def _build_header_part(header_title: str | None, policy_code: str | None) -> etree._Element:
    """Build the root <w:hdr> element for the running header."""
    nsmap = {
        "w": NSMAP["w"],
        "xml": "http://www.w3.org/XML/1998/namespace",
    }
    hdr = etree.Element(W + "hdr", nsmap=nsmap)
    p = etree.SubElement(hdr, W + "p")

    pPr = etree.SubElement(p, W + "pPr")
    tabs = etree.SubElement(pPr, W + "tabs")
    tab = etree.SubElement(tabs, W + "tab")
    tab.set(W + "val", "right")
    tab.set(W + "pos", str(inches_to_twips(6.3)))
    jc = etree.SubElement(pPr, W + "jc")
    jc.set(W + "val", "left")

    title = (header_title or "").strip()
    code = (policy_code or "").strip()
    if title and code:
        _add_run(p, title)
        _add_tab(p)
        _add_run(p, code)
    elif title:
        _add_run(p, title)
    elif code:
        _add_run(p, code)
    # else: empty header paragraph (still valid).

    return hdr


def _add_run(p: etree._Element, text: str) -> None:
    r = etree.SubElement(p, W + "r")
    apply_run_format(r, _HEADER_FMT)
    t = etree.SubElement(r, W + "t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text


def _add_tab(p: etree._Element) -> None:
    r = etree.SubElement(p, W + "r")
    apply_run_format(r, _HEADER_FMT)
    etree.SubElement(r, W + "tab")


def apply(doc: Doc, parts: ResolvedParts) -> Doc:
    """Apply Rule 3 in place. Mutates doc.document and sets doc.header.

    The caller is responsible for serializing doc.header into the docx
    zip at word/header_corgi.xml and updating the relationships /
    content-types parts (those live outside this rule's purview because
    they cross the in-memory tree boundary).
    """
    body = doc.document.find(W + "body")
    if body is None:
        raise ValueError("document.xml has no <w:body>")
    sectPr = body.find(W + "sectPr")
    if sectPr is None:
        sectPr = make_element("sectPr")
        body.append(sectPr)
    _set_section_geometry(sectPr)
    _replace_header_reference(sectPr, HEADER_REL_ID)

    doc.header = _build_header_part(parts.header_title_text, parts.policy_code)
    return doc
