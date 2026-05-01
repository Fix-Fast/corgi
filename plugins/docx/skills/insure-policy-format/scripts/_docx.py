"""Low-level OOXML / python-docx helpers.

No business logic — just primitives the rules compose. Each helper either
returns a value or mutates the passed-in tree element. Rules are
responsible for ordering and side-effect choreography.

Conventions:

- Everything in the wordprocessingml namespace is referenced via the W
  prefix below. Use ``W + "tagname"`` as the lxml element tag.
- Twips are 1/20 of a point. Word stores spacing/indent in twips and
  font size in half-points; the converter helpers below make this
  explicit.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterator

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NSMAP = {"w": W_NS}


def pt_to_twips(pt: float) -> int:
    return int(round(pt * 20))


def pt_to_half_pt(pt: float) -> int:
    return int(round(pt * 2))


def inches_to_twips(inches: float) -> int:
    return int(round(inches * 1440))


@dataclass
class RunFormat:
    """Run-level character formatting block applied as a unit.

    Any field set to None is left untouched on the run; any field set to
    a value overrides the run's existing setting. Bold/bCs use True to
    set, False to clear, None to leave alone.
    """

    font: str | None = None
    size_pt: float | None = None
    bold: bool | None = None
    color_hex: str | None = None  # 6-digit hex e.g. "000000"


def make_element(tag: str, attrs: dict[str, str] | None = None) -> etree._Element:
    el = etree.Element(W + tag)
    if attrs:
        for k, v in attrs.items():
            el.set(W + k, v)
    return el


def iter_paragraphs(body: etree._Element) -> Iterator[etree._Element]:
    """Yield every non-empty <w:p> in body order, including descents into <w:tbl>/<w:sdt>.

    Empty paragraphs (no <w:t> text and no break elements) are skipped:
    they are layout artifacts (manual blank lines) and not addressable
    by parts.json substring matching, so they should not occupy an
    index slot in the document order downstream rules see.
    """
    for el in body.iter(W + "p"):
        if not paragraph_text(el).strip():
            continue
        yield el


def paragraph_text(p: etree._Element) -> str:
    """Concatenate every <w:t> descendant of the paragraph as a single string.

    Tabs and breaks are reduced to spaces. Result is whitespace-collapsed
    upstream by the caller (see parts.normalize_text).
    """
    out: list[str] = []
    for el in p.iter():
        tag = el.tag
        if tag == W + "t":
            out.append(el.text or "")
        elif tag == W + "tab":
            out.append(" ")
        elif tag in (W + "br", W + "cr"):
            out.append(" ")
    return "".join(out)


def get_or_create_pPr(p: etree._Element) -> etree._Element:
    pPr = p.find(W + "pPr")
    if pPr is None:
        pPr = make_element("pPr")
        p.insert(0, pPr)
    return pPr


def set_pstyle(p: etree._Element, style_id: str) -> None:
    """Set <w:pStyle w:val="style_id"/> on the paragraph, replacing any existing."""
    pPr = get_or_create_pPr(p)
    existing = pPr.find(W + "pStyle")
    if existing is not None:
        pPr.remove(existing)
    pStyle = make_element("pStyle", {"val": style_id})
    pPr.insert(0, pStyle)


def clear_pPr_child(p: etree._Element, tag: str) -> None:
    """Remove every direct child of pPr with the given tag."""
    pPr = p.find(W + "pPr")
    if pPr is None:
        return
    for el in pPr.findall(W + tag):
        pPr.remove(el)


def set_alignment(p: etree._Element, val: str) -> None:
    """Set <w:jc w:val="..."/> ('left', 'center', 'right', 'both')."""
    pPr = get_or_create_pPr(p)
    existing = pPr.find(W + "jc")
    if existing is not None:
        pPr.remove(existing)
    pPr.append(make_element("jc", {"val": val}))


def set_spacing(p: etree._Element, before_pt: float | None, after_pt: float | None) -> None:
    """Set <w:spacing w:before/w:after> in twips. None leaves the field absent."""
    pPr = get_or_create_pPr(p)
    existing = pPr.find(W + "spacing")
    if existing is not None:
        pPr.remove(existing)
    attrs: dict[str, str] = {}
    if before_pt is not None:
        attrs["before"] = str(pt_to_twips(before_pt))
    if after_pt is not None:
        attrs["after"] = str(pt_to_twips(after_pt))
    if attrs:
        pPr.append(make_element("spacing", attrs))


def set_indent(p: etree._Element, left_twips: int | None, hanging_twips: int | None) -> None:
    """Set <w:ind w:left/w:hanging>. Pass None to remove the override."""
    pPr = get_or_create_pPr(p)
    existing = pPr.find(W + "ind")
    if existing is not None:
        pPr.remove(existing)
    if left_twips is None and hanging_twips is None:
        return
    attrs: dict[str, str] = {}
    if left_twips is not None:
        attrs["left"] = str(left_twips)
    if hanging_twips is not None:
        attrs["hanging"] = str(hanging_twips)
    pPr.append(make_element("ind", attrs))


def set_numPr(p: etree._Element, num_id: int, ilvl: int) -> None:
    """Set <w:numPr> on the paragraph (replacing existing)."""
    pPr = get_or_create_pPr(p)
    existing = pPr.find(W + "numPr")
    if existing is not None:
        pPr.remove(existing)
    numPr = make_element("numPr")
    numPr.append(make_element("ilvl", {"val": str(ilvl)}))
    numPr.append(make_element("numId", {"val": str(num_id)}))
    pPr.append(numPr)


def has_numPr(p: etree._Element) -> bool:
    pPr = p.find(W + "pPr")
    return pPr is not None and pPr.find(W + "numPr") is not None


def get_pstyle(p: etree._Element) -> str | None:
    pPr = p.find(W + "pPr")
    if pPr is None:
        return None
    pStyle = pPr.find(W + "pStyle")
    if pStyle is None:
        return None
    return pStyle.get(W + "val")


def iter_runs(p: etree._Element) -> Iterator[etree._Element]:
    return iter(p.findall(W + "r"))


def get_or_create_rPr(r: etree._Element) -> etree._Element:
    rPr = r.find(W + "rPr")
    if rPr is None:
        rPr = make_element("rPr")
        r.insert(0, rPr)
    return rPr


def apply_run_format(r: etree._Element, fmt: RunFormat) -> None:
    """Apply RunFormat to a single <w:r>. Non-None fields override; None leaves alone.

    Existing fields not addressed by RunFormat (e.g. <w:i/>, language tags)
    are left in place.
    """
    rPr = get_or_create_rPr(r)
    if fmt.font is not None:
        existing = rPr.find(W + "rFonts")
        if existing is not None:
            rPr.remove(existing)
        rPr.append(
            make_element("rFonts", {
                "ascii": fmt.font,
                "hAnsi": fmt.font,
                "eastAsia": fmt.font,
                "cs": fmt.font,
            })
        )
    if fmt.bold is True:
        if rPr.find(W + "b") is None:
            rPr.append(make_element("b"))
        if rPr.find(W + "bCs") is None:
            rPr.append(make_element("bCs"))
    elif fmt.bold is False:
        for tag in ("b", "bCs"):
            for el in rPr.findall(W + tag):
                rPr.remove(el)
    if fmt.color_hex is not None:
        existing = rPr.find(W + "color")
        if existing is not None:
            rPr.remove(existing)
        rPr.append(make_element("color", {"val": fmt.color_hex}))
    if fmt.size_pt is not None:
        for tag in ("sz", "szCs"):
            existing = rPr.find(W + tag)
            if existing is not None:
                rPr.remove(existing)
        sz_val = str(pt_to_half_pt(fmt.size_pt))
        rPr.append(make_element("sz", {"val": sz_val}))
        rPr.append(make_element("szCs", {"val": sz_val}))


def apply_run_format_to_all(p: etree._Element, fmt: RunFormat) -> None:
    """Apply RunFormat to every <w:r> in the paragraph."""
    for r in iter_runs(p):
        apply_run_format(r, fmt)


def get_body(doc_root: etree._Element) -> etree._Element:
    body = doc_root.find(W + "body")
    if body is None:
        raise ValueError("document.xml has no <w:body>")
    return body


def get_sectPr(body: etree._Element) -> etree._Element | None:
    """Return the body-level sectPr, if any. (Per-paragraph sectPrs are not handled here.)"""
    return body.find(W + "sectPr")


def remove_paragraph(p: etree._Element) -> None:
    parent = p.getparent()
    if parent is not None:
        parent.remove(p)


def insert_after(reference: etree._Element, new: etree._Element) -> None:
    parent = reference.getparent()
    if parent is None:
        raise ValueError("reference element has no parent")
    parent.insert(list(parent).index(reference) + 1, new)


def replace_paragraph_text(p: etree._Element, new_text: str) -> None:
    """Replace all <w:t> content in p with a single run containing new_text.

    Wipes existing runs (and their formatting). Caller must reapply
    RunFormat afterward if formatting is wanted.
    """
    for r in p.findall(W + "r"):
        p.remove(r)
    for hyperlink in p.findall(W + "hyperlink"):
        p.remove(hyperlink)
    r = make_element("r")
    t = make_element("t")
    t.text = new_text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    p.append(r)
