"""Rule 1: text hierarchy + heading/body styling.

Per format.md §1:

    title -> Heading 1
    section heading -> Heading 2
    subheading, coverage heading, or insuring-agreement heading -> Heading 3
    body text -> ordinary paragraph

    Title styling:        Bricolage Grotesque, 26pt, bold, centered, 18pt space after
    Section heading:      Bricolage Grotesque, 14pt, bold, left,    16pt before, 8pt after
    Subheading:           Bricolage Grotesque, 12pt, bold, left,    10pt before, 6pt after
    Body text:            Inter, 11pt, black, left,                 6pt after

`ignored_body_texts` paragraphs are removed from the document.
"""
from __future__ import annotations

from lxml import etree

from _docx import (
    RunFormat,
    apply_run_format_to_all,
    clear_pPr_child,
    get_body,
    iter_paragraphs,
    remove_paragraph,
    set_alignment,
    set_pstyle,
    set_spacing,
)
from parts import ResolvedParts


TITLE_STYLE_ID = "Heading1"
SECTION_STYLE_ID = "Heading2"
SUBHEAD_STYLE_ID = "Heading3"
BODY_STYLE_ID = "BodyText"


_TITLE_FMT = RunFormat(font="Bricolage Grotesque", size_pt=26, bold=True, color_hex="000000")
_SECTION_FMT = RunFormat(font="Bricolage Grotesque", size_pt=14, bold=True, color_hex="000000")
_SUBHEAD_FMT = RunFormat(font="Bricolage Grotesque", size_pt=12, bold=True, color_hex="000000")
_BODY_FMT = RunFormat(font="Inter", size_pt=11, color_hex="000000")


def _style_paragraph(
    p: etree._Element,
    *,
    style_id: str,
    fmt: RunFormat,
    alignment: str,
    space_before_pt: float | None,
    space_after_pt: float,
    clear_indent: bool,
) -> None:
    set_pstyle(p, style_id)
    set_alignment(p, alignment)
    set_spacing(p, space_before_pt, space_after_pt)
    if clear_indent:
        clear_pPr_child(p, "ind")
    apply_run_format_to_all(p, fmt)


def apply(doc_root: etree._Element, parts: ResolvedParts) -> etree._Element:
    """Apply Rule 1 in place. Returns the same root."""
    body = get_body(doc_root)
    paragraphs = list(iter_paragraphs(body))

    # Drop ignored paragraphs first so subsequent indices still match —
    # we resolved indices against the original document, so we must use
    # the original list and remove elements without re-indexing.
    for idx in sorted(parts.ignored_indices, reverse=True):
        remove_paragraph(paragraphs[idx])

    for idx, p in enumerate(paragraphs):
        if idx in parts.ignored_indices:
            continue
        if idx in parts.title_indices:
            _style_paragraph(
                p, style_id=TITLE_STYLE_ID, fmt=_TITLE_FMT,
                alignment="center", space_before_pt=None, space_after_pt=18,
                clear_indent=True,
            )
        elif idx in parts.section_indices:
            _style_paragraph(
                p, style_id=SECTION_STYLE_ID, fmt=_SECTION_FMT,
                alignment="left", space_before_pt=16, space_after_pt=8,
                clear_indent=True,
            )
        elif idx in parts.subheading_indices:
            _style_paragraph(
                p, style_id=SUBHEAD_STYLE_ID, fmt=_SUBHEAD_FMT,
                alignment="left", space_before_pt=10, space_after_pt=6,
                clear_indent=True,
            )
        else:
            # Body paragraphs keep their indent — Rule 2 owns indent on
            # list items and continuation paragraphs.
            _style_paragraph(
                p, style_id=BODY_STYLE_ID, fmt=_BODY_FMT,
                alignment="left", space_before_pt=None, space_after_pt=6,
                clear_indent=False,
            )

    return doc_root
