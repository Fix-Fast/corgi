"""Rule 1: text hierarchy + heading/body styling.

Per format.md §1:

    title -> Heading 1
    section heading -> Heading 2
    subheading, coverage heading, or insuring-agreement heading -> Heading 3
    body text -> ordinary paragraph

    Title styling:        Bricolage Grotesque ExtraBold, 23pt, centered, 10pt after
    Section heading:      Bricolage Grotesque, 14pt, bold, left,    16pt before, 8pt after
                          (carried by the Heading 2 style def, not run-level)
    Subheading:           Bricolage Grotesque, 13pt, bold, left,    10pt after
    Body text:            Inter, 11pt, black, left,                 10pt after

`ignored_body_texts` paragraphs are removed from the document.
"""
from __future__ import annotations

from lxml import etree

from _docx import (
    RunFormat,
    append_page_break,
    apply_run_format_to_all,
    clear_pPr_child,
    clear_run_format_overrides,
    get_body,
    iter_paragraphs,
    remove_paragraph,
    set_alignment,
    set_pstyle,
    set_spacing,
    strip_following_redundant_page_breaks,
    upsert_doc_default_run_format,
    upsert_paragraph_style,
)
from parts import ResolvedParts


TITLE_STYLE_ID = "Heading1"
SECTION_STYLE_ID = "Heading2"
SUBHEAD_STYLE_ID = "Heading3"
BODY_STYLE_ID = "BodyText"


_TITLE_FMT = RunFormat(font="Bricolage Grotesque ExtraBold", size_pt=23, bold=False, color_hex="000000")
_SECTION_FMT = RunFormat(font="Bricolage Grotesque", size_pt=14, bold=True, color_hex="000000")
_SUBHEAD_FMT = RunFormat(font="Bricolage Grotesque", size_pt=13, bold=True, color_hex="000000")
_BODY_FMT = RunFormat(font="Inter", size_pt=11, color_hex="000000")
_NOTICES_FMT = RunFormat(font="Inter", size_pt=13, bold=True, color_hex="000000")


def _style_paragraph(
    p: etree._Element,
    *,
    style_id: str,
    fmt: RunFormat | None,
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
    if fmt is not None:
        apply_run_format_to_all(p, fmt)
    else:
        # Style-def driven: strip any source run-level overrides so the
        # style is the single source of truth.
        clear_run_format_overrides(p)


def _inject_doc_defaults(styles_root: etree._Element) -> None:
    """Set the doc-default rPr to body styling (Inter 11pt black).

    Every element in the doc inherits this unless it overrides — most
    importantly the list markers, whose canonical abstractNum levels
    omit run formatting and rely on inheritance to pick up the body
    look. Headings (Heading 1/2/3) override font/size/bold so they're
    unaffected.
    """
    upsert_doc_default_run_format(styles_root, _BODY_FMT)


def _inject_section_style_def(styles_root: etree._Element) -> None:
    """Upsert the Heading 2 style def carrying section-heading formatting.

    Section headings are styled via this definition rather than via run-level
    overrides — the runs themselves are left bare so the style def is the
    single source of truth for section-heading appearance.
    """
    upsert_paragraph_style(
        styles_root,
        style_id=SECTION_STYLE_ID,
        name="heading 2",
        based_on="Normal",
        next_style="Normal",
        fmt=_SECTION_FMT,
        alignment="left",
        space_before_pt=16,
        space_after_pt=8,
    )


def apply(
    doc_root: etree._Element,
    parts: ResolvedParts,
    styles_root: etree._Element,
) -> etree._Element:
    """Apply Rule 1 in place. Mutates doc_root and styles_root, returns doc_root."""
    _inject_doc_defaults(styles_root)
    _inject_section_style_def(styles_root)

    body = get_body(doc_root)
    paragraphs = list(iter_paragraphs(body))

    # Drop ignored paragraphs first so subsequent indices still match —
    # we resolved indices against the original document, so we must use
    # the original list and remove elements without re-indexing.
    for idx in sorted(parts.ignored_indices, reverse=True):
        remove_paragraph(paragraphs[idx])

    # Identify the last notices-block paragraph by document order so we
    # can append a hard page break inside it (per format.md §1).
    last_notice_idx = max(parts.notices_block_indices) if parts.notices_block_indices else None

    for idx, p in enumerate(paragraphs):
        if idx in parts.ignored_indices:
            continue
        if idx in parts.title_indices:
            _style_paragraph(
                p, style_id=TITLE_STYLE_ID, fmt=_TITLE_FMT,
                alignment="center", space_before_pt=None, space_after_pt=10,
                clear_indent=True,
            )
        elif idx in parts.section_indices:
            # Section headings are style-def driven (see _inject_section_style_def).
            # The runs are left bare; appearance is carried by the Heading 2 style.
            _style_paragraph(
                p, style_id=SECTION_STYLE_ID, fmt=None,
                alignment="left", space_before_pt=16, space_after_pt=8,
                clear_indent=True,
            )
        elif idx in parts.subheading_indices:
            _style_paragraph(
                p, style_id=SUBHEAD_STYLE_ID, fmt=_SUBHEAD_FMT,
                alignment="left", space_before_pt=None, space_after_pt=10,
                clear_indent=True,
            )
        elif idx in parts.notices_block_indices:
            _style_paragraph(
                p, style_id=BODY_STYLE_ID, fmt=_NOTICES_FMT,
                alignment="left", space_before_pt=None, space_after_pt=10,
                clear_indent=True,
            )
            if idx == last_notice_idx:
                append_page_break(p)
                # Source docs often carry their own empty page-break
                # paragraph after a notices block; remove its redundant
                # break so we don't render a blank page.
                strip_following_redundant_page_breaks(p)
        else:
            # Body paragraphs keep their indent — Rule 2 owns indent on
            # list items and continuation paragraphs.
            _style_paragraph(
                p, style_id=BODY_STYLE_ID, fmt=_BODY_FMT,
                alignment="left", space_before_pt=None, space_after_pt=10,
                clear_indent=False,
            )

    return doc_root
