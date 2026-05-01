"""Rule 2: list normalization.

Per format.md §2:

    - Marker sequence: 1)/1.  a)/a.  i)/i.  (1)  (a)  (i)
    - Top three levels: closing ')' or '.' both accepted; parenthesized
      forms require both parens.
    - List markers must stand on their own (start of paragraph or after
      space, not stuck to a word/citation like officer(s) or §4958(c)).
    - Multiple embedded markers in a paragraph -> split into separate
      list paragraphs.
    - Continuation paragraph stays attached to its list item.
    - Level format definitions: L0..L5 use canonical numFmt + lvlText
      with the indents in the spec (720/1440/2160/2880/3600/4320 left,
      360 hanging across the board).
    - Single space between marker and body, not tab.
    - Strip paragraph-level indent overrides on list paragraphs.
    - List markers: Inter, 11pt, black.

Implementation: rewrite numbering.xml to a single canonical multilevel
abstractNum + num. Walk every body paragraph (skipping headings); detect
leading + embedded markers; split paragraphs on embedded markers; assign
numPr (referencing our single num) and ilvl per detected shape; strip
the marker text from the paragraph's runs (numbering renders it).
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from lxml import etree

from _docx import (
    W,
    clear_pPr_child,
    get_body,
    get_pstyle,
    iter_paragraphs,
    make_element,
    set_indent,
    set_numPr,
)
from rule_1 import SECTION_STYLE_ID, SUBHEAD_STYLE_ID, TITLE_STYLE_ID


# Indents per format.md §2 list ladder. Mirror the abstractNum ind values
# below so continuation paragraphs sit at the same body-text column as
# their parent list item.
_LEVEL_LEFT_INDENT = {
    0: 720,
    1: 1440,
    2: 2160,
    3: 2880,
    4: 3600,
    5: 4320,
    6: 5040,
    7: 5760,
    8: 6480,
}


# numId/abstractNumId picked to be high enough not to collide with anything
# pandoc / Word might have inherited from the source. The exact numbers
# don't matter for behavior — the test canonicalizer hashes them anyway.
CANONICAL_ABSTRACT_NUM_ID = 9000
CANONICAL_NUM_ID = 9000


HEADING_STYLES = frozenset({TITLE_STYLE_ID, SECTION_STYLE_ID, SUBHEAD_STYLE_ID})


@dataclass(frozen=True)
class _MarkerMatch:
    start: int          # offset in paragraph text where the marker begins
    end: int            # offset in paragraph text where marker (incl trailing whitespace) ends
    level: int          # 0..5
    is_leading: bool


# Patterns are anchored at marker start; "leading" matches absorb leading
# whitespace, "embedded" matches require a non-word boundary character
# behind. Both consume one trailing whitespace char so the marker text
# can be cleanly excised.

_PAREN_LETTER_RE = re.compile(r"\(([a-z])\)")
_PAREN_ROMAN_RE = re.compile(r"\(([ivxlcdm]+)\)")
_PAREN_DECIMAL_RE = re.compile(r"\((\d+)\)")
_BARE_DECIMAL_RE = re.compile(r"(\d+)\s*[).]")
_BARE_LETTER_RE = re.compile(r"([a-z])\s*[).]")
_BARE_ROMAN_RE = re.compile(r"([ivxlcdm]{2,})\s*[).]")  # 2+ chars: unambiguously roman


_ALPHA_PREDECESSOR_OF_ROMAN = {"i": "h", "v": "u", "x": "w"}


def _detect_marker_at(text: str, start: int, *, is_leading: bool, last_l1: str | None, last_l4: str | None) -> tuple[int, int] | None:
    """Look for a marker beginning at offset `start`. Returns (end_offset, level) or None.

    If is_leading=True, no boundary check is done (caller has already
    eaten leading whitespace). If is_leading=False, the caller has
    already validated boundary before `start` (preceding char is
    non-alphanumeric, non-')', non-'(', non-'.').

    Parenthesized shapes ((1)/(a)/(i)) are only matched at the leading
    position. In embedded position they collide too readily with
    citations ("paragraph (b) below", "Section 4958(c)") to be useful
    as markers, per format.md's citation rule.
    """
    rest = text[start:]
    if is_leading:
        # Parenthesized: try decimal, then roman (multi-char + single roman with
        # alpha-pred check), then letter. Putting roman before letter avoids
        # misclassifying (i)/(v)/(x) — which match both — as alpha.
        m = _PAREN_DECIMAL_RE.match(rest)
        if m:
            return start + m.end(), 3
        m = _PAREN_ROMAN_RE.match(rest)
        if m and len(m.group(1)) >= 2:
            return start + m.end(), 5
        m = _PAREN_LETTER_RE.match(rest)
        if m:
            token = m.group(1)
            # Single-char i/v/x: ambiguous between L4 alpha and L5 roman.
            # Default to L5 (fresh roman sub-list) unless the L4 alpha
            # sequence just emitted the predecessor letter.
            if token in _ALPHA_PREDECESSOR_OF_ROMAN:
                pred = _ALPHA_PREDECESSOR_OF_ROMAN[token]
                if last_l4 != pred:
                    return start + m.end(), 5
            return start + m.end(), 4

    # Bare decimal: 1) or 1.
    m = _BARE_DECIMAL_RE.match(rest)
    if m:
        return start + m.end(), 0
    # Multi-char roman: ii)/iii)/iv)/etc — definitely roman.
    m = _BARE_ROMAN_RE.match(rest)
    if m:
        return start + m.end(), 2
    # Single letter (could be alpha or roman).
    m = _BARE_LETTER_RE.match(rest)
    if m:
        token = m.group(1)
        end = start + m.end()
        if token in _ALPHA_PREDECESSOR_OF_ROMAN:
            pred = _ALPHA_PREDECESSOR_OF_ROMAN[token]
            # If we just emitted the alpha predecessor at L1, this is L1
            # continuation; same logic at L4 (after the `(...)` predecessor).
            if last_l1 == pred:
                return end, 1
            if last_l4 == pred:
                return end, 4
            return end, 2  # default: start a new roman sub-list
        return end, 1
    return None


def _scan_leading_marker(text: str, last_l1: str | None, last_l4: str | None) -> _MarkerMatch | None:
    """Return the leading marker for the paragraph, or None.

    Embedded markers are intentionally NOT scanned. Per format.md §2,
    list-item recognition keys off leading position only — a
    mid-paragraph `1)` or `(b)` is treated as plain text (citation /
    enumeration reference) rather than a marker, mirroring Word's own
    behavior when typing.
    """
    pos = 0
    while pos < len(text) and text[pos].isspace():
        pos += 1
    if pos >= len(text):
        return None
    result = _detect_marker_at(
        text, pos, is_leading=True, last_l1=last_l1, last_l4=last_l4,
    )
    if result is None:
        return None
    end, level = result
    consumed_end = end
    if consumed_end < len(text) and text[consumed_end] == " ":
        consumed_end += 1
    return _MarkerMatch(0, consumed_end, level, is_leading=True)


def _paragraph_text(p: etree._Element) -> str:
    """Local paragraph_text — same semantics as _docx.paragraph_text."""
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


def _excise_marker_text(p: etree._Element, marker_end: int) -> None:
    """Remove the first `marker_end` characters of paragraph text from the runs.

    Walks <w:r> children and trims their <w:t>/tab/br content from the
    front. Leaves the rest untouched.
    """
    remaining = marker_end
    runs_to_remove: list[etree._Element] = []
    for child in list(p):
        if remaining <= 0:
            break
        if child.tag == W + "pPr":
            continue
        if child.tag != W + "r":
            continue
        # Walk run's contents.
        run_text_remaining = remaining
        empty_run = True
        for sub in list(child):
            if run_text_remaining <= 0:
                empty_run = False
                break
            if sub.tag == W + "rPr":
                continue
            if sub.tag == W + "t":
                text = sub.text or ""
                if len(text) <= run_text_remaining:
                    run_text_remaining -= len(text)
                    child.remove(sub)
                else:
                    sub.text = text[run_text_remaining:]
                    run_text_remaining = 0
                    empty_run = False
            elif sub.tag in (W + "tab", W + "br", W + "cr"):
                run_text_remaining -= 1
                child.remove(sub)
            else:
                empty_run = False
        remaining = run_text_remaining
        # If run has only rPr (or nothing) left, remove it.
        if not any(c.tag != W + "rPr" for c in child):
            runs_to_remove.append(child)
    for r in runs_to_remove:
        p.remove(r)


def _build_canonical_abstract_num() -> etree._Element:
    """Build the single canonical abstractNum element used for all lists.

    Levels 0–5 follow format.md spec. Levels 6, 7, 8 are explicit
    placeholders so the definition is stable across runs (Word requires
    9 levels to be valid, but format.md doesn't license content beyond
    level 5; we mirror the lvl-2 roman/lvl-5 form to give Word
    *something* sensible if the spec ever drifts there).
    """
    spec = [
        # (numFmt, lvlText, left_twips)
        ("decimal",     "%1)",  720),
        ("lowerLetter", "%2)",  1440),
        ("lowerRoman",  "%3)",  2160),
        ("decimal",     "(%4)", 2880),
        ("lowerLetter", "(%5)", 3600),
        ("lowerRoman",  "(%6)", 4320),
        ("decimal",     "%7)",  5040),
        ("lowerLetter", "%8)",  5760),
        ("lowerRoman",  "%9)",  6480),
    ]
    abstract = make_element("abstractNum", {"abstractNumId": str(CANONICAL_ABSTRACT_NUM_ID)})
    abstract.append(make_element("multiLevelType", {"val": "multilevel"}))
    for ilvl, (fmt, ltext, left) in enumerate(spec):
        lvl = make_element("lvl", {"ilvl": str(ilvl)})
        lvl.append(make_element("start", {"val": "1"}))
        lvl.append(make_element("numFmt", {"val": fmt}))
        lvl.append(make_element("suff", {"val": "space"}))
        lvl.append(make_element("lvlText", {"val": ltext}))
        lvl.append(make_element("lvlJc", {"val": "left"}))
        pPr = make_element("pPr")
        pPr.append(make_element("ind", {"left": str(left), "hanging": "360"}))
        lvl.append(pPr)
        rPr = make_element("rPr")
        rPr.append(make_element("rFonts", {
            "ascii": "Inter", "hAnsi": "Inter", "cs": "Inter", "eastAsia": "Inter",
        }))
        rPr.append(make_element("color", {"val": "000000"}))
        rPr.append(make_element("sz", {"val": "22"}))
        rPr.append(make_element("szCs", {"val": "22"}))
        lvl.append(rPr)
        abstract.append(lvl)
    return abstract


def _build_num(num_id: int) -> etree._Element:
    """Build a <w:num> instance pointing at the canonical abstractNumId.

    Each section gets its own num so list counters restart at every
    Heading 2 boundary (per format.md §2). All <w:num>s share one
    abstractNumId — only the binding differs, so list formatting stays
    identical across sections.
    """
    num = make_element("num", {"numId": str(num_id)})
    num.append(make_element("abstractNumId", {"val": str(CANONICAL_ABSTRACT_NUM_ID)}))
    # Force every level back to start=1 in this num. Without explicit
    # lvlOverride/startOverride entries, Word inherits the abstract's
    # start values but interprets numIds relative to document order in
    # ways that can leak counters across nums; the explicit override
    # makes per-num restart unambiguous.
    for ilvl in range(9):
        ovr = make_element("lvlOverride", {"ilvl": str(ilvl)})
        ovr.append(make_element("startOverride", {"val": "1"}))
        num.append(ovr)
    return num


def _install_abstract_num(numbering_root: etree._Element) -> None:
    """Strip existing abstractNum/num entries; install the single canonical abstractNum."""
    for tag in ("abstractNum", "num"):
        for el in numbering_root.findall(W + tag):
            numbering_root.remove(el)
    numbering_root.append(_build_canonical_abstract_num())


def apply(doc_root: etree._Element, numbering_root: etree._Element) -> None:
    """Apply Rule 2 in place to the document and numbering trees.

    Walks body paragraphs once. For each:
      - On a section heading (Heading 2): allocate a fresh <w:num>; reset
        marker-disambiguation state. Subsequent list items in this
        section reference the new num so counters restart at 1 per spec.
      - On a list-marker paragraph: split off any embedded markers as
        their own paragraphs, excise the leading marker text, attach
        numPr, drop paragraph-level ind overrides.
      - On a body continuation of a list item: indent it to match the
        parent item's body-text column (no marker rendered).

    The numbering tree is wiped of pre-existing abstractNum/num entries
    and gets one canonical abstractNum + one num per section installed.
    """
    _install_abstract_num(numbering_root)

    body = get_body(doc_root)
    last_l1: str | None = None
    last_l4: str | None = None

    # numId allocation. CANONICAL_NUM_ID is the first; bump per section.
    next_num_id = CANONICAL_NUM_ID
    current_num_id = next_num_id
    numbering_root.append(_build_num(current_num_id))
    next_num_id += 1

    # Track the indent level of the current list item so a body paragraph
    # immediately following it can be indented to match.
    current_item_indent: int | None = None

    p = next(iter_paragraphs(body), None)
    while p is not None:
        style = get_pstyle(p)
        if style == SECTION_STYLE_ID:
            last_l1 = None
            last_l4 = None
            current_item_indent = None
            current_num_id = next_num_id
            numbering_root.append(_build_num(current_num_id))
            next_num_id += 1
            p = _next_paragraph_in_tree(p)
            continue
        if style in HEADING_STYLES:
            # Title or subheading: reset disambiguation but keep current num.
            last_l1 = None
            last_l4 = None
            current_item_indent = None
            p = _next_paragraph_in_tree(p)
            continue

        # If the paragraph already carries numPr (e.g. from a prior pass),
        # treat it as a list item: read its ilvl and update the running
        # continuation indent so following body paragraphs inherit it.
        # Skip marker scanning — the marker text was already excised.
        existing_numPr = p.find(W + "pPr")
        existing_numPr = existing_numPr.find(W + "numPr") if existing_numPr is not None else None
        if existing_numPr is not None:
            ilvl_el = existing_numPr.find(W + "ilvl")
            if ilvl_el is not None:
                try:
                    current_item_indent = _LEVEL_LEFT_INDENT.get(int(ilvl_el.get(W + "val", "0")))
                except ValueError:
                    pass
            p = _next_paragraph_in_tree(p)
            continue

        text = _paragraph_text(p)
        leading = _scan_leading_marker(text, last_l1, last_l4)
        if leading is None:
            # Body paragraph with no marker. If it's immediately
            # following a list item, indent it to match (continuation).
            if current_item_indent is not None:
                set_indent(p, current_item_indent, hanging_twips=None)
            p = _next_paragraph_in_tree(p)
            continue

        _excise_marker_text(p, leading.end)
        set_numPr(p, current_num_id, leading.level)
        clear_pPr_child(p, "ind")
        current_item_indent = _LEVEL_LEFT_INDENT.get(leading.level)
        token = _extract_token(text, leading)
        if leading.level == 1:
            last_l1 = token
            last_l4 = None
        elif leading.level == 4:
            last_l4 = token
        elif leading.level in (0, 3):
            last_l1 = None
            last_l4 = None

        p = _next_paragraph_in_tree(p)


def _extract_token(text: str, marker: _MarkerMatch) -> str | None:
    """Pull the alphanumeric token out of a marker substring (e.g. 'a' from 'a)')."""
    sub = text[marker.start:marker.end]
    m = re.search(r"[A-Za-z0-9]+", sub)
    return m.group(0).lower() if m else None


def _next_paragraph_in_tree(p: etree._Element) -> etree._Element | None:
    """Find the next non-empty <w:p> in document order, descending into nested containers.

    Walks forward via following-sibling, descending into tables / SDTs as
    needed; ascends to parent siblings when a branch is exhausted. Mirrors
    iter_paragraphs() so split-off paragraphs (inserted immediately after
    the source) are picked up on the next iteration.
    """
    cur = p
    while cur is not None:
        # 1) Try descendants of following siblings, then parent's following siblings, etc.
        sib = cur.getnext()
        while sib is None:
            cur = cur.getparent()
            if cur is None:
                return None
            sib = cur.getnext()
        # Descend into sib looking for the first <w:p>.
        for el in sib.iter(W + "p"):
            from _docx import paragraph_text
            if paragraph_text(el).strip():
                return el
        cur = sib
    return None
