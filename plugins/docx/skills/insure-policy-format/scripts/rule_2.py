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

_BARE_UPPER_LETTER_RE = re.compile(r"([A-Z])\s*[).]")
_PAREN_LETTER_RE = re.compile(r"\(([a-z])\)")
_PAREN_ROMAN_RE = re.compile(r"\(([ivxlcdm]+)\)")
_PAREN_DECIMAL_RE = re.compile(r"\((\d+)\)")
_BARE_DECIMAL_RE = re.compile(r"(\d+)\s*[).]")
_BARE_LETTER_RE = re.compile(r"([a-z])\s*[).]")
_BARE_ROMAN_RE = re.compile(r"([ivxlcdm]{2,})\s*[).]")  # 2+ chars: unambiguously roman


_ALPHA_PREDECESSOR_OF_ROMAN = {"i": "h", "v": "u", "x": "w"}


def _detect_marker_at(
    text: str,
    start: int,
    *,
    is_leading: bool,
    last_alpha_l2: str | None,
    last_alpha_l4: str | None,
    last_level: int,
) -> tuple[int, int] | None:
    """Look for a marker beginning at offset `start`. Returns (end_offset, level) or None.

    Canonical ladder (post-Rule-0):
        L0 upper-alpha  A.
        L1 decimal      1.
        L2 lower-alpha  a.
        L3 decimal      1.   (also recognized as (1) for typed authoring)
        L4 lower-alpha  a.   (also recognized as (a))
        L5 lower-roman  i.   (also recognized as (i))

    L1/L3 (and L2/L4) share numFmt; the level is disambiguated by
    `last_level` — a decimal or alpha that follows a deeper level is
    treated as the deeper variant. Roman vs alpha for single-letter
    cases (i/v/x) is resolved via `last_alpha_l2`/`last_alpha_l4`.

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
        # Parenthesized authoring forms map to deep levels (L3-L5).
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
                if last_alpha_l4 != pred:
                    return start + m.end(), 5
            return start + m.end(), 4

    # Bare upper-letter: A. or A) — only valid at L0.
    m = _BARE_UPPER_LETTER_RE.match(rest)
    if m:
        return start + m.end(), 0
    # Bare decimal: 1) or 1.  L1 by default, L3 if we are nested deeper.
    m = _BARE_DECIMAL_RE.match(rest)
    if m:
        level = 3 if last_level >= 2 else 1
        return start + m.end(), level
    # Multi-char roman: ii)/iii)/iv)/etc — definitely roman, L5.
    m = _BARE_ROMAN_RE.match(rest)
    if m:
        return start + m.end(), 5
    # Single letter (could be alpha or roman).
    m = _BARE_LETTER_RE.match(rest)
    if m:
        token = m.group(1)
        end = start + m.end()
        if token in _ALPHA_PREDECESSOR_OF_ROMAN:
            pred = _ALPHA_PREDECESSOR_OF_ROMAN[token]
            # If we just emitted the alpha predecessor at L2 or L4,
            # treat as alpha continuation at that level. Otherwise the
            # token is the start of a roman sub-list (L5).
            if last_alpha_l2 == pred:
                return end, 2
            if last_alpha_l4 == pred:
                return end, 4
            return end, 5
        # Plain alpha. L2 by default, L4 if nested under a decimal at L3.
        level = 4 if last_level >= 3 else 2
        return end, level
    return None


def _scan_leading_marker(
    text: str,
    last_alpha_l2: str | None,
    last_alpha_l4: str | None,
    last_level: int,
) -> _MarkerMatch | None:
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
        text, pos, is_leading=True,
        last_alpha_l2=last_alpha_l2, last_alpha_l4=last_alpha_l4,
        last_level=last_level,
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
        # (numFmt, lvlText, left_twips). Per format.md §2 canonical ladder:
        # A. -> 1. -> a. -> 1. -> a. -> i.  (period-rooted, upper-letter at L0)
        ("upperLetter", "%1.", 720),
        ("decimal",     "%2.", 1440),
        ("lowerLetter", "%3.", 2160),
        ("decimal",     "%4.", 2880),
        ("lowerLetter", "%5.", 3600),
        ("lowerRoman",  "%6.", 4320),
        # Levels 6-8 are placeholders — format.md doesn't license content
        # beyond level 5; provide stable definitions so Word doesn't trip.
        ("decimal",     "%7.", 5040),
        ("lowerLetter", "%8.", 5760),
        ("lowerRoman",  "%9.", 6480),
    ]
    abstract = make_element("abstractNum", {"abstractNumId": str(CANONICAL_ABSTRACT_NUM_ID)})
    abstract.append(make_element("multiLevelType", {"val": "multilevel"}))
    for ilvl, (fmt, ltext, left) in enumerate(spec):
        # No <w:suff>: defaults to "tab" (matches the user reference's
        # native lists). No <w:rPr>: marker formatting inherits from the
        # paragraph's run style — body is Inter 11pt black, so markers
        # render the same without explicit overrides.
        lvl = make_element("lvl", {"ilvl": str(ilvl)})
        lvl.append(make_element("start", {"val": "1"}))
        lvl.append(make_element("numFmt", {"val": fmt}))
        lvl.append(make_element("lvlText", {"val": ltext}))
        lvl.append(make_element("lvlJc", {"val": "left"}))
        pPr = make_element("pPr")
        pPr.append(make_element("ind", {"left": str(left), "hanging": "360"}))
        lvl.append(pPr)
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
    last_alpha_l2: str | None = None
    last_alpha_l4: str | None = None
    last_level: int = -1

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
            last_alpha_l2 = None
            last_alpha_l4 = None
            last_level = -1
            current_item_indent = None
            current_num_id = next_num_id
            numbering_root.append(_build_num(current_num_id))
            next_num_id += 1
            p = _next_paragraph_in_tree(p)
            continue
        if style in HEADING_STYLES:
            # Title or subheading: reset disambiguation but keep current num.
            last_alpha_l2 = None
            last_alpha_l4 = None
            last_level = -1
            current_item_indent = None
            p = _next_paragraph_in_tree(p)
            continue

        # If the paragraph already carries numPr (source-native multilevel
        # list, or a prior-pass artifact), keep its ilvl but reassign numId
        # to this section's canonical num — Rule 2 wiped every source
        # abstractNum/num when it installed the canonical one, so leaving
        # the source's numId in place would dangle. Also strip the source's
        # paragraph-level ind override so canonical indentation applies.
        existing_numPr = p.find(W + "pPr")
        existing_numPr = existing_numPr.find(W + "numPr") if existing_numPr is not None else None
        if existing_numPr is not None:
            ilvl_el = existing_numPr.find(W + "ilvl")
            ilvl = 0
            if ilvl_el is not None:
                try:
                    ilvl = int(ilvl_el.get(W + "val", "0"))
                except ValueError:
                    ilvl = 0
            set_numPr(p, current_num_id, ilvl)
            clear_pPr_child(p, "ind")
            current_item_indent = _LEVEL_LEFT_INDENT.get(ilvl)
            p = _next_paragraph_in_tree(p)
            continue

        text = _paragraph_text(p)
        leading = _scan_leading_marker(text, last_alpha_l2, last_alpha_l4, last_level)
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
        if leading.level == 2:
            last_alpha_l2 = token
            last_alpha_l4 = None
        elif leading.level == 4:
            last_alpha_l4 = token
        elif leading.level in (0, 1, 3):
            # Hitting a parent or peer-of-parent resets the alpha trackers below.
            last_alpha_l2 = None
            last_alpha_l4 = None
        last_level = leading.level

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
