"""Rule 2: lists.

Per format.md §2 — Lists.

This rule has two phases:

1. Outline marker normalization (pre-pass). Source documents flagged
   in `outline_normalizations` use non-canonical marker styles (e.g.
   uppercase Roman at the top level). Rewrite their markers in-place
   to the canonical ladder before list detection runs.

2. List detection. Walk body paragraphs (skipping headings); detect
   leading list markers; assign `numPr` per detected level; strip the
   marker text from the paragraph's runs (numbering renders it). A
   single canonical multilevel abstractNum is installed in
   numbering.xml; one `<w:num>` instance is allocated per Heading 2
   so list counters restart at every section boundary.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from lxml import etree

from _docx import (
    Doc,
    W,
    clear_pPr_child,
    get_body,
    get_pstyle,
    iter_paragraphs,
    make_element,
    paragraph_text,
    set_indent,
    set_numPr,
)
from parts import LevelSpec, ResolvedParts
from rule_1 import SECTION_STYLE_ID, SUBHEAD_STYLE_ID, TITLE_STYLE_ID


# ---------------------------------------------------------------------
# Phase 1: outline marker normalization (pre-pass).
#
# Rewrite source markers in flagged sections to the canonical ladder
# (`A.` -> `1.` -> `a.` -> `1.` -> `a.` -> `i.`) so phase 2 sees only
# canonical markers. Per format.md §2 / "Outline marker normalization":
#
# - Range = from a flagged section heading up to (but excluding) the
#   next section heading. Section headings are identified by the
#   Heading 2 pStyle that Rule 1 has already applied.
# - Counters reset to 0 for all levels deeper than the one that just
#   fired.
# - Both leading and embedded markers in a paragraph are rewritten.
# - Embedded scans use a non-word-boundary guard so citations
#   (`Section IV.A.`, `officer(s)`, `§4958(c)`, `sixty (60) days`)
#   stay as plain text.
# ---------------------------------------------------------------------


_CANONICAL_LEVEL_FORMS: tuple[tuple[str, str, str], ...] = (
    ("upper_alpha", "{}", "."),
    ("decimal", "{}", "."),
    ("lower_alpha", "{}", "."),
    ("decimal", "{}", "."),
    ("lower_alpha", "{}", "."),
    ("lower_roman", "{}", "."),
)


_ROMAN_PAIRS = (
    (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"),
    (100, "c"), (90, "xc"), (50, "l"), (40, "xl"),
    (10, "x"), (9, "ix"), (5, "v"), (4, "iv"),
    (1, "i"),
)


def _to_roman(n: int, upper: bool) -> str:
    if n < 1:
        raise ValueError(f"roman index must be >= 1, got {n}")
    out = ""
    for value, sym in _ROMAN_PAIRS:
        while n >= value:
            out += sym
            n -= value
    return out.upper() if upper else out


def _to_alpha(n: int, upper: bool) -> str:
    if n < 1:
        raise ValueError(f"alpha index must be >= 1, got {n}")
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(ord("a") + rem) + out
    return out.upper() if upper else out


def _format_token(sequence: str, n: int) -> str:
    if sequence == "decimal":
        return str(n)
    if sequence == "lower_alpha":
        return _to_alpha(n, upper=False)
    if sequence == "upper_alpha":
        return _to_alpha(n, upper=True)
    if sequence == "lower_roman":
        return _to_roman(n, upper=False)
    if sequence == "upper_roman":
        return _to_roman(n, upper=True)
    raise ValueError(f"unknown sequence type: {sequence!r}")


def _canonical_marker(level_idx: int, counter: int) -> str:
    if level_idx < 0 or level_idx >= len(_CANONICAL_LEVEL_FORMS):
        raise ValueError(
            f"canonical marker only defined for levels 0..{len(_CANONICAL_LEVEL_FORMS) - 1}, got {level_idx}"
        )
    sequence, template, suffix = _CANONICAL_LEVEL_FORMS[level_idx]
    return template.format(_format_token(sequence, counter)) + suffix


def _scan_marker_matches(
    text: str, source_levels: list[LevelSpec]
) -> list[tuple[int, int, int]]:
    raw: list[tuple[int, int, int]] = []
    for level_idx, level in enumerate(source_levels):
        for m in level.compiled_embedded.finditer(text):
            raw.append((m.start(), m.end(), level_idx))
    leading: tuple[int, int, int] | None = None
    for level_idx, level in enumerate(source_levels):
        m = level.compiled_leading.match(text)
        if m and (leading is None or level_idx < leading[2]):
            leading = (m.start(), m.end(), level_idx)
    if leading is not None:
        raw = [(s, e, lv) for (s, e, lv) in raw if s >= leading[1]]
        raw.append(leading)
    raw.sort(key=lambda x: (x[0], x[2]))
    chosen: list[tuple[int, int, int]] = []
    last_end = -1
    for start, end, lv in raw:
        if start < last_end:
            continue
        chosen.append((start, end, lv))
        last_end = end
    return chosen


def _replace_in_paragraph(
    p: etree._Element, replacements: list[tuple[int, int, str]]
) -> None:
    if not replacements:
        return
    pieces: list[tuple[etree._Element, str, str]] = []
    for el in p.iter():
        tag = el.tag
        if tag == W + "t":
            pieces.append((el, "t", el.text or ""))
        elif tag == W + "tab":
            pieces.append((el, "tab", " "))
        elif tag in (W + "br", W + "cr"):
            pieces.append((el, "br", " "))

    spans: list[tuple[int, int]] = []
    cursor = 0
    for _el, _kind, text in pieces:
        spans.append((cursor, cursor + len(text)))
        cursor += len(text)

    edits: dict[int, list[tuple[int, int, str]]] = {}
    for r_start, r_end, new_str in sorted(replacements):
        wrote = False
        for i, (s, e) in enumerate(spans):
            if e <= r_start:
                continue
            if s >= r_end:
                break
            local_start = max(0, r_start - s)
            local_end = min(e - s, r_end - s)
            edits.setdefault(i, []).append((local_start, local_end, new_str if not wrote else ""))
            wrote = True

    for i, piece_edits in edits.items():
        el, kind, text = pieces[i]
        if kind != "t":
            continue
        piece_edits.sort(key=lambda x: x[0])
        out = []
        prev = 0
        for local_start, local_end, repl in piece_edits:
            out.append(text[prev:local_start])
            out.append(repl)
            prev = local_end
        out.append(text[prev:])
        el.text = "".join(out)


def _rewrite_paragraph(
    p: etree._Element, source_levels: list[LevelSpec], counters: list[int]
) -> None:
    text = paragraph_text(p)
    matches = _scan_marker_matches(text, source_levels)
    if not matches:
        return
    replacements: list[tuple[int, int, str]] = []
    for start, end, level_idx in matches:
        counters[level_idx] += 1
        for j in range(level_idx + 1, len(counters)):
            counters[j] = 0
        canonical = _canonical_marker(level_idx, counters[level_idx])
        is_leading = start == 0 or text[:start].strip() == ""
        if is_leading:
            replacement = text[:start] + canonical + " "
            replacements.append((0, end, replacement))
        else:
            replacements.append((start, end, " " + canonical + " "))
    _replace_in_paragraph(p, replacements)


def _normalize_ws(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _apply_outline_normalizations(
    doc_root: etree._Element, normalizations: list
) -> None:
    """Pre-pass: rewrite non-canonical markers to canonical for flagged sections.

    Sections are located by Heading 2 pStyle (set by Rule 1) plus a
    text match against `section_text`. The source-paragraph index that
    `parts.resolve()` computed is intentionally NOT used — Rule 1 may
    have removed paragraphs by the time this runs, so absolute indices
    don't line up. The substring-uniqueness invariant from
    `parts.resolve()` still guarantees a unique match here.
    """
    if not normalizations:
        return
    body = get_body(doc_root)
    paragraphs = list(iter_paragraphs(body))
    section_positions = [
        i for i, p in enumerate(paragraphs)
        if get_pstyle(p) == SECTION_STYLE_ID
    ]
    section_texts = {
        i: _normalize_ws(paragraph_text(paragraphs[i]))
        for i in section_positions
    }
    for norm in normalizations:
        target = _normalize_ws(norm.section_text)
        matches = [i for i in section_positions if target in section_texts[i]]
        if len(matches) != 1:
            raise ValueError(
                f"outline normalization: section_text {norm.section_text!r} "
                f"resolved to {len(matches)} Heading 2 paragraphs post-Rule 1; "
                f"expected exactly 1"
            )
        sec_pos = matches[0]
        next_pos = next(
            (i for i in section_positions if i > sec_pos),
            len(paragraphs),
        )
        counters = [0] * len(norm.source_levels)
        for p in paragraphs[sec_pos + 1 : next_pos]:
            _rewrite_paragraph(p, norm.source_levels, counters)


# ---------------------------------------------------------------------
# Phase 2: list detection and structure.
# ---------------------------------------------------------------------


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


def apply(doc: Doc, parts: ResolvedParts) -> Doc:
    """Apply Rule 2 in place to the document and numbering trees.

    Two phases:
      1. Outline marker normalization (pre-pass) — rewrite source
         markers in flagged sections to the canonical ladder.
      2. List detection — walk body paragraphs, detect leading list
         markers, attach numPr/ilvl, install the canonical multilevel
         numbering definition.

    Phase 2 walks body paragraphs once. For each:
      - On a section heading (Heading 2): allocate a fresh <w:num>; reset
        marker-disambiguation state. Subsequent list items in this
        section reference the new num so counters restart at 1 per spec.
      - On a list-marker paragraph: excise the leading marker text,
        attach numPr, drop paragraph-level ind overrides.
      - On a body continuation of a list item: indent it to match the
        parent item's body-text column (no marker rendered).

    The numbering tree is wiped of pre-existing abstractNum/num entries
    and gets one canonical abstractNum + one num per section installed.
    """
    _apply_outline_normalizations(doc.document, parts.outline_normalizations)
    _install_abstract_num(doc.numbering)

    body = get_body(doc.document)
    last_alpha_l2: str | None = None
    last_alpha_l4: str | None = None
    last_level: int = -1

    # numId allocation. CANONICAL_NUM_ID is the first; bump per section.
    next_num_id = CANONICAL_NUM_ID
    current_num_id = next_num_id
    doc.numbering.append(_build_num(current_num_id))
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
            doc.numbering.append(_build_num(current_num_id))
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

    return doc


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
