"""Rule 0: outline normalization.

Per format.md §0:

    Rule 0 rewrites such outlines to the canonical sequence
    (1) -> a) -> i) -> (1) -> (a) -> (i)). [...] The rewrite is
    content-level, not styling. After Rule 0 finishes, the document
    looks as if it had been authored in canonical form, so Rules 1/2/3
    see only canonical markers and stay commutative with each other.

For each outline_normalizations entry, walk the paragraphs from the
named section heading up to (but not including) the next section
heading, tracking per-level counters with parent-aware resets. For each
paragraph, both the leading marker and any embedded markers are
rewritten in the paragraph's <w:t> runs.
"""
from __future__ import annotations

from lxml import etree

from _docx import iter_paragraphs, paragraph_text, get_body, W
from parts import LevelSpec, ResolvedParts


_CANONICAL_LEVEL_FORMS: tuple[tuple[str, str, str], ...] = (
    ("decimal", "{}", ")"),
    ("lower_alpha", "{}", ")"),
    ("lower_roman", "{}", ")"),
    ("decimal", "({})", ""),
    ("lower_alpha", "({})", ""),
    ("lower_roman", "({})", ""),
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
    """Return non-overlapping (start, end, level_idx) matches in text.

    Leading match is selected with the leading regex (anchored at start
    after optional whitespace, lower-numbered level wins on tie).
    Embedded matches use the boundary-anchored regex. The two are merged
    and de-overlapped left-to-right.
    """
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
    """Apply (start, end, new_str) replacements in paragraph-text coordinates to <w:t> runs.

    Replacements must be non-overlapping. Walks <w:t>/<w:tab>/<w:br>/<w:cr>
    in document order (matching paragraph_text() semantics) and edits each
    overlapped <w:t>'s text. Replacements that span multiple <w:t> elements
    write the new string into the first overlapped <w:t> and blank the
    overlapped slice of subsequent ones.
    """
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


def apply(doc_root: etree._Element, parts: ResolvedParts) -> etree._Element:
    """Apply Rule 0 to the document, in place. Returns the same root."""
    if not parts.outline_normalizations:
        return doc_root
    body = get_body(doc_root)
    paragraphs = list(iter_paragraphs(body))
    section_indices = sorted(parts.section_indices)
    for norm in parts.outline_normalizations:
        next_idx = next(
            (i for i in section_indices if i > norm.section_index),
            len(paragraphs),
        )
        counters = [0] * len(norm.source_levels)
        for p in paragraphs[norm.section_index + 1 : next_idx]:
            _rewrite_paragraph(p, norm.source_levels, counters)
    return doc_root
