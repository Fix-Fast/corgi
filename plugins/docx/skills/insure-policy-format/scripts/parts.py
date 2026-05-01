"""parts.json schema, validation, and substring resolution.

The formatter takes a `parts.json` manifest that points at specific
source paragraphs by verbatim substring (per the SKILL.md prompt
contract). This module:

  1. Parses and validates the JSON shape (raising on schema violations).
  2. Resolves each substring to a unique paragraph index in the source
     document by matching against the corpus of paragraph texts joined
     with newlines, with whitespace collapsed.

No pandoc. The corpus is built directly from the source DOCX's OOXML
via _docx.iter_paragraphs / paragraph_text.
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

from _docx import iter_paragraphs, paragraph_text


SEQUENCE_TYPES = (
    "decimal",
    "upper_alpha",
    "lower_alpha",
    "upper_roman",
    "lower_roman",
)


@dataclass
class LevelSpec:
    pattern: str
    sequence: str
    compiled_leading: re.Pattern[str]
    compiled_embedded: re.Pattern[str]


@dataclass
class OutlineNormalization:
    section_text: str
    section_index: int  # resolved
    source_levels: list[LevelSpec]


@dataclass
class ResolvedParts:
    """Parts.json after substring resolution. All *_indices fields point at paragraphs in source order."""

    ignored_indices: set[int] = field(default_factory=set)
    title_indices: set[int] = field(default_factory=set)
    section_indices: set[int] = field(default_factory=set)
    subheading_indices: set[int] = field(default_factory=set)
    outline_normalizations: list[OutlineNormalization] = field(default_factory=list)
    header_title_text: str | None = None
    policy_code: str | None = None


def _normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _normalize_match_string(s: str) -> str:
    return "\n".join(_normalize_text(chunk) for chunk in s.split("\n"))


def _parse_level_spec(raw: object, field_name: str) -> LevelSpec:
    if not isinstance(raw, dict):
        raise TypeError(
            f"{field_name}: each level entry must be an object with 'pattern' and "
            f"'sequence' keys, got {type(raw).__name__}"
        )
    pattern = raw.get("pattern")
    sequence = raw.get("sequence")
    if not isinstance(pattern, str) or not pattern:
        raise ValueError(f"{field_name}: missing or non-string 'pattern' in {raw!r}")
    if sequence not in SEQUENCE_TYPES:
        raise ValueError(
            f"{field_name}: 'sequence' must be one of {list(SEQUENCE_TYPES)}, got {sequence!r}"
        )
    unexpected = set(raw) - {"pattern", "sequence"}
    if unexpected:
        raise ValueError(f"{field_name}: unknown keys: {sorted(unexpected)}")
    try:
        compiled_leading = re.compile(r"^\s*" + pattern)
        # Embedded markers must (a) start at a non-word boundary and
        # (b) be followed by whitespace. The trailing whitespace check
        # rejects citation-style usages like "and B." (followed by ',')
        # or "Section A's" while still catching real inline markers
        # like "1) ... or 2) baz".
        compiled_embedded = re.compile(r"(?<![A-Za-z0-9).(])" + pattern + r"(?=\s)")
    except re.error as e:
        raise ValueError(f"{field_name}: invalid regex {pattern!r}: {e}")
    if compiled_leading.groups != 1:
        raise ValueError(
            f"{field_name}: pattern {pattern!r} must have exactly one capture group "
            f"(the enumeration token); got {compiled_leading.groups}"
        )
    return LevelSpec(
        pattern=pattern,
        sequence=sequence,
        compiled_leading=compiled_leading,
        compiled_embedded=compiled_embedded,
    )


def _parse_outline_normalization(raw: object, idx: int) -> tuple[str, list[LevelSpec]]:
    field_name = f"outline_normalizations[{idx}]"
    if not isinstance(raw, dict):
        raise TypeError(f"{field_name}: must be an object, got {type(raw).__name__}")
    section_text = raw.get("section_text")
    source_levels = raw.get("source_levels")
    if not isinstance(section_text, str) or not section_text.strip():
        raise ValueError(f"{field_name}: missing or empty 'section_text'")
    if not isinstance(source_levels, list) or not source_levels:
        raise ValueError(f"{field_name}: 'source_levels' must be a non-empty list")
    unexpected = set(raw) - {"section_text", "source_levels"}
    if unexpected:
        raise ValueError(f"{field_name}: unknown keys: {sorted(unexpected)}")
    levels = [
        _parse_level_spec(level, f"{field_name}.source_levels[{i}]")
        for i, level in enumerate(source_levels)
    ]
    return section_text, levels


def load_raw(path: Path) -> dict[str, object]:
    raw = json.loads(path.read_text())
    if not isinstance(raw, dict):
        raise TypeError(f"parts.json must be a JSON object at the top level, got {type(raw).__name__}")
    return raw


def build_corpus(doc_root: etree._Element) -> tuple[list[str], str, list[tuple[int, int]]]:
    """Return (paragraph_texts, joined_corpus, paragraph_spans).

    paragraph_texts[i] is the normalized (whitespace-collapsed) text of
    the i-th paragraph in document order. joined_corpus is the texts
    joined by '\\n'. paragraph_spans[i] is (start, end) of paragraph i
    inside joined_corpus.
    """
    body = doc_root.find(f"{{{etree.QName(doc_root).namespace}}}body")
    if body is None:
        # fallback for docs that aren't a w:document root
        body = doc_root
    texts: list[str] = []
    spans: list[tuple[int, int]] = []
    cursor = 0
    for p in iter_paragraphs(body):
        text = _normalize_text(paragraph_text(p))
        start = cursor
        end = cursor + len(text)
        spans.append((start, end))
        texts.append(text)
        cursor = end + 1  # +1 for the joining '\n'
    return texts, "\n".join(texts), spans


def _find_all(haystack: str, needle: str) -> list[int]:
    out: list[int] = []
    start = 0
    while True:
        i = haystack.find(needle, start)
        if i < 0:
            return out
        out.append(i)
        start = i + 1


def _paragraph_for_position(pos: int, spans: list[tuple[int, int]]) -> int:
    for idx, (start, end) in enumerate(spans):
        if start <= pos < end:
            return idx
        if pos < start:
            return idx
    return len(spans) - 1


def resolve_substring(
    entry: object,
    field_name: str,
    texts: list[str],
    corpus: str,
    spans: list[tuple[int, int]],
) -> int:
    if not isinstance(entry, str):
        raise TypeError(
            f"{field_name}: entries must be strings (verbatim text drawn from the "
            f"target paragraph, optionally extended with neighbor paragraph text "
            f"separated by newlines), got {type(entry).__name__}"
        )
    needle = _normalize_match_string(entry)
    if not needle.strip():
        raise ValueError(f"{field_name}: empty match string is not allowed")

    positions = _find_all(corpus, needle)
    if not positions:
        raise ValueError(
            f"{field_name}: no match for {entry!r}. Provide a substring drawn "
            f"verbatim from the target paragraph (whitespace is collapsed). To "
            f"disambiguate identical paragraphs, extend the string across "
            f"paragraph boundaries with '\\n' and include some neighbor text."
        )
    if len(positions) > 1:
        targets = [_paragraph_for_position(p, spans) for p in positions]
        previews = "; ".join(
            f"#{idx}: {texts[idx][:80]}{'...' if len(texts[idx]) > 80 else ''}"
            for idx in targets
        )
        raise ValueError(
            f"{field_name}: match string {entry!r} matches {len(positions)} positions "
            f"in the corpus and is ambiguous. Extend the string with neighbor "
            f"paragraph text (using '\\n') to make it unique. Candidate paragraphs: {previews}"
        )
    return _paragraph_for_position(positions[0], spans)


def resolve(parts_path: Path, doc_root: etree._Element) -> ResolvedParts:
    raw = load_raw(parts_path)
    texts, corpus, spans = build_corpus(doc_root)

    def resolve_list(key: str) -> set[int]:
        out: set[int] = set()
        for entry in raw.get(key, []):
            out.add(resolve_substring(entry, key, texts, corpus, spans))
        return out

    resolved = ResolvedParts(
        ignored_indices=resolve_list("ignored_body_texts"),
        title_indices=resolve_list("title_texts"),
        section_indices=resolve_list("section_heading_texts"),
        subheading_indices=resolve_list("subheading_texts"),
        header_title_text=raw.get("header_title_text"),
        policy_code=raw.get("policy_code"),
    )

    outline_raw = raw.get("outline_normalizations", [])
    if not isinstance(outline_raw, list):
        raise TypeError(
            f"outline_normalizations: must be a list, got {type(outline_raw).__name__}"
        )
    for i, entry in enumerate(outline_raw):
        section_text, levels = _parse_outline_normalization(entry, i)
        section_idx = resolve_substring(
            section_text, f"outline_normalizations[{i}].section_text",
            texts, corpus, spans,
        )
        resolved.outline_normalizations.append(
            OutlineNormalization(
                section_text=section_text,
                section_index=section_idx,
                source_levels=levels,
            )
        )

    if resolved.header_title_text is not None and not isinstance(resolved.header_title_text, str):
        raise TypeError("header_title_text must be a string or null")
    if resolved.policy_code is not None and not isinstance(resolved.policy_code, str):
        raise TypeError("policy_code must be a string or null")

    return resolved
