from __future__ import annotations

import json
import re
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from lxml import etree

from _ooxml import patch_list_levels, set_list_suffix, strip_list_ind_overrides



LEVEL_PATTERNS = [
    re.compile(r"^\s*(\d+)\)(\s+|(?=[A-Z(]))"),
    re.compile(r"^\s*([a-z])\)(\s+|(?=[A-Z(]))"),
    re.compile(r"^\s*([ivx]+)\)(\s+|(?=[A-Z(]))"),
    re.compile(r"^\s*\((\d+)\)(\s+|(?=[A-Z(]))"),
    re.compile(r"^\s*\(([a-z])\)(\s+|(?=[A-Z(]))"),
    re.compile(r"^\s*\(([ivx]+)\)(\s+|(?=[A-Z(]))"),
]

EMBEDDED_MARKER_PATTERNS = [
    re.compile(r"(\(\d+\)(\s+|(?=[A-Z(])))"),
    re.compile(r"(\([a-z]\)(\s+|(?=[A-Z(])))"),
    re.compile(r"(\([ivx]+\)(\s+|(?=[A-Z(])))"),
    re.compile(r"(?<=\s)(\d+\)(\s+|(?=[A-Z(])))"),
    re.compile(r"(?<=\s)([a-z]\)(\s+|(?=[A-Z(])))"),
    re.compile(r"(?<=\s)([ivx]+\)(\s+|(?=[A-Z(])))"),
]

LEVEL_TO_OL_ATTR = {
    0: [1, {"t": "Decimal"}, {"t": "OneParen"}],
    1: [1, {"t": "LowerAlpha"}, {"t": "OneParen"}],
    2: [1, {"t": "LowerRoman"}, {"t": "OneParen"}],
    3: [1, {"t": "Decimal"}, {"t": "TwoParens"}],
    4: [1, {"t": "LowerAlpha"}, {"t": "TwoParens"}],
    5: [1, {"t": "LowerRoman"}, {"t": "TwoParens"}],
}


def inlines_to_text(inlines: List[dict]) -> str:
    out: List[str] = []
    for x in inlines:
        t = x["t"]
        if t == "Str":
            out.append(x["c"])
        elif t in ("Space", "SoftBreak", "LineBreak"):
            out.append(" ")
        elif t in (
            "Strong",
            "Emph",
            "Underline",
            "SmallCaps",
            "Strikeout",
            "Superscript",
            "Subscript",
        ):
            out.append(inlines_to_text(x["c"]))
        elif t == "Span":
            out.append(inlines_to_text(x["c"][1]))
        elif t == "Quoted":
            out.append(inlines_to_text(x["c"][1]))
    return "".join(out)


def strip_chars_from_inlines(inlines: List[dict], n: int) -> List[dict]:
    if n <= 0:
        return list(inlines)
    remaining = n
    out: List[dict] = []
    for x in inlines:
        if remaining <= 0:
            out.append(x)
            continue
        t = x["t"]
        if t == "Str":
            s = x["c"]
            if len(s) <= remaining:
                remaining -= len(s)
                continue
            out.append({"t": "Str", "c": s[remaining:]})
            remaining = 0
        elif t in ("Space", "SoftBreak", "LineBreak"):
            remaining -= 1
        elif t in ("Strong", "Emph", "Underline", "SmallCaps", "Strikeout"):
            inner = strip_chars_from_inlines(x["c"], remaining)
            consumed = min(remaining, len(inlines_to_text(x["c"])))
            remaining -= consumed
            if inner:
                out.append({"t": t, "c": inner})
        elif t == "Span":
            inner = strip_chars_from_inlines(x["c"][1], remaining)
            consumed = min(remaining, len(inlines_to_text(x["c"][1])))
            remaining -= consumed
            if inner:
                out.append({"t": "Span", "c": [x["c"][0], inner]})
        else:
            out.append(x)
    while out and out[0]["t"] in ("Space", "SoftBreak"):
        out.pop(0)
    if out and out[0]["t"] == "Str":
        stripped = out[0]["c"].lstrip()
        if stripped:
            out[0] = {"t": "Str", "c": stripped}
        else:
            out.pop(0)
    return out


def split_inlines_at(inlines: List[dict], offset: int) -> Tuple[List[dict], List[dict]]:
    if offset <= 0:
        return [], list(inlines)
    left: List[dict] = []
    right: List[dict] = []
    remaining = offset
    for x in inlines:
        if remaining <= 0:
            right.append(x)
            continue
        t = x["t"]
        if t == "Str":
            s = x["c"]
            if len(s) <= remaining:
                left.append(x)
                remaining -= len(s)
            else:
                left.append({"t": "Str", "c": s[:remaining]})
                right.append({"t": "Str", "c": s[remaining:]})
                remaining = 0
        elif t in ("Space", "SoftBreak", "LineBreak"):
            left.append(x)
            remaining -= 1
        elif t in ("Strong", "Emph", "Underline", "SmallCaps", "Strikeout", "Superscript", "Subscript"):
            inner_text = inlines_to_text(x["c"])
            if len(inner_text) <= remaining:
                left.append(x)
                remaining -= len(inner_text)
            else:
                l_in, r_in = split_inlines_at(x["c"], remaining)
                if l_in:
                    left.append({"t": t, "c": l_in})
                if r_in:
                    right.append({"t": t, "c": r_in})
                remaining = 0
        elif t == "Span":
            inner_text = inlines_to_text(x["c"][1])
            if len(inner_text) <= remaining:
                left.append(x)
                remaining -= len(inner_text)
            else:
                l_in, r_in = split_inlines_at(x["c"][1], remaining)
                if l_in:
                    left.append({"t": "Span", "c": [x["c"][0], l_in]})
                if r_in:
                    right.append({"t": "Span", "c": [x["c"][0], r_in]})
                remaining = 0
        else:
            left.append(x)
    while right and right[0]["t"] in ("Space", "SoftBreak"):
        right.pop(0)
    if right and right[0]["t"] == "Str":
        stripped = right[0]["c"].lstrip()
        if stripped:
            right[0] = {"t": "Str", "c": stripped}
        else:
            right.pop(0)
    return left, right


def trim_trailing_chars_from_inlines(inlines: List[dict], n: int) -> List[dict]:
    if n <= 0:
        return list(inlines)
    remaining = n
    out: List[dict] = []
    for x in reversed(inlines):
        if remaining <= 0:
            out.append(x)
            continue
        t = x["t"]
        if t == "Str":
            s = x["c"]
            if len(s) <= remaining:
                remaining -= len(s)
                continue
            out.append({"t": "Str", "c": s[:-remaining]})
            remaining = 0
        elif t in ("Space", "SoftBreak", "LineBreak"):
            remaining -= 1
        elif t in ("Strong", "Emph", "Underline", "SmallCaps", "Strikeout", "Superscript", "Subscript"):
            inner_text = inlines_to_text(x["c"])
            if len(inner_text) <= remaining:
                remaining -= len(inner_text)
                continue
            trimmed = trim_trailing_chars_from_inlines(x["c"], remaining)
            consumed = len(inner_text) - len(inlines_to_text(trimmed))
            remaining -= consumed
            if trimmed:
                out.append({"t": t, "c": trimmed})
        elif t == "Span":
            inner_text = inlines_to_text(x["c"][1])
            if len(inner_text) <= remaining:
                remaining -= len(inner_text)
                continue
            trimmed = trim_trailing_chars_from_inlines(x["c"][1], remaining)
            consumed = len(inner_text) - len(inlines_to_text(trimmed))
            remaining -= consumed
            if trimmed:
                out.append({"t": "Span", "c": [x["c"][0], trimmed]})
        else:
            out.append(x)
    out.reverse()
    while out and out[-1]["t"] in ("Space", "SoftBreak"):
        out.pop()
    if out and out[-1]["t"] == "Str":
        stripped = out[-1]["c"].rstrip()
        if stripped:
            out[-1] = {"t": "Str", "c": stripped}
        else:
            out.pop()
    return out


@dataclass
class SourcePara:
    inlines: List[dict]
    text: str
    quote_depth: int
    index: int


@dataclass
class Block:
    id: str
    kind: str
    level: Optional[int] = None
    inlines: List[dict] = field(default_factory=list)
    continuations: List[List[dict]] = field(default_factory=list)
    source_index: int = -1
    quote_depth: int = 0
    flags: List[str] = field(default_factory=list)

    @property
    def text(self) -> str:
        return inlines_to_text(self.inlines).strip()


@dataclass
class HeaderMetadata:
    left_text: str
    code: str


def extract_policy_code(text: str) -> Optional[str]:
    upper = text.upper()
    for prefix in ("CORGI-TECH-", "CORG-TECH-"):
        start = upper.find(prefix)
        if start < 0:
            continue
        end = start + len(prefix)
        while end < len(upper) and upper[end].isdigit():
            end += 1
        if end == start + len(prefix):
            continue
        return upper[start:end]
    return None


def is_running_header_text(text: str) -> bool:
    return extract_policy_code(text) is not None


def derive_header_left_text(title_text: str) -> str:
    title = re.sub(r"\s+", " ", title_text).strip()
    if not title:
        raise ValueError("empty title text passed to derive_header_left_text")
    title = title.title()
    title = re.sub(r"\bInsurance Policy\b", "Policy", title, flags=re.IGNORECASE)
    title = re.sub(r"\bInsurance\b", "", title, flags=re.IGNORECASE)
    title = re.sub(r"\s+", " ", title).strip()
    if not title:
        raise ValueError(f"title collapsed to empty after normalization: {title_text!r}")
    return title


def extract_header_metadata(source_paras: List["SourcePara"], blocks: List["Block"]) -> HeaderMetadata:
    code: Optional[str] = None
    for para in source_paras:
        code = extract_policy_code(para.text)
        if code:
            break
    if not code:
        raise ValueError(
            "no CORGI-TECH-* / CORG-TECH-* policy code found in source running header; "
            "this does not look like a Corgi-Tech policy DOCX"
        )

    title_block = next((block for block in blocks if block.kind == "title" and block.text), None)
    if title_block is None:
        raise ValueError("no title block classified; source is missing an ALL-CAPS '... INSURANCE POLICY' heading")
    left_text = derive_header_left_text(title_block.text)

    return HeaderMetadata(left_text=left_text, code=code)


def flatten_paras(blocks: List[dict], depth: int = 0) -> List[SourcePara]:
    out: List[SourcePara] = []
    for block in blocks:
        t = block["t"]
        if t == "Para":
            text = inlines_to_text(block["c"]).strip()
            if text:
                out.append(SourcePara(block["c"], text, depth, -1))
        elif t == "BlockQuote":
            out.extend(flatten_paras(block["c"], depth + 1))
    for i, para in enumerate(out):
        para.index = i
    return out


def is_title(text: str, seen_title: bool) -> bool:
    if seen_title:
        return False
    return text.isupper() and "INSURANCE POLICY" in text


def is_section_heading(text: str) -> bool:
    return bool(re.match(r"^SECTION\s+[IVXLC]+\s*:", text))


def is_subheading(text: str) -> bool:
    if re.match(r"^Coverage\s+[A-Z]\s+[—-]", text):
        return True
    if text.endswith(":") and len(text.split()) <= 10:
        return True
    if text.isupper() and len(text.split()) <= 14:
        return True
    return False


def parse_leading_marker(
    inlines: List[dict],
    prev_l1: Optional[str],
    prev_l4: Optional[str],
) -> Tuple[Optional[int], Optional[str], List[dict], bool]:
    text = inlines_to_text(inlines)
    matches = []
    for level, regex in enumerate(LEVEL_PATTERNS):
        m = regex.match(text)
        if m:
            matches.append((level, m.end(), m.group(1)))
    if not matches:
        return None, None, inlines, False

    ambiguous = False
    if len(matches) == 1:
        level, end, token = matches[0]
        return level, token, strip_chars_from_inlines(inlines, end), False

    romans = [m for m in matches if m[0] in (2, 5)]
    alpha = [m for m in matches if m[0] in (1, 4)]
    if romans and any(len(m[2]) > 1 for m in romans):
        level, end, token = next(m for m in romans if len(m[2]) > 1)
        return level, token, strip_chars_from_inlines(inlines, end), False

    token = matches[0][2]
    if token == "i":
        if prev_l1 == "h":
            level, end, token = next(m for m in alpha if m[0] == 1)
            return level, token, strip_chars_from_inlines(inlines, end), False
        if prev_l4 == "h":
            level, end, token = next(m for m in alpha if m[0] == 4)
            return level, token, strip_chars_from_inlines(inlines, end), False
    if token == "v":
        if prev_l1 == "u":
            level, end, token = next(m for m in alpha if m[0] == 1)
            return level, token, strip_chars_from_inlines(inlines, end), False
        if prev_l4 == "u":
            level, end, token = next(m for m in alpha if m[0] == 4)
            return level, token, strip_chars_from_inlines(inlines, end), False
    if token == "x":
        if prev_l1 == "w":
            level, end, token = next(m for m in alpha if m[0] == 1)
            return level, token, strip_chars_from_inlines(inlines, end), False
        if prev_l4 == "w":
            level, end, token = next(m for m in alpha if m[0] == 4)
            return level, token, strip_chars_from_inlines(inlines, end), False

    ambiguous = len(alpha) > 0 and len(romans) > 0
    level, end, token = romans[0] if romans else matches[0]
    return level, token, strip_chars_from_inlines(inlines, end), ambiguous


def find_embedded_marker_offset(inlines: List[dict]) -> Optional[int]:
    text = inlines_to_text(inlines)
    best: Optional[int] = None
    for regex in EMBEDDED_MARKER_PATTERNS:
        for match in regex.finditer(text):
            if match.start() <= 0:
                continue
            if best is None or match.start() < best:
                best = match.start()
    if best is not None:
        return best
    return None


def cleanup_left_split_fragment(inlines: List[dict]) -> List[dict]:
    text = inlines_to_text(inlines)
    new_text = re.sub(r"\s*\(\s*$", "", text)
    new_text = re.sub(r"\s+(and|or)\s*$", "", new_text, flags=re.IGNORECASE)
    new_text = re.sub(r"\s+$", "", new_text)
    trim = len(text) - len(new_text)
    if trim <= 0:
        return inlines
    return trim_trailing_chars_from_inlines(inlines, trim)


def split_on_embedded_markers(source_paras: List[SourcePara]) -> List[SourcePara]:
    out: List[SourcePara] = []
    for para in source_paras:
        pending = [(para.inlines, para.text)]
        while pending:
            cur_inlines, cur_text = pending.pop(0)
            offset = find_embedded_marker_offset(cur_inlines)
            if offset is None:
                out.append(SourcePara(cur_inlines, inlines_to_text(cur_inlines).strip(), para.quote_depth, para.index))
                continue
            left, right = split_inlines_at(cur_inlines, offset)
            left = cleanup_left_split_fragment(left)
            left_text = inlines_to_text(left).strip()
            right_text = inlines_to_text(right).strip()
            if not left_text or not right_text:
                out.append(SourcePara(cur_inlines, inlines_to_text(cur_inlines).strip(), para.quote_depth, para.index))
                continue
            pending.insert(0, (right, right_text))
            pending.insert(0, (left, left_text))
    for i, para in enumerate(out):
        para.index = i
    return out


def merge_or_continue(prev_block: Block, para: SourcePara) -> None:
    prev_text = prev_block.text
    starts_lower = para.text[:1].islower()
    ends_mid_sentence = bool(prev_text) and not re.search(r"[.!?;:]$", prev_text)
    if starts_lower or ends_mid_sentence:
        prev_block.inlines.append({"t": "Space"})
        prev_block.inlines.extend(para.inlines)
        return
    prev_block.continuations.append(para.inlines)


def classify(source_paras: List[SourcePara]) -> List[Block]:
    blocks: List[Block] = []
    seen_title = False
    prev_l1: Optional[str] = None
    prev_l4: Optional[str] = None
    last_list: Optional[Block] = None
    next_id = 0

    for para in source_paras:
        text = para.text.strip()
        if not text:
            continue
        if is_running_header_text(text):
            continue

        if is_title(text, seen_title):
            seen_title = True
            blocks.append(Block(f"b{next_id:04d}", "title", inlines=para.inlines, source_index=para.index, quote_depth=para.quote_depth))
            next_id += 1
            last_list = None
            prev_l1 = None
            prev_l4 = None
            continue
        if is_section_heading(text):
            blocks.append(Block(f"b{next_id:04d}", "section_heading", inlines=para.inlines, source_index=para.index, quote_depth=para.quote_depth))
            next_id += 1
            last_list = None
            prev_l1 = None
            prev_l4 = None
            continue

        level, token, stripped, ambiguous = parse_leading_marker(para.inlines, prev_l1, prev_l4)
        if level is not None:
            block = Block(
                id=f"b{next_id:04d}",
                kind="list_item",
                level=level,
                inlines=stripped,
                source_index=para.index,
                quote_depth=para.quote_depth,
                flags=["ambiguous_marker"] if ambiguous else [],
            )
            blocks.append(block)
            next_id += 1
            last_list = block
            if level == 1:
                prev_l1 = token
            elif level == 4:
                prev_l4 = token
            continue

        if is_subheading(text):
            blocks.append(Block(f"b{next_id:04d}", "subheading", inlines=para.inlines, source_index=para.index, quote_depth=para.quote_depth))
            next_id += 1
            last_list = None
            prev_l1 = None
            prev_l4 = None
            continue

        if last_list is not None and (para.quote_depth > 0 or last_list.level is not None):
            merge_or_continue(last_list, para)
            continue

        blocks.append(Block(f"b{next_id:04d}", "body", inlines=para.inlines, source_index=para.index, quote_depth=para.quote_depth))
        next_id += 1
        last_list = None
        prev_l1 = None
        prev_l4 = None

    return blocks


def item_to_pandoc_blocks(block: Block) -> List[dict]:
    if not block.continuations:
        return [{"t": "Para", "c": block.inlines}]
    merged = list(block.inlines)
    for cont in block.continuations:
        merged.extend([{"t": "LineBreak"}, {"t": "LineBreak"}])
        merged.extend(cont)
    return [{"t": "Para", "c": merged}]


def compose_list(blocks: List[Block], i: int) -> Tuple[dict, int]:
    base_level = blocks[i].level or 0
    items: List[List[dict]] = []
    while i < len(blocks) and blocks[i].kind == "list_item" and (blocks[i].level or 0) >= base_level:
        cur = blocks[i]
        if (cur.level or 0) == base_level:
            item_blocks = item_to_pandoc_blocks(cur)
            items.append(item_blocks)
            i += 1
            if i < len(blocks) and blocks[i].kind == "list_item" and (blocks[i].level or 0) > base_level:
                nested, i = compose_list(blocks, i)
                items[-1].append(nested)
        else:
            nested, i = compose_list(blocks, i)
            if items:
                items[-1].append(nested)
            else:
                items.append([nested])
    return {"t": "OrderedList", "c": [LEVEL_TO_OL_ATTR[base_level], items]}, i


def compose_ast(blocks: List[Block]) -> dict:
    pandoc_blocks: List[dict] = []
    i = 0
    while i < len(blocks):
        block = blocks[i]
        if block.kind == "title":
            pandoc_blocks.append({"t": "Header", "c": [1, ["", [], []], block.inlines]})
            i += 1
        elif block.kind == "section_heading":
            pandoc_blocks.append({"t": "Header", "c": [2, ["", [], []], block.inlines]})
            i += 1
        elif block.kind == "subheading":
            pandoc_blocks.append({"t": "Header", "c": [3, ["", [], []], block.inlines]})
            i += 1
        elif block.kind == "list_item":
            ol, i = compose_list(blocks, i)
            pandoc_blocks.append(ol)
        else:
            pandoc_blocks.append({"t": "Para", "c": block.inlines})
            i += 1
    return {"pandoc-api-version": [1, 23, 1], "meta": {}, "blocks": pandoc_blocks}


def save_blocks(blocks: List[Block], path: Path) -> None:
    payload = []
    for block in blocks:
        payload.append({
            "id": block.id,
            "kind": block.kind,
            "level": block.level,
            "text": block.text,
            "continuations": [inlines_to_text(c).strip() for c in block.continuations],
            "source_index": block.source_index,
            "quote_depth": block.quote_depth,
            "flags": block.flags,
            "_inlines": block.inlines,
            "_continuations_inlines": block.continuations,
        })
    path.write_text(json.dumps(payload, indent=2) + "\n")


def load_blocks(path: Path) -> List[Block]:
    raw = json.loads(path.read_text())
    return [
        Block(
            id=item["id"],
            kind=item["kind"],
            level=item.get("level"),
            inlines=item.get("_inlines", []),
            continuations=item.get("_continuations_inlines", []),
            source_index=item.get("source_index", -1),
            quote_depth=item.get("quote_depth", 0),
            flags=item.get("flags", []),
        )
        for item in raw
    ]


def ensure_style(doc: Document, name: str) -> None:
    if name in doc.styles:
        return
    doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)


def apply_header(doc: Document, header_meta: HeaderMetadata) -> None:
    for section in doc.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.header_distance = Inches(0.35)
        section_header = section.header
        para = section_header.paragraphs[0] if section_header.paragraphs else section_header.add_paragraph()
        para.clear()
        para.paragraph_format.tab_stops.add_tab_stop(Inches(6.3), WD_TAB_ALIGNMENT.RIGHT)
        header_text = f"{header_meta.left_text}\t{header_meta.code}"
        run = para.add_run(header_text)
        run.font.name = "Inter"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Inter")
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(128, 128, 128)
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT


def style_docx(path: Path, header: HeaderMetadata) -> None:
    doc = Document(path)
    ensure_style(doc, "CorgiTitle")
    ensure_style(doc, "CorgiSection")
    ensure_style(doc, "CorgiSubheading")
    ensure_style(doc, "CorgiBody")
    apply_header(doc, header)

    for para in doc.paragraphs:
        style_name = para.style.name if para.style else ""
        if style_name == "Heading 1":
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_after = Pt(18)
            for run in para.runs:
                run.font.name = "Bricolage Grotesque"
                run._element.rPr.rFonts.set(qn("w:eastAsia"), "Bricolage Grotesque")
                run.font.size = Pt(26)
                run.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)
        elif style_name == "Heading 2":
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            para.paragraph_format.space_before = Pt(16)
            para.paragraph_format.space_after = Pt(8)
            for run in para.runs:
                run.font.name = "Bricolage Grotesque"
                run._element.rPr.rFonts.set(qn("w:eastAsia"), "Bricolage Grotesque")
                run.font.size = Pt(14)
                run.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)
        elif style_name == "Heading 3":
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            para.paragraph_format.space_before = Pt(10)
            para.paragraph_format.space_after = Pt(6)
            for run in para.runs:
                run.font.name = "Bricolage Grotesque"
                run._element.rPr.rFonts.set(qn("w:eastAsia"), "Bricolage Grotesque")
                run.font.size = Pt(12)
                run.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)
        else:
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            para.paragraph_format.space_after = Pt(6)
            for run in para.runs:
                run.font.name = "Inter"
                run._element.rPr.rFonts.set(qn("w:eastAsia"), "Inter")
                run.font.size = Pt(11)
                run.font.color.rgb = RGBColor(0, 0, 0)
    doc.save(path)


def patch_list_marker_style(path: Path) -> None:
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with zipfile.ZipFile(path, "r") as zin:
        items = [(item, zin.read(item.filename)) for item in zin.infolist()]
        root = etree.fromstring(zin.read("word/numbering.xml"))

    for lvl in root.xpath(".//w:lvl", namespaces=ns):
        for rpr in lvl.xpath("./w:rPr", namespaces=ns):
            lvl.remove(rpr)
        rpr = etree.SubElement(lvl, f"{{{ns['w']}}}rPr")
        etree.SubElement(
            rpr,
            f"{{{ns['w']}}}rFonts",
            attrib={
                f"{{{ns['w']}}}ascii": "Inter",
                f"{{{ns['w']}}}hAnsi": "Inter",
                f"{{{ns['w']}}}cs": "Inter",
                f"{{{ns['w']}}}eastAsia": "Inter",
            },
        )
        etree.SubElement(rpr, f"{{{ns['w']}}}color", attrib={f"{{{ns['w']}}}val": "000000"})
        etree.SubElement(rpr, f"{{{ns['w']}}}sz", attrib={f"{{{ns['w']}}}val": "22"})
        etree.SubElement(rpr, f"{{{ns['w']}}}szCs", attrib={f"{{{ns['w']}}}val": "22"})

    numbering = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

    tmp = path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item, data in items:
            if item.filename == "word/numbering.xml":
                data = numbering
            zout.writestr(item, data)
    tmp.replace(path)


def patch_ooxml(path: Path) -> None:
    patch_list_levels(path)
    set_list_suffix(path, "space")
    strip_list_ind_overrides(path)
    patch_list_marker_style(path)


def render(ast_path: Path, output_docx: Path) -> None:
    subprocess.check_call(["pandoc", str(ast_path), "-f", "json", "-t", "docx", "-o", str(output_docx)])


def build_report(blocks: List[Block], path: Path, output_docx: Path) -> None:
    doc = Document(output_docx)
    list_paras = 0
    heading1 = 0
    heading2 = 0
    heading3 = 0
    first_paras: List[str] = []
    for para in doc.paragraphs:
        if para.text.strip() and len(first_paras) < 20:
            first_paras.append(para.text.strip())
        style_name = para.style.name if para.style else ""
        if style_name == "Heading 1":
            heading1 += 1
        elif style_name == "Heading 2":
            heading2 += 1
        elif style_name == "Heading 3":
            heading3 += 1
        if "List" in style_name:
            list_paras += 1
    kind_counts: Dict[str, int] = {}
    flags: Dict[str, int] = {}
    for block in blocks:
        kind_counts[block.kind] = kind_counts.get(block.kind, 0) + 1
        for flag in block.flags:
            flags[flag] = flags.get(flag, 0) + 1
    path.write_text(json.dumps({
        "output": str(output_docx),
        "total_blocks": len(blocks),
        "kind_counts": kind_counts,
        "flags": flags,
        "docx": {
            "paragraphs": len(doc.paragraphs),
            "list_style_paragraphs": list_paras,
            "heading1": heading1,
            "heading2": heading2,
            "heading3": heading3,
            "first_nonempty_paragraphs": first_paras,
        },
    }, indent=2) + "\n")


def default_artifact_paths(output_docx: Path, artifacts_dir: Path) -> Tuple[Path, Path, Path]:
    return (
        artifacts_dir / f"{output_docx.stem}.blocks.json",
        artifacts_dir / f"{output_docx.stem}.ast.json",
        artifacts_dir / f"{output_docx.stem}.report.json",
    )


def run(
    source_docx: Path,
    output_docx: Path,
    *,
    artifacts_dir: Optional[Path] = None,
    blocks_in: Optional[Path] = None,
) -> None:
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    if artifacts_dir is not None:
        artifacts_dir.mkdir(parents=True, exist_ok=True)

    pandoc_json = subprocess.check_output(["pandoc", str(source_docx), "-t", "json"])
    source_ast = json.loads(pandoc_json)
    source_paras = flatten_paras(source_ast["blocks"])

    if blocks_in:
        blocks = load_blocks(blocks_in)
    else:
        paras = split_on_embedded_markers(source_paras)
        blocks = classify(paras)

    header = extract_header_metadata(source_paras, blocks)
    ast = compose_ast(blocks)

    if artifacts_dir is not None:
        blocks_out, ast_out, report_out = default_artifact_paths(output_docx, artifacts_dir)
        save_blocks(blocks, blocks_out)
        ast_out.write_text(json.dumps(ast) + "\n")
        render(ast_out, output_docx)
    else:
        with tempfile.NamedTemporaryFile(mode="w", suffix=".ast.json", delete=False) as tf:
            tf.write(json.dumps(ast))
            ast_tmp = Path(tf.name)
        try:
            render(ast_tmp, output_docx)
        finally:
            ast_tmp.unlink(missing_ok=True)

    style_docx(output_docx, header)
    patch_ooxml(output_docx)

    if artifacts_dir is not None:
        build_report(blocks, report_out, output_docx)
