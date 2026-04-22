"""Byte-level OOXML patches applied to the rendered DOCX.

These functions were originally three separate CLI scripts invoked via
nested `uv run` subprocesses. They're now in-process library calls so the
package is self-contained and portable.

All three use byte-level regex surgery (not XML parsing) to preserve
namespace declarations (mc, w14, w15, ...) that a namespace-unaware parser
would drop.
"""
from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from pathlib import Path


LEVEL_SPEC: dict[int, tuple[str, str, int]] = {
    0: ("decimal",     "%1)",   720),
    1: ("lowerLetter", "%2)",  1440),
    2: ("lowerRoman",  "%3)",  2160),
    3: ("decimal",     "(%4)", 2880),
    4: ("lowerLetter", "(%5)", 3600),
    5: ("lowerRoman",  "(%6)", 4320),
}
HANGING_TWIPS = 360

RE_ABSTRACT_NUM = re.compile(
    rb'<w:abstractNum\s+[^>]*w:abstractNumId="(\d+)"[^>]*>(.*?)</w:abstractNum>',
    re.DOTALL,
)
RE_LVL = re.compile(
    rb'<w:lvl\s+[^>]*w:ilvl="(\d+)"[^>]*>(.*?)</w:lvl>', re.DOTALL
)
RE_NUMFMT = re.compile(rb'<w:numFmt\s+[^/>]*/>')
RE_LVLTEXT = re.compile(rb'<w:lvlText\s+[^/>]*/>')
RE_PPR = re.compile(rb'<w:pPr\b[^>]*>.*?</w:pPr>', re.DOTALL)
RE_IND = re.compile(rb'<w:ind\b[^/>]*/>|<w:ind\b[^>]*>.*?</w:ind>', re.DOTALL)
RE_RPR = re.compile(rb'<w:rPr\b[^>]*>.*?</w:rPr>|<w:rPr\b[^/>]*/>', re.DOTALL)
RE_NUM = re.compile(rb'<w:num\s+[^>]*w:numId="(\d+)"[^>]*>(.*?)</w:num>', re.DOTALL)
RE_ABSTRACT_NUM_REF = re.compile(rb'<w:abstractNumId\s+[^>]*w:val="(\d+)"')
RE_SUFF = re.compile(rb'<w:suff\s+[^/>]*/>')
RE_PARA = re.compile(rb"<w:p\b[^>]*>.*?</w:p>", re.DOTALL)
RE_NUMPR = re.compile(rb"<w:numPr\b")


def _rewrite_part(docx: Path, part: str, new_bytes: bytes) -> None:
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td) / "out.docx"
        with zipfile.ZipFile(docx, "r") as zin, zipfile.ZipFile(
            tmp, "w", zipfile.ZIP_DEFLATED
        ) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == part:
                    data = new_bytes
                zout.writestr(item, data)
        shutil.move(str(tmp), str(docx))


def _replace_or_insert(
    body: bytes,
    pattern: re.Pattern[bytes],
    new_tag: bytes,
    anchor_after: re.Pattern[bytes] | None = None,
) -> bytes:
    m = pattern.search(body)
    if m:
        return body[: m.start()] + new_tag + body[m.end():]
    if anchor_after:
        am = anchor_after.search(body)
        if am:
            return body[: am.end()] + new_tag + body[am.end():]
    return new_tag + body


def _patch_ind_inside_ppr(lvl_body: bytes, left: int, hanging: int) -> bytes:
    new_ind = f'<w:ind w:left="{left}" w:hanging="{hanging}"/>'.encode()
    ppm = RE_PPR.search(lvl_body)
    if ppm:
        ppr = ppm.group(0)
        new_ppr, n = RE_IND.subn(new_ind, ppr, count=1)
        if n == 0:
            close = ppr.rfind(b"</w:pPr>")
            new_ppr = ppr[:close] + new_ind + ppr[close:]
        return lvl_body[: ppm.start()] + new_ppr + lvl_body[ppm.end():]
    ppr_block = b"<w:pPr>" + new_ind + b"</w:pPr>"
    return lvl_body + ppr_block


def _patch_lvl_body(lvl_body: bytes, ilvl: int) -> bytes:
    fmt, text, left = LEVEL_SPEC[ilvl]
    new_numfmt = f'<w:numFmt w:val="{fmt}"/>'.encode()
    new_lvltext = f'<w:lvlText w:val="{text}"/>'.encode()
    lvl_body = _replace_or_insert(lvl_body, RE_NUMFMT, new_numfmt)
    lvl_body = _replace_or_insert(
        lvl_body,
        RE_LVLTEXT,
        new_lvltext,
        anchor_after=re.compile(re.escape(new_numfmt)),
    )
    lvl_body = _patch_ind_inside_ppr(lvl_body, left, HANGING_TWIPS)
    lvl_body = RE_RPR.sub(b"", lvl_body)
    return lvl_body


def _patch_abstract_body_levels(body: bytes) -> bytes:
    def replace(m: re.Match[bytes]) -> bytes:
        ilvl = int(m.group(1).decode())
        if ilvl not in LEVEL_SPEC:
            return m.group(0)
        inner_start = m.start(2) - m.start(0)
        inner_end = m.end(2) - m.start(0)
        full = m.group(0)
        new_body = _patch_lvl_body(m.group(2), ilvl)
        return full[:inner_start] + new_body + full[inner_end:]

    return RE_LVL.sub(replace, body)


def patch_list_levels(docx: Path) -> None:
    """Force canonical L0-L5 definitions on every abstractNum."""
    with zipfile.ZipFile(docx, "r") as z:
        xml = z.read("word/numbering.xml")

    def replace(m: re.Match[bytes]) -> bytes:
        inner_start = m.start(2) - m.start(0)
        inner_end = m.end(2) - m.start(0)
        full = m.group(0)
        new_body = _patch_abstract_body_levels(m.group(2))
        return full[:inner_start] + new_body + full[inner_end:]

    new_xml = RE_ABSTRACT_NUM.sub(replace, xml)
    if new_xml != xml:
        _rewrite_part(docx, "word/numbering.xml", new_xml)


def _patch_lvl_suff(lvl_body: bytes, suff_val: str) -> bytes:
    new_suff = f'<w:suff w:val="{suff_val}"/>'.encode()
    if RE_SUFF.search(lvl_body):
        return RE_SUFF.sub(new_suff, lvl_body, count=1)
    nm = RE_NUMFMT.search(lvl_body)
    insert_at = nm.end() if nm else 0
    return lvl_body[:insert_at] + new_suff + lvl_body[insert_at:]


def set_list_suffix(docx: Path, suff_val: str = "space") -> None:
    """Set <w:suff> on every level of every list to suff_val."""
    if suff_val not in ("tab", "space", "nothing"):
        raise ValueError(f"invalid suff_val: {suff_val!r}")
    with zipfile.ZipFile(docx, "r") as z:
        xml = z.read("word/numbering.xml")

    def patch_abstract_body(body: bytes) -> bytes:
        def replace(m: re.Match[bytes]) -> bytes:
            inner_start = m.start(2) - m.start(0)
            inner_end = m.end(2) - m.start(0)
            full = m.group(0)
            new_body = _patch_lvl_suff(m.group(2), suff_val)
            return full[:inner_start] + new_body + full[inner_end:]

        return RE_LVL.sub(replace, body)

    def replace_abs(m: re.Match[bytes]) -> bytes:
        inner_start = m.start(2) - m.start(0)
        inner_end = m.end(2) - m.start(0)
        full = m.group(0)
        return full[:inner_start] + patch_abstract_body(m.group(2)) + full[inner_end:]

    new_xml = RE_ABSTRACT_NUM.sub(replace_abs, xml)
    if new_xml != xml:
        _rewrite_part(docx, "word/numbering.xml", new_xml)


def strip_list_ind_overrides(docx: Path) -> None:
    """Remove <w:ind> from every list-item paragraph's pPr."""
    with zipfile.ZipFile(docx, "r") as z:
        xml = z.read("word/document.xml")

    def patch_para(pm: re.Match[bytes]) -> bytes:
        para = pm.group(0)
        ppm = RE_PPR.search(para)
        if not ppm:
            return para
        ppr = ppm.group(0)
        if not RE_NUMPR.search(ppr):
            return para
        new_ppr, n = RE_IND.subn(b"", ppr)
        if n == 0:
            return para
        return para[: ppm.start()] + new_ppr + para[ppm.end():]

    new_xml = RE_PARA.sub(patch_para, xml)
    if new_xml != xml:
        _rewrite_part(docx, "word/document.xml", new_xml)
