#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.11,<3.13"
# dependencies = [
#   "lxml>=5.4.0",
# ]
# ///
"""Corgi insurance-policy DOCX formatter — CLI orchestrator.

Reads the source DOCX as raw OOXML (no pandoc), applies the four rules
defined in format.md in numerical order (0 -> 1 -> 2 -> 3), and writes
the result. Each rule mutates the in-memory XML trees; this module
handles the docx zip I/O and orchestration.

Usage:
    uv run format.py <input.docx> -o <output.docx> --parts-in <parts.json>
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

import rule_0
import rule_1
import rule_2
import rule_3
from parts import resolve


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

HEADER_PART = f"word/{rule_3.HEADER_PART_NAME}"
HEADER_RELS_PART = f"word/_rels/{rule_3.HEADER_PART_NAME}.rels"
HEADER_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
)
HEADER_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
)


def _read_part(zf: zipfile.ZipFile, name: str) -> bytes:
    return zf.read(name)


def _ensure_relationship(rels_xml: bytes, rel_id: str, target: str, rel_type: str) -> bytes:
    root = etree.fromstring(rels_xml)
    for rel in root.findall(f"{{{RELS_NS}}}Relationship"):
        if rel.get("Id") == rel_id:
            rel.set("Type", rel_type)
            rel.set("Target", target)
            break
    else:
        new = etree.SubElement(root, f"{{{RELS_NS}}}Relationship")
        new.set("Id", rel_id)
        new.set("Type", rel_type)
        new.set("Target", target)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _ensure_header_content_type(ct_xml: bytes) -> bytes:
    root = etree.fromstring(ct_xml)
    for ovr in root.findall(f"{{{CT_NS}}}Override"):
        if ovr.get("PartName") == "/" + HEADER_PART:
            ovr.set("ContentType", HEADER_CONTENT_TYPE)
            return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    new = etree.SubElement(root, f"{{{CT_NS}}}Override")
    new.set("PartName", "/" + HEADER_PART)
    new.set("ContentType", HEADER_CONTENT_TYPE)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _serialize_tree(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def format_docx(source: Path, output: Path, parts_path: Path) -> None:
    with zipfile.ZipFile(source, "r") as zin:
        members: dict[str, bytes] = {n: _read_part(zin, n) for n in zin.namelist()}

    if "word/document.xml" not in members:
        raise ValueError("source docx is missing word/document.xml")
    doc_root = etree.fromstring(members["word/document.xml"])

    # Resolve parts against the *unmodified* source tree so substring
    # matching sees the document as the user authored it.
    resolved = resolve(parts_path, doc_root)

    # Numbering tree may not exist in source; create a minimal one if absent.
    if "word/numbering.xml" in members:
        numbering_root = etree.fromstring(members["word/numbering.xml"])
    else:
        numbering_root = etree.Element(W + "numbering", nsmap={"w": W_NS})

    # Apply the rules in numerical order. Each rule is a pure (tree, parts)
    # -> tree transformation; downstream rules read upstream effects from
    # the tree itself (e.g. Rule 2 reads pStyle on paragraphs to skip
    # headings styled by Rule 1).
    rule_0.apply(doc_root, resolved)
    rule_1.apply(doc_root, resolved)
    rule_2.apply(doc_root, numbering_root)
    _, header_bytes = rule_3.apply(doc_root, resolved)

    # Reassemble the docx.
    members["word/document.xml"] = _serialize_tree(doc_root)
    members["word/numbering.xml"] = _serialize_tree(numbering_root)
    members[HEADER_PART] = header_bytes
    members[HEADER_RELS_PART] = _build_header_rels()
    if "word/_rels/document.xml.rels" in members:
        members["word/_rels/document.xml.rels"] = _ensure_relationship(
            members["word/_rels/document.xml.rels"],
            rel_id=rule_3.HEADER_REL_ID,
            target=rule_3.HEADER_RELS_TARGET,
            rel_type=HEADER_REL_TYPE,
        )
    if "[Content_Types].xml" in members:
        members["[Content_Types].xml"] = _ensure_header_content_type(members["[Content_Types].xml"])

    # Ensure numbering.xml is registered as a part if we just created one.
    members = _ensure_numbering_registration(members)

    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in members.items():
                zout.writestr(name, data)
        shutil.move(str(tmp_path), str(output))
    finally:
        if tmp_path.exists():
            tmp_path.unlink()


def _build_header_rels() -> bytes:
    root = etree.Element(f"{{{RELS_NS}}}Relationships", nsmap={None: RELS_NS})
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _ensure_numbering_registration(members: dict[str, bytes]) -> dict[str, bytes]:
    """If word/numbering.xml is present but not registered, add the relationship + content-type."""
    if "word/numbering.xml" not in members:
        return members
    rels = members.get("word/_rels/document.xml.rels")
    if rels is not None:
        rels_root = etree.fromstring(rels)
        has_numbering = any(
            r.get("Type", "").endswith("/numbering")
            for r in rels_root.findall(f"{{{RELS_NS}}}Relationship")
        )
        if not has_numbering:
            new = etree.SubElement(rels_root, f"{{{RELS_NS}}}Relationship")
            new.set("Id", "rIdCorgiNumbering")
            new.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering")
            new.set("Target", "numbering.xml")
            members["word/_rels/document.xml.rels"] = etree.tostring(
                rels_root, xml_declaration=True, encoding="UTF-8", standalone=True,
            )
    ct = members.get("[Content_Types].xml")
    if ct is not None:
        ct_root = etree.fromstring(ct)
        has_numbering = any(
            o.get("PartName") == "/word/numbering.xml"
            for o in ct_root.findall(f"{{{CT_NS}}}Override")
        )
        if not has_numbering:
            new = etree.SubElement(ct_root, f"{{{CT_NS}}}Override")
            new.set("PartName", "/word/numbering.xml")
            new.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")
            members["[Content_Types].xml"] = etree.tostring(
                ct_root, xml_declaration=True, encoding="UTF-8", standalone=True,
            )
    return members


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="insure-policy-format",
        description="Format a Corgi-Tech insurance policy DOCX into canonical form.",
    )
    parser.add_argument("source_docx", type=Path)
    parser.add_argument("-o", "--output", type=Path, help="Output DOCX (default: source.formatted.docx)")
    parser.add_argument("--parts-in", type=Path, required=True)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    output = args.output or args.source_docx.with_name(f"{args.source_docx.stem}.formatted.docx")
    format_docx(args.source_docx, output, args.parts_in)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
