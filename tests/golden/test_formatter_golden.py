#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.11,<3.13"
# dependencies = [
#   "lxml>=5.4.0",
# ]
# ///
"""Golden test for insure-policy-format.

Runs two assertions against the bundled CGL fixture:

    1. format(cgl_original.docx, parts=cgl_original.parts.json) == cgl_original_formatted.docx
       (the formatter reproduces the canonical golden from the raw source)

    2. format(cgl_original_formatted.docx, parts=cgl_original.parts.json) == cgl_original_formatted.docx
       (running the formatter on the golden is a no-op — every rule is a fixed point)

Comparison is OOXML-level: each .xml member of the .docx zip is canonicalized
and compared, with volatile bits stripped (rsid revision IDs, paraId/textId,
docProps timestamps). Non-XML members (fonts, etc.) are compared by exact
bytes.

Run directly:

    uv run tests/golden/test_formatter_golden.py
"""
from __future__ import annotations

import argparse
import difflib
import hashlib
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

REPO_ROOT = Path(__file__).resolve().parents[2]
FIXTURE_DIR = Path(__file__).resolve().parent
FORMATTER = REPO_ROOT / "plugins/docx/skills/insure-policy-format/scripts/format.py"


@dataclass
class Fixture:
    name: str
    original: Path
    golden: Path
    parts: Path
    parts_formatted: Path | None  # None disables the idempotency check

FIXTURES = [
    Fixture(
        name="cgl",
        original=FIXTURE_DIR / "cgl_original.docx",
        golden=FIXTURE_DIR / "cgl_original_formatted.docx",
        parts=FIXTURE_DIR / "cgl_original.parts.json",
        parts_formatted=FIXTURE_DIR / "cgl_original_formatted.parts.json",
    ),
    Fixture(
        name="seic_do",
        original=FIXTURE_DIR / "seic_do_original.docx",
        golden=FIXTURE_DIR / "seic_do_original_formatted.docx",
        parts=FIXTURE_DIR / "seic_do_original.parts.json",
        parts_formatted=FIXTURE_DIR / "seic_do_original_formatted.parts.json",
    ),
]

VOLATILE_ATTR_RE = re.compile(r"\{[^}]*}(rsidR|rsidRDefault|rsidRPr|rsidP|rsidTr|rsidSect|rsidRoot|paraId|textId)$")
SKIP_MEMBERS = {"docProps/core.xml"}


def run_formatter(source: Path, output: Path, parts: Path) -> None:
    subprocess.run(
        ["uv", "run", str(FORMATTER), str(source), "-o", str(output), "--parts-in", str(parts)],
        check=True,
        capture_output=True,
    )


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


_PPR_TAG = W + "pPr"
_RPR_TAG = W + "rPr"
# Sort the children of pPr / rPr deterministically. OOXML doesn't render
# differently based on the order of these property children, but a string
# diff of pretty-printed XML does, so canonicalize before comparing.
_SORT_TAGS = {_PPR_TAG, _RPR_TAG}


def _strip_volatile_attrs(root: etree._Element) -> None:
    for el in root.iter():
        for attr in list(el.attrib):
            if VOLATILE_ATTR_RE.search(attr):
                del el.attrib[attr]


def _sort_property_children(root: etree._Element) -> None:
    """Sort children of every pPr/rPr by their tag, stably.

    The schema technically prescribes an order, but Word renders the same
    paragraph regardless. Sorting by tag makes the diff insensitive to
    formatter implementation choices about insertion order.
    """
    for el in root.iter():
        if el.tag in _SORT_TAGS:
            children = sorted(el, key=lambda c: c.tag)
            for c in list(el):
                el.remove(c)
            for c in children:
                el.append(c)


def _hash_element(el: etree._Element) -> str:
    """Stable content hash for an OOXML element, ignoring its own ID-bearing attrs.

    Also strips the *inner* `<w:abstractNumId>` reference inside a `<w:num>`
    so that a `<w:num>`'s hash depends only on its own structure (e.g.
    lvlOverrides), not on which abstractNum it points at. Otherwise a
    pure content change to the canonical abstractNum (e.g. updating
    lvlText) would cascade into different numId hashes in document.xml.
    """
    clone = etree.fromstring(etree.tostring(el))
    for attr in (f"{W}abstractNumId", f"{W}numId"):
        if attr in clone.attrib:
            del clone.attrib[attr]
    for child in clone.findall(f"{W}abstractNumId"):
        clone.remove(child)
    canonical = etree.tostring(clone, method="c14n2")
    return hashlib.sha1(canonical).hexdigest()[:12]


def renumber_ids(numbering_xml: bytes, document_xml: bytes) -> tuple[bytes, bytes]:
    """Rewrite abstractNumId / numId values to content-derived hashes.

    Word's `numId` and `abstractNumId` are arbitrary integers; allocation order
    is not part of the document's meaning. Two formatters that produce the same
    numbering definitions but in different allocation orders are equivalent.
    Normalize so the diff doesn't flag pure ID reshuffles.
    """
    num_root = etree.fromstring(numbering_xml)
    abstract_id_map: dict[str, str] = {}
    for ab in num_root.findall(f"{W}abstractNum"):
        old = ab.get(f"{W}abstractNumId")
        abstract_id_map[old] = "A" + _hash_element(ab)

    # First pass: rename abstractNumId references inside <w:num> so num hashes
    # are computed against the canonicalized form.
    for num in num_root.findall(f"{W}num"):
        ref = num.find(f"{W}abstractNumId")
        if ref is not None:
            old = ref.get(f"{W}val")
            if old in abstract_id_map:
                ref.set(f"{W}val", abstract_id_map[old])

    num_id_map: dict[str, str] = {}
    for num in num_root.findall(f"{W}num"):
        old = num.get(f"{W}numId")
        num_id_map[old] = "N" + _hash_element(num)

    # Apply renames in numbering.xml.
    for ab in num_root.findall(f"{W}abstractNum"):
        old = ab.get(f"{W}abstractNumId")
        ab.set(f"{W}abstractNumId", abstract_id_map[old])
    for num in num_root.findall(f"{W}num"):
        old = num.get(f"{W}numId")
        num.set(f"{W}numId", num_id_map[old])

    # Sort children for stable ordering.
    abstracts = sorted(num_root.findall(f"{W}abstractNum"), key=lambda e: e.get(f"{W}abstractNumId"))
    nums = sorted(num_root.findall(f"{W}num"), key=lambda e: e.get(f"{W}numId"))
    others = [c for c in num_root if c.tag not in (f"{W}abstractNum", f"{W}num")]
    for child in list(num_root):
        num_root.remove(child)
    for c in others + abstracts + nums:
        num_root.append(c)

    # Apply numId renames in document.xml.
    doc_root = etree.fromstring(document_xml)
    for el in doc_root.iter(f"{W}numId"):
        old = el.get(f"{W}val")
        if old in num_id_map:
            el.set(f"{W}val", num_id_map[old])

    return (
        etree.tostring(num_root, xml_declaration=True, encoding="UTF-8", standalone=True),
        etree.tostring(doc_root, xml_declaration=True, encoding="UTF-8", standalone=True),
    )


def canonicalize_xml(data: bytes) -> str:
    root = etree.fromstring(data)
    _strip_volatile_attrs(root)
    _sort_property_children(root)
    return etree.tostring(root, pretty_print=True, encoding="unicode")


def load_canonical(path: Path) -> dict[str, object]:
    """Read a docx, return {member_name: canonical_str_or_bytes}, with numbering IDs normalized."""
    with zipfile.ZipFile(path) as zf:
        members = {n: zf.read(n) for n in zf.namelist() if n not in SKIP_MEMBERS}
    if "word/numbering.xml" in members and "word/document.xml" in members:
        members["word/numbering.xml"], members["word/document.xml"] = renumber_ids(
            members["word/numbering.xml"], members["word/document.xml"]
        )
    out: dict[str, object] = {}
    for name, data in members.items():
        if name.endswith(".xml") or name.endswith(".rels"):
            out[name] = canonicalize_xml(data)
        else:
            out[name] = data
    return out


def diff_docx(produced: Path, golden: Path) -> list[str]:
    p_members = load_canonical(produced)
    g_members = load_canonical(golden)
    problems: list[str] = []
    for n in sorted(set(p_members) - set(g_members)):
        problems.append(f"only in produced: {n}")
    for n in sorted(set(g_members) - set(p_members)):
        problems.append(f"only in golden: {n}")
    for name in sorted(set(p_members) & set(g_members)):
        p_v, g_v = p_members[name], g_members[name]
        if p_v == g_v:
            continue
        if isinstance(p_v, str):
            diff = list(
                difflib.unified_diff(
                    g_v.splitlines(keepends=True),
                    p_v.splitlines(keepends=True),
                    fromfile=f"golden:{name}",
                    tofile=f"produced:{name}",
                    n=2,
                )
            )
            problems.append(f"--- {name} ---\n" + "".join(diff[:200]))
        else:
            problems.append(f"binary mismatch: {name} (produced={len(p_v)}B golden={len(g_v)}B)")
    return problems


def assert_equal(label: str, produced: Path, golden: Path, max_lines: int) -> bool:
    problems = diff_docx(produced, golden)
    if not problems:
        print(f"  PASS: {label}")
        return True
    print(f"  FAIL: {label} ({len(problems)} mismatching parts)")
    for chunk in problems:
        for line in chunk.splitlines()[:max_lines]:
            print(f"    {line}")
        print()
    return False


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--max-diff-lines",
        type=int,
        default=40,
        help="Cap on lines of unified diff per failing part (default: 40).",
    )
    parser.add_argument(
        "--keep",
        action="store_true",
        help="Keep produced files in /tmp for inspection.",
    )
    parser.add_argument(
        "--regenerate",
        action="store_true",
        help="Overwrite each golden with the formatter's current output. "
             "Use only when an intentional spec/script change has been "
             "reviewed and the new output is the desired golden.",
    )
    args = parser.parse_args()

    workdir = Path(tempfile.mkdtemp(prefix="formatter_golden_"))
    print(f"workdir: {workdir}")
    all_ok = True

    for fx in FIXTURES:
        print(f"\n=== fixture: {fx.name} ===")
        produced_from_orig = workdir / f"{fx.name}_from_original.docx"
        run_formatter(fx.original, produced_from_orig, fx.parts)
        if args.regenerate:
            shutil.copyfile(produced_from_orig, fx.golden)
            print(f"  REGENERATED: {fx.golden}")
        ok1 = assert_equal(
            f"{fx.name}: reproduces golden",
            produced_from_orig, fx.golden, args.max_diff_lines,
        )
        all_ok = all_ok and ok1

        if fx.parts_formatted is None:
            print(f"  SKIP: {fx.name}: idempotency check disabled")
            continue
        produced_from_golden = workdir / f"{fx.name}_from_golden.docx"
        run_formatter(fx.golden, produced_from_golden, fx.parts_formatted)
        ok2 = assert_equal(
            f"{fx.name}: idempotent on golden",
            produced_from_golden, fx.golden, args.max_diff_lines,
        )
        all_ok = all_ok and ok2

    if not args.keep:
        shutil.rmtree(workdir, ignore_errors=True)

    return 0 if all_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
