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
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

REPO_ROOT = Path(__file__).resolve().parents[2]
FIXTURE_DIR = Path(__file__).resolve().parent
ORIGINAL = FIXTURE_DIR / "cgl_original.docx"
GOLDEN = FIXTURE_DIR / "cgl_original_formatted.docx"
PARTS = FIXTURE_DIR / "cgl_original.parts.json"
PARTS_FORMATTED = FIXTURE_DIR / "cgl_original_formatted.parts.json"
FORMATTER = REPO_ROOT / "plugins/docx/skills/insure-policy-format/scripts/format.py"

VOLATILE_ATTR_RE = re.compile(r"\{[^}]*}(rsidR|rsidRDefault|rsidRPr|rsidP|rsidTr|rsidSect|rsidRoot|paraId|textId)$")
SKIP_MEMBERS = {"docProps/core.xml"}


def run_formatter(source: Path, output: Path, parts: Path) -> None:
    subprocess.run(
        ["uv", "run", str(FORMATTER), str(source), "-o", str(output), "--parts-in", str(parts)],
        check=True,
        capture_output=True,
    )


def canonicalize_xml(data: bytes) -> str:
    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(data, parser=parser)
    for el in root.iter():
        for attr in list(el.attrib):
            if VOLATILE_ATTR_RE.search(attr):
                del el.attrib[attr]
    return etree.tostring(root, pretty_print=True, encoding="unicode")


def diff_docx(produced: Path, golden: Path) -> list[str]:
    with zipfile.ZipFile(produced) as p_zip, zipfile.ZipFile(golden) as g_zip:
        p_names = set(p_zip.namelist()) - SKIP_MEMBERS
        g_names = set(g_zip.namelist()) - SKIP_MEMBERS
        problems: list[str] = []
        only_p = sorted(p_names - g_names)
        only_g = sorted(g_names - p_names)
        for n in only_p:
            problems.append(f"only in produced: {n}")
        for n in only_g:
            problems.append(f"only in golden: {n}")
        for name in sorted(p_names & g_names):
            p_bytes = p_zip.read(name)
            g_bytes = g_zip.read(name)
            if name.endswith(".xml") or name.endswith(".rels"):
                p_norm = canonicalize_xml(p_bytes)
                g_norm = canonicalize_xml(g_bytes)
                if p_norm != g_norm:
                    diff = list(
                        difflib.unified_diff(
                            g_norm.splitlines(keepends=True),
                            p_norm.splitlines(keepends=True),
                            fromfile=f"golden:{name}",
                            tofile=f"produced:{name}",
                            n=2,
                        )
                    )
                    problems.append(f"--- {name} ---\n" + "".join(diff[:200]))
            else:
                if p_bytes != g_bytes:
                    problems.append(f"binary mismatch: {name} (produced={len(p_bytes)}B golden={len(g_bytes)}B)")
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
    args = parser.parse_args()

    workdir = Path(tempfile.mkdtemp(prefix="cgl_golden_"))
    print(f"workdir: {workdir}")
    produced_from_orig = workdir / "from_original.docx"
    produced_from_golden = workdir / "from_golden.docx"

    print("[1/2] format(cgl_original.docx) vs golden")
    run_formatter(ORIGINAL, produced_from_orig, PARTS)
    ok1 = assert_equal("reproduces golden", produced_from_orig, GOLDEN, args.max_diff_lines)

    print("[2/2] format(cgl_original_formatted.docx) vs golden (idempotency)")
    run_formatter(GOLDEN, produced_from_golden, PARTS_FORMATTED)
    ok2 = assert_equal("idempotent on golden", produced_from_golden, GOLDEN, args.max_diff_lines)

    if not args.keep:
        shutil.rmtree(workdir, ignore_errors=True)

    return 0 if (ok1 and ok2) else 1


if __name__ == "__main__":
    raise SystemExit(main())
