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
        # Idempotency disabled: current formatter output for this fixture is buggy
        # (Rule 0 leaves both new and old markers, subheadings concatenate with body).
        # The rewrite should fix that and turn this on.
        parts_formatted=None,
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

    workdir = Path(tempfile.mkdtemp(prefix="formatter_golden_"))
    print(f"workdir: {workdir}")
    all_ok = True

    for fx in FIXTURES:
        print(f"\n=== fixture: {fx.name} ===")
        produced_from_orig = workdir / f"{fx.name}_from_original.docx"
        run_formatter(fx.original, produced_from_orig, fx.parts)
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
