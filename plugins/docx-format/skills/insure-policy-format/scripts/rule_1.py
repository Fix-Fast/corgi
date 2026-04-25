#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.11,<3.13"
# dependencies = [
#   "python-docx>=1.2.0",
#   "lxml>=5.4.0",
# ]
# ///
from __future__ import annotations

import argparse
from pathlib import Path

from formatter import build_text_hierarchy_docx


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="insure-policy-format-rule-1",
        description="Converge the document text hierarchy: title, section headings, subheadings, and body text.",
    )
    parser.add_argument("source_docx", type=Path)
    parser.add_argument("-o", "--output", type=Path, required=True, help="Output DOCX")
    parser.add_argument("--parts-in", type=Path, required=True, help="Document-parts manifest JSON")
    parser.add_argument("--ast-out", type=Path, default=None, help="Optional pandoc AST JSON output")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    build_text_hierarchy_docx(
        args.source_docx,
        args.output,
        parts_in=args.parts_in,
        ast_out=args.ast_out,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
