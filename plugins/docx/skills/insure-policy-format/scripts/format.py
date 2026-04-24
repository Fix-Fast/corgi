#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.11,<3.13"
# dependencies = [
#   "python-docx>=1.2.0",
#   "lxml>=5.4.0",
# ]
# ///
"""Corgi insurance-policy DOCX formatter.

Invoked by the `insure-policy-format` skill via `uv run`. Reformats a
Corgi-Tech policy DOCX into the canonical heading / list / layout
conventions.

Usage:
    uv run format.py <input.docx> -o <output.docx> --parts-in <parts.json>
"""
from __future__ import annotations

import argparse
from pathlib import Path

from formatter import (
    apply_page_layout_and_header,
    build_text_hierarchy_docx,
    normalize_list_formatting,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="insure-policy-format",
        description="Format a Corgi-Tech insurance policy DOCX into canonical form.",
    )
    parser.add_argument("source_docx", type=Path, help="Input DOCX to format.")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Formatted output DOCX. Defaults next to the source as *.formatted.docx.",
    )
    parser.add_argument(
        "--parts-in",
        type=Path,
        required=True,
        help="Document-parts manifest JSON produced by Claude or another caller.",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    output_docx = args.output or args.source_docx.with_name(
        f"{args.source_docx.stem}.formatted.docx"
    )
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    build_text_hierarchy_docx(
        args.source_docx,
        output_docx,
        parts_in=args.parts_in,
    )
    apply_page_layout_and_header(
        args.source_docx,
        output_docx,
        parts_in=args.parts_in,
    )
    normalize_list_formatting(output_docx)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
