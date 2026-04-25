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

from formatter import normalize_outlines_to_docx


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="insure-policy-format-rule-0",
        description=(
            "Rewrite non-canonical outline markers in a Corgi-Tech policy DOCX "
            "(e.g. uppercase-letter top level 'A./B./C.') to the canonical "
            "sequence '1)/a)/i)/(1)/(a)/(i)'. Pre-pass before Rule 1."
        ),
    )
    parser.add_argument("source_docx", type=Path)
    parser.add_argument("-o", "--output", type=Path, required=True, help="Output DOCX")
    parser.add_argument(
        "--parts-in",
        type=Path,
        required=True,
        help=(
            "Document-parts manifest JSON. Must include 'outline_normalizations' "
            "describing which sections to remap."
        ),
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    normalize_outlines_to_docx(
        args.source_docx,
        args.output,
        parts_in=args.parts_in,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
