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

from formatter import apply_page_layout_and_header


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="insure-policy-format-rule-3",
        description="Converge page layout and the running header to the canonical state.",
    )
    parser.add_argument("source_docx", type=Path)
    parser.add_argument("target_docx", type=Path)
    parser.add_argument("--parts-in", type=Path, required=True)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    apply_page_layout_and_header(
        args.source_docx,
        args.target_docx,
        parts_in=args.parts_in,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
