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
    uv run format.py <input.docx> -o <output.docx> [--artifacts-dir DIR]
"""
from __future__ import annotations

import argparse
from pathlib import Path

from formatter import run


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
        "--artifacts-dir",
        type=Path,
        default=None,
        help="If set, write debug artifacts (blocks.json, ast.json, report.json) here.",
    )
    parser.add_argument(
        "--blocks-in",
        type=Path,
        default=None,
        help="Optional precomputed blocks.json to render from.",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    output_docx = args.output or args.source_docx.with_name(
        f"{args.source_docx.stem}.formatted.docx"
    )
    run(
        args.source_docx,
        output_docx,
        artifacts_dir=args.artifacts_dir,
        blocks_in=args.blocks_in,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
