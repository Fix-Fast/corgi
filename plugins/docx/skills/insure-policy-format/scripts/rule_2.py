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

from formatter import normalize_list_formatting


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="insure-policy-format-rule-2",
        description="Converge list structure and list formatting to the canonical state.",
    )
    parser.add_argument("target_docx", type=Path)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    normalize_list_formatting(args.target_docx)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
