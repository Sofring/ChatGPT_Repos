"""Command line interface for the svg_to_pptx package."""

from __future__ import annotations

import argparse
from pathlib import Path

from .parser import parse_svg
from .pptx_builder import build_presentation


def create_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Convert SVG files into PowerPoint presentations")
    parser.add_argument("input", type=Path, help="Path to the source SVG file")
    parser.add_argument("output", type=Path, help="Desired output PPTX path")
    return parser


def main(argv: list[str] | None = None) -> Path:
    parser = create_parser()
    args = parser.parse_args(argv)
    document = parse_svg(args.input)
    return build_presentation(document, args.output)


if __name__ == "__main__":  # pragma: no cover
    main()
