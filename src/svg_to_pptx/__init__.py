"""Utilities for converting SVG graphics into PowerPoint presentations."""

from .parser import parse_svg
from .pptx_builder import build_presentation

__all__ = ["parse_svg", "build_presentation"]
