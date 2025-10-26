"""Data structures used during the SVG to PPTX conversion pipeline."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional, Sequence, Tuple


Color = Optional[str]


class ShapeType(str, Enum):
    """Supported shape types."""

    RECT = "rect"
    CIRCLE = "circle"
    ELLIPSE = "ellipse"
    LINE = "line"
    POLYLINE = "polyline"
    POLYGON = "polygon"
    PATH = "path"
    TEXT = "text"


@dataclass
class Style:
    """Common styling attributes shared across shapes."""

    fill_color: Color = None
    stroke_color: Color = None
    stroke_width: float = 1.0
    opacity: float = 1.0


@dataclass
class TextStyle(Style):
    """Style options that only apply to text nodes."""

    font_family: Optional[str] = None
    font_size: Optional[float] = None


@dataclass
class Shape:
    """Base shape definition."""

    shape_type: ShapeType
    style: Style


@dataclass
class Rect(Shape):
    x: float
    y: float
    width: float
    height: float
    rx: float = 0.0
    ry: float = 0.0


@dataclass
class Circle(Shape):
    cx: float
    cy: float
    r: float


@dataclass
class Ellipse(Shape):
    cx: float
    cy: float
    rx: float
    ry: float


@dataclass
class Line(Shape):
    x1: float
    y1: float
    x2: float
    y2: float


@dataclass
class Polyline(Shape):
    points: Sequence[Tuple[float, float]]
    closed: bool = False


@dataclass
class PathSegment:
    """A simplified path consisting of straight line segments."""

    points: List[Tuple[float, float]] = field(default_factory=list)
    closed: bool = False


@dataclass
class Path(Shape):
    segments: Sequence[PathSegment]


@dataclass
class Text(Shape):
    text: str
    x: float
    y: float


@dataclass
class Document:
    """The parsed SVG document."""

    width: Optional[float]
    height: Optional[float]
    shapes: Sequence[Shape]

