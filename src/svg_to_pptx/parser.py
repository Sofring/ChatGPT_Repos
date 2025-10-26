"""Parse SVG files into an intermediate representation suitable for pptx generation."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from pathlib import Path as PathlibPath
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from .models import (
    Circle,
    Document,
    Ellipse,
    Line,
    Path as PathShape,
    PathSegment,
    Polyline,
    Rect,
    Shape,
    ShapeType,
    Style,
    Text,
    TextStyle,
)

_LENGTH_RE = re.compile(r"(?P<value>-?\d+(?:\.\d+)?)(?P<unit>[a-z%]*)", re.IGNORECASE)
_COLOR_NAMES: Dict[str, str] = {
    "black": "#000000",
    "white": "#ffffff",
    "red": "#ff0000",
    "green": "#008000",
    "blue": "#0000ff",
    "yellow": "#ffff00",
    "cyan": "#00ffff",
    "magenta": "#ff00ff",
    "none": "none",
}


def _parse_float(value: Optional[str], default: float = 0.0) -> float:
    if value is None:
        return default
    value = value.strip()
    if not value:
        return default
    match = _LENGTH_RE.fullmatch(value)
    if not match:
        try:
            return float(value)
        except ValueError:
            return default
    number = float(match.group("value"))
    unit = match.group("unit").lower()
    if unit in {"", "px"}:
        return number
    if unit == "pt":
        return number * 1.3333
    if unit == "cm":
        return number * 37.7953
    if unit == "mm":
        return number * 3.77953
    if unit == "in":
        return number * 96.0
    if unit == "%":
        return number
    return number


def _parse_color(value: Optional[str]) -> Optional[str]:
    if value is None:
        return None
    value = value.strip()
    if not value:
        return None
    normalized = _COLOR_NAMES.get(value.lower(), value)
    if normalized.lower() == "none":
        return None
    if normalized.startswith("#") and len(normalized) == 7:
        return normalized.lower()
    # rgb(r,g,b)
    if normalized.startswith("rgb"):
        values = re.findall(r"\d+", normalized)
        if len(values) == 3:
            return "#" + "".join(f"{int(v):02x}" for v in values)
    return normalized


def _merge_style(element: ET.Element) -> Dict[str, str]:
    style: Dict[str, str] = {}
    if "style" in element.attrib:
        declarations = element.attrib["style"].split(";")
        for decl in declarations:
            if ":" in decl:
                key, value = decl.split(":", 1)
                style[key.strip()] = value.strip()
    return style


def _extract_style(element: ET.Element, base_style: Optional[Style] = None) -> Style:
    inline = _merge_style(element)
    fill = element.get("fill", inline.get("fill"))
    stroke = element.get("stroke", inline.get("stroke"))
    stroke_width = element.get("stroke-width", inline.get("stroke-width"))
    opacity = element.get("opacity", inline.get("opacity"))
    style = Style(
        fill_color=_parse_color(fill),
        stroke_color=_parse_color(stroke),
        stroke_width=_parse_float(stroke_width, base_style.stroke_width if base_style else 1.0),
        opacity=float(opacity) if opacity is not None else (base_style.opacity if base_style else 1.0),
    )
    return style


def _extract_text_style(element: ET.Element) -> TextStyle:
    inline = _merge_style(element)
    font_family = element.get("font-family", inline.get("font-family"))
    font_size = element.get("font-size", inline.get("font-size"))
    base_style = _extract_style(element, base_style=TextStyle())
    return TextStyle(
        fill_color=base_style.fill_color,
        stroke_color=base_style.stroke_color,
        stroke_width=base_style.stroke_width,
        opacity=base_style.opacity,
        font_family=font_family,
        font_size=_parse_float(font_size) if font_size else None,
    )


def _parse_points(points_str: str) -> List[Tuple[float, float]]:
    points = []
    for pair in re.findall(r"-?\d+(?:\.\d+)?,-?\d+(?:\.\d+)?", points_str.strip()):
        x_str, y_str = pair.split(",")
        points.append((_parse_float(x_str), _parse_float(y_str)))
    if not points:
        numbers = re.findall(r"-?\d+(?:\.\d+)?", points_str)
        for i in range(0, len(numbers), 2):
            try:
                points.append((_parse_float(numbers[i]), _parse_float(numbers[i + 1])))
            except IndexError:
                break
    return points


def _path_tokens(path_data: str) -> Iterable[str]:
    pattern = re.compile(r"([MmLlHhVvZz])|(-?\d*\.?\d+)")
    for match in pattern.finditer(path_data):
        token = match.group(0)
        if token:
            yield token


def _parse_path(path_data: str) -> Sequence[PathSegment]:
    tokens = list(_path_tokens(path_data))
    segments: List[PathSegment] = []
    cursor = (0.0, 0.0)
    start_point: Optional[Tuple[float, float]] = None
    i = 0
    command = None
    while i < len(tokens):
        token = tokens[i]
        if re.fullmatch(r"[A-Za-z]", token):
            command = token
            i += 1
            continue
        if command is None:
            raise ValueError("Path data missing command")
        if command in {"M", "L"}:
            x = _parse_float(token)
            y = _parse_float(tokens[i + 1]) if i + 1 < len(tokens) else 0.0
            point = (x, y)
            i += 2
        elif command in {"m", "l"}:
            x = cursor[0] + _parse_float(token)
            y = cursor[1] + _parse_float(tokens[i + 1]) if i + 1 < len(tokens) else cursor[1]
            point = (x, y)
            i += 2
        elif command in {"H"}:
            x = _parse_float(token)
            point = (x, cursor[1])
            i += 1
        elif command in {"h"}:
            x = cursor[0] + _parse_float(token)
            point = (x, cursor[1])
            i += 1
        elif command in {"V"}:
            y = _parse_float(token)
            point = (cursor[0], y)
            i += 1
        elif command in {"v"}:
            y = cursor[1] + _parse_float(token)
            point = (cursor[0], y)
            i += 1
        elif command in {"Z", "z"}:
            if segments and start_point is not None:
                segments[-1].closed = True
            i += 1
            cursor = start_point if start_point else cursor
            command = None
            continue
        else:
            i += 1
            continue
        if command in {"M", "m"}:
            segments.append(PathSegment(points=[point]))
            start_point = point
            command = "L" if command == "M" else "l"
        else:
            if not segments:
                segments.append(PathSegment(points=[cursor, point]))
            else:
                segments[-1].points.append(point)
        cursor = point
    return segments


def _parse_rect(element: ET.Element) -> Rect:
    style = _extract_style(element)
    return Rect(
        shape_type=ShapeType.RECT,
        style=style,
        x=_parse_float(element.get("x")),
        y=_parse_float(element.get("y")),
        width=_parse_float(element.get("width")),
        height=_parse_float(element.get("height")),
        rx=_parse_float(element.get("rx")),
        ry=_parse_float(element.get("ry")),
    )


def _parse_circle(element: ET.Element) -> Circle:
    style = _extract_style(element)
    return Circle(
        shape_type=ShapeType.CIRCLE,
        style=style,
        cx=_parse_float(element.get("cx")),
        cy=_parse_float(element.get("cy")),
        r=_parse_float(element.get("r")),
    )


def _parse_ellipse(element: ET.Element) -> Ellipse:
    style = _extract_style(element)
    return Ellipse(
        shape_type=ShapeType.ELLIPSE,
        style=style,
        cx=_parse_float(element.get("cx")),
        cy=_parse_float(element.get("cy")),
        rx=_parse_float(element.get("rx")),
        ry=_parse_float(element.get("ry")),
    )


def _parse_line(element: ET.Element) -> Line:
    style = _extract_style(element)
    return Line(
        shape_type=ShapeType.LINE,
        style=style,
        x1=_parse_float(element.get("x1")),
        y1=_parse_float(element.get("y1")),
        x2=_parse_float(element.get("x2")),
        y2=_parse_float(element.get("y2")),
    )


def _parse_polyline(element: ET.Element, closed: bool) -> Polyline:
    style = _extract_style(element)
    points = _parse_points(element.get("points", ""))
    return Polyline(shape_type=ShapeType.POLYGON if closed else ShapeType.POLYLINE, style=style, points=points, closed=closed)


def _parse_path_element(element: ET.Element) -> PathShape:
    style = _extract_style(element)
    data = element.get("d", "")
    segments = _parse_path(data)
    return PathShape(shape_type=ShapeType.PATH, style=style, segments=segments)


def _parse_text(element: ET.Element) -> Text:
    text_style = _extract_text_style(element)
    text_content = " ".join(t.strip() for t in element.itertext() if t.strip())
    x = _parse_float(element.get("x"))
    y = _parse_float(element.get("y"))
    return Text(shape_type=ShapeType.TEXT, style=text_style, text=text_content, x=x, y=y)


_SHAPE_PARSERS = {
    "rect": _parse_rect,
    "circle": _parse_circle,
    "ellipse": _parse_ellipse,
    "line": _parse_line,
    "polyline": lambda el: _parse_polyline(el, closed=False),
    "polygon": lambda el: _parse_polyline(el, closed=True),
    "path": _parse_path_element,
    "text": _parse_text,
}


def parse_svg(svg_file: PathlibPath | str) -> Document:
    """Parse *svg_file* and return a :class:`Document` instance."""

    svg_path = PathlibPath(svg_file)
    tree = ET.parse(svg_path)
    root = tree.getroot()
    width = _parse_float(root.get("width")) if root.get("width") else None
    height = _parse_float(root.get("height")) if root.get("height") else None

    shapes: List[Shape] = []
    for element in root.iter():
        tag = element.tag.split("}")[-1]
        if tag in _SHAPE_PARSERS:
            parser = _SHAPE_PARSERS[tag]
            try:
                shape = parser(element)
            except Exception as exc:  # pragma: no cover - safeguard
                raise ValueError(f"Unable to parse element '{tag}': {exc}") from exc
            shapes.append(shape)
    return Document(width=width, height=height, shapes=shapes)

