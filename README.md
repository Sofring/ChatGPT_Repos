# SVG to PPTX Converter

This project provides a small Python utility that converts SVG files into PowerPoint presentations. It exposes a Python API as well as a command line interface for converting artwork into slides that can be edited with Microsoft PowerPoint or LibreOffice Impress.

## Installation

1. Create and activate a virtual environment (recommended):

   ```bash
   python -m venv .venv
   source .venv/bin/activate
   ```

2. Install the package and its dependencies:

   ```bash

   pip install -e ".[test]"
   # or, equivalently
   pip install -e '.[test]'
   ```

   > **Note:** The quotes prevent shells such as `zsh` from interpreting the square
   > brackets as a glob pattern. If you see `no matches found: .[test]`, re-run the
   > command with quotes (or escape the brackets as `pip install -e .\[test]`).


## Command Line Usage

The `svg2pptx` command accepts an input SVG file and an output PPTX path:

```bash
svg2pptx input.svg output.pptx
```

The generated presentation contains a single blank slide sized according to the SVG's width and height (when present) with shapes positioned using the SVG coordinate system. Basic support is provided for rectangles, circles, ellipses, lines, polylines, polygons, simple paths (M/L/H/V commands), and text.

## Python API

```python
from svg_to_pptx import parse_svg, build_presentation

document = parse_svg("diagram.svg")
build_presentation(document, "diagram.pptx")
```

## Tests

Run the integration tests with `pytest`:

```bash
pytest
```

## Known Limitations

- Only a subset of SVG path commands are supported (M, L, H, V). Curves are not yet implemented.
- Transformations (e.g., `transform="translate(...)"`) are ignored.
- Text boxes are sized with a fixed width/height; complex text layout is not preserved.
