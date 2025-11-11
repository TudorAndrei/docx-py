# DOCX.js Python Port

> ⚠️ **WARNING**: This project is `Vibe Coded`

A Python port of the docx.js library for rendering DOCX files to HTML using BeautifulSoup4 and lxml.

## Installation

```bash
pip install -e .
```

Or using uv:
```bash
uv sync
```

## Usage

### Basic Usage

```python
from docx_preview import render_async
from bs4 import BeautifulSoup

with open('document.docx', 'rb') as f:
    docx_data = f.read()

soup = BeautifulSoup('<html><body></body></html>', 'html.parser')
body = soup.body

doc = render_async(docx_data, body, body)
print(soup.prettify())
```

### Streamlit App

Run the Streamlit app for an interactive interface:

```bash
streamlit run streamlit_app.py
```

This will open a web interface where you can:
- Upload DOCX files
- View the rendered HTML output
- Download the rendered HTML
- View document metadata

## API

### `render_async(data: bytes, body_container, style_container=None, user_options=None)`

Renders a DOCX file to HTML.

**Parameters:**
- `data`: DOCX file as bytes
- `body_container`: BeautifulSoup Tag element for the body
- `style_container`: BeautifulSoup Tag element for styles (defaults to body_container)
- `user_options`: Dictionary of options (see below)

**Returns:** WordDocument object

### Options

- `in_wrapper`: Wrap document in a container div (default: True)
- `class_name`: CSS class prefix (default: "docx")
- `ignore_fonts`: Don't render embedded fonts (default: False)
- `render_headers`: Render document headers (default: True)
- `render_footers`: Render document footers (default: True)
- `render_footnotes`: Render footnotes (default: True)
- `render_endnotes`: Render endnotes (default: True)
- `ignore_width`: Ignore width constraints (default: False)
- `ignore_height`: Ignore height constraints (default: False)
- `break_pages`: Enable page breaking (default: True)
- `debug`: Enable debug logging (default: False)

## Features

- ✅ Paragraph and text rendering
- ✅ Tables with borders and styling
- ✅ Images and drawings
- ✅ Headers and footers
- ✅ Footnotes and endnotes
- ✅ Numbering and lists
- ✅ Styles and formatting
- ✅ Hyperlinks
- ✅ Math equations (MathML)
- ✅ VML graphics
- ✅ Comments (experimental)

## Dependencies

- `beautifulsoup4` - HTML generation
- `lxml` - XML parsing
- `streamlit` - Web interface (optional, for streamlit_app.py)

## License

Same as the original docx.js library (Apache-2.0)

