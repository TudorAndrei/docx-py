"""Docx preview module for parsing and rendering Word documents."""

from typing import TYPE_CHECKING, Any

from bs4 import BeautifulSoup

if TYPE_CHECKING:
    from bs4 import Tag

from docxjs.document_parser import DocumentParser, DocumentParserOptions
from docxjs.html_renderer import HtmlRenderer
from docxjs.word_document import WordDocument


class Options:
    """Options for document parsing and rendering."""

    def __init__(self) -> None:
        """Initialize options with default values."""
        self.in_wrapper: bool = True
        self.hide_wrapper_on_print: bool = False
        self.ignore_width: bool = False
        self.ignore_height: bool = False
        self.ignore_fonts: bool = False
        self.break_pages: bool = True
        self.debug: bool = False
        self.experimental: bool = False
        self.class_name: str = "docx"
        self.trim_xml_declaration: bool = True
        self.render_headers: bool = True
        self.render_footers: bool = True
        self.render_footnotes: bool = True
        self.render_endnotes: bool = True
        self.ignore_last_rendered_page_break: bool = True
        self.use_base64_url: bool = False
        self.render_changes: bool = False
        self.render_comments: bool = False
        self.render_alt_chunks: bool = True


default_options = Options()


def parse_async(
    data: bytes,
    user_options: dict[str, Any] | None = None,
) -> WordDocument:
    """Parse Word document asynchronously."""
    options_dict = {**vars(default_options)}
    if user_options:
        options_dict.update(user_options)

    parser_options = DocumentParserOptions(
        ignore_width=options_dict.get("ignore_width", False),
        debug=options_dict.get("debug", False),
    )
    parser = DocumentParser(parser_options)

    return WordDocument.load(data, parser, options_dict)


def render_document(
    document: WordDocument,
    body_container: "Tag",
    style_container: "Tag | None" = None,
    user_options: dict[str, Any] | None = None,
) -> WordDocument:
    """Render document to HTML."""
    options_dict = {**vars(default_options)}
    if user_options:
        options_dict.update(user_options)

    soup = BeautifulSoup("<html></html>", "html.parser")
    renderer = HtmlRenderer(soup)
    renderer.render(document, body_container, style_container, options_dict)
    return document


def render_async(
    data: bytes,
    body_container: "Tag",
    style_container: "Tag | None" = None,
    user_options: dict[str, Any] | None = None,
) -> WordDocument:
    """Render Word document asynchronously."""
    doc = parse_async(data, user_options)
    render_document(doc, body_container, style_container, user_options)
    return doc
