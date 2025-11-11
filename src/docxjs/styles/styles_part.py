"""Styles part module."""

from typing import TYPE_CHECKING

from docxjs.document_parser import DocumentParser
from docxjs.part import Part

if TYPE_CHECKING:
    from lxml import etree

    from docxjs.document.style import IDomStyle
    from docxjs.open_xml_package import OpenXmlPackage


class StylesPart(Part):
    """Styles part handler."""

    def __init__(
        self, pkg: "OpenXmlPackage", path: str, parser: DocumentParser,
    ) -> None:
        """Initialize styles part."""
        super().__init__(pkg, path)
        self._document_parser = parser
        self.styles: list[IDomStyle] = []

    def parse_xml(self, root: "etree._Element") -> None:
        """Parse XML for styles part."""
        self.styles = self._document_parser.parse_styles_file(root)

