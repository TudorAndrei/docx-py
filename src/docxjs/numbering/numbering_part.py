"""Numbering part module."""

from typing import TYPE_CHECKING

from docxjs.document_parser import DocumentParser
from docxjs.part import Part

if TYPE_CHECKING:
    from lxml import etree

    from docxjs.document.dom import IDomNumbering
    from docxjs.open_xml_package import OpenXmlPackage


class NumberingPart(Part):
    """Numbering part handler."""

    def __init__(
        self, pkg: "OpenXmlPackage", path: str, parser: DocumentParser,
    ) -> None:
        """Initialize numbering part."""
        super().__init__(pkg, path)
        self._document_parser = parser
        self.dom_numberings: list[IDomNumbering] = []

    def parse_xml(self, root: "etree._Element") -> None:
        """Parse XML for numbering part."""
        self.dom_numberings = self._document_parser.parse_numbering_file(
            root,
        )

