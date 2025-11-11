"""Document part module."""

from typing import TYPE_CHECKING

from docxjs.document_parser import DocumentParser
from docxjs.part import Part

if TYPE_CHECKING:
    from lxml import etree

    from docxjs.document.document import DocumentElement
    from docxjs.open_xml_package import OpenXmlPackage


class DocumentPart(Part):
    """Document part handler."""

    def __init__(
        self,
        pkg: "OpenXmlPackage",
        path: str,
        parser: DocumentParser,
    ) -> None:
        """Initialize document part."""
        super().__init__(pkg, path)
        self._document_parser = parser
        self.body: DocumentElement | None = None

    def parse_xml(self, root: "etree._Element") -> None:
        """Parse XML for document part."""
        self.body = self._document_parser.parse_document_file(root)

