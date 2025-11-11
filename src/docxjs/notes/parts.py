"""Notes parts module."""

from typing import TYPE_CHECKING

from docxjs.document.dom import OpenXmlElement
from docxjs.document_parser import DocumentParser
from docxjs.part import Part

if TYPE_CHECKING:
    from lxml import etree

    from docxjs.open_xml_package import OpenXmlPackage


class BaseNotePart(Part):
    """Base note part handler."""

    def __init__(
        self, pkg: "OpenXmlPackage", path: str, parser: DocumentParser,
    ) -> None:
        """Initialize note part."""
        super().__init__(pkg, path)
        self._document_parser = parser
        self.notes: list[OpenXmlElement] = []

    def parse_xml(self, root: "etree._Element") -> None:
        """Parse XML for note part."""
        elem_name = "footnote" if "footnote" in self.path else "endnote"
        self.notes = self._document_parser.parse_notes(
            root, elem_name, OpenXmlElement,
        )


class FootnotesPart(BaseNotePart):
    """Footnotes part handler."""


class EndnotesPart(BaseNotePart):
    """Endnotes part handler."""

