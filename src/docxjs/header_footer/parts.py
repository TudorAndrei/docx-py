"""Header and footer parts module."""

from typing import TYPE_CHECKING

from docxjs.document.dom import DomType, OpenXmlElement
from docxjs.document_parser import DocumentParser
from docxjs.part import Part

if TYPE_CHECKING:
    from lxml import etree

    from docxjs.open_xml_package import OpenXmlPackage


class BaseHeaderFooterPart(Part):
    """Base header/footer part handler."""

    def __init__(
        self, pkg: "OpenXmlPackage", path: str, parser: DocumentParser,
    ) -> None:
        """Initialize header/footer part."""
        super().__init__(pkg, path)
        self._document_parser = parser
        self.root_element: OpenXmlElement | None = None

    def parse_xml(self, root: "etree._Element") -> None:
        """Parse XML for header/footer part."""
        elements = self._document_parser.parse_body_elements(root)
        self.root_element = OpenXmlElement()
        self.root_element.type = (
            DomType.Header if "header" in self.path else DomType.Footer
        )
        self.root_element.children = elements


class HeaderPart(BaseHeaderFooterPart):
    """Header part handler."""


class FooterPart(BaseHeaderFooterPart):
    """Footer part handler."""

