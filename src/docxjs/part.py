"""Part module for handling document parts."""

from typing import TYPE_CHECKING

from lxml import etree

from docxjs.open_xml_package import OpenXmlPackage
from docxjs.xml_parser import serialize_xml_string

if TYPE_CHECKING:
    from docxjs.relationship import Relationship


class Part:
    """Base part handler."""

    def __init__(self, package: OpenXmlPackage, path: str) -> None:
        """Initialize part."""
        self._package = package
        self.path = path
        self._xml_document: etree._Element | None = None
        self.rels: list[Relationship] | None = None

    def load(self) -> None:
        """Load part from package."""
        self.rels = self._package.load_relationships(self.path)

        xml_text = self._package.load_file(self.path)
        if xml_text is None:
            return

        xml_doc = self._package.parse_xml_document(xml_text)

        if self._package.options.keep_origin:
            self._xml_document = xml_doc

        self.parse_xml(xml_doc)

    def save(self) -> None:
        """Save part to package."""
        if self._xml_document is not None:
            content = serialize_xml_string(self._xml_document)
            self._package.update(self.path, content)

    def parse_xml(self, root: etree._Element) -> None:
        """Parse XML for part."""

