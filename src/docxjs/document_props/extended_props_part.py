"""Extended properties part module."""

from typing import TYPE_CHECKING

from docxjs.part import Part
from docxjs.xml_parser import global_xml_parser

if TYPE_CHECKING:
    from lxml import etree


class ExtendedPropsPart(Part):
    """Extended properties part handler."""

    def __init__(self, package, path: str) -> None:
        """Initialize extended properties part."""
        super().__init__(package, path)
        self.appVersion: str = ""
        self.application: str = ""
        self.characters: int = 0
        self.company: str = ""
        self.lines: int = 0
        self.pages: int = 0
        self.paragraphs: int = 0
        self.template: str = ""
        self.words: int = 0

    def parse_xml(self, root: "etree._Element") -> None:
        """Parse XML for extended properties part."""
        xml = global_xml_parser

        app_elem = xml.element(root, "Application")
        if app_elem is not None and app_elem.text:
            self.application = app_elem.text

        version_elem = xml.element(root, "AppVersion")
        if version_elem is not None and version_elem.text:
            self.appVersion = version_elem.text

        chars_elem = xml.element(root, "Characters")
        if chars_elem is not None and chars_elem.text:
            try:
                self.characters = int(chars_elem.text)
            except ValueError:
                pass

        company_elem = xml.element(root, "Company")
        if company_elem is not None and company_elem.text:
            self.company = company_elem.text

        lines_elem = xml.element(root, "Lines")
        if lines_elem is not None and lines_elem.text:
            try:
                self.lines = int(lines_elem.text)
            except ValueError:
                pass

        pages_elem = xml.element(root, "Pages")
        if pages_elem is not None and pages_elem.text:
            try:
                self.pages = int(pages_elem.text)
            except ValueError:
                pass

        paragraphs_elem = xml.element(root, "Paragraphs")
        if paragraphs_elem is not None and paragraphs_elem.text:
            try:
                self.paragraphs = int(paragraphs_elem.text)
            except ValueError:
                pass

        template_elem = xml.element(root, "Template")
        if template_elem is not None and template_elem.text:
            self.template = template_elem.text

        words_elem = xml.element(root, "Words")
        if words_elem is not None and words_elem.text:
            try:
                self.words = int(words_elem.text)
            except ValueError:
                pass

