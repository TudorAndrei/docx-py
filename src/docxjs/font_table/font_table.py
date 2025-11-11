"""Font table part module."""

from typing import TYPE_CHECKING

from docxjs.part import Part

if TYPE_CHECKING:
    from lxml import etree

    from docxjs.open_xml_package import OpenXmlPackage


class FontDeclaration:
    """Font declaration."""

    def __init__(self) -> None:
        """Initialize font declaration."""
        self.name: str = ""
        self.embed_font_refs: list[dict] = []


class FontTablePart(Part):
    """Font table part handler."""

    def __init__(self, pkg: "OpenXmlPackage", path: str) -> None:
        """Initialize font table part."""
        super().__init__(pkg, path)
        self.fonts: list[FontDeclaration] = []

    def parse_xml(self, root: "etree._Element") -> None:
        """Parse XML for font table part."""
        _ = root  # Currently unused but required by interface
        self.fonts = []

