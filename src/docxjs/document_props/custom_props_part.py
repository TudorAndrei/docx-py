"""Custom properties part module."""

from typing import TYPE_CHECKING

from docxjs.part import Part

if TYPE_CHECKING:
    from lxml import etree


class CustomPropsPart(Part):
    """Custom properties part handler."""

    def parse_xml(self, root: "etree._Element") -> None:
        """Parse XML for custom properties part."""

