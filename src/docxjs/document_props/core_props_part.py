"""Core properties part module."""

from typing import TYPE_CHECKING

from docxjs.part import Part

if TYPE_CHECKING:
    from lxml import etree


class CorePropsPart(Part):
    """Core properties part handler."""

    def parse_xml(self, root: "etree._Element") -> None:
        """Parse XML for core properties part."""

