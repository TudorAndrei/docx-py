"""Comments extended part module."""

from typing import TYPE_CHECKING

from docxjs.part import Part

if TYPE_CHECKING:
    from lxml import etree


class CommentsExtendedPart(Part):
    """Extended comments part handler."""

    def parse_xml(self, root: "etree._Element") -> None:
        """Parse XML for extended comments."""

