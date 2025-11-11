"""Comments part module."""

from typing import TYPE_CHECKING

from docxjs.utils import key_by

from docxjs.document_parser import DocumentParser
from docxjs.part import Part

if TYPE_CHECKING:
    from lxml import etree

    from docxjs.open_xml_package import OpenXmlPackage


class CommentsPart(Part):
    """Comments part handler."""

    def __init__(
        self,
        pkg: "OpenXmlPackage",
        path: str,
        parser: DocumentParser,
    ) -> None:
        """Initialize comments part."""
        super().__init__(pkg, path)
        self._document_parser = parser
        self.comments = []
        self.comment_map = {}

    def parse_xml(self, root: "etree._Element") -> None:
        """Parse XML for comments."""
        self.comments = self._document_parser.parse_comments(root)
        self.comment_map = key_by(self.comments, lambda x: x.id)

