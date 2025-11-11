"""Bookmarks module for handling Word document bookmarks."""

from typing import TYPE_CHECKING

from docxjs.document.dom import DomType, OpenXmlElement
from docxjs.xml_parser import XmlParser

if TYPE_CHECKING:
    from lxml import etree


class WmlBookmarkStart(OpenXmlElement):
    """Word ML bookmark start element."""

    def __init__(self) -> None:
        """Initialize bookmark start element."""
        super().__init__()
        self.type = DomType.BookmarkStart
        self.id: str = ""
        self.name: str = ""
        self.col_first: int | None = None
        self.col_last: int | None = None


class WmlBookmarkEnd(OpenXmlElement):
    """Word ML bookmark end element."""

    def __init__(self) -> None:
        """Initialize bookmark end element."""
        super().__init__()
        self.type = DomType.BookmarkEnd
        self.id: str = ""


def parse_bookmark_start(elem: "etree._Element", xml: XmlParser) -> WmlBookmarkStart:
    """Parse bookmark start element."""
    result = WmlBookmarkStart()
    result.id = xml.attr(elem, "id") or ""
    result.name = xml.attr(elem, "name") or ""
    result.col_first = xml.int_attr(elem, "colFirst")
    result.col_last = xml.int_attr(elem, "colLast")
    return result


def parse_bookmark_end(elem: "etree._Element", xml: XmlParser) -> WmlBookmarkEnd:
    """Parse bookmark end element."""
    result = WmlBookmarkEnd()
    result.id = xml.attr(elem, "id") or ""
    return result

