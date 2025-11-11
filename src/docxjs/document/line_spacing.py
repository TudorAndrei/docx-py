"""Line spacing module."""

from typing import TYPE_CHECKING

from docxjs.xml_parser import XmlParser

if TYPE_CHECKING:
    from lxml import etree

    from docxjs.document.common import Length


class LineSpacing:
    """Line spacing properties."""

    def __init__(self) -> None:
        """Initialize line spacing properties."""
        self.after: Length | None = None
        self.before: Length | None = None
        self.line: int | None = None
        self.line_rule: str | None = None


def parse_line_spacing(
    elem: "etree._Element", xml: XmlParser,
) -> LineSpacing:
    """Parse line spacing element."""
    result = LineSpacing()
    result.before = xml.length_attr(elem, "before")
    result.after = xml.length_attr(elem, "after")
    result.line = xml.int_attr(elem, "line")
    result.line_rule = xml.attr(elem, "lineRule")
    return result

