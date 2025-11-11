"""Border module for handling Word document borders."""

from typing import TYPE_CHECKING

from docxjs.document.common import Length, LengthUsage
from docxjs.xml_parser import XmlParser

if TYPE_CHECKING:
    from lxml import etree


class Border:
    """Border properties."""

    def __init__(self) -> None:
        """Initialize border properties."""
        self.color: str | None = None
        self.type: str | None = None
        self.size: Length | None = None
        self.frame: bool = False
        self.shadow: bool = False
        self.offset: Length | None = None


class Borders:
    """Borders collection."""

    def __init__(self) -> None:
        """Initialize borders collection."""
        self.top: Border | None = None
        self.left: Border | None = None
        self.right: Border | None = None
        self.bottom: Border | None = None


def parse_border(elem: "etree._Element", xml: XmlParser) -> Border:
    """Parse border element."""
    result = Border()
    result.type = xml.attr(elem, "val")
    result.color = xml.attr(elem, "color")
    result.size = xml.length_attr(elem, "sz", LengthUsage.Border)
    result.offset = xml.length_attr(elem, "space", LengthUsage.Point)
    result.frame = xml.bool_attr(elem, "frame", default_value=False)
    result.shadow = xml.bool_attr(
        elem, "shadow", default_value=False,
    )
    return result


def parse_borders(elem: "etree._Element", xml: XmlParser) -> Borders:
    """Parse borders element."""
    result = Borders()

    for e in xml.elements(elem):
        local_name = e.tag.split("}")[-1] if "}" in e.tag else e.tag
        if local_name == "left":
            result.left = parse_border(e, xml)
        elif local_name == "top":
            result.top = parse_border(e, xml)
        elif local_name == "right":
            result.right = parse_border(e, xml)
        elif local_name == "bottom":
            result.bottom = parse_border(e, xml)

    return result

