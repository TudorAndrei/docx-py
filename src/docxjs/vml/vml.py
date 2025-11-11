"""VML parser module."""

import re
from typing import TYPE_CHECKING

from lxml import etree

from docxjs.document.common import LengthUsage, convert_length
from docxjs.document.dom import DomType, OpenXmlElement
from docxjs.xml_parser import global_xml_parser

if TYPE_CHECKING:
    from docxjs.document_parser import DocumentParser


class VmlElement(OpenXmlElement):
    """VML element."""

    def __init__(self) -> None:
        """Initialize VML element."""
        super().__init__()
        self.type = DomType.VmlElement
        self.tag_name: str = ""
        self.css_style_text: str | None = None
        self.attrs: dict[str, str] = {}
        self.wrap_type: str | None = None
        self.image_href: dict[str, str] | None = None


def parse_vml_element(  # noqa: C901, PLR0912
    elem: "etree._Element",
    parser: "DocumentParser",
) -> VmlElement | None:
    """Parse VML element from XML."""
    result = VmlElement()
    local_name = etree.QName(elem).localname

    if local_name == "rect":
        result.tag_name = "rect"
        result.attrs.update({"width": "100%", "height": "100%"})
    elif local_name == "oval":
        result.tag_name = "ellipse"
        result.attrs.update({"cx": "50%", "cy": "50%", "rx": "50%", "ry": "50%"})
    elif local_name == "line":
        result.tag_name = "line"
    elif local_name == "shape":
        result.tag_name = "g"
    elif local_name == "textbox":
        result.tag_name = "foreignObject"
        result.attrs.update({"width": "100%", "height": "100%"})
    else:
        return None

    xml = global_xml_parser
    for at_name, at_value in xml.attrs(elem):
        if at_name == "style":
            result.css_style_text = at_value
        elif at_name == "fillcolor":
            result.attrs["fill"] = at_value
        elif at_name == "from":
            x1, y1 = parse_point(at_value)
            result.attrs.update({"x1": x1, "y1": y1})
        elif at_name == "to":
            x2, y2 = parse_point(at_value)
            result.attrs.update({"x2": x2, "y2": y2})

    for el in xml.elements(elem):
        el_local_name = etree.QName(el).localname
        if el_local_name == "stroke":
            result.attrs.update(parse_stroke(el))
        elif el_local_name == "fill":
            result.attrs.update(parse_fill(el))
        elif el_local_name == "imagedata":
            result.tag_name = "image"
            result.attrs.update({"width": "100%", "height": "100%"})
            result.image_href = {
                "id": xml.attr(el, "id") or "",
                "title": xml.attr(el, "title") or "",
            }
        elif el_local_name == "txbxContent":
            result.children.extend(parser.parse_body_elements(el))
        else:
            child = parse_vml_element(el, parser)
            if child:
                result.children.append(child)

    return result


def parse_stroke(el: "etree._Element") -> dict[str, str]:
    """Parse stroke attributes from VML element."""
    xml = global_xml_parser
    result = {}
    color = xml.attr(el, "color")
    if color:
        result["stroke"] = color
    weight = xml.length_attr(el, "weight", LengthUsage.Emu)
    result["stroke-width"] = weight or "1px"
    return result


def parse_fill(_el: "etree._Element") -> dict[str, str]:
    """Parse fill attributes from VML element."""
    return {}


def parse_point(val: str) -> tuple[str, str]:
    """Parse point coordinates from string."""
    parts = val.split(",")
    return (parts[0] if len(parts) > 0 else "", parts[1] if len(parts) > 1 else "")


def convert_path(path: str) -> str:
    """Convert VML path to SVG path."""

    def replace(match: re.Match[str]) -> str:
        m = match.group(0)
        if re.match(r"[-\d]", m):
            return convert_length(m, LengthUsage.VmlEmu) or m
        if re.match(r"[ml,]", m):
            return m
        return ""

    return re.sub(r"([mlxe])|([-\d]+)|([,])", replace, path)
