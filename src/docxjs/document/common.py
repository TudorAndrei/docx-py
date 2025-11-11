"""Common document utilities and types."""

import re
from typing import TYPE_CHECKING

from lxml import etree
from docxjs.utils import clamp

if TYPE_CHECKING:
    from docxjs.xml_parser import XmlParser

ns = {
    "wordml": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "drawingml": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "picture": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "compatibility": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "math": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}

LengthType = str
Length = str


class Font:
    """Font properties."""

    name: str
    family: str


class CommonProperties:
    """Common document properties."""

    font_size: Length | None = None
    color: str | None = None


class LengthUsageType:
    """Length usage type configuration."""

    def __init__(
        self,
        mul: float,
        unit: LengthType,
        min_val: float | None = None,
        max_val: float | None = None,
    ) -> None:
        """Initialize length usage type."""
        self.mul = mul
        self.unit = unit
        self.min = min_val
        self.max = max_val


class LengthUsage:
    """Length usage constants."""

    Dxa = LengthUsageType(0.05, "pt")
    Emu = LengthUsageType(1 / 12700, "pt")
    FontSize = LengthUsageType(0.5, "pt")
    Border = LengthUsageType(0.125, "pt", 0.25, 12)
    Point = LengthUsageType(1, "pt")
    Percent = LengthUsageType(0.02, "%")
    LineHeight = LengthUsageType(1 / 240, "")
    VmlEmu = LengthUsageType(1 / 12700, "")


def convert_length(
    val: str | None,
    usage: LengthUsageType = LengthUsage.Dxa,
) -> str | None:
    """Convert length value to formatted string."""
    if val is None:
        return None

    if re.match(r".+(p[xt]|[%])$", val):
        return val

    try:
        num = int(val) * usage.mul

        if usage.min is not None and usage.max is not None:
            num = clamp(num, usage.min, usage.max)
    except (ValueError, TypeError):
        return val
    else:
        return f"{num:.2f}{usage.unit}"


def convert_boolean(
    v: str | None,
    *,
    default_value: bool = False,
) -> bool:
    """Convert string to boolean value."""
    if v is None:
        return default_value

    v_lower = v.lower()
    if v_lower in ("1", "on", "true"):
        return True
    if v_lower in ("0", "off", "false"):
        return False
    return default_value


def convert_percentage(val: str | None) -> float | None:
    """Convert percentage string to float."""
    if not val:
        return None
    try:
        return int(val) / 100
    except (ValueError, TypeError):
        return None


def parse_common_property(
    elem: "etree._Element",
    props: CommonProperties,
    xml: "XmlParser",
) -> bool:
    """Parse common property from XML element."""
    if not isinstance(elem, etree.Element):
        return False

    namespace_uri = elem.nsmap.get(elem.prefix) if elem.prefix else None
    if namespace_uri != ns["wordml"]:
        return False

    local_name = etree.QName(elem).localname

    if local_name == "color":
        props.color = xml.attr(elem, "val")
        return True
    if local_name == "sz":
        props.font_size = xml.length_attr(elem, "val", LengthUsage.FontSize)
        return True

    return False

