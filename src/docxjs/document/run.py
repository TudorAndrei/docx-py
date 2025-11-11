"""Run module for handling Word document runs."""

from typing import TYPE_CHECKING

from docxjs.document.common import CommonProperties, parse_common_property
from docxjs.document.dom import DomType, OpenXmlElement
from docxjs.xml_parser import XmlParser

if TYPE_CHECKING:
    from lxml import etree


class RunProperties(CommonProperties):
    """Run properties."""


class WmlRun(OpenXmlElement):
    """Word ML run element."""

    def __init__(self) -> None:
        """Initialize run element."""
        super().__init__()
        self.type = DomType.Run
        self.id: str | None = None
        self.vertical_align: str | None = None
        self.field_run: bool = False


def parse_run_properties(
    elem: "etree._Element", xml: XmlParser,
) -> RunProperties:
    """Parse run properties element."""
    result = RunProperties()

    for el in xml.elements(elem):
        parse_run_property(el, result, xml)

    return result


def parse_run_property(
    elem: "etree._Element",
    props: RunProperties,
    xml: XmlParser,
) -> bool:
    """Parse run property element."""
    return bool(parse_common_property(elem, props, xml))

