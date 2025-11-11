"""Paragraph module for handling Word document paragraphs."""

from typing import TYPE_CHECKING

from lxml import etree

from docxjs.document.common import CommonProperties, ns, parse_common_property
from docxjs.document.dom import DomType, OpenXmlElement
from docxjs.document.line_spacing import LineSpacing, parse_line_spacing
from docxjs.document.run import RunProperties, parse_run_properties
from docxjs.document.section import SectionProperties, parse_section_properties
from docxjs.xml_parser import XmlParser

if TYPE_CHECKING:
    from docxjs.document.border import Borders


class ParagraphTab:
    """Paragraph tab properties."""

    def __init__(self) -> None:
        """Initialize paragraph tab properties."""
        self.style: str | None = None
        self.leader: str | None = None
        self.position: str | None = None


class ParagraphNumbering:
    """Paragraph numbering properties."""

    def __init__(self) -> None:
        """Initialize paragraph numbering properties."""
        self.id: str | None = None
        self.level: int | None = None


class ParagraphProperties(CommonProperties):
    """Paragraph properties."""

    def __init__(self) -> None:
        """Initialize paragraph properties."""
        super().__init__()
        self.section_props: SectionProperties | None = None
        self.tabs: list[ParagraphTab] | None = None
        self.numbering: ParagraphNumbering | None = None
        self.border: Borders | None = None
        self.text_alignment: str | None = None
        self.line_spacing: LineSpacing | None = None
        self.keep_lines: bool = False
        self.keep_next: bool = False
        self.page_break_before: bool = False
        self.outline_level: int | None = None
        self.style_name: str | None = None
        self.run_props: RunProperties | None = None


class WmlParagraph(OpenXmlElement, ParagraphProperties):
    """Word ML paragraph element."""

    def __init__(self) -> None:
        """Initialize paragraph element."""
        OpenXmlElement.__init__(self)
        ParagraphProperties.__init__(self)
        self.type = DomType.Paragraph


def parse_paragraph_properties(
    elem: "etree._Element",
    xml: XmlParser,
) -> ParagraphProperties:
    """Parse paragraph properties element."""
    result = ParagraphProperties()

    for el in xml.elements(elem):
        parse_paragraph_property(el, result, xml)

    return result


def parse_paragraph_property(  # noqa: C901, PLR0911, PLR0912
    elem: "etree._Element",
    props: ParagraphProperties,
    xml: XmlParser,
) -> bool:
    """Parse paragraph property element."""
    namespace_uri = elem.nsmap.get(elem.prefix) if elem.prefix else None
    if namespace_uri != ns["wordml"]:
        return False

    if parse_common_property(elem, props, xml):
        return True

    local_name = etree.QName(elem).localname

    if local_name == "tabs":
        props.tabs = parse_tabs(elem, xml)
        return True
    if local_name == "sectPr":
        props.section_props = parse_section_properties(elem, xml)
        return True
    if local_name == "numPr":
        props.numbering = parse_numbering(elem, xml)
        return True
    if local_name == "spacing":
        props.line_spacing = parse_line_spacing(elem, xml)
        return False
    if local_name == "textAlignment":
        props.text_alignment = xml.attr(elem, "val")
        return False
    if local_name == "keepLines":
        props.keep_lines = xml.bool_attr(elem, "val", default_value=True)
        return True
    if local_name == "keepNext":
        props.keep_next = xml.bool_attr(elem, "val", default_value=True)
        return True
    if local_name == "pageBreakBefore":
        props.page_break_before = xml.bool_attr(
            elem,
            "val",
            default_value=True,
        )
        return True
    if local_name == "outlineLvl":
        props.outline_level = xml.int_attr(elem, "val")
        return True
    if local_name == "pStyle":
        props.style_name = xml.attr(elem, "val")
        return True
    if local_name == "rPr":
        props.run_props = parse_run_properties(elem, xml)
        return True

    return False


def parse_tabs(
    elem: "etree._Element",
    xml: XmlParser,
) -> list[ParagraphTab]:
    """Parse tabs element."""
    result = []
    for e in xml.elements(elem, "tab"):
        tab = ParagraphTab()
        tab.position = xml.length_attr(e, "pos")
        tab.leader = xml.attr(e, "leader")
        tab.style = xml.attr(e, "val")
        result.append(tab)
    return result


def parse_numbering(
    elem: "etree._Element",
    xml: XmlParser,
) -> ParagraphNumbering:
    """Parse numbering element."""
    result = ParagraphNumbering()

    for e in xml.elements(elem):
        local_name = e.tag.split("}")[-1] if "}" in e.tag else e.tag
        if local_name == "numId":
            result.id = xml.attr(e, "val")
        elif local_name == "ilvl":
            result.level = xml.int_attr(e, "val")

    return result
