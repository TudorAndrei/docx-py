"""Section module for handling Word document sections."""

from typing import TYPE_CHECKING

from lxml import etree

from docxjs.document.border import Borders, parse_borders
from docxjs.xml_parser import XmlParser

if TYPE_CHECKING:
    from docxjs.document.common import Length


class Column:
    """Column properties."""

    def __init__(self) -> None:
        """Initialize column properties."""
        self.space: Length | None = None
        self.width: Length | None = None


class Columns:
    """Columns properties."""

    def __init__(self) -> None:
        """Initialize columns properties."""
        self.space: Length | None = None
        self.number_of_columns: int | None = None
        self.separator: bool = False
        self.equal_width: bool = True
        self.columns: list[Column] = []


class PageSize:
    """Page size properties."""

    def __init__(self) -> None:
        """Initialize page size properties."""
        self.width: Length | None = None
        self.height: Length | None = None
        self.orientation: str | None = None


class PageNumber:
    """Page number properties."""

    def __init__(self) -> None:
        """Initialize page number properties."""
        self.start: int | None = None
        self.chap_sep: str | None = None
        self.chap_style: str | None = None
        self.format: str | None = None


class PageMargins:
    """Page margins properties."""

    def __init__(self) -> None:
        """Initialize page margins properties."""
        self.top: Length | None = None
        self.right: Length | None = None
        self.bottom: Length | None = None
        self.left: Length | None = None
        self.header: Length | None = None
        self.footer: Length | None = None
        self.gutter: Length | None = None


class FooterHeaderReference:
    """Footer/header reference."""

    def __init__(self) -> None:
        """Initialize footer/header reference."""
        self.id: str = ""
        self.type: str = "default"


class SectionProperties:
    """Section properties."""

    def __init__(self) -> None:
        """Initialize section properties."""
        self.type: str | None = None
        self.page_size: PageSize | None = None
        self.page_margins: PageMargins | None = None
        self.page_borders: Borders | None = None
        self.page_number: PageNumber | None = None
        self.columns: Columns | None = None
        self.footer_refs: list[FooterHeaderReference] = []
        self.header_refs: list[FooterHeaderReference] = []
        self.title_page: bool = False


def parse_section_properties(  # noqa: C901
    elem: "etree._Element",
    xml: XmlParser,
) -> SectionProperties:
    """Parse section properties element."""
    section = SectionProperties()

    for e in xml.elements(elem):
        local_name = etree.QName(e).localname

        if local_name == "pgSz":
            section.page_size = PageSize()
            section.page_size.width = xml.length_attr(e, "w")
            section.page_size.height = xml.length_attr(e, "h")
            section.page_size.orientation = xml.attr(e, "orient")
        elif local_name == "type":
            section.type = xml.attr(e, "val")
        elif local_name == "pgMar":
            section.page_margins = PageMargins()
            section.page_margins.left = xml.length_attr(e, "left")
            section.page_margins.right = xml.length_attr(e, "right")
            section.page_margins.top = xml.length_attr(e, "top")
            section.page_margins.bottom = xml.length_attr(e, "bottom")
            section.page_margins.header = xml.length_attr(e, "header")
            section.page_margins.footer = xml.length_attr(e, "footer")
            section.page_margins.gutter = xml.length_attr(e, "gutter")
        elif local_name == "cols":
            section.columns = parse_columns(e, xml)
        elif local_name == "headerReference":
            if section.header_refs is None:
                section.header_refs = []
            section.header_refs.append(parse_footer_header_reference(e, xml))
        elif local_name == "footerReference":
            if section.footer_refs is None:
                section.footer_refs = []
            section.footer_refs.append(parse_footer_header_reference(e, xml))
        elif local_name == "titlePg":
            section.title_page = xml.bool_attr(e, "val", default_value=True)
        elif local_name == "pgBorders":
            section.page_borders = parse_borders(e, xml)
        elif local_name == "pgNumType":
            section.page_number = parse_page_number(e, xml)

    return section


def parse_columns(
    elem: "etree._Element",
    xml: XmlParser,
) -> Columns:
    """Parse columns element."""
    result = Columns()
    result.number_of_columns = xml.int_attr(elem, "num")
    result.space = xml.length_attr(elem, "space")
    result.separator = xml.bool_attr(elem, "sep", default_value=False)
    result.equal_width = xml.bool_attr(
        elem,
        "equalWidth",
        default_value=True,
    )

    for e in xml.elements(elem, "col"):
        col = Column()
        col.width = xml.length_attr(e, "w")
        col.space = xml.length_attr(e, "space")
        result.columns.append(col)

    return result


def parse_page_number(
    elem: "etree._Element",
    xml: XmlParser,
) -> PageNumber:
    """Parse page number element."""
    result = PageNumber()
    result.chap_sep = xml.attr(elem, "chapSep")
    result.chap_style = xml.attr(elem, "chapStyle")
    result.format = xml.attr(elem, "fmt")
    result.start = xml.int_attr(elem, "start")
    return result


def parse_footer_header_reference(
    elem: "etree._Element",
    xml: XmlParser,
) -> FooterHeaderReference:
    """Parse footer/header reference element."""
    result = FooterHeaderReference()
    result.id = xml.attr(elem, "id") or ""
    result.type = xml.attr(elem, "type") or "default"
    return result
