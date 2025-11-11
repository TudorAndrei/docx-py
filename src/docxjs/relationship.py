"""Relationship module for handling document relationships."""

from enum import Enum

from lxml import etree

from docxjs.xml_parser import XmlParser


class RelationshipTypes(Enum):
    """Relationship types enumeration."""

    OfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    FontTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
    Image = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Numbering = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    Styles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    StylesWithEffects = "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects"
    Theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    Settings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    WebSettings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
    Hyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Footnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
    Endnotes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
    Footer = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    Header = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    ExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
    CoreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
    CustomProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/custom-properties"
    Comments = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    CommentsExtended = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
    AltChunk = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk"


class Relationship:
    """Relationship between document parts."""

    def __init__(
        self, id_val: str, rel_type: str, target: str, target_mode: str = "",
    ) -> None:
        """Initialize relationship."""
        self.id = id_val
        self.type = rel_type
        self.target = target
        self.target_mode = target_mode


def parse_relationships(
    root: etree._Element, xml: XmlParser,
) -> list[Relationship]:
    """Parse relationships from XML."""
    result = []
    for e in xml.elements(root):
        rel = Relationship(
            xml.attr(e, "Id"),
            xml.attr(e, "Type"),
            xml.attr(e, "Target"),
            xml.attr(e, "TargetMode") or "",
        )
        result.append(rel)
    return result

