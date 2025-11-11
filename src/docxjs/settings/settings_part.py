"""Settings part module."""

from typing import TYPE_CHECKING

from docxjs.part import Part
from docxjs.xml_parser import global_xml_parser

if TYPE_CHECKING:
    from lxml import etree

    from docxjs.open_xml_package import OpenXmlPackage


class SettingsPart(Part):
    """Settings part handler."""

    def __init__(self, pkg: "OpenXmlPackage", path: str) -> None:
        """Initialize settings part."""
        super().__init__(pkg, path)
        self.settings: dict[str, str] = {}

    def parse_xml(self, root: "etree._Element") -> None:
        """Parse XML for settings part."""
        default_tab_stop = global_xml_parser.element_attr(
            root, "defaultTabStop", "val",
        )
        if default_tab_stop:
            self.settings = {"defaultTabStop": default_tab_stop}

