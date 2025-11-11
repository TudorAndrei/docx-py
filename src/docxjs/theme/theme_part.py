"""Theme part module."""

from typing import TYPE_CHECKING, Any

from lxml import etree

from docxjs.part import Part

if TYPE_CHECKING:
    from docxjs.open_xml_package import OpenXmlPackage


class DmlTheme:
    """Theme data structure."""

    def __init__(self) -> None:
        """Initialize theme."""
        self.color_scheme: dict[str, Any] | None = None
        self.font_scheme: dict[str, Any] | None = None


class ThemePart(Part):
    """Theme part handler."""

    def __init__(self, pkg: "OpenXmlPackage", path: str) -> None:
        """Initialize theme part."""
        super().__init__(pkg, path)
        self.theme: DmlTheme | None = None

    def parse_xml(self, root: etree._Element) -> None:
        """Parse XML for theme part."""
        result = DmlTheme()
        xml = self._package.xml_parser
        theme_elements = xml.element(root, "themeElements")

        if theme_elements is not None:
            for el in xml.elements(theme_elements):
                qname = etree.QName(el)
                if qname.localname == "clrScheme":
                    result.color_scheme = self._parse_color_scheme(el, xml)
                elif qname.localname == "fontScheme":
                    result.font_scheme = self._parse_font_scheme(el, xml)

        self.theme = result

    def _parse_color_scheme(
        self, elem: etree._Element, xml: Any,
    ) -> dict[str, Any]:
        """Parse color scheme."""
        result: dict[str, Any] = {
            "name": xml.attr(elem, "name"),
            "colors": {},
        }

        for el in xml.elements(elem):
            qname = etree.QName(el)
            srgb_clr = xml.element(el, "srgbClr")
            sys_clr = xml.element(el, "sysClr")

            if srgb_clr is not None:
                result["colors"][qname.localname] = xml.attr(srgb_clr, "val")
            elif sys_clr is not None:
                result["colors"][qname.localname] = xml.attr(sys_clr, "lastClr")

        return result

    def _parse_font_scheme(
        self, elem: etree._Element, xml: Any,
    ) -> dict[str, Any]:
        """Parse font scheme."""
        result: dict[str, Any] = {
            "name": xml.attr(elem, "name"),
        }

        for el in xml.elements(elem):
            qname = etree.QName(el)
            if qname.localname == "majorFont":
                result["major_font"] = self._parse_font_info(el, xml)
            elif qname.localname == "minorFont":
                result["minor_font"] = self._parse_font_info(el, xml)

        return result

    def _parse_font_info(
        self, elem: etree._Element, xml: Any,
    ) -> dict[str, str]:
        """Parse font info."""
        return {
            "latin_typeface": xml.element_attr(elem, "latin", "typeface") or "",
            "ea_typeface": xml.element_attr(elem, "ea", "typeface") or "",
            "cs_typeface": xml.element_attr(elem, "cs", "typeface") or "",
        }

