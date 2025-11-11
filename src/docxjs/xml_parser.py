"""XML parser utilities."""
import re

from lxml import etree

from docxjs.document.common import (
    Length,
    LengthUsage,
    LengthUsageType,
    convert_boolean,
    convert_length,
)

UTF8_BOM = 0xFEFF


def parse_xml_string(
    xml_string: str, trim_xml_declaration: bool = False,  # noqa: FBT001, FBT002
) -> etree._ElementTree:
    """Parse XML string into element tree."""
    if trim_xml_declaration:
        xml_string = re.sub(r"<\?.*?\?>", "", xml_string, count=1)

    xml_string = remove_utf8_bom(xml_string)

    try:
        parser = etree.XMLParser(recover=True)
        return etree.fromstring(xml_string.encode("utf-8"), parser=parser)
    except etree.XMLSyntaxError as e:
        msg = f"XML parsing error: {e}"
        raise ValueError(msg) from e


def remove_utf8_bom(data: str) -> str:
    """Remove UTF-8 BOM from string."""
    if len(data) > 0 and ord(data[0]) == UTF8_BOM:
        return data[1:]
    return data


def serialize_xml_string(elem: etree._Element) -> str:
    """Serialize XML element to string."""
    return etree.tostring(elem, encoding="unicode", method="xml")


class XmlParser:
    """XML parser utility class."""

    def elements(
        self, elem: etree._Element, local_name: str | None = None,
    ) -> list[etree._Element]:
        """Get child elements matching local name."""
        result = []

        for child in elem:
            if isinstance(child, etree._Element):  # noqa: SLF001
                qname = etree.QName(child)
                if local_name is None or qname.localname == local_name:
                    result.append(child)

        return result

    def element(
        self, elem: etree._Element, local_name: str,
    ) -> etree._Element | None:
        """Get first child element matching local name."""
        for child in elem:
            if isinstance(child, etree._Element):  # noqa: SLF001
                qname = etree.QName(child)
                if qname.localname == local_name:
                    return child
        return None

    def element_attr(
        self,
        elem: etree._Element,
        local_name: str,
        attr_local_name: str,
    ) -> str | None:
        """Get attribute value from child element."""
        el = self.element(elem, local_name)
        return self.attr(el, attr_local_name) if el is not None else None

    def attrs(self, elem: etree._Element) -> list[tuple[str, str]]:
        """Get all attributes as list of tuples."""
        return list(elem.attrib.items())

    def attr(self, elem: etree._Element | None, local_name: str) -> str | None:
        """Get attribute value by local name."""
        if elem is None:
            return None

        qname = etree.QName(local_name)
        return elem.get(qname.localname) or elem.get(local_name)

    def int_attr(
        self,
        node: etree._Element,
        attr_name: str,
        default_value: int | None = None,
    ) -> int | None:
        """Get integer attribute value."""
        val = self.attr(node, attr_name)
        if val:
            try:
                return int(val)
            except ValueError:
                return default_value
        return default_value

    def hex_attr(
        self,
        node: etree._Element,
        attr_name: str,
        default_value: int | None = None,
    ) -> int | None:
        """Get hexadecimal attribute value."""
        val = self.attr(node, attr_name)
        if val:
            try:
                return int(val, 16)
            except ValueError:
                return default_value
        return default_value

    def float_attr(
        self,
        node: etree._Element,
        attr_name: str,
        default_value: float | None = None,
    ) -> float | None:
        """Get float attribute value."""
        val = self.attr(node, attr_name)
        if val:
            try:
                return float(val)
            except ValueError:
                return default_value
        return default_value

    def bool_attr(
        self,
        node: etree._Element,
        attr_name: str,
        default_value: bool | None = None,  # noqa: FBT001
    ) -> bool:
        """Get boolean attribute value."""
        return convert_boolean(
            self.attr(node, attr_name),
            default_value=default_value if default_value is not None else False,
        )

    def length_attr(
        self,
        node: etree._Element,
        attr_name: str,
        usage: LengthUsageType = LengthUsage.Dxa,
    ) -> Length | None:
        """Get length attribute value."""
        val = self.attr(node, attr_name)
        return convert_length(val, usage) if val else None


global_xml_parser = XmlParser()

