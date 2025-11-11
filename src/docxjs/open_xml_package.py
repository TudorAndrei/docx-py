"""Open XML package module for handling ZIP-based document packages."""

import io
import zipfile
from pathlib import Path

from lxml import etree

from docxjs.relationship import Relationship, parse_relationships
from docxjs.utils import split_path
from docxjs.xml_parser import XmlParser, parse_xml_string


class OpenXmlPackageOptions:
    """Options for Open XML package."""

    def __init__(
        self,
        *,
        trim_xml_declaration: bool = True,
        keep_origin: bool = False,
    ) -> None:
        """Initialize package options."""
        self.trim_xml_declaration = trim_xml_declaration
        self.keep_origin = keep_origin


class OpenXmlPackage:
    """Open XML package handler."""

    def __init__(
        self, zip_file: zipfile.ZipFile, options: OpenXmlPackageOptions,
    ) -> None:
        """Initialize Open XML package."""
        self._zip = zip_file
        self.options = options
        self.xml_parser = XmlParser()

    def get(self, path: str) -> zipfile.ZipInfo | None:
        """Get file info from package."""
        p = self._normalize_path(path)
        try:
            return self._zip.getinfo(p)
        except KeyError:
            try:
                return self._zip.getinfo(p.replace("/", "\\"))
            except KeyError:
                return None

    def update(self, path: str, content: str | bytes) -> None:
        """Update file in package."""
        if isinstance(content, str):
            content = content.encode("utf-8")
        self._zip.writestr(path, content)

    @staticmethod
    def load(
        input_data: bytes | str | io.BytesIO,
        options: OpenXmlPackageOptions,
    ) -> "OpenXmlPackage":
        """Load Open XML package from input data."""
        if isinstance(input_data, str):
            path = Path(input_data)
            zip_data = io.BytesIO(path.read_bytes())
        elif isinstance(input_data, bytes):
            zip_data = io.BytesIO(input_data)
        else:
            zip_data = input_data

        zip_file = zipfile.ZipFile(zip_data, "r")
        return OpenXmlPackage(zip_file, options)

    def load_file(
        self, path: str, *, as_bytes: bool = False,
    ) -> str | bytes | None:
        """Load file from package."""
        zip_info = self.get(path)
        if zip_info is None:
            return None

        data = self._zip.read(zip_info)
        if as_bytes:
            return data
        try:
            return data.decode("utf-8")
        except UnicodeDecodeError:
            return data.decode("latin-1")

    def load_relationships(
        self, path: str | None = None,
    ) -> list[Relationship] | None:
        """Load relationships from package."""
        if path is None:
            rels_path = "_rels/.rels"
        else:
            folder, file_name = split_path(path)
            rels_path = f"{folder}_rels/{file_name}.rels"

        txt = self.load_file(rels_path)
        if txt is None:
            return None

        xml_doc = self.parse_xml_document(txt)
        if xml_doc is None:
            return None

        return parse_relationships(xml_doc, self.xml_parser)

    def parse_xml_document(
        self, txt: str,
    ) -> etree._Element | None:
        """Parse XML document from string."""
        return parse_xml_string(txt, self.options.trim_xml_declaration)

    def _normalize_path(self, path: str) -> str:
        """Normalize file path."""
        return path.removeprefix("/")

