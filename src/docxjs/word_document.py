"""Word document parser module."""

import base64
from typing import Any

from docxjs.comments.comments_extended_part import CommentsExtendedPart
from docxjs.comments.comments_part import CommentsPart
from docxjs.document.document_part import DocumentPart
from docxjs.document_parser import DocumentParser
from docxjs.document_props.core_props_part import CorePropsPart
from docxjs.document_props.custom_props_part import CustomPropsPart
from docxjs.document_props.extended_props_part import ExtendedPropsPart
from docxjs.font_table.font_table import FontTablePart
from docxjs.header_footer.parts import FooterPart, HeaderPart
from docxjs.notes.parts import EndnotesPart, FootnotesPart
from docxjs.numbering.numbering_part import NumberingPart
from docxjs.open_xml_package import OpenXmlPackage, OpenXmlPackageOptions
from docxjs.part import Part
from docxjs.relationship import Relationship, RelationshipTypes
from docxjs.settings.settings_part import SettingsPart
from docxjs.styles.styles_part import StylesPart
from docxjs.theme.theme_part import ThemePart
from docxjs.utils import bytes_to_base64, resolve_path, split_path

top_level_rels = [
    {"type": RelationshipTypes.OfficeDocument.value, "target": "word/document.xml"},
    {"type": RelationshipTypes.ExtendedProperties.value, "target": "docProps/app.xml"},
    {"type": RelationshipTypes.CoreProperties.value, "target": "docProps/core.xml"},
    {"type": RelationshipTypes.CustomProperties.value, "target": "docProps/custom.xml"},
]


class WordDocument:
    """Word document parser."""

    def __init__(self) -> None:
        """Initialize Word document."""
        self._package: OpenXmlPackage | None = None
        self._parser: DocumentParser | None = None
        self._options = None

        self.rels: list[Relationship] = []
        self.parts: list[Part] = []
        self.parts_map: dict[str, Part] = {}

        self.document_part: DocumentPart | None = None
        self.font_table_part: FontTablePart | None = None
        self.numbering_part: NumberingPart | None = None
        self.styles_part: StylesPart | None = None
        self.footnotes_part = None
        self.endnotes_part = None
        self.theme_part = None
        self.core_props_part = None
        self.extended_props_part = None
        self.settings_part = None
        self.comments_part = None
        self.comments_extended_part = None

    @staticmethod
    def load(
        input_data: bytes,
        parser: DocumentParser,
        options: dict[str, Any],
    ) -> "WordDocument":
        """Load Word document from input data."""
        d = WordDocument()
        d._options = options
        d._parser = parser

        package_options = OpenXmlPackageOptions(
            trim_xml_declaration=options.get("trim_xml_declaration", True),
            keep_origin=options.get("keep_origin", False),
        )
        d._package = OpenXmlPackage.load(input_data, package_options)
        d.rels = d._package.load_relationships() or []

        for rel in top_level_rels:
            r = next((x for x in d.rels if x.type == rel["type"]), None) or rel
            if isinstance(r, Relationship):
                target = r.target
                rel_type = r.type
            else:
                target = r.get("target", rel["target"])
                rel_type = r.get("type", rel["type"])
            d.load_relationship_part(
                target,
                rel_type,
            )

        return d

    def load_relationship_part(  # noqa: C901, PLR0912
        self,
        path: str,
        rel_type: str,
    ) -> Part | None:
        """Load relationship part."""
        if path in self.parts_map:
            return self.parts_map[path]

        if self._package is None:
            return None
        if self._parser is None:
            return None

        if not self._package.get(path):
            return None

        part: Part | None = None

        if rel_type == RelationshipTypes.OfficeDocument.value:
            self.document_part = part = DocumentPart(
                self._package,
                path,
                self._parser,
            )
        elif rel_type == RelationshipTypes.FontTable.value:
            self.font_table_part = part = FontTablePart(
                self._package,
                path,
            )
        elif rel_type == RelationshipTypes.Numbering.value:
            self.numbering_part = part = NumberingPart(
                self._package,
                path,
                self._parser,
            )
        elif rel_type == RelationshipTypes.Styles.value:
            self.styles_part = part = StylesPart(
                self._package,
                path,
                self._parser,
            )
        elif rel_type == RelationshipTypes.Theme.value:
            self.theme_part = part = ThemePart(self._package, path)
        elif rel_type == RelationshipTypes.Footnotes.value:
            self.footnotes_part = part = FootnotesPart(
                self._package,
                path,
                self._parser,
            )
        elif rel_type == RelationshipTypes.Endnotes.value:
            self.endnotes_part = part = EndnotesPart(
                self._package,
                path,
                self._parser,
            )
        elif rel_type == RelationshipTypes.Footer.value:
            part = FooterPart(self._package, path, self._parser)
            if not part:
                return None
        elif rel_type == RelationshipTypes.Header.value:
            part = HeaderPart(self._package, path, self._parser)
            if not part:
                return None
        elif rel_type == RelationshipTypes.CoreProperties.value:
            self.core_props_part = part = CorePropsPart(self._package, path)
        elif rel_type == RelationshipTypes.ExtendedProperties.value:
            self.extended_props_part = part = ExtendedPropsPart(
                self._package,
                path,
            )
        elif rel_type == RelationshipTypes.CustomProperties.value:
            part = CustomPropsPart(self._package, path)
        elif rel_type == RelationshipTypes.Settings.value:
            self.settings_part = part = SettingsPart(self._package, path)
        elif rel_type == RelationshipTypes.Comments.value:
            self.comments_part = part = CommentsPart(
                self._package,
                path,
                self._parser,
            )
        elif rel_type == RelationshipTypes.CommentsExtended.value:
            self.comments_extended_part = part = CommentsExtendedPart(
                self._package,
                path,
            )

        if part is None:
            return None

        self.parts_map[path] = part
        self.parts.append(part)

        part.load()

        if part.rels:
            folder, _ = split_path(part.path)
            for rel in part.rels:
                resolved_path = resolve_path(rel.target, folder)
                self.load_relationship_part(resolved_path, rel.type)

        return part

    def load_document_image(
        self,
        id_val: str,
        part: Part | None = None,
    ) -> str | None:
        """Load document image."""
        data = self.load_resource(
            part or self.document_part,
            id_val,
            as_bytes=True,
        )
        return self.blob_to_url(data) if data else None

    def load_numbering_image(self, id_val: str) -> str | None:
        """Load numbering image."""
        data = self.load_resource(
            self.numbering_part,
            id_val,
            as_bytes=True,
        )
        return self.blob_to_url(data) if data else None

    def load_font(self, id_val: str, key: str) -> str | None:
        """Load font."""
        data = self.load_resource(
            self.font_table_part,
            id_val,
            as_bytes=True,
        )
        if data:
            deobfuscated = deobfuscate_font(data, key)
            return self.blob_to_url(deobfuscated) if deobfuscated else None
        return None

    def load_alt_chunk(
        self,
        id_val: str,
        part: Part | None = None,
    ) -> str | None:
        """Load alt chunk."""
        data = self.load_resource(
            part or self.document_part,
            id_val,
            as_bytes=False,
        )
        return data.decode("utf-8") if isinstance(data, bytes) else data

    def blob_to_url(self, blob: bytes) -> str | None:
        """Convert blob to URL."""
        if not blob:
            return None

        if isinstance(self._options, dict) and self._options.get(
            "use_base64_url", False
        ):
            return bytes_to_base64(blob)
        encoded = base64.b64encode(blob).decode("utf-8")
        return f"data:application/octet-stream;base64,{encoded}"

    def find_part_by_rel_id(
        self,
        id_val: str,
        base_part: Part | None = None,
    ) -> Part | None:
        """Find part by relationship ID."""
        rels = (base_part.rels if base_part else self.rels) or []
        rel = next((r for r in rels if r.id == id_val), None)
        if not rel:
            return None
        folder = split_path(base_part.path)[0] if base_part else ""
        resolved_path = resolve_path(rel.target, folder)
        return self.parts_map.get(resolved_path)

    def get_path_by_id(self, part: Part, id_val: str) -> str | None:
        """Get path by ID."""
        rel = (
            next((x for x in part.rels if x.id == id_val), None) if part.rels else None
        )
        if not rel:
            return None
        folder, _ = split_path(part.path)
        return resolve_path(rel.target, folder)

    def load_resource(
        self,
        part: Part | None,
        id_val: str,
        as_bytes: bool = False,  # noqa: FBT001, FBT002
    ) -> bytes | None:
        """Load resource from part."""
        if not part:
            return None
        path = self.get_path_by_id(part, id_val)
        if not path:
            return None
        if self._package is None:
            return None
        result = self._package.load_file(path, as_bytes=as_bytes)
        if result is None:
            return None
        if isinstance(result, str):
            return result.encode("utf-8")
        return result


def deobfuscate_font(data: bytes, guid_key: str) -> bytes:
    """Deobfuscate font data."""
    length = 16
    trimmed = guid_key.replace("{", "").replace("}", "").replace("-", "")
    numbers = [0] * length

    for i in range(length):
        numbers[length - i - 1] = int(trimmed[i * 2 : i * 2 + 2], 16)

    result = bytearray(data)
    for i in range(min(32, len(result))):
        result[i] = result[i] ^ numbers[i % length]

    return bytes(result)
