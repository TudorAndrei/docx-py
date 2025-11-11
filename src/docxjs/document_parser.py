"""Document parser module for parsing Word documents."""

from collections.abc import Callable
from typing import TYPE_CHECKING, Any

from lxml import etree

if TYPE_CHECKING:
    from lxml.etree import _Element

from docxjs.comments.elements import (
    WmlComment,
    WmlCommentRangeEnd,
    WmlCommentRangeStart,
    WmlCommentReference,
)
from docxjs.document.bookmarks import parse_bookmark_end, parse_bookmark_start
from docxjs.document.common import LengthUsage, convert_length
from docxjs.document.document import DocumentElement
from docxjs.document.dom import (
    DomType,
    IDomImage,
    IDomNumbering,
    NumberingPicBullet,
    OpenXmlElement,
    WmlAltChunk,
    WmlBreak,
    WmlHyperlink,
    WmlNoteReference,
    WmlSmartTag,
    WmlSymbol,
    WmlTable,
    WmlTableCell,
    WmlTableColumn,
    WmlTableRow,
    WmlText,
)
from docxjs.document.fields import (
    WmlFieldChar,
    WmlFieldSimple,
    WmlInstructionText,
)
from docxjs.document.paragraph import (
    WmlParagraph,
    parse_paragraph_properties,
    parse_paragraph_property,
)
from docxjs.document.run import WmlRun, parse_run_properties
from docxjs.document.section import SectionProperties, parse_section_properties
from docxjs.document.style import IDomStyle, IDomSubStyle
from docxjs.utils import enclose_font_family
from docxjs.vml.vml import parse_vml_element
from docxjs.xml_parser import global_xml_parser

autos = {
    "shd": "inherit",
    "color": "black",
    "borderColor": "black",
    "highlight": "transparent",
}

supported_namespace_uris = []

mml_tag_map = {
    "oMath": DomType.MmlMath,
    "oMathPara": DomType.MmlMathParagraph,
    "f": DomType.MmlFraction,
    "func": DomType.MmlFunction,
    "fName": DomType.MmlFunctionName,
    "num": DomType.MmlNumerator,
    "den": DomType.MmlDenominator,
    "rad": DomType.MmlRadical,
    "deg": DomType.MmlDegree,
    "e": DomType.MmlBase,
    "sSup": DomType.MmlSuperscript,
    "sSub": DomType.MmlSubscript,
    "sPre": DomType.MmlPreSubSuper,
    "sup": DomType.MmlSuperArgument,
    "sub": DomType.MmlSubArgument,
    "d": DomType.MmlDelimiter,
    "nary": DomType.MmlNary,
    "eqArr": DomType.MmlEquationArray,
    "lim": DomType.MmlLimit,
    "limLow": DomType.MmlLimitLower,
    "m": DomType.MmlMatrix,
    "mr": DomType.MmlMatrixRow,
    "box": DomType.MmlBox,
    "bar": DomType.MmlBar,
    "groupChr": DomType.MmlGroupChar,
}


class DocumentParserOptions:
    """Document parser options."""

    def __init__(
        self,
        *,
        ignore_width: bool = False,
        debug: bool = False,
    ) -> None:
        """Initialize document parser options."""
        self.ignore_width = ignore_width
        self.debug = debug


class DocumentParser:
    """Document parser for parsing Word documents."""

    def __init__(
        self,
        options: DocumentParserOptions | None = None,
    ) -> None:
        """Initialize document parser."""
        self.options = options or DocumentParserOptions()
        self.xml = global_xml_parser

    def parse_notes(
        self,
        xml_doc: "_Element",
        elem_name: str,
        elem_class: type,
    ) -> list[OpenXmlElement]:
        """Parse notes from XML document."""
        result = []
        for el in self.xml.elements(xml_doc, elem_name):
            node = elem_class()
            node.id = self.xml.attr(el, "id")
            node.note_type = self.xml.attr(el, "type")
            node.children = self.parse_body_elements(el)
            result.append(node)
        return result

    def parse_comments(self, xml_doc: "_Element") -> list:
        """Parse comments from XML document."""
        result = []
        for el in self.xml.elements(xml_doc, "comment"):
            item = WmlComment()
            item.id = self.xml.attr(el, "id")
            item.author = self.xml.attr(el, "author")
            item.initials = self.xml.attr(el, "initials")
            item.date = self.xml.attr(el, "date")
            item.children = self.parse_body_elements(el)
            result.append(item)
        return result

    def parse_document_file(
        self,
        xml_doc: "_Element",
    ) -> DocumentElement:
        """Parse document file from XML."""
        xbody = self.xml.element(xml_doc, "body")
        background = self.xml.element(xml_doc, "background")
        sect_pr = self.xml.element(xbody, "sectPr") if xbody is not None else None

        result = DocumentElement()
        result.type = DomType.Document
        result.children = self.parse_body_elements(xbody) if xbody is not None else []
        result.props = (
            parse_section_properties(sect_pr, self.xml)
            if sect_pr is not None
            else SectionProperties()
        )
        result.css_style = (
            self.parse_background(background) if background is not None else {}
        )
        return result

    def parse_background(
        self,
        elem: "_Element",
    ) -> dict[str, str]:
        """Parse background element."""
        result = {}
        color = self._color_attr(elem, "color")
        if color:
            result["background-color"] = color
        return result

    def parse_body_elements(
        self,
        element: "_Element | None",
    ) -> list[OpenXmlElement]:
        """Parse body elements from XML."""
        children = []
        if element is None:
            return children

        for elem in self.xml.elements(element):
            local_name = etree.QName(elem).localname
            if local_name == "p":
                children.append(self.parse_paragraph(elem))
            elif local_name == "altChunk":
                children.append(self.parse_alt_chunk(elem))
            elif local_name == "tbl":
                children.append(self.parse_table(elem))
            elif local_name == "sdt":
                children.extend(
                    self.parse_sdt(
                        elem,
                        lambda e: self.parse_body_elements(e),
                    ),
                )
        return children

    def parse_styles_file(
        self,
        xstyles: "_Element",
    ) -> list[IDomStyle]:
        """Parse styles file from XML."""
        result = []
        for n in self.xml.elements(xstyles):
            local_name = etree.QName(n).localname
            if local_name == "style":
                result.append(self.parse_style(n))
            elif local_name == "docDefaults":
                result.append(self.parse_default_styles(n))
        return result

    def parse_default_styles(self, node: "_Element") -> IDomStyle:
        """Parse default styles from XML node."""
        result = IDomStyle()
        result.id = None
        result.name = None
        result.target = None
        result.based_on = None
        result.styles = []

        for c in self.xml.elements(node):
            local_name = etree.QName(c).localname
            if local_name == "rPrDefault":
                r_pr = self.xml.element(c, "rPr")
                if r_pr is not None:
                    sub_style = IDomSubStyle()
                    sub_style.target = "span"
                    sub_style.values = self.parse_default_properties(
                        r_pr,
                        {},
                    )
                    result.styles.append(sub_style)
            elif local_name == "pPrDefault":
                p_pr = self.xml.element(c, "pPr")
                if p_pr is not None:
                    sub_style = IDomSubStyle()
                    sub_style.target = "p"
                    sub_style.values = self.parse_default_properties(
                        p_pr,
                        {},
                    )
                    result.styles.append(sub_style)
        return result

    def parse_style(self, node: "_Element") -> IDomStyle:  # noqa: C901, PLR0912
        """Parse style from XML node."""
        result = IDomStyle()
        result.id = self.xml.attr(node, "styleId")
        result.is_default = self.xml.bool_attr(
            node,
            "default",
            default_value=False,
        )
        result.name = None
        result.target = None
        result.based_on = None
        result.styles = []
        result.linked = None

        style_type = self.xml.attr(node, "type")
        if style_type == "paragraph":
            result.target = "p"
        elif style_type == "table":
            result.target = "table"
        elif style_type == "character":
            result.target = "span"

        for n in self.xml.elements(node):
            local_name = etree.QName(n).localname
            if local_name == "basedOn":
                result.based_on = self.xml.attr(n, "val")
            elif local_name == "name":
                result.name = self.xml.attr(n, "val")
            elif local_name == "link":
                result.linked = self.xml.attr(n, "val")
            elif local_name == "next":
                result.next = self.xml.attr(n, "val")
            elif local_name == "aliases":
                aliases_str = self.xml.attr(n, "val")
                result.aliases = aliases_str.split(",") if aliases_str else []
            elif local_name == "pPr":
                sub_style = IDomSubStyle()
                sub_style.target = "p"
                sub_style.values = self.parse_default_properties(n, {})
                result.styles.append(sub_style)
                result.paragraph_props = parse_paragraph_properties(n, self.xml)
            elif local_name == "rPr":
                sub_style = IDomSubStyle()
                sub_style.target = "span"
                sub_style.values = self.parse_default_properties(n, {})
                result.styles.append(sub_style)
                result.run_props = parse_run_properties(n, self.xml)
            elif local_name in ["tblPr", "tcPr"]:
                sub_style = IDomSubStyle()
                sub_style.target = "td"
                sub_style.values = self.parse_default_properties(n, {})
                result.styles.append(sub_style)
            elif local_name == "tblStylePr":
                for s in self.parse_table_style(n):
                    result.styles.append(s)
        return result

    def parse_table_style(self, node: "_Element") -> list[IDomSubStyle]:  # noqa: C901, PLR0912
        """Parse table style from XML node."""
        result = []
        style_type = self.xml.attr(node, "type")
        selector = ""
        modificator = ""

        if style_type == "firstRow":
            modificator = ".first-row"
            selector = "tr.first-row td"
        elif style_type == "lastRow":
            modificator = ".last-row"
            selector = "tr.last-row td"
        elif style_type == "firstCol":
            modificator = ".first-col"
            selector = "td.first-col"
        elif style_type == "lastCol":
            modificator = ".last-col"
            selector = "td.last-col"
        elif style_type == "band1Vert":
            modificator = ":not(.no-vband)"
            selector = "td.odd-col"
        elif style_type == "band2Vert":
            modificator = ":not(.no-vband)"
            selector = "td.even-col"
        elif style_type == "band1Horz":
            modificator = ":not(.no-hband)"
            selector = "tr.odd-row"
        elif style_type == "band2Horz":
            modificator = ":not(.no-hband)"
            selector = "tr.even-row"
        else:
            return []

        for n in self.xml.elements(node):
            local_name = etree.QName(n).localname
            if local_name == "pPr":
                sub_style = IDomSubStyle()
                sub_style.target = f"{selector} p"
                sub_style.mod = modificator
                sub_style.values = self.parse_default_properties(n, {})
                result.append(sub_style)
            elif local_name == "rPr":
                sub_style = IDomSubStyle()
                sub_style.target = f"{selector} span"
                sub_style.mod = modificator
                sub_style.values = self.parse_default_properties(n, {})
                result.append(sub_style)
            elif local_name in ["tblPr", "tcPr"]:
                sub_style = IDomSubStyle()
                sub_style.target = selector
                sub_style.mod = modificator
                sub_style.values = self.parse_default_properties(n, {})
                result.append(sub_style)
        return result

    def parse_numbering_file(
        self,
        node: "_Element",
    ) -> list[IDomNumbering]:
        """Parse numbering file from XML node."""
        result = []
        mapping = {}
        bullets = []

        for n in self.xml.elements(node):
            local_name = etree.QName(n).localname
            if local_name == "abstractNum":
                result.extend(self.parse_abstract_numbering(n, bullets))
            elif local_name == "numPicBullet":
                bullet = self.parse_numbering_pic_bullet(n)
                if bullet:
                    bullets.append(bullet)
            elif local_name == "num":
                num_id = self.xml.attr(n, "numId")
                abstract_num_id = self.xml.element_attr(
                    n,
                    "abstractNumId",
                    "val",
                )
                if abstract_num_id:
                    mapping[abstract_num_id] = num_id

        for x in result:
            if x.id in mapping:
                x.id = mapping[x.id]
        return result

    def parse_numbering_pic_bullet(
        self,
        elem: "_Element",
    ) -> NumberingPicBullet | None:
        """Parse numbering picture bullet from XML element."""
        pict = self.xml.element(elem, "pict")
        if not pict:
            return None
        shape = self.xml.element(pict, "shape")
        if not shape:
            return None
        imagedata = self.xml.element(shape, "imagedata")
        if not imagedata:
            return None

        result = NumberingPicBullet()
        result.id = self.xml.int_attr(elem, "numPicBulletId", 0)
        result.src = self.xml.attr(imagedata, "id") or ""
        result.style = self.xml.attr(shape, "style")
        return result

    def parse_abstract_numbering(
        self,
        node: "_Element",
        bullets: list[NumberingPicBullet],
    ) -> list[IDomNumbering]:
        """Parse abstract numbering from XML node."""
        result = []
        id_val = self.xml.attr(node, "abstractNumId") or ""

        for n in self.xml.elements(node):
            local_name = etree.QName(n).localname
            if local_name == "lvl":
                result.append(
                    self.parse_numbering_level(id_val, n, bullets),
                )
        return result

    def parse_numbering_level(
        self,
        id_val: str,
        node: "_Element",
        bullets: list[NumberingPicBullet],
    ) -> IDomNumbering:
        """Parse numbering level from XML node."""
        result = IDomNumbering()
        result.id = id_val
        result.level = self.xml.int_attr(node, "ilvl", 0)
        result.start = 1
        result.p_style_name = None
        result.p_style = {}
        result.r_style = {}
        result.suff = "tab"

        for n in self.xml.elements(node):
            local_name = etree.QName(n).localname
            if local_name == "start":
                result.start = self.xml.int_attr(n, "val", 1)
            elif local_name == "pPr":
                self.parse_default_properties(n, result.p_style)
            elif local_name == "rPr":
                self.parse_default_properties(n, result.r_style)
            elif local_name == "lvlPicBulletId":
                bullet_id = self.xml.int_attr(n, "val")
                result.bullet = next(
                    (x for x in bullets if x and x.id == bullet_id),
                    None,
                )
            elif local_name == "lvlText":
                result.level_text = self.xml.attr(n, "val")
            elif local_name == "pStyle":
                result.p_style_name = self.xml.attr(n, "val")
            elif local_name == "numFmt":
                result.format = self.xml.attr(n, "val")
            elif local_name == "suff":
                result.suff = self.xml.attr(n, "val") or "tab"
        return result

    def parse_sdt(
        self,
        node: "_Element",
        parser: Callable,
    ) -> list[OpenXmlElement]:
        """Parse SDT element from XML node."""
        sdt_content = self.xml.element(node, "sdtContent")
        return parser(sdt_content) if sdt_content is not None else []

    def parse_inserted(
        self,
        node: "_Element",
        parent_parser: Callable,
    ) -> OpenXmlElement:
        """Parse inserted element from XML node."""
        result = OpenXmlElement()
        result.type = DomType.Inserted
        parsed = parent_parser(node)
        result.children = parsed.children if parsed and parsed.children else []
        return result

    def parse_deleted(
        self,
        node: "_Element",
        parent_parser: Callable,
    ) -> OpenXmlElement:
        """Parse deleted element from XML node."""
        result = OpenXmlElement()
        result.type = DomType.Deleted
        parsed = parent_parser(node)
        result.children = parsed.children if parsed and parsed.children else []
        return result

    def parse_alt_chunk(self, node: "_Element") -> WmlAltChunk:
        """Parse alt chunk element from XML node."""
        result = WmlAltChunk()
        result.id = self.xml.attr(node, "id")
        return result

    def parse_paragraph(self, node: "_Element") -> WmlParagraph:  # noqa: C901, PLR0912
        """Parse paragraph from XML node."""
        result = WmlParagraph()

        for el in self.xml.elements(node):
            local_name = etree.QName(el).localname
            if local_name == "pPr":
                self.parse_paragraph_properties(el, result)
            elif local_name == "r":
                result.children.append(self.parse_run(el, result))
            elif local_name == "hyperlink":
                result.children.append(self.parse_hyperlink(el, result))
            elif local_name == "smartTag":
                result.children.append(self.parse_smart_tag(el, result))
            elif local_name == "bookmarkStart":
                result.children.append(parse_bookmark_start(el, self.xml))
            elif local_name == "bookmarkEnd":
                result.children.append(parse_bookmark_end(el, self.xml))
            elif local_name == "commentRangeStart":
                result.children.append(
                    WmlCommentRangeStart(self.xml.attr(el, "id") or ""),
                )
            elif local_name == "commentRangeEnd":
                result.children.append(
                    WmlCommentRangeEnd(self.xml.attr(el, "id") or ""),
                )
            elif local_name in ["oMath", "oMathPara"]:
                result.children.append(self.parse_math_element(el))
            elif local_name == "sdt":
                result.children.extend(
                    self.parse_sdt(
                        el,
                        lambda e: self.parse_paragraph(e).children,
                    ),
                )
            elif local_name == "ins":
                result.children.append(
                    self.parse_inserted(
                        el,
                        lambda e: self.parse_paragraph(e),
                    ),
                )
            elif local_name == "del":
                result.children.append(
                    self.parse_deleted(
                        el,
                        lambda e: self.parse_paragraph(e),
                    ),
                )
        return result

    def parse_paragraph_properties(
        self,
        elem: "_Element",
        paragraph: WmlParagraph,
    ) -> None:
        """Parse paragraph properties from XML element."""
        if paragraph.css_style is None:
            paragraph.css_style = {}

        def handler(c: "_Element") -> bool:
            if parse_paragraph_property(c, paragraph, self.xml):
                return True
            local_name = etree.QName(c).localname
            if local_name == "pStyle":
                paragraph.style_name = self.xml.attr(c, "val")
                return True
            if local_name == "cnfStyle":
                paragraph.class_name = self._class_name_of_cnf_style(c)
                return True
            if local_name == "framePr":
                self.parse_frame(c, paragraph)
                return True
            return local_name == "rPr"

        self.parse_default_properties(
            elem,
            paragraph.css_style,
            None,
            handler,
        )

    def parse_frame(
        self,
        node: "_Element",
        paragraph: WmlParagraph,
    ) -> None:
        """Parse frame properties from XML node."""
        drop_cap = self.xml.attr(node, "dropCap")
        if drop_cap == "drop":
            paragraph.css_style["float"] = "left"

    def parse_hyperlink(
        self,
        node: "_Element",
        parent: OpenXmlElement | None = None,
    ) -> WmlHyperlink:
        """Parse hyperlink from XML node."""
        result = WmlHyperlink()
        result.parent = parent
        result.anchor = self.xml.attr(node, "anchor")
        result.id = self.xml.attr(node, "id")

        for c in self.xml.elements(node):
            local_name = etree.QName(c).localname
            if local_name == "r":
                result.children.append(self.parse_run(c, result))
        return result

    def parse_smart_tag(
        self,
        node: "_Element",
        parent: OpenXmlElement | None = None,
    ) -> WmlSmartTag:
        """Parse smart tag from XML node."""
        result = WmlSmartTag()
        result.parent = parent
        uri = self.xml.attr(node, "uri")
        element = self.xml.attr(node, "element")
        if uri:
            result.uri = uri
        if element:
            result.element = element

        for c in self.xml.elements(node):
            local_name = etree.QName(c).localname
            if local_name == "r":
                result.children.append(self.parse_run(c, result))
        return result

    def parse_run(  # noqa: C901, PLR0912, PLR0915
        self,
        node: "_Element",
        parent: OpenXmlElement | None = None,
    ) -> WmlRun:
        """Parse run from XML node."""
        result = WmlRun()
        result.parent = parent

        for c in self.xml.elements(node):
            checked_c = self.check_alternate_content(c)
            local_name = etree.QName(checked_c).localname

            if local_name == "t":
                text_elem = WmlText()
                text_elem.text = checked_c.text or ""
                result.children.append(text_elem)
            elif local_name == "delText":
                text_elem = WmlText()
                text_elem.type = DomType.DeletedText
                text_elem.text = checked_c.text or ""
                result.children.append(text_elem)
            elif local_name == "commentReference":
                result.children.append(
                    WmlCommentReference(self.xml.attr(checked_c, "id") or ""),
                )
            elif local_name == "fldSimple":
                field = WmlFieldSimple()
                field.instruction = self.xml.attr(checked_c, "instr") or ""
                field.lock = self.xml.bool_attr(
                    checked_c,
                    "lock",
                    default_value=False,
                )
                field.dirty = self.xml.bool_attr(
                    checked_c,
                    "dirty",
                    default_value=False,
                )
                result.children.append(field)
            elif local_name == "instrText":
                result.field_run = True
                instr = WmlInstructionText()
                instr.text = checked_c.text or ""
                result.children.append(instr)
            elif local_name == "fldChar":
                result.field_run = True
                field_char = WmlFieldChar()
                field_char.char_type = self.xml.attr(checked_c, "fldCharType") or ""
                field_char.lock = self.xml.bool_attr(
                    checked_c,
                    "lock",
                    default_value=False,
                )
                field_char.dirty = self.xml.bool_attr(
                    checked_c,
                    "dirty",
                    default_value=False,
                )
                result.children.append(field_char)
            elif local_name == "noBreakHyphen":
                elem = OpenXmlElement()
                elem.type = DomType.NoBreakHyphen
                result.children.append(elem)
            elif local_name == "br":
                break_elem = WmlBreak()
                break_type = self.xml.attr(checked_c, "type") or "textWrapping"
                break_elem.break_type = break_type
                result.children.append(break_elem)
            elif local_name == "lastRenderedPageBreak":
                break_elem = WmlBreak()
                break_elem.break_type = "lastRenderedPageBreak"
                result.children.append(break_elem)
            elif local_name == "sym":
                symbol = WmlSymbol()
                font = self.xml.attr(checked_c, "font") or ""
                symbol.font = enclose_font_family(font)
                symbol.char = self.xml.attr(checked_c, "char") or ""
                result.children.append(symbol)
            elif local_name == "tab":
                elem = OpenXmlElement()
                elem.type = DomType.Tab
                result.children.append(elem)
            elif local_name == "footnoteReference":
                note_ref = WmlNoteReference()
                note_ref.type = DomType.FootnoteReference
                note_ref.id = self.xml.attr(checked_c, "id") or ""
                result.children.append(note_ref)
            elif local_name == "endnoteReference":
                note_ref = WmlNoteReference()
                note_ref.type = DomType.EndnoteReference
                note_ref.id = self.xml.attr(checked_c, "id") or ""
                result.children.append(note_ref)
            elif local_name == "drawing":
                d = self.parse_drawing(checked_c)
                if d:
                    result.children = [d]
            elif local_name == "pict":
                result.children.append(self.parse_vml_picture(checked_c))
            elif local_name == "rPr":
                self.parse_run_properties(checked_c, result)
        return result

    def parse_math_element(self, elem: "_Element") -> OpenXmlElement:
        """Parse math element from XML."""
        local_name = etree.QName(elem).localname
        props_tag = f"{local_name}Pr"
        result = OpenXmlElement()
        result.type = mml_tag_map.get(local_name, DomType.MmlMath)
        result.children = []

        for el in self.xml.elements(elem):
            el_local_name = etree.QName(el).localname
            child_type = mml_tag_map.get(el_local_name)
            if child_type:
                result.children.append(self.parse_math_element(el))
            elif el_local_name == "r":
                run = self.parse_run(el)
                run.type = DomType.MmlRun
                result.children.append(run)
            elif el_local_name == props_tag:
                result.props = self.parse_math_properties(el)
        return result

    def parse_math_properties(
        self,
        elem: "_Element",
    ) -> dict[str, Any]:
        """Parse math properties from XML element."""
        result = {}
        for el in self.xml.elements(elem):
            local_name = etree.QName(el).localname
            if local_name == "chr":
                result["char"] = self.xml.attr(el, "val")
            elif local_name == "vertJc":
                result["verticalJustification"] = self.xml.attr(el, "val")
            elif local_name == "pos":
                result["position"] = self.xml.attr(el, "val")
            elif local_name == "degHide":
                result["hideDegree"] = self.xml.bool_attr(
                    el,
                    "val",
                    default_value=False,
                )
            elif local_name == "begChr":
                result["beginChar"] = self.xml.attr(el, "val")
            elif local_name == "endChr":
                result["endChar"] = self.xml.attr(el, "val")
        return result

    def parse_run_properties(
        self,
        elem: "_Element",
        run: WmlRun,
    ) -> None:
        """Parse run properties from XML element."""
        if run.css_style is None:
            run.css_style = {}

        def handler(c: "_Element") -> bool:
            local_name = etree.QName(c).localname
            if local_name == "rStyle":
                run.style_name = self.xml.attr(c, "val")
                return True
            if local_name == "vertAlign":
                run.vertical_align = self._value_of_vert_align(
                    c,
                    as_tag_name=True,
                )
                return True
            return False

        self.parse_default_properties(elem, run.css_style, None, handler)

    def parse_vml_picture(self, elem: "_Element") -> OpenXmlElement:
        """Parse VML picture from XML element."""
        result = OpenXmlElement()
        result.type = DomType.VmlPicture
        result.children = []

        for el in self.xml.elements(elem):
            child = parse_vml_element(el, self)
            if child:
                result.children.append(child)
        return result

    def check_alternate_content(
        self,
        elem: "_Element",
    ) -> "_Element":
        """Check and handle alternate content element."""
        local_name = etree.QName(elem).localname
        if local_name != "AlternateContent":
            return elem

        choice = self.xml.element(elem, "Choice")
        if choice:
            requires = self.xml.attr(choice, "Requires")
            namespace_uri = elem.nsmap.get(requires) if requires else None
            if namespace_uri in supported_namespace_uris:
                return choice[0] if len(choice) > 0 else choice

        fallback = self.xml.element(elem, "Fallback")
        return fallback[0] if fallback and len(fallback) > 0 else elem

    def parse_drawing(self, node: "_Element") -> OpenXmlElement | None:
        """Parse drawing from XML node."""
        for n in self.xml.elements(node):
            local_name = etree.QName(n).localname
            if local_name in ["inline", "anchor"]:
                return self.parse_drawing_wrapper(n)
        return None

    def parse_drawing_wrapper(  # noqa: C901, PLR0912
        self,
        node: "_Element",
    ) -> OpenXmlElement:
        """Parse drawing wrapper from XML node."""
        result = OpenXmlElement()
        result.type = DomType.Drawing
        result.children = []
        result.css_style = {}
        is_anchor = etree.QName(node).localname == "anchor"

        wrap_type = None
        simple_pos = self.xml.bool_attr(
            node,
            "simplePos",
            default_value=False,
        )

        pos_x = {"relative": "page", "align": "left", "offset": "0"}
        pos_y = {"relative": "page", "align": "top", "offset": "0"}

        for n in self.xml.elements(node):
            local_name = etree.QName(n).localname
            if local_name == "simplePos" and simple_pos:
                pos_x["offset"] = self.xml.length_attr(n, "x", LengthUsage.Emu) or "0"
                pos_y["offset"] = self.xml.length_attr(n, "y", LengthUsage.Emu) or "0"
            elif local_name == "extent":
                result.css_style["width"] = (
                    self.xml.length_attr(n, "cx", LengthUsage.Emu) or ""
                )
                result.css_style["height"] = (
                    self.xml.length_attr(n, "cy", LengthUsage.Emu) or ""
                )
            elif local_name in ["positionH", "positionV"] and not simple_pos:
                pos = pos_x if local_name == "positionH" else pos_y
                align_node = self.xml.element(n, "align")
                offset_node = self.xml.element(n, "posOffset")
                pos["relative"] = self.xml.attr(n, "relativeFrom") or pos["relative"]
                if align_node:
                    pos["align"] = align_node.text or pos["align"]
                if offset_node:
                    pos["offset"] = (
                        convert_length(
                            offset_node.text,
                            LengthUsage.Emu,
                        )
                        or pos["offset"]
                    )
            elif local_name == "wrapTopAndBottom":
                wrap_type = "wrapTopAndBottom"
            elif local_name == "wrapNone":
                wrap_type = "wrapNone"
            elif local_name == "graphic":
                g = self.parse_graphic(n)
                if g:
                    result.children.append(g)

        if wrap_type == "wrapTopAndBottom":
            result.css_style["display"] = "block"
            if pos_x.get("align"):
                result.css_style["text-align"] = pos_x["align"]
                result.css_style["width"] = "100%"
        elif wrap_type == "wrapNone":
            result.css_style["display"] = "block"
            result.css_style["position"] = "relative"
            result.css_style["width"] = "0px"
            result.css_style["height"] = "0px"
            if pos_x.get("offset"):
                result.css_style["left"] = pos_x["offset"]
            if pos_y.get("offset"):
                result.css_style["top"] = pos_y["offset"]
        elif is_anchor and pos_x.get("align") in ["left", "right"]:
            result.css_style["float"] = pos_x["align"]

        return result

    def parse_graphic(
        self,
        elem: "_Element",
    ) -> OpenXmlElement | None:
        """Parse graphic from XML element."""
        graphic_data = self.xml.element(elem, "graphicData")
        if graphic_data is None:
            return None

        for n in self.xml.elements(graphic_data):
            local_name = etree.QName(n).localname
            if local_name == "pic":
                return self.parse_picture(n)
        return None

    def parse_picture(self, elem: "_Element") -> IDomImage:
        """Parse picture from XML element."""
        result = IDomImage()
        blip_fill = self.xml.element(elem, "blipFill")
        blip = self.xml.element(blip_fill, "blip") if blip_fill is not None else None
        src_rect = (
            self.xml.element(blip_fill, "srcRect") if blip_fill is not None else None
        )

        result.src = self.xml.attr(blip, "embed") or "" if blip is not None else ""

        if src_rect is not None:
            result.src_rect = [
                (self.xml.int_attr(src_rect, "l", 0) or 0) / 100000,
                (self.xml.int_attr(src_rect, "t", 0) or 0) / 100000,
                (self.xml.int_attr(src_rect, "r", 0) or 0) / 100000,
                (self.xml.int_attr(src_rect, "b", 0) or 0) / 100000,
            ]

        sp_pr = self.xml.element(elem, "spPr")
        xfrm = self.xml.element(sp_pr, "xfrm") if sp_pr is not None else None

        result.css_style["position"] = "relative"

        if xfrm is not None:
            result.rotation = (self.xml.int_attr(xfrm, "rot", 0) or 0) / 60000

            for n in self.xml.elements(xfrm):
                local_name = etree.QName(n).localname
                if local_name == "ext":
                    result.css_style["width"] = (
                        self.xml.length_attr(
                            n,
                            "cx",
                            LengthUsage.Emu,
                        )
                        or ""
                    )
                    result.css_style["height"] = (
                        self.xml.length_attr(
                            n,
                            "cy",
                            LengthUsage.Emu,
                        )
                        or ""
                    )
                elif local_name == "off":
                    result.css_style["left"] = (
                        self.xml.length_attr(n, "x", LengthUsage.Emu) or ""
                    )
                    result.css_style["top"] = (
                        self.xml.length_attr(n, "y", LengthUsage.Emu) or ""
                    )

        return result

    def parse_table(self, node: "_Element") -> WmlTable:
        """Parse table from XML node."""
        result = WmlTable()

        for c in self.xml.elements(node):
            local_name = etree.QName(c).localname
            if local_name == "tr":
                result.children.append(self.parse_table_row(c))
            elif local_name == "tblGrid":
                result.columns = self.parse_table_columns(c)
            elif local_name == "tblPr":
                self.parse_table_properties(c, result)
        return result

    def parse_table_columns(
        self,
        node: "_Element",
    ) -> list[WmlTableColumn]:
        """Parse table columns from XML node."""
        result = []
        for n in self.xml.elements(node):
            local_name = etree.QName(n).localname
            if local_name == "gridCol":
                col = WmlTableColumn()
                col.width = self.xml.length_attr(n, "w")
                result.append(col)
        return result

    def parse_table_properties(
        self,
        elem: "_Element",
        table: WmlTable,
    ) -> None:
        """Parse table properties from XML element."""
        table.css_style = {}
        table.cell_style = {}

        def handler(c: "_Element") -> bool:  # noqa: PLR0911
            local_name = etree.QName(c).localname
            if local_name == "tblStyle":
                table.style_name = self.xml.attr(c, "val")
                return True
            if local_name == "tblLook":
                table.class_name = self._class_name_of_tbl_look(c)
                return True
            if local_name == "tblpPr":
                self.parse_table_position(c, table)
                return True
            if local_name == "tblStyleColBandSize":
                table.col_band_size = self.xml.int_attr(c, "val")
                return True
            if local_name == "tblStyleRowBandSize":
                table.row_band_size = self.xml.int_attr(c, "val")
                return True
            if local_name == "hidden":
                table.css_style["display"] = "none"
                return True
            return False

        self.parse_default_properties(elem, table.css_style, table.cell_style, handler)

        if table.css_style.get("text-align") == "center":
            del table.css_style["text-align"]
            table.css_style["margin-left"] = "auto"
            table.css_style["margin-right"] = "auto"
        elif table.css_style.get("text-align") == "right":
            del table.css_style["text-align"]
            table.css_style["margin-left"] = "auto"

    def parse_table_position(
        self,
        node: "_Element",
        table: WmlTable,
    ) -> None:
        """Parse table position from XML node."""
        top_from_text = self.xml.length_attr(node, "topFromText")
        bottom_from_text = self.xml.length_attr(node, "bottomFromText")
        right_from_text = self.xml.length_attr(node, "rightFromText")
        left_from_text = self.xml.length_attr(node, "leftFromText")

        table.css_style["float"] = "left"
        if bottom_from_text:
            table.css_style["margin-bottom"] = self._add_size(
                table.css_style.get("margin-bottom"),
                bottom_from_text,
            )
        if left_from_text:
            table.css_style["margin-left"] = self._add_size(
                table.css_style.get("margin-left"),
                left_from_text,
            )
        if right_from_text:
            table.css_style["margin-right"] = self._add_size(
                table.css_style.get("margin-right"),
                right_from_text,
            )
        if top_from_text:
            table.css_style["margin-top"] = self._add_size(
                table.css_style.get("margin-top"),
                top_from_text,
            )

    def parse_table_row(self, node: "_Element") -> WmlTableRow:
        """Parse table row from XML node."""
        result = WmlTableRow()

        for c in self.xml.elements(node):
            local_name = etree.QName(c).localname
            if local_name == "tc":
                result.children.append(self.parse_table_cell(c))
            elif local_name in ["trPr", "tblPrEx"]:
                self.parse_table_row_properties(c, result)
        return result

    def parse_table_row_properties(
        self,
        elem: "_Element",
        row: WmlTableRow,
    ) -> None:
        """Parse table row properties from XML element."""
        row.css_style = self.parse_default_properties(
            elem,
            {},
            None,
            lambda c: (
                self._handle_table_row_property(c, row)
                if etree.QName(c).localname == "cnfStyle"
                or etree.QName(c).localname == "tblHeader"
                or etree.QName(c).localname in ["gridBefore", "gridAfter"]
                else False
            ),
        )

    def _handle_table_row_property(
        self,
        c: "_Element",
        row: WmlTableRow,
    ) -> bool:
        """Handle table row property."""
        local_name = etree.QName(c).localname
        if local_name == "cnfStyle":
            row.class_name = self._class_name_of_cnf_style(c)
            return True
        if local_name == "tblHeader":
            row.is_header = self.xml.bool_attr(
                c,
                "val",
                default_value=False,
            )
            return True
        if local_name == "gridBefore":
            row.grid_before = self.xml.int_attr(c, "val")
            return True
        if local_name == "gridAfter":
            row.grid_after = self.xml.int_attr(c, "val")
            return True
        return False

    def parse_table_cell(self, node: "_Element") -> WmlTableCell:
        """Parse table cell from XML node."""
        result = WmlTableCell()

        for c in self.xml.elements(node):
            local_name = etree.QName(c).localname
            if local_name == "tbl":
                result.children.append(self.parse_table(c))
            elif local_name == "p":
                result.children.append(self.parse_paragraph(c))
            elif local_name == "tcPr":
                self.parse_table_cell_properties(c, result)
        return result

    def parse_table_cell_properties(
        self,
        elem: "_Element",
        cell: WmlTableCell,
    ) -> None:
        """Parse table cell properties from XML element."""

        def handler(c: "_Element") -> bool:
            local_name = etree.QName(c).localname
            if local_name == "gridSpan":
                cell.span = self.xml.int_attr(c, "val")
                return True
            if local_name == "vMerge":
                cell.vertical_merge = self.xml.attr(c, "val") or "continue"
                return True
            if local_name == "cnfStyle":
                cell.class_name = self._class_name_of_cnf_style(c)
                return True
            return False

        cell.css_style = self.parse_default_properties(elem, {}, None, handler)
        self.parse_table_cell_vertical_text(elem, cell)

    def parse_table_cell_vertical_text(
        self,
        elem: "_Element",
        cell: WmlTableCell,
    ) -> None:
        """Parse table cell vertical text from XML element."""
        direction_map = {
            "btLr": {
                "writingMode": "vertical-rl",
                "transform": "rotate(180deg)",
            },
            "lrTb": {
                "writingMode": "vertical-lr",
                "transform": "none",
            },
            "tbRl": {
                "writingMode": "vertical-rl",
                "transform": "none",
            },
        }

        for c in self.xml.elements(elem):
            local_name = etree.QName(c).localname
            if local_name == "textDirection":
                direction = self.xml.attr(c, "val")
                style = direction_map.get(
                    direction,
                    {"writingMode": "horizontal-tb"},
                )
                cell.css_style["writing-mode"] = style["writingMode"]
                cell.css_style["transform"] = style["transform"]

    def parse_default_properties(  # noqa: C901, PLR0912, PLR0915
        self,
        elem: "_Element",
        style: dict[str, str] | None = None,
        child_style: dict[str, str] | None = None,
        handler: Callable | None = None,
    ) -> dict[str, str]:
        """Parse default properties from XML element."""
        if style is None:
            style = {}

        for c in self.xml.elements(elem):
            if handler and handler(c):
                continue

            local_name = etree.QName(c).localname

            if local_name == "jc":
                style["text-align"] = self._value_of_jc(c)
            elif local_name == "textAlignment":
                style["vertical-align"] = self._value_of_text_alignment(c)
            elif local_name == "color":
                style["color"] = self._color_attr(c, "val", None, autos["color"])
            elif local_name == "sz":
                size = self.xml.length_attr(c, "val", LengthUsage.FontSize)
                if size:
                    style["font-size"] = size
                    style["min-height"] = size
            elif local_name == "shd":
                style["background-color"] = self._color_attr(
                    c,
                    "fill",
                    None,
                    autos["shd"],
                )
            elif local_name == "highlight":
                style["background-color"] = self._color_attr(
                    c,
                    "val",
                    None,
                    autos["highlight"],
                )
            elif local_name == "position":
                style["verticalAlign"] = self.xml.length_attr(
                    c,
                    "val",
                    LengthUsage.FontSize,
                )
            elif local_name == "tcW":
                if not self.options.ignore_width:
                    style["width"] = self._value_of_size(c, "w")
            elif local_name == "tblW":
                style["width"] = self._value_of_size(c, "w")
            elif local_name == "trHeight":
                self.parse_tr_height(c, style)
            elif local_name == "strike":
                style["text-decoration"] = (
                    "line-through"
                    if self.xml.bool_attr(c, "val", default_value=True)
                    else "none"
                )
            elif local_name == "b":
                style["font-weight"] = (
                    "bold"
                    if self.xml.bool_attr(c, "val", default_value=True)
                    else "normal"
                )
            elif local_name == "i":
                style["font-style"] = (
                    "italic"
                    if self.xml.bool_attr(c, "val", default_value=True)
                    else "normal"
                )
            elif local_name == "caps":
                style["text-transform"] = (
                    "uppercase"
                    if self.xml.bool_attr(c, "val", default_value=True)
                    else "none"
                )
            elif local_name == "smallCaps":
                style["font-variant"] = (
                    "small-caps"
                    if self.xml.bool_attr(c, "val", default_value=True)
                    else "none"
                )
            elif local_name == "u":
                self.parse_underline(c, style)
            elif local_name in ["ind", "tblInd"]:
                self.parse_indentation(c, style)
            elif local_name == "rFonts":
                self.parse_font(c, style)
            elif local_name == "tblBorders":
                self.parse_border_properties(c, child_style or style)
            elif local_name == "tblCellSpacing":
                style["border-spacing"] = self._value_of_margin(c)
                style["border-collapse"] = "separate"
            elif local_name == "pBdr":
                self.parse_border_properties(c, style)
            elif local_name == "bdr":
                style["border"] = self._value_of_border(c)
            elif local_name == "tcBorders":
                self.parse_border_properties(c, style)
            elif local_name == "vanish":
                if self.xml.bool_attr(c, "val", default_value=True):
                    style["display"] = "none"
            elif local_name in ["tblCellMar", "tcMar"]:
                self.parse_margin_properties(c, child_style or style)
            elif local_name == "tblLayout":
                style["table-layout"] = self._value_of_tbl_layout(c)
            elif local_name == "vAlign":
                style["vertical-align"] = self._value_of_text_alignment(c)
            elif local_name == "spacing":
                if etree.QName(elem).localname == "pPr":
                    self.parse_spacing(c, style)
            elif local_name == "wordWrap":
                if self.xml.bool_attr(c, "val", default_value=False):
                    style["overflow-wrap"] = "break-word"
            elif local_name == "suppressAutoHyphens":
                style["hyphens"] = (
                    "none"
                    if self.xml.bool_attr(c, "val", default_value=True)
                    else "auto"
                )
            elif local_name == "lang":
                style["$lang"] = self.xml.attr(c, "val")
            elif local_name in ["rtl", "bidi"]:
                if self.xml.bool_attr(c, "val", default_value=True):
                    style["direction"] = "rtl"
            elif self.options.debug:
                pass

        return style

    def parse_underline(
        self,
        node: "_Element",
        style: dict[str, str],
    ) -> None:
        """Parse underline from XML node."""
        val = self.xml.attr(node, "val")
        if val is None:
            return

        underline_styles = {
            "dash": "underline dashed",
            "dashDotDotHeavy": "underline dashed",
            "dashDotHeavy": "underline dashed",
            "dashedHeavy": "underline dashed",
            "dashLong": "underline dashed",
            "dashLongHeavy": "underline dashed",
            "dotDash": "underline dashed",
            "dotDotDash": "underline dashed",
            "dotted": "underline dotted",
            "dottedHeavy": "underline dotted",
            "double": "underline double",
            "single": "underline",
            "thick": "underline",
            "wave": "underline wavy",
            "wavyDouble": "underline wavy",
            "wavyHeavy": "underline wavy",
            "words": "underline",
            "none": "none",
        }

        style["text-decoration"] = underline_styles.get(val, "underline")

        col = self._color_attr(node, "color")
        if col:
            style["text-decoration-color"] = col

    def parse_font(
        self,
        node: "_Element",
        style: dict[str, str],
    ) -> None:
        """Parse font from XML node."""
        ascii_font = self.xml.attr(node, "ascii")
        ascii_theme = self._theme_value(node, "asciiTheme")
        east_asia = self.xml.attr(node, "eastAsia")
        fonts = [f for f in [ascii_font, ascii_theme, east_asia] if f]
        fonts = [enclose_font_family(f) for f in fonts]

        if fonts:
            unique_fonts = list(dict.fromkeys(fonts))
            style["font-family"] = ", ".join(unique_fonts)

    def parse_indentation(
        self,
        node: "_Element",
        style: dict[str, str],
    ) -> None:
        """Parse indentation from XML node."""
        first_line = self.xml.length_attr(node, "firstLine")
        hanging = self.xml.length_attr(node, "hanging")
        left = self.xml.length_attr(node, "left")
        start = self.xml.length_attr(node, "start")
        right = self.xml.length_attr(node, "right")
        end = self.xml.length_attr(node, "end")

        if first_line:
            style["text-indent"] = first_line
        if hanging:
            style["text-indent"] = f"-{hanging}"
        if left or start:
            style["margin-inline-start"] = left or start
        if right or end:
            style["margin-inline-end"] = right or end

    def parse_spacing(
        self,
        node: "_Element",
        style: dict[str, str],
    ) -> None:
        """Parse spacing from XML node."""
        before = self.xml.length_attr(node, "before")
        after = self.xml.length_attr(node, "after")
        line = self.xml.int_attr(node, "line")
        line_rule = self.xml.attr(node, "lineRule")

        if before:
            style["margin-top"] = before
        if after:
            style["margin-bottom"] = after

        if line is not None:
            if line_rule == "auto":
                style["line-height"] = f"{(line / 240):.2f}"
            elif line_rule == "atLeast":
                style["line-height"] = f"calc(100% + {line / 20}pt)"
            else:
                style["line-height"] = f"{line / 20}pt"
                style["min-height"] = f"{line / 20}pt"

    def parse_margin_properties(
        self,
        node: "_Element",
        output: dict[str, str],
    ) -> None:
        """Parse margin properties from XML node."""
        for c in self.xml.elements(node):
            local_name = etree.QName(c).localname
            if local_name == "left":
                output["padding-left"] = self._value_of_margin(c)
            elif local_name == "right":
                output["padding-right"] = self._value_of_margin(c)
            elif local_name == "top":
                output["padding-top"] = self._value_of_margin(c)
            elif local_name == "bottom":
                output["padding-bottom"] = self._value_of_margin(c)

    def parse_tr_height(
        self,
        node: "_Element",
        output: dict[str, str],
    ) -> None:
        """Parse table row height from XML node."""
        h_rule = self.xml.attr(node, "hRule")
        val = self.xml.length_attr(node, "val")
        if val:
            if h_rule == "exact":
                output["height"] = val
            else:
                output["height"] = val

    def parse_border_properties(
        self,
        node: "_Element",
        output: dict[str, str],
    ) -> None:
        """Parse border properties from XML node."""
        for c in self.xml.elements(node):
            local_name = etree.QName(c).localname
            if local_name in ["start", "left"]:
                output["border-left"] = self._value_of_border(c)
            elif local_name in ["end", "right"]:
                output["border-right"] = self._value_of_border(c)
            elif local_name == "top":
                output["border-top"] = self._value_of_border(c)
            elif local_name == "bottom":
                output["border-bottom"] = self._value_of_border(c)

    def _color_attr(
        self,
        node: "_Element",
        attr_name: str,
        def_value: str | None = None,
        auto_color: str = "black",
    ) -> str | None:
        """Get color attribute from XML node."""
        known_colors = [
            "black",
            "blue",
            "cyan",
            "darkBlue",
            "darkCyan",
            "darkGray",
            "darkGreen",
            "darkMagenta",
            "darkRed",
            "darkYellow",
            "green",
            "lightGray",
            "magenta",
            "none",
            "red",
            "white",
            "yellow",
        ]
        v = self.xml.attr(node, attr_name)
        if v:
            if v == "auto":
                return auto_color
            if v in known_colors:
                return v
            return f"#{v}"
        theme_color = self.xml.attr(node, "themeColor")
        return f"var(--docx-{theme_color}-color)" if theme_color else def_value

    def _theme_value(
        self,
        c: "_Element",
        attr: str,
    ) -> str | None:
        """Get theme value from XML element."""
        val = self.xml.attr(c, attr)
        return f"var(--docx-{val}-font)" if val else None

    def _value_of_size(
        self,
        c: "_Element",
        attr: str,
    ) -> str | None:
        """Get size value from XML element."""
        size_type = self.xml.attr(c, "type")
        if size_type == "pct":
            usage = LengthUsage.Percent
        elif size_type == "auto":
            return "auto"
        else:
            usage = LengthUsage.Dxa
        return self.xml.length_attr(c, attr, usage)

    def _value_of_margin(self, c: "_Element") -> str | None:
        """Get margin value from XML element."""
        return self.xml.length_attr(c, "w", LengthUsage.Dxa)

    def _value_of_border(self, c: "_Element") -> str:
        """Get border value from XML element."""
        border_type = self._parse_border_type(self.xml.attr(c, "val"))
        if border_type == "none":
            return "none"
        color = self._color_attr(c, "color")
        size = self.xml.length_attr(c, "sz", LengthUsage.Border) or "1pt"
        border_color = color if color != "auto" else autos["borderColor"]
        return f"{size} {border_type} {border_color}"

    def _parse_border_type(self, border_type: str | None) -> str:
        if not border_type:
            return "solid"
        border_map = {
            "single": "solid",
            "dashDotStroked": "solid",
            "dashed": "dashed",
            "dashSmallGap": "dashed",
            "dotDash": "dotted",
            "dotDotDash": "dotted",
            "dotted": "dotted",
            "double": "double",
            "doubleWave": "double",
            "inset": "inset",
            "nil": "none",
            "none": "none",
            "outset": "outset",
            "thick": "solid",
            "thickThinLargeGap": "solid",
            "thickThinMediumGap": "solid",
            "thickThinSmallGap": "solid",
            "thinThickLargeGap": "solid",
            "thinThickMediumGap": "solid",
            "thinThickSmallGap": "solid",
            "thinThickThinLargeGap": "solid",
            "thinThickThinMediumGap": "solid",
            "thinThickThinSmallGap": "solid",
            "threeDEmboss": "solid",
            "threeDEngrave": "solid",
            "triple": "double",
            "wave": "solid",
        }
        return border_map.get(border_type, "solid")

    def _value_of_tbl_layout(self, c: "_Element") -> str:
        """Get table layout value from XML element."""
        layout_type = self.xml.attr(c, "val")
        return "fixed" if layout_type == "fixed" else "auto"

    def _class_name_of_cnf_style(self, c: "_Element") -> str:
        """Get class name from CNF style XML element."""
        val = self.xml.attr(c, "val")
        if not val:
            return ""
        classes = [
            "first-row",
            "last-row",
            "first-col",
            "last-col",
            "odd-col",
            "even-col",
            "odd-row",
            "even-row",
            "ne-cell",
            "nw-cell",
            "se-cell",
            "sw-cell",
        ]
        result = []
        for i, cls in enumerate(classes):
            if i < len(val) and val[i] == "1":
                result.append(cls)
        return " ".join(result)

    def _value_of_jc(self, c: "_Element") -> str:
        """Get justification value from XML element."""
        jc_type = self.xml.attr(c, "val")
        jc_map = {
            "start": "left",
            "left": "left",
            "center": "center",
            "end": "right",
            "right": "right",
            "both": "justify",
        }
        return jc_map.get(jc_type, jc_type or "left")

    def _value_of_vert_align(
        self,
        c: "_Element",
        *,
        as_tag_name: bool = False,
    ) -> str | None:
        """Get vertical alignment value from XML element."""
        align_type = self.xml.attr(c, "val")
        if align_type == "subscript":
            return "sub"
        if align_type == "superscript":
            return "sup" if as_tag_name else "super"
        return "sup" if as_tag_name and align_type else None

    def _value_of_text_alignment(self, c: "_Element") -> str:
        """Get text alignment value from XML element."""
        align_type = self.xml.attr(c, "val")
        align_map = {
            "auto": "baseline",
            "baseline": "baseline",
            "top": "top",
            "center": "middle",
            "bottom": "bottom",
        }
        return align_map.get(align_type, align_type or "baseline")

    def _add_size(self, a: str | None, b: str | None) -> str:
        """Add two size values together."""
        if a is None:
            return b or ""
        if b is None:
            return a
        return f"calc({a} + {b})"

    def _class_name_of_tbl_look(self, c: "_Element") -> str:
        """Get class name from table look XML element."""
        val = self.xml.hex_attr(c, "val", 0)
        class_name = ""
        if self.xml.bool_attr(c, "firstRow", default_value=False) or (val & 0x0020):
            class_name += " first-row"
        if self.xml.bool_attr(c, "lastRow", default_value=False) or (val & 0x0040):
            class_name += " last-row"
        if self.xml.bool_attr(
            c,
            "firstColumn",
            default_value=False,
        ) or (val & 0x0080):
            class_name += " first-col"
        if self.xml.bool_attr(
            c,
            "lastColumn",
            default_value=False,
        ) or (val & 0x0100):
            class_name += " last-col"
        if self.xml.bool_attr(c, "noHBand", default_value=False) or (val & 0x0200):
            class_name += " no-hband"
        if self.xml.bool_attr(c, "noVBand", default_value=False) or (val & 0x0400):
            class_name += " no-vband"
        return class_name.strip()
