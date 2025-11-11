"""HTML renderer module for rendering Word documents to HTML."""

import re
from typing import TYPE_CHECKING, Any

from bs4 import BeautifulSoup, Comment, NavigableString, Tag

from docxjs.document.bookmarks import WmlBookmarkStart
from docxjs.document.document import DocumentElement
from docxjs.document.dom import (
    DomType,
    IDomImage,
    IDomNumbering,
    OpenXmlElement,
    WmlAltChunk,
    WmlBreak,
    WmlHyperlink,
    WmlNoteReference,
    WmlSymbol,
    WmlTable,
    WmlTableCell,
    WmlTableColumn,
    WmlTableRow,
    WmlText,
)
from docxjs.document.paragraph import WmlParagraph
from docxjs.document.run import WmlRun
from docxjs.document.section import FooterHeaderReference, SectionProperties
from docxjs.document.style import IDomStyle, IDomSubStyle
from docxjs.header_footer.parts import BaseHeaderFooterPart
from docxjs.utils import (
    enclose_font_family,
    escape_class_name,
    key_by,
    merge_deep,
)
from docxjs.vml.vml import VmlElement
from docxjs.word_document import WordDocument

if TYPE_CHECKING:
    from docxjs.font_table.font_table import FontTablePart
    from docxjs.part import Part
    from docxjs.theme.theme_part import ThemePart

ns = {
    "svg": "http://www.w3.org/2000/svg",
    "mathML": "http://www.w3.org/1998/Math/MathML",
}


class HtmlRenderer:
    """HTML renderer for Word documents."""

    def __init__(self, html_document: BeautifulSoup) -> None:
        """Initialize HTML renderer."""
        self.html_document = html_document
        self.class_name = "docx"
        self.root_selector = ""
        self.document: WordDocument | None = None
        self.options = None
        self.style_map: dict[str, IDomStyle] | None = None
        self.current_part: Part | None = None

        self.table_vertical_merges: list[dict[int, Tag]] = []
        self.current_vertical_merge: dict[int, Tag] = {}
        self.table_cell_positions: list[dict[str, int]] = []
        self.current_cell_position: dict[str, int] = {"col": 0, "row": 0}

        self.footnote_map: dict[str, Any] = {}
        self.endnote_map: dict[str, Any] = {}
        self.current_footnote_ids: list[str] = []
        self.current_endnote_ids: list[str] = []
        self.used_header_footer_parts: list[str] = []

        self.default_tab_size: str | None = None
        self.current_tabs: list[dict] = []

        self.comment_map: dict[str, Tag] = {}

        self.tasks: list[Any] = []
        self.post_render_tasks: list[Any] = []

    def render(  # noqa: C901, PLR0912
        self,
        document: WordDocument,
        body_container: Tag,
        style_container: Tag | None = None,
        options: dict[str, Any] | None = None,
    ) -> None:
        """Render document to HTML."""
        if options is None:
            options = {}
        self.document = document
        self.options = options
        self.class_name = options.get("class_name", "docx")
        in_wrapper = options.get("in_wrapper", True)
        self.root_selector = (
            f".{self.class_name}-wrapper" if in_wrapper else ":root"
        )
        self.style_map = None
        self.tasks = []

        style_container = style_container or body_container

        if hasattr(style_container, "clear"):
            style_container.clear()
        if hasattr(body_container, "clear"):
            body_container.clear()

        style_container.append(self.create_comment("docxjs library predefined styles"))
        style_container.append(self.render_default_style())

        if document.theme_part:
            style_container.append(self.create_comment("docxjs document theme values"))
            self.render_theme(document.theme_part, style_container)

        if document.styles_part:
            self.style_map = self.process_styles(document.styles_part.styles)
            style_container.append(self.create_comment("docxjs document styles"))
            style_container.append(self.render_styles(document.styles_part.styles))

        if document.numbering_part:
            self.process_numberings(document.numbering_part.dom_numberings)
            style_container.append(
                self.create_comment("docxjs document numbering styles"),
            )
            style_container.append(
                self.render_numbering(
                    document.numbering_part.dom_numberings, style_container,
                ),
            )

        if document.footnotes_part:
            self.footnote_map = key_by(document.footnotes_part.notes, lambda x: x.id)

        if document.endnotes_part:
            self.endnote_map = key_by(document.endnotes_part.notes, lambda x: x.id)

        if document.settings_part and document.settings_part.settings:
            self.default_tab_size = document.settings_part.settings.get(
                "defaultTabStop",
            )

        if not options.get("ignore_fonts", False) and document.font_table_part:
            self.render_font_table(document.font_table_part, style_container)

        section_elements = self.render_sections(document.document_part.body)

        if options.get("in_wrapper", True):
            body_container.append(self.render_wrapper(section_elements))
        else:
            for elem in section_elements:
                body_container.append(elem)

        for task in self.post_render_tasks:
            task()

        self.refresh_tab_stops()

    def render_theme(
        self, theme_part: "ThemePart", style_container: Tag,
    ) -> None:
        """Render theme to style container."""
        variables = {}
        if theme_part.theme:
            font_scheme = getattr(theme_part.theme, "font_scheme", None)
            if font_scheme:
                if hasattr(font_scheme, "major_font") and font_scheme.major_font:
                    variables["--docx-majorHAnsi-font"] = getattr(
                        font_scheme.major_font, "latin_typeface", "",
                    )
                if hasattr(font_scheme, "minor_font") and font_scheme.minor_font:
                    variables["--docx-minorHAnsi-font"] = getattr(
                        font_scheme.minor_font, "latin_typeface", "",
                    )

            color_scheme = getattr(theme_part.theme, "color_scheme", None)
            if color_scheme:
                colors = getattr(color_scheme, "colors", {})
                for k, v in colors.items():
                    variables[f"--docx-{k}-color"] = f"#{v}"

        css_text = self.style_to_string(f".{self.class_name}", variables)
        style_container.append(self.create_style_element(css_text))

    def render_font_table(
        self, fonts_part: "FontTablePart", style_container: Tag,
    ) -> None:
        """Render font table to style container."""
        for f in fonts_part.fonts:
            for ref in getattr(f, "embed_font_refs", []):
                font_data = self.document.load_font(
                    ref.get("id", ""), ref.get("key", ""),
                )
                if font_data:
                    css_values = {
                        "font-family": enclose_font_family(f.name),
                        "src": f"url({font_data})",
                    }
                    ref_type = ref.get("type", "")
                    if ref_type in ["bold", "boldItalic"]:
                        css_values["font-weight"] = "bold"
                    if ref_type in ["italic", "boldItalic"]:
                        css_values["font-style"] = "italic"
                    css_text = self.style_to_string("@font-face", css_values)
                    style_container.append(
                        self.create_comment(f"docxjs {f.name} font"),
                    )
                    style_container.append(
                        self.create_style_element(css_text),
                    )

    def process_style_name(self, class_name: str | None) -> str:
        """Process style name."""
        return (
            f"{self.class_name}_{escape_class_name(class_name)}"
            if class_name
            else self.class_name
        )

    def process_styles(self, styles: list[IDomStyle]) -> dict[str, IDomStyle]:
        """Process styles."""
        styles_map = key_by([s for s in styles if s.id], lambda x: x.id)

        for style in [s for s in styles if s.based_on]:
            base_style = styles_map.get(style.based_on)
            if base_style:
                if style.paragraph_props and base_style.paragraph_props:
                    style.paragraph_props = merge_deep(
                        style.paragraph_props, base_style.paragraph_props,
                    )
                if style.run_props and base_style.run_props:
                    style.run_props = merge_deep(
                        style.run_props, base_style.run_props,
                    )

                for base_values in base_style.styles:
                    style_values = next(
                        (
                            x
                            for x in style.styles
                            if x.target == base_values.target
                        ),
                        None,
                    )
                    if style_values:
                        self.copy_style_properties(
                            base_values.values, style_values.values,
                        )
                    else:
                        new_style = IDomSubStyle()
                        new_style.target = base_values.target
                        new_style.values = {**base_values.values}
                        style.styles.append(new_style)
            elif self.options.get("debug", False):
                pass

        for style in styles:
            style.css_name = self.process_style_name(style.id)

        return styles_map

    def process_numberings(
        self, numberings: list[IDomNumbering],
    ) -> None:
        """Process numberings."""
        for num in [n for n in numberings if n.p_style_name]:
            style = self.find_style(num.p_style_name)
            if (
                style
                and style.paragraph_props
                and style.paragraph_props.numbering
            ):
                style.paragraph_props.numbering.level = num.level

    def process_element(self, element: OpenXmlElement) -> None:
        """Process element."""
        if element.children:
            for e in element.children:
                e.parent = element
                if e.type == DomType.Table:
                    self.process_table(e)
                else:
                    self.process_element(e)

    def process_table(self, table: WmlTable) -> None:
        """Process table element."""
        for r in table.children:
            for c in r.children:
                if table.cell_style:
                    c.css_style = self.copy_style_properties(
                        table.cell_style,
                        c.css_style or {},
                        [
                            "border-left",
                            "border-right",
                            "border-top",
                            "border-bottom",
                            "padding-left",
                            "padding-right",
                            "padding-top",
                            "padding-bottom",
                        ],
                    )
                self.process_element(c)

    def copy_style_properties(
        self,
        input_style: dict[str, str],
        output_style: dict[str, str],
        attrs: list[str] | None = None,
    ) -> dict[str, str]:
        """Copy style properties from input to output."""
        if not input_style:
            return output_style or {}
        if output_style is None:
            output_style = {}
        if attrs is None:
            attrs = list(input_style.keys())

        for key in attrs:
            if key in input_style and key not in output_style:
                output_style[key] = input_style[key]

        return output_style

    def create_page_element(
        self, class_name: str, props: SectionProperties,
    ) -> Tag:
        """Create page element."""
        elem = self.create_element("section", {"class": class_name})

        if props and props.page_margins:
            if props.page_margins.left:
                style = elem.get("style", "")
                elem["style"] = (
                    f"{style}padding-left: {props.page_margins.left}; "
                ).strip()
            if props.page_margins.right:
                style = elem.get("style", "")
                elem["style"] = (
                    f"{style}padding-right: {props.page_margins.right}; "
                ).strip()
            if props.page_margins.top:
                style = elem.get("style", "")
                elem["style"] = (
                    f"{style}padding-top: {props.page_margins.top}; "
                ).strip()
            if props.page_margins.bottom:
                style = elem.get("style", "")
                elem["style"] = (
                    f"{style}padding-bottom: {props.page_margins.bottom}; "
                ).strip()

        if props and props.page_size:
            if (
                not self.options.get("ignore_width", False)
                and props.page_size.width
            ):
                style = elem.get("style", "")
                elem["style"] = (
                    f"{style}width: {props.page_size.width}; "
                ).strip()
            if (
                not self.options.get("ignore_height", False)
                and props.page_size.height
            ):
                style = elem.get("style", "")
                elem["style"] = (
                    f"{style}min-height: {props.page_size.height}; "
                ).strip()

        return elem

    def create_section_content(
        self, props: SectionProperties,
    ) -> Tag:
        """Create section content element."""
        elem = self.create_element("article")

        if props and props.columns and props.columns.number_of_columns:
            style = f"column-count: {props.columns.number_of_columns}; "
            if props.columns.space:
                style += f"column-gap: {props.columns.space}; "
            if props.columns.separator:
                style += "column-rule: 1px solid black; "
            elem["style"] = style.strip()

        return elem

    def render_sections(
        self, document: DocumentElement,
    ) -> list[Tag]:
        """Render document sections."""
        self.process_element(document)
        sections = self.split_by_section(document.children, document.props)
        pages = self.group_by_page_breaks(sections)
        prev_props = None
        result = []

        for i, page_sections in enumerate(pages):
            self.current_footnote_ids = []

            section = page_sections[0]
            props = section["sectProps"]
            page_element = self.create_page_element(self.class_name, props)
            self.render_style_values(document.css_style or {}, page_element)

            if self.options.get("render_headers", True):
                self.render_header_footer(
                    props.header_refs or [],
                    props,
                    len(result),
                    first_of_section=prev_props != props,
                    into=page_element,
                )

            for sect in page_sections:
                content_element = self.create_section_content(
                    sect["sectProps"],
                )
                self.render_elements(sect["elements"], content_element)
                page_element.append(content_element)
                props = sect["sectProps"]

            if self.options.get("render_footnotes", True):
                self.render_notes(
                    self.current_footnote_ids,
                    self.footnote_map,
                    page_element,
                )

            if (
                self.options.get("render_endnotes", True)
                and i == len(pages) - 1
            ):
                self.render_notes(
                    self.current_endnote_ids,
                    self.endnote_map,
                    page_element,
                )

            if self.options.get("render_footers", True):
                self.render_header_footer(
                    props.footer_refs or [],
                    props,
                    len(result),
                    first_of_section=prev_props != props,
                    into=page_element,
                )

            result.append(page_element)
            prev_props = props

        return result

    def render_header_footer(
        self,
        refs: list[FooterHeaderReference],
        props: SectionProperties,
        page: int,
        *,
        first_of_section: bool,
        into: Tag,
    ) -> None:
        """Render header or footer."""
        if not refs:
            return

        ref = None
        if props.title_page and first_of_section:
            ref = next((x for x in refs if x.type == "first"), None)
        if not ref and page % 2 == 1:
            ref = next((x for x in refs if x.type == "even"), None)
        if not ref:
            ref = next((x for x in refs if x.type == "default"), None)

        if ref:
            part = self.document.find_part_by_rel_id(
                ref.id, self.document.document_part,
            )
            if part and isinstance(part, BaseHeaderFooterPart):
                self.current_part = part
                if part.path not in self.used_header_footer_parts:
                    self.process_element(part.root_element)
                    self.used_header_footer_parts.append(part.path)
                self.render_elements([part.root_element], into)[0]
                self.current_part = None

    def split_by_section(  # noqa: C901
        self,
        elements: list[OpenXmlElement],
        default_props: SectionProperties,
    ) -> list[dict]:
        """Split elements by section."""
        current = {"sectProps": None, "elements": [], "pageBreak": False}
        result = [current]

        for elem in elements:
            if elem.type == DomType.Paragraph:
                p = elem
                style = self.find_style(p.style_name) if p.style_name else None
                if (
                    style
                    and style.paragraph_props
                    and style.paragraph_props.page_break_before
                ):
                    current["sectProps"] = default_props
                    current["pageBreak"] = True
                    current = {"sectProps": None, "elements": [], "pageBreak": False}
                    result.append(current)

            current["elements"].append(elem)

            if elem.type == DomType.Paragraph:
                p = elem
                sect_props = getattr(p, "section_props", None)
                p_break_index = -1
                r_break_index = -1

                if self.options.get("break_pages", True) and p.children:
                    for i, r in enumerate(p.children):
                        r_break_index = next(
                            (
                                j
                                for j, child in enumerate(r.children or [])
                                if self.is_page_break_element(child)
                            ),
                            -1,
                        )
                        if r_break_index != -1:
                            p_break_index = i
                            break

                if sect_props or p_break_index != -1:
                    current["sectProps"] = sect_props
                    current["pageBreak"] = p_break_index != -1
                    current = {"sectProps": None, "elements": [], "pageBreak": False}
                    result.append(current)

        current_sect_props = None
        for i in range(len(result) - 1, -1, -1):
            if result[i]["sectProps"] is None:
                result[i]["sectProps"] = current_sect_props or default_props
            else:
                current_sect_props = result[i]["sectProps"]

        return result

    def group_by_page_breaks(
        self, sections: list[dict],
    ) -> list[list[dict]]:
        """Group sections by page breaks."""
        current = []
        prev = None
        result = [current]

        for s in sections:
            current.append(s)
            ignore_break = self.options.get(
                "ignore_last_rendered_page_break", True,
            )
            if (
                ignore_break
                or s["pageBreak"]
                or self.is_page_break_section(prev, s["sectProps"])
            ):
                result.append(current := [])
            prev = s["sectProps"]

        return [x for x in result if x]

    def is_page_break_element(self, elem: OpenXmlElement) -> bool:
        """Check if element is a page break."""
        if elem.type != DomType.Break:
            return False
        break_elem = elem
        if break_elem.break_type == "lastRenderedPageBreak":
            return not self.options.get(
                "ignore_last_rendered_page_break", True,
            )
        return break_elem.break_type == "page"

    def is_page_break_section(
        self,
        prev: SectionProperties | None,
        next_props: SectionProperties | None,
    ) -> bool:
        """Check if section should break page."""
        if not prev or not next_props:
            return False
        return (
            prev.page_size
            and next_props.page_size
            and (
                prev.page_size.orientation
                != next_props.page_size.orientation
                or prev.page_size.width != next_props.page_size.width
                or prev.page_size.height != next_props.page_size.height
            )
        )

    def render_wrapper(self, children: list[Tag]) -> Tag:
        """Render wrapper element."""
        wrapper = self.create_element(
            "div", {"class": f"{self.class_name}-wrapper"},
        )
        for child in children:
            wrapper.append(child)
        return wrapper

    def render_default_style(self) -> Tag:
        """Render default styles."""
        c = self.class_name
        wrapper_style = f"""
.{c}-wrapper {{ background: gray; padding: 30px; padding-bottom: 0px;
    display: flex; flex-flow: column; align-items: center; }}
.{c}-wrapper>section.{c} {{ background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5);
    margin-bottom: 30px; }}"""
        if self.options.get("hide_wrapper_on_print", False):
            wrapper_style = f"@media not print {{ {wrapper_style} }}"

        style_text = f"""{wrapper_style}
.{c} {{ color: black; hyphens: auto; text-underline-position: from-font; }}
section.{c} {{ box-sizing: border-box; display: flex; flex-flow: column nowrap;
    position: relative; overflow: hidden; }}
section.{c}>article {{ margin-bottom: auto; z-index: 1; }}
section.{c}>footer {{ z-index: 1; }}
.{c} table {{ border-collapse: collapse; }}
.{c} table td, .{c} table th {{ vertical-align: top; }}
.{c} p {{ margin: 0pt; min-height: 1em; }}
.{c} span {{ white-space: pre-wrap; overflow-wrap: break-word; }}
.{c} a {{ color: inherit; text-decoration: inherit; }}
.{c} svg {{ fill: transparent; }}"""

        return self.create_style_element(style_text)

    def render_numbering(
        self, numberings: list[IDomNumbering], style_container: Tag,
    ) -> Tag:
        """Render numbering styles."""
        style_text = ""
        reset_counters = []

        for num in numberings:
            selector = f"p.{self.numbering_class(num.id, num.level)}"
            list_style_type = "none"

            if num.bullet:
                variable = f"--{self.class_name}-{num.bullet.src}".lower()
                style_text += self.style_to_string(
                    f"{selector}:before",
                    {
                        "content": "' '",
                        "display": "inline-block",
                        "background": f"var({variable})",
                    },
                    num.bullet.style or "",
                )

                image_data = self.document.load_numbering_image(
                    num.bullet.src,
                )
                if image_data:
                    text = (
                        f"{self.root_selector} {{ {variable}: "
                        f"url({image_data}) }}"
                    )
                    style_container.append(
                        self.create_style_element(text),
                    )
            elif num.level_text:
                counter = self.numbering_counter(num.id, num.level)
                counter_reset = f"{counter} {num.start - 1}"
                if num.level > 0:
                    style_text += self.style_to_string(
                        f"p.{self.numbering_class(num.id, num.level - 1)}",
                        {
                            "counter-set": counter_reset,
                        },
                    )
                reset_counters.append(counter_reset)

                style_text += self.style_to_string(
                    f"{selector}:before",
                    {
                        "content": self.level_text_to_content(
                            num.level_text,
                            num.suff,
                            num.id,
                            self.num_format_to_css_value(num.format or ""),
                        ),
                        "counter-increment": counter,
                        **num.r_style,
                    },
                )
            else:
                list_style_type = self.num_format_to_css_value(num.format or "")

            style_text += self.style_to_string(selector, {
                "display": "list-item",
                "list-style-position": "inside",
                "list-style-type": list_style_type,
                **num.p_style,
            })

        if reset_counters:
            style_text += self.style_to_string(self.root_selector, {
                "counter-reset": " ".join(reset_counters),
            })

        return self.create_style_element(style_text)

    def render_styles(self, styles: list[IDomStyle]) -> Tag:
        """Render styles."""
        style_text = ""
        styles_map = self.style_map
        default_styles = key_by(
            [s for s in styles if s.is_default], lambda s: s.target,
        )

        for style in styles:
            sub_styles = style.styles

            if style.linked:
                linked_style = (
                    styles_map.get(style.linked) if style.linked else None
                )
                if linked_style:
                    sub_styles = sub_styles + linked_style.styles
                elif self.options.get("debug", False):
                    pass

            for sub_style in sub_styles:
                selector = f"{style.target or ''}.{style.css_name}"
                if style.target != sub_style.target:
                    selector += f" {sub_style.target}"
                if default_styles.get(style.target) == style:
                    selector = (
                        f".{self.class_name} {style.target}, {selector}"
                    )

                style_text += self.style_to_string(
                    selector, sub_style.values,
                )

        return self.create_style_element(style_text)

    def render_notes(
        self, note_ids: list[str], notes_map: dict[str, Any], into: Tag,
    ) -> None:
        """Render notes."""
        notes = [
            notes_map.get(id_val)
            for id_val in note_ids
            if notes_map.get(id_val)
        ]
        if notes:
            result = self.create_element("ol")
            self.render_elements(notes, result)
            into.append(result)

    def render_element(  # noqa: C901, PLR0911, PLR0912
        self, elem: OpenXmlElement,
    ) -> Tag | NavigableString | None:
        """Render element to HTML."""
        if elem.type == DomType.Paragraph:
            return self.render_paragraph(elem)
        if elem.type == DomType.BookmarkStart:
            return self.render_bookmark_start(elem)
        if elem.type == DomType.BookmarkEnd:
            return None
        if elem.type == DomType.Run:
            return self.render_run(elem)
        if elem.type == DomType.Table:
            return self.render_table(elem)
        if elem.type == DomType.Row:
            return self.render_table_row(elem)
        if elem.type == DomType.Cell:
            return self.render_table_cell(elem)
        if elem.type == DomType.Hyperlink:
            return self.render_hyperlink(elem)
        if elem.type == DomType.SmartTag:
            return self.render_smart_tag(elem)
        if elem.type == DomType.Drawing:
            return self.render_drawing(elem)
        if elem.type == DomType.Image:
            return self.render_image(elem)
        if elem.type == DomType.Text:
            return self.render_text(elem)
        if elem.type == DomType.DeletedText:
            return self.render_deleted_text(elem)
        if elem.type == DomType.Tab:
            return self.render_tab(elem)
        if elem.type == DomType.Symbol:
            return self.render_symbol(elem)
        if elem.type == DomType.Break:
            return self.render_break(elem)
        if elem.type == DomType.Footer:
            return self.render_container(elem, "footer")
        if elem.type == DomType.Header:
            return self.render_container(elem, "header")
        if elem.type in [DomType.Footnote, DomType.Endnote]:
            return self.render_container(elem, "li")
        if elem.type == DomType.FootnoteReference:
            return self.render_footnote_reference(elem)
        if elem.type == DomType.EndnoteReference:
            return self.render_endnote_reference(elem)
        if elem.type == DomType.NoBreakHyphen:
            return self.create_element("wbr")
        if elem.type == DomType.VmlPicture:
            return self.render_vml_picture(elem)
        if elem.type == DomType.VmlElement:
            return self.render_vml_element(elem)
        if elem.type == DomType.MmlMath:
            return self.render_container_ns(
                elem, ns["mathML"], "math", {"xmlns": ns["mathML"]},
            )
        if elem.type == DomType.MmlMathParagraph:
            return self.render_container(elem, "span")
        if elem.type == DomType.MmlFraction:
            return self.render_container_ns(elem, ns["mathML"], "mfrac")
        if elem.type == DomType.MmlBase:
            parent_type = getattr(elem.parent, "type", None) if elem.parent else None
            tag_name = "mtd" if parent_type == DomType.MmlMatrixRow else "mrow"
            return self.render_container_ns(elem, ns["mathML"], tag_name)
        if elem.type in [
            DomType.MmlNumerator,
            DomType.MmlDenominator,
            DomType.MmlFunction,
            DomType.MmlLimit,
            DomType.MmlBox,
        ]:
            return self.render_container_ns(elem, ns["mathML"], "mrow")
        if elem.type == DomType.MmlGroupChar:
            return self.render_mml_group_char(elem)
        if elem.type == DomType.MmlLimitLower:
            return self.render_container_ns(elem, ns["mathML"], "munder")
        if elem.type == DomType.MmlMatrix:
            return self.render_container_ns(elem, ns["mathML"], "mtable")
        if elem.type == DomType.MmlMatrixRow:
            return self.render_container_ns(elem, ns["mathML"], "mtr")
        if elem.type == DomType.MmlRadical:
            return self.render_mml_radical(elem)
        if elem.type == DomType.MmlSuperscript:
            return self.render_container_ns(elem, ns["mathML"], "msup")
        if elem.type == DomType.MmlSubscript:
            return self.render_container_ns(elem, ns["mathML"], "msub")
        if elem.type in [
            DomType.MmlDegree,
            DomType.MmlSuperArgument,
            DomType.MmlSubArgument,
        ]:
            return self.render_container_ns(elem, ns["mathML"], "mn")
        if elem.type == DomType.MmlFunctionName:
            return self.render_container_ns(elem, ns["mathML"], "ms")
        if elem.type == DomType.MmlDelimiter:
            return self.render_mml_delimiter(elem)
        if elem.type == DomType.MmlRun:
            return self.render_mml_run(elem)
        if elem.type == DomType.MmlNary:
            return self.render_mml_nary(elem)
        if elem.type == DomType.MmlPreSubSuper:
            return self.render_mml_pre_sub_super(elem)
        if elem.type == DomType.MmlBar:
            return self.render_mml_bar(elem)
        if elem.type == DomType.MmlEquationArray:
            return self.render_mml_list(elem)
        if elem.type == DomType.Inserted:
            return self.render_inserted(elem)
        if elem.type == DomType.Deleted:
            return self.render_deleted(elem)
        if elem.type == DomType.CommentRangeStart:
            return self.render_comment_range_start(elem)
        if elem.type == DomType.CommentRangeEnd:
            return self.render_comment_range_end(elem)
        if elem.type == DomType.CommentReference:
            return self.render_comment_reference(elem)
        if elem.type == DomType.AltChunk:
            return self.render_alt_chunk(elem)
        return None

    def render_elements(
        self, elems: list[OpenXmlElement], into: Tag | None = None,
    ) -> list[Tag]:
        """Render elements to HTML."""
        if elems is None:
            return []

        result = []
        for e in elems:
            rendered = self.render_element(e)
            if rendered:
                if isinstance(rendered, list):
                    result.extend(rendered)
                else:
                    result.append(rendered)

        if into:
            for r in result:
                into.append(r)

        return result

    def render_container(
        self,
        elem: OpenXmlElement,
        tag_name: str,
        props: dict | None = None,
    ) -> Tag:
        """Render container element."""
        result = self.create_element(tag_name, props or {})
        self.render_elements(elem.children or [], result)
        return result

    def render_container_ns(
        self,
        elem: OpenXmlElement,
        namespace: str,
        tag_name: str,
        props: dict | None = None,
    ) -> Tag:
        """Render container element with namespace."""
        result = self.create_element_ns(namespace, tag_name, props or {})
        self.render_elements(elem.children or [], result)
        return result

    def render_paragraph(self, elem: WmlParagraph) -> Tag:
        """Render paragraph element."""
        result = self.render_container(elem, "p")

        style = self.find_style(elem.style_name)
        if style and style.paragraph_props:
            elem.tabs = (
                getattr(elem, "tabs", None) or style.paragraph_props.tabs
            )

        self.render_class(elem, result)
        self.render_style_values(elem.css_style or {}, result)
        self.render_common_properties(result, elem)

        numbering = (
            getattr(elem, "numbering", None)
            or (
                style.paragraph_props.numbering
                if style and style.paragraph_props
                else None
            )
        )
        if numbering:
            class_name = self.numbering_class(
                numbering.id, numbering.level,
            )
            result["class"] = (
                result.get("class", "") + f" {class_name}"
            ).strip()

        return result

    def render_hyperlink(self, elem: WmlHyperlink) -> Tag:
        """Render hyperlink element."""
        result = self.render_container(elem, "a")
        self.render_style_values(elem.css_style or {}, result)

        href = ""
        if elem.id:
            rels = self.document.document_part.rels or []
            rel = next(
                (
                    it
                    for it in rels
                    if it.id == elem.id and it.target_mode == "External"
                ),
                None,
            )
            href = rel.target if rel else href

        if elem.anchor:
            href += f"#{elem.anchor}"

        result["href"] = href
        return result

    def render_smart_tag(
        self, elem: OpenXmlElement,
    ) -> Tag:
        """Render smart tag element."""
        return self.render_container(elem, "span")

    def render_alt_chunk(self, elem: WmlAltChunk) -> Tag | None:
        """Render alt chunk element."""
        if not self.options.get("render_alt_chunks", True):
            return None

        result = self.create_element("iframe")
        alt_chunk_data = self.document.load_alt_chunk(elem.id, self.current_part)
        if alt_chunk_data:
            result["srcdoc"] = alt_chunk_data
        return result

    def render_drawing(self, elem: OpenXmlElement) -> Tag:
        """Render drawing element."""
        result = self.render_container(elem, "div")
        base_style = (
            "display: inline-block; position: relative; "
            "text-indent: 0px; "
        )
        result["style"] = (
            base_style + (result.get("style", "") or "")
        )
        self.render_style_values(elem.css_style or {}, result)
        return result

    def render_image(self, elem: IDomImage) -> Tag:
        """Render image element."""
        result = self.create_element("img")
        transform = (
            elem.css_style.get("transform") if elem.css_style else None
        )

        self.render_style_values(elem.css_style or {}, result)

        if elem.src_rect and any(x != 0 for x in elem.src_rect):
            left, top, right, bottom = elem.src_rect
            transform = (
                f"scale({1 / (1 - left - right)}, "
                f"{1 / (1 - top - bottom)})"
            )
            clip_path = (
                f"clip-path: rect({100 * top:.2f}% "
                f"{100 * (1 - right):.2f}% "
                f"{100 * (1 - bottom):.2f}% "
                f"{100 * left:.2f}%);"
            )
            result["style"] = (
                result.get("style", "") + f" {clip_path}"
            ).strip()

        if elem.rotation:
            transform = (
                f"rotate({elem.rotation}deg) {transform or ''}"
            )

        if transform:
            result["style"] = (
                result.get("style", "") + f" transform: {transform.strip()};"
            ).strip()

        if self.document:
            image_data = self.document.load_document_image(
                elem.src, self.current_part,
            )
            if image_data:
                result["src"] = image_data

        return result

    def render_text(self, elem: WmlText) -> NavigableString:
        """Render text element."""
        return NavigableString(elem.text)

    def render_deleted_text(
        self, elem: WmlText,
    ) -> NavigableString | None:
        """Render deleted text element."""
        return (
            self.render_text(elem)
            if self.options.get("render_changes", False)
            else None
        )

    def render_inserted(
        self, elem: OpenXmlElement,
    ) -> Tag | list[Tag] | None:
        """Render inserted element."""
        if self.options.get("render_changes", False):
            return self.render_container(elem, "ins")
        rendered = self.render_elements(elem.children or [])
        if len(rendered) == 1:
            return rendered[0]
        return rendered if rendered else None

    def render_deleted(
        self, elem: OpenXmlElement,
    ) -> Tag | None:
        """Render deleted element."""
        if self.options.get("render_changes", False):
            return self.render_container(elem, "del")
        return None

    def render_comment_range_start(
        self, elem: OpenXmlElement,
    ) -> Tag | None:
        """Render comment range start element."""
        if not self.options.get("render_comments", False):
            return None

        comment_id = getattr(elem, "id", None)
        result = self.create_comment(f"start of comment #{comment_id}")
        
        if comment_id:
            self.comment_map[comment_id] = result
            
            def set_range_start():
                pass
            
            self.later(set_range_start)

        return result

    def render_comment_range_end(
        self, elem: OpenXmlElement,
    ) -> Tag | None:
        """Render comment range end element."""
        if not self.options.get("render_comments", False):
            return None

        comment_id = getattr(elem, "id", None)
        result = self.create_comment(f"end of comment #{comment_id}")
        
        if comment_id:
            comment_range = self.comment_map.get(comment_id)
            
            def set_range_end():
                pass
            
            if comment_range:
                self.later(set_range_end)

        return result

    def render_comment_reference(
        self, elem: OpenXmlElement,
    ) -> Tag | None:
        """Render comment reference element."""
        if not self.options.get("render_comments", False):
            return None
        comment_id = getattr(elem, "id", None)
        if not comment_id or not self.document or not self.document.comments_part:
            return None
        comment = self.document.comments_part.comment_map.get(comment_id)
        if not comment:
            return None

        container = self.create_element("div")
        
        comment_ref_el = self.create_element("span", {"class": f"{self.class_name}-comment-ref"})
        comment_ref_el.string = "💬"
        
        comments_container_el = self.create_element("div", {"class": f"{self.class_name}-comment-popover"})
        
        self.render_comment_content(comment, comments_container_el)
        
        comment_id_val = getattr(comment, "id", comment_id)
        author = getattr(comment, "author", "")
        date = getattr(comment, "date", "")
        container.append(self.create_comment(f"comment #{comment_id_val} by {author} on {date}"))
        container.append(comment_ref_el)
        container.append(comments_container_el)
        
        return container

    def render_comment_content(
        self, comment: Any, container: Tag,
    ) -> None:
        """Render comment content."""
        author_el = self.create_element("div", {"class": f"{self.class_name}-comment-author"})
        author_el.string = getattr(comment, "author", "")
        container.append(author_el)
        
        date_el = self.create_element("div", {"class": f"{self.class_name}-comment-date"})
        date_val = getattr(comment, "date", "")
        if date_val:
            try:
                from datetime import datetime
                dt = datetime.fromisoformat(date_val.replace("Z", "+00:00"))
                date_el.string = dt.strftime("%Y-%m-%d %H:%M:%S")
            except (ValueError, AttributeError):
                date_el.string = str(date_val)
        container.append(date_el)
        
        if hasattr(comment, "children"):
            self.render_elements(comment.children or [], container)

    def render_break(self, elem: WmlBreak) -> Tag | None:
        """Render break element."""
        if elem.break_type == "textWrapping":
            return self.create_element("br")
        return None

    def render_mml_radical(self, elem: OpenXmlElement) -> Tag:
        """Render MathML radical element."""
        base = next(
            (el for el in (elem.children or []) if el.type == DomType.MmlBase),
            None,
        )

        props = getattr(elem, "props", None)
        if props and props.get("hideDegree"):
            result = self.create_element_ns(ns["mathML"], "msqrt")
            if base:
                self.render_elements([base], result)
            return result

        degree = next(
            (el for el in (elem.children or []) if el.type == DomType.MmlDegree),
            None,
        )
        if base and degree:
            result = self.create_element_ns(ns["mathML"], "mroot")
            self.render_elements([base, degree], result)
            return result
        if base:
            result = self.create_element_ns(ns["mathML"], "msqrt")
            self.render_elements([base], result)
            return result
        return self.create_element_ns(ns["mathML"], "msqrt")

    def render_mml_delimiter(self, elem: OpenXmlElement) -> Tag:
        """Render MathML delimiter element."""
        children = []
        props = getattr(elem, "props", None) or {}
        begin_char = props.get("beginChar", "(")
        end_char = props.get("endChar", ")")

        begin_elem = self.create_element_ns(ns["mathML"], "mo")
        begin_elem.string = begin_char
        children.append(begin_elem)

        rendered_children = self.render_elements(elem.children or [])
        children.extend(rendered_children)

        end_elem = self.create_element_ns(ns["mathML"], "mo")
        end_elem.string = end_char
        children.append(end_elem)

        result = self.create_element_ns(ns["mathML"], "mrow")
        for child in children:
            result.append(child)
        return result

    def render_mml_nary(self, elem: OpenXmlElement) -> Tag:
        """Render MathML n-ary element."""
        children = []
        grouped = key_by(elem.children or [], lambda x: x.type)

        sup = grouped.get(DomType.MmlSuperArgument)
        sub = grouped.get(DomType.MmlSubArgument)
        props = getattr(elem, "props", None) or {}
        char_val = props.get("char", "\u222B")

        char_elem = self.create_element_ns(ns["mathML"], "mo")
        char_elem.string = char_val

        if sup and sub:
            sup_elem = self.render_element(sup)
            sub_elem = self.render_element(sub)
            if isinstance(sup_elem, list):
                sup_elem = sup_elem[0] if sup_elem else None
            if isinstance(sub_elem, list):
                sub_elem = sub_elem[0] if sub_elem else None
            munderover = self.create_element_ns(ns["mathML"], "munderover")
            if char_elem:
                munderover.append(char_elem)
            if sub_elem:
                munderover.append(sub_elem)
            if sup_elem:
                munderover.append(sup_elem)
            children.append(munderover)
        elif sup:
            sup_elem = self.render_element(sup)
            if isinstance(sup_elem, list):
                sup_elem = sup_elem[0] if sup_elem else None
            mover = self.create_element_ns(ns["mathML"], "mover")
            if char_elem:
                mover.append(char_elem)
            if sup_elem:
                mover.append(sup_elem)
            children.append(mover)
        elif sub:
            sub_elem = self.render_element(sub)
            if isinstance(sub_elem, list):
                sub_elem = sub_elem[0] if sub_elem else None
            munder = self.create_element_ns(ns["mathML"], "munder")
            if char_elem:
                munder.append(char_elem)
            if sub_elem:
                munder.append(sub_elem)
            children.append(munder)
        else:
            children.append(char_elem)

        base = grouped.get(DomType.MmlBase)
        if base:
            base_children = self.render_elements(base.children or [])
            children.extend(base_children)

        result = self.create_element_ns(ns["mathML"], "mrow")
        for child in children:
            result.append(child)
        return result

    def render_mml_pre_sub_super(self, elem: OpenXmlElement) -> Tag:
        """Render MathML pre-sub-super element."""
        children = []
        grouped = key_by(elem.children or [], lambda x: x.type)

        sup = grouped.get(DomType.MmlSuperArgument)
        sub = grouped.get(DomType.MmlSubArgument)

        stub_elem = self.create_element_ns(ns["mathML"], "mo")

        if sup and sub:
            sup_elem = self.render_element(sup)
            sub_elem = self.render_element(sub)
            if isinstance(sup_elem, list):
                sup_elem = sup_elem[0] if sup_elem else None
            if isinstance(sub_elem, list):
                sub_elem = sub_elem[0] if sub_elem else None
            msubsup = self.create_element_ns(ns["mathML"], "msubsup")
            msubsup.append(stub_elem)
            if sub_elem:
                msubsup.append(sub_elem)
            if sup_elem:
                msubsup.append(sup_elem)
            children.append(msubsup)

        base = grouped.get(DomType.MmlBase)
        if base:
            base_children = self.render_elements(base.children or [])
            children.extend(base_children)

        result = self.create_element_ns(ns["mathML"], "mrow")
        for child in children:
            result.append(child)
        return result

    def render_mml_group_char(self, elem: OpenXmlElement) -> Tag:
        """Render MathML group character element."""
        props = getattr(elem, "props", None) or {}
        vertical_justification = props.get("verticalJustification", "")
        tag_name = "mover" if vertical_justification == "bot" else "munder"
        result = self.render_container_ns(elem, ns["mathML"], tag_name)

        char_val = props.get("char")
        if char_val:
            char_elem = self.create_element_ns(ns["mathML"], "mo")
            char_elem.string = char_val
            result.append(char_elem)

        return result

    def render_mml_bar(self, elem: OpenXmlElement) -> Tag:
        """Render MathML bar element."""
        result = self.render_container_ns(elem, ns["mathML"], "mrow")
        props = getattr(elem, "props", None) or {}
        position = props.get("position", "")
        current_style = result.get("style", "")
        if position == "top":
            result["style"] = (
                current_style + "text-decoration: overline; "
            ).strip()
        elif position == "bottom":
            result["style"] = (
                current_style + "text-decoration: underline; "
            ).strip()
        return result

    def render_mml_run(self, elem: OpenXmlElement) -> Tag:
        """Render MathML run element."""
        result = self.create_element_ns(ns["mathML"], "ms")
        self.render_elements(elem.children or [], result)
        self.render_class(elem, result)
        self.render_style_values(elem.css_style or {}, result)
        return result

    def render_mml_list(self, elem: OpenXmlElement) -> Tag:
        """Render MathML list element."""
        result = self.create_element_ns(ns["mathML"], "mtable")
        self.render_class(elem, result)
        self.render_style_values(elem.css_style or {}, result)

        for child in self.render_elements(elem.children or []):
            mtr = self.create_element_ns(ns["mathML"], "mtr")
            mtd = self.create_element_ns(ns["mathML"], "mtd")
            mtd.append(child)
            mtr.append(mtd)
            result.append(mtr)

        return result

    def render_symbol(self, elem: WmlSymbol) -> Tag:
        """Render symbol element."""
        span = self.create_element("span")
        span["style"] = f"font-family: {elem.font};"
        char_code = elem.char.lstrip("0") if elem.char.startswith("0") else elem.char
        span.string = f"&#x{char_code};"
        return span

    def render_footnote_reference(
        self, elem: WmlNoteReference,
    ) -> Tag:
        """Render footnote reference element."""
        result = self.create_element("sup")
        self.current_footnote_ids.append(elem.id)
        result.string = str(len(self.current_footnote_ids))
        return result

    def render_endnote_reference(
        self, elem: WmlNoteReference,
    ) -> Tag:
        """Render endnote reference element."""
        result = self.create_element("sup")
        self.current_endnote_ids.append(elem.id)
        result.string = str(len(self.current_endnote_ids))
        return result

    def render_tab(self, _elem: OpenXmlElement) -> Tag:
        """Render tab element."""
        tab_span = self.create_element("span")
        tab_span.string = "\u2003"

        if self.options.get("experimental", False):
            tab_span["class"] = self.tab_stop_class()
            parent = _elem.parent
            while parent and parent.type != DomType.Paragraph:
                parent = parent.parent
            stops = getattr(parent, "tabs", None) if parent else None
            self.current_tabs.append({"stops": stops, "span": tab_span})

        return tab_span

    def render_bookmark_start(self, elem: WmlBookmarkStart) -> Tag:
        """Render bookmark start element."""
        return self.create_element("span", {"id": elem.name})

    def render_run(self, elem: WmlRun) -> Tag | None:
        """Render run element."""
        if elem.field_run:
            return None

        result = self.create_element("span")
        if elem.id:
            result["id"] = elem.id

        self.render_class(elem, result)
        self.render_style_values(elem.css_style or {}, result)

        if elem.vertical_align:
            wrapper = self.create_element(elem.vertical_align)
            self.render_elements(elem.children or [], wrapper)
            result.append(wrapper)
        else:
            self.render_elements(elem.children or [], result)

        return result

    def render_table(self, elem: WmlTable) -> Tag:
        """Render table element."""
        result = self.create_element("table")

        self.table_cell_positions.append(
            self.current_cell_position.copy(),
        )
        self.table_vertical_merges.append(
            self.current_vertical_merge.copy(),
        )
        self.current_vertical_merge = {}
        self.current_cell_position = {"col": 0, "row": 0}

        if elem.columns:
            result.append(self.render_table_columns(elem.columns))

        self.render_class(elem, result)
        self.render_elements(elem.children or [], result)
        self.render_style_values(elem.css_style or {}, result)

        self.current_vertical_merge = self.table_vertical_merges.pop()
        self.current_cell_position = self.table_cell_positions.pop()

        return result

    def render_table_columns(
        self, columns: list[WmlTableColumn],
    ) -> Tag:
        """Render table columns."""
        result = self.create_element("colgroup")
        for col in columns:
            col_elem = self.create_element("col")
            if col.width:
                col_elem["style"] = f"width: {col.width};"
            result.append(col_elem)
        return result

    def render_table_row(self, elem: WmlTableRow) -> Tag:
        """Render table row element."""
        result = self.create_element("tr")
        self.current_cell_position["col"] = 0

        if elem.grid_before:
            result.append(
                self.render_table_cell_placeholder(elem.grid_before),
            )

        self.render_class(elem, result)
        self.render_elements(elem.children or [], result)
        self.render_style_values(elem.css_style or {}, result)

        if elem.grid_after:
            result.append(
                self.render_table_cell_placeholder(elem.grid_after),
            )

        self.current_cell_position["row"] += 1
        return result

    def render_table_cell_placeholder(self, col_span: int) -> Tag:
        """Render table cell placeholder."""
        result = self.create_element("td", {"colspan": str(col_span)})
        result["style"] = "border: none;"
        return result

    def render_table_cell(self, elem: WmlTableCell) -> Tag:
        """Render table cell element."""
        result = self.render_container(elem, "td")

        key = self.current_cell_position["col"]

        if elem.vertical_merge:
            if elem.vertical_merge == "restart":
                self.current_vertical_merge[key] = result
                result["rowspan"] = "1"
            elif self.current_vertical_merge.get(key):
                current_rowspan = int(
                    self.current_vertical_merge[key].get("rowspan", "1"),
                )
                self.current_vertical_merge[key]["rowspan"] = str(
                    current_rowspan + 1,
                )
                result["style"] = "display: none;"
        else:
            self.current_vertical_merge[key] = None

        self.render_class(elem, result)
        self.render_style_values(elem.css_style or {}, result)

        if elem.span:
            result["colspan"] = str(elem.span)

        self.current_cell_position["col"] += int(
            result.get("colspan", "1"),
        )

        return result

    def render_vml_picture(self, elem: OpenXmlElement) -> Tag:
        """Render VML picture element."""
        return self.render_container(elem, "div")

    def render_vml_element(self, elem: VmlElement) -> Tag:
        """Render VML element."""
        container = self.create_element_ns(ns["svg"], "svg")
        if elem.css_style_text:
            container["style"] = elem.css_style_text

        result = self.render_vml_child_element(elem)
        container.append(result)

        if elem.image_href and elem.image_href.get("id"):
            image_data = self.document.load_document_image(
                elem.image_href["id"], self.current_part,
            )
            if image_data:
                result["href"] = image_data

        return container

    def render_vml_child_element(self, elem: VmlElement) -> Tag:
        """Render VML child element."""
        result = self.create_element_ns(ns["svg"], elem.tag_name)
        for k, v in elem.attrs.items():
            result[k] = v

        for child in elem.children:
            if isinstance(child, VmlElement):
                result.append(self.render_vml_child_element(child))
            else:
                rendered = self.render_element(child)
                if rendered:
                    result.append(rendered)

        return result

    def render_style_values(
        self, style: dict[str, str], output: Tag,
    ) -> None:
        """Render style values to output tag."""
        for k, v in style.items():
            if k.startswith("$"):
                output[k[1:]] = v
            else:
                current_style = output.get("style", "")
                output["style"] = (
                    current_style + f"{k}: {v}; "
                ).strip()

    def render_common_properties(
        self, output: Tag, props: Any,
    ) -> None:
        """Render common properties to output tag."""
        if props is None:
            return

        current_style = output.get("style", "")
        if hasattr(props, "color") and props.color:
            output["style"] = (
                current_style + f"color: {props.color}; "
            ).strip()
            current_style = output.get("style", "")

        if hasattr(props, "font_size") and props.font_size:
            output["style"] = (
                current_style + f"font-size: {props.font_size}; "
            ).strip()

    def render_class(
        self, input_elem: OpenXmlElement, output: Tag,
    ) -> None:
        """Render class to output tag."""
        if input_elem.class_name:
            output["class"] = input_elem.class_name
        if input_elem.style_name:
            current_class = output.get("class", "")
            style_name = self.process_style_name(input_elem.style_name)
            output["class"] = (
                current_class + f" {style_name}"
            ).strip()

    def find_style(self, style_name: str | None) -> IDomStyle | None:
        """Find style by name."""
        return self.style_map.get(style_name) if (style_name and self.style_map) else None

    def numbering_class(self, id_val: str, lvl: int) -> str:
        """Get numbering class name."""
        return f"{self.class_name}-num-{id_val}-{lvl}"

    def tab_stop_class(self) -> str:
        """Get tab stop class name."""
        return f"{self.class_name}-tab-stop"

    def style_to_string(
        self, selectors: str, values: dict[str, str], css_text: str = "",
    ) -> str:
        """Convert style dictionary to CSS string."""
        result = f"{selectors} {{\n"
        for key, val in values.items():
            if not key.startswith("$"):
                result += f"  {key}: {val};\n"
        if css_text:
            result += css_text
        return result + "}\n"

    def numbering_counter(self, id_val: str, lvl: int) -> str:
        """Get numbering counter name."""
        return f"{self.class_name}-num-{id_val}-{lvl}"

    def level_text_to_content(
        self, text: str, suff: str, id_val: str, num_format: str,
    ) -> str:
        """Convert level text to CSS content."""
        suff_map = {
            "tab": "\\9",
            "space": "\\a0",
        }

        def replace_counter(m: re.Match[str]) -> str:
            counter_name = self.numbering_counter(
                id_val, int(m.group(0)[1:]) - 1,
            )
            return f'"counter({counter_name}, {num_format})"'

        result = re.sub(r"%\d*", replace_counter, text)
        return f'"{result}{suff_map.get(suff, "")}"'

    def num_format_to_css_value(self, format_val: str) -> str:
        """Convert number format to CSS value."""
        mapping = {
            "none": "none",
            "bullet": "disc",
            "decimal": "decimal",
            "lowerLetter": "lower-alpha",
            "upperLetter": "upper-alpha",
            "lowerRoman": "lower-roman",
            "upperRoman": "upper-roman",
            "decimalZero": "decimal-leading-zero",
            "aiueo": "katakana",
            "aiueoFullWidth": "katakana",
            "chineseCounting": "simp-chinese-informal",
            "chineseCountingThousand": "simp-chinese-informal",
            "chineseLegalSimplified": "simp-chinese-formal",
            "chosung": "hangul-consonant",
            "ideographDigital": "cjk-ideographic",
            "ideographTraditional": "cjk-heavenly-stem",
            "ideographLegalTraditional": "trad-chinese-formal",
            "ideographZodiac": "cjk-earthly-branch",
            "iroha": "katakana-iroha",
            "irohaFullWidth": "katakana-iroha",
            "japaneseCounting": "japanese-informal",
            "japaneseDigitalTenThousand": "cjk-decimal",
            "japaneseLegal": "japanese-formal",
            "thaiNumbers": "thai",
            "koreanCounting": "korean-hangul-formal",
            "koreanDigital": "korean-hangul-formal",
            "koreanDigital2": "korean-hanja-informal",
            "hebrew1": "hebrew",
            "hebrew2": "hebrew",
            "hindiNumbers": "devanagari",
            "ganada": "hangul",
            "taiwaneseCounting": "cjk-ideographic",
            "taiwaneseCountingThousand": "cjk-ideographic",
            "taiwaneseDigital": "cjk-decimal",
        }
        return mapping.get(format_val, format_val)

    def create_element(
        self, tag_name: str, props: dict | None = None,
    ) -> Tag:
        """Create HTML element."""
        result = self.html_document.new_tag(tag_name)
        if props:
            for k, v in props.items():
                result[k] = v
        return result

    def create_element_ns(
        self,
        namespace: str,
        tag_name: str,
        props: dict | None = None,
    ) -> Tag:
        """Create HTML element with namespace."""
        result = self.html_document.new_tag(
            tag_name, namespace=namespace,
        )
        if props:
            for k, v in props.items():
                result[k] = v
        return result

    def create_style_element(self, css_text: str) -> Tag:
        """Create style element."""
        style = self.create_element("style")
        css_text_stripped = css_text.rstrip()
        if css_text_stripped and not css_text_stripped.endswith(" "):
            css_text_stripped += " "
        style.string = css_text_stripped
        return style

    def create_comment(self, text: str) -> Tag:
        """Create HTML comment."""
        return Comment(text)

    def later(self, func: Any) -> None:
        """Defer function execution to post-render phase."""
        self.post_render_tasks.append(func)

    def render_run_properties(
        self, output: Tag, props: Any,
    ) -> None:
        """Render run properties."""
        self.render_common_properties(output, props)

    def refresh_tab_stops(self) -> None:
        """Refresh tab stops (experimental feature)."""
        if not self.options.get("experimental", False):
            return

