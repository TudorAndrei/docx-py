"""DOM types and elements module."""

from enum import Enum
from typing import Any


class DomType(Enum):
    """Document object model type enumeration."""

    Document = "document"
    Paragraph = "paragraph"
    Run = "run"
    Break = "break"
    NoBreakHyphen = "noBreakHyphen"
    Table = "table"
    Row = "row"
    Cell = "cell"
    Hyperlink = "hyperlink"
    SmartTag = "smartTag"
    Drawing = "drawing"
    Image = "image"
    Text = "text"
    Tab = "tab"
    Symbol = "symbol"
    BookmarkStart = "bookmarkStart"
    BookmarkEnd = "bookmarkEnd"
    Footer = "footer"
    Header = "header"
    FootnoteReference = "footnoteReference"
    EndnoteReference = "endnoteReference"
    Footnote = "footnote"
    Endnote = "endnote"
    SimpleField = "simpleField"
    ComplexField = "complexField"
    Instruction = "instruction"
    VmlPicture = "vmlPicture"
    MmlMath = "mmlMath"
    MmlMathParagraph = "mmlMathParagraph"
    MmlFraction = "mmlFraction"
    MmlFunction = "mmlFunction"
    MmlFunctionName = "mmlFunctionName"
    MmlNumerator = "mmlNumerator"
    MmlDenominator = "mmlDenominator"
    MmlRadical = "mmlRadical"
    MmlBase = "mmlBase"
    MmlDegree = "mmlDegree"
    MmlSuperscript = "mmlSuperscript"
    MmlSubscript = "mmlSubscript"
    MmlPreSubSuper = "mmlPreSubSuper"
    MmlSubArgument = "mmlSubArgument"
    MmlSuperArgument = "mmlSuperArgument"
    MmlNary = "mmlNary"
    MmlDelimiter = "mmlDelimiter"
    MmlRun = "mmlRun"
    MmlEquationArray = "mmlEquationArray"
    MmlLimit = "mmlLimit"
    MmlLimitLower = "mmlLimitLower"
    MmlMatrix = "mmlMatrix"
    MmlMatrixRow = "mmlMatrixRow"
    MmlBox = "mmlBox"
    MmlBar = "mmlBar"
    MmlGroupChar = "mmlGroupChar"
    VmlElement = "vmlElement"
    Inserted = "inserted"
    Deleted = "deleted"
    DeletedText = "deletedText"
    Comment = "comment"
    CommentReference = "commentReference"
    CommentRangeStart = "commentRangeStart"
    CommentRangeEnd = "commentRangeEnd"
    AltChunk = "altChunk"


class OpenXmlElement:
    """Base class for OpenXML elements."""

    def __init__(self) -> None:
        """Initialize OpenXML element."""
        self.type: DomType | None = None
        self.children: list[OpenXmlElement] | None = []
        self.css_style: dict[str, str] | None = {}
        self.props: dict[str, Any] | None = None
        self.style_name: str | None = None
        self.class_name: str | None = None
        self.parent: OpenXmlElement | None = None


class WmlHyperlink(OpenXmlElement):
    """Word ML hyperlink element."""

    def __init__(self) -> None:
        """Initialize hyperlink element."""
        super().__init__()
        self.type = DomType.Hyperlink
        self.id: str | None = None
        self.anchor: str | None = None


class WmlAltChunk(OpenXmlElement):
    """Word ML alt chunk element."""

    def __init__(self) -> None:
        """Initialize alt chunk element."""
        super().__init__()
        self.type = DomType.AltChunk
        self.id: str | None = None


class WmlSmartTag(OpenXmlElement):
    """Word ML smart tag element."""

    def __init__(self) -> None:
        """Initialize smart tag element."""
        super().__init__()
        self.type = DomType.SmartTag
        self.uri: str | None = None
        self.element: str | None = None


class WmlNoteReference(OpenXmlElement):
    """Word ML note reference element."""

    def __init__(self) -> None:
        """Initialize note reference element."""
        super().__init__()
        self.id: str = ""


class WmlBreak(OpenXmlElement):
    """Word ML break element."""

    def __init__(self) -> None:
        """Initialize break element."""
        super().__init__()
        self.type = DomType.Break
        self.break_type: str = "textWrapping"


class WmlText(OpenXmlElement):
    """Word ML text element."""

    def __init__(self) -> None:
        """Initialize text element."""
        super().__init__()
        self.type = DomType.Text
        self.text: str = ""


class WmlSymbol(OpenXmlElement):
    """Word ML symbol element."""

    def __init__(self) -> None:
        """Initialize symbol element."""
        super().__init__()
        self.type = DomType.Symbol
        self.font: str = ""
        self.char: str = ""


class WmlTableColumn:
    """Word ML table column."""

    def __init__(self) -> None:
        """Initialize table column."""
        self.width: str | None = None


class WmlTable(OpenXmlElement):
    """Word ML table element."""

    def __init__(self) -> None:
        """Initialize table element."""
        super().__init__()
        self.type = DomType.Table
        self.columns: list[WmlTableColumn] | None = None
        self.cell_style: dict[str, str] | None = None
        self.col_band_size: int | None = None
        self.row_band_size: int | None = None


class WmlTableRow(OpenXmlElement):
    """Word ML table row element."""

    def __init__(self) -> None:
        """Initialize table row element."""
        super().__init__()
        self.type = DomType.Row
        self.is_header: bool | None = None
        self.grid_before: int | None = None
        self.grid_after: int | None = None


class WmlTableCell(OpenXmlElement):
    """Word ML table cell element."""

    def __init__(self) -> None:
        """Initialize table cell element."""
        super().__init__()
        self.type = DomType.Cell
        self.vertical_merge: str | None = None
        self.span: int | None = None


class IDomImage(OpenXmlElement):
    """DOM image element."""

    def __init__(self) -> None:
        """Initialize image element."""
        super().__init__()
        self.type = DomType.Image
        self.src: str = ""
        self.src_rect: list[float] | None = None
        self.rotation: float = 0.0


class NumberingPicBullet:
    """Numbering picture bullet."""

    def __init__(self) -> None:
        """Initialize numbering picture bullet."""
        self.id: int = 0
        self.src: str = ""
        self.style: str | None = None


class IDomNumbering:
    """DOM numbering element."""

    def __init__(self) -> None:
        """Initialize numbering element."""
        self.id: str = ""
        self.level: int = 0
        self.start: int = 1
        self.p_style_name: str | None = None
        self.p_style: dict[str, str] = {}
        self.r_style: dict[str, str] = {}
        self.level_text: str | None = None
        self.suff: str = "tab"
        self.format: str | None = None
        self.bullet: NumberingPicBullet | None = None

