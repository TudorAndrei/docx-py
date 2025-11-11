"""Comment elements module."""

from docxjs.document.dom import DomType, OpenXmlElement


class WmlComment(OpenXmlElement):
    """Word ML comment element."""

    def __init__(self) -> None:
        """Initialize comment element."""
        super().__init__()
        self.type = DomType.Comment
        self.id: str = ""
        self.author: str = ""
        self.initials: str = ""
        self.date: str = ""


class WmlCommentReference(OpenXmlElement):
    """Word ML comment reference element."""

    def __init__(self, id_val: str = "") -> None:
        """Initialize comment reference element."""
        super().__init__()
        self.type = DomType.CommentReference
        self.id = id_val


class WmlCommentRangeStart(OpenXmlElement):
    """Word ML comment range start element."""

    def __init__(self, id_val: str = "") -> None:
        """Initialize comment range start element."""
        super().__init__()
        self.type = DomType.CommentRangeStart
        self.id = id_val


class WmlCommentRangeEnd(OpenXmlElement):
    """Word ML comment range end element."""

    def __init__(self, id_val: str = "") -> None:
        """Initialize comment range end element."""
        super().__init__()
        self.type = DomType.CommentRangeEnd
        self.id = id_val

