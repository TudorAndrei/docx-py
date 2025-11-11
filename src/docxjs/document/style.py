"""Style module for handling Word document styles."""

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docxjs.document.paragraph import ParagraphProperties
    from docxjs.document.run import RunProperties


class IDomSubStyle:
    """DOM sub-style."""

    def __init__(self) -> None:
        """Initialize sub-style."""
        self.target: str = ""
        self.mod: str | None = None
        self.values: dict[str, str] = {}


class IDomStyle:
    """DOM style."""

    def __init__(self) -> None:
        """Initialize style."""
        self.id: str | None = None
        self.name: str | None = None
        self.css_name: str | None = None
        self.aliases: list[str] | None = None
        self.target: str | None = None
        self.based_on: str | None = None
        self.is_default: bool = False
        self.styles: list[IDomSubStyle] = []
        self.linked: str | None = None
        self.next: str | None = None
        self.paragraph_props: ParagraphProperties | None = None
        self.run_props: RunProperties | None = None

