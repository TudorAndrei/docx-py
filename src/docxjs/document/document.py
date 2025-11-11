"""Document element module."""

from typing import TYPE_CHECKING

from docxjs.document.dom import OpenXmlElement

if TYPE_CHECKING:
    from docxjs.document.section import SectionProperties


class DocumentElement(OpenXmlElement):
    """Document element."""

    def __init__(self) -> None:
        """Initialize document element."""
        super().__init__()
        self.props: SectionProperties | None = None

