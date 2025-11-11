"""Fields module for handling Word document fields."""

from docxjs.document.dom import DomType, OpenXmlElement


class WmlInstructionText(OpenXmlElement):
    """Word ML instruction text element."""

    def __init__(self) -> None:
        """Initialize instruction text element."""
        super().__init__()
        self.type = DomType.Instruction
        self.text: str = ""


class WmlFieldChar(OpenXmlElement):
    """Word ML field character element."""

    def __init__(self) -> None:
        """Initialize field character element."""
        super().__init__()
        self.type = DomType.ComplexField
        self.char_type: str = ""
        self.lock: bool = False
        self.dirty: bool = False


class WmlFieldSimple(OpenXmlElement):
    """Word ML simple field element."""

    def __init__(self) -> None:
        """Initialize simple field element."""
        super().__init__()
        self.type = DomType.SimpleField
        self.instruction: str = ""
        self.lock: bool = False
        self.dirty: bool = False

