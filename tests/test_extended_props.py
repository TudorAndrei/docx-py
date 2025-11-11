"""Tests for extended properties."""

from pathlib import Path

import pytest

from docxjs import parse_async

EXPECTED_CHARACTERS = 393
EXPECTED_LINES = 3
EXPECTED_PAGES = 3
EXPECTED_WORDS = 68


def test_extended_props() -> None:
    """Test extended properties parsing."""
    test_dir = Path(__file__).parent / "extended-props-test"
    docx_path = test_dir / "document.docx"

    if not docx_path.exists():
        pytest.skip(f"Test document not found: {docx_path}")

    with docx_path.open("rb") as f:
        docx_data = f.read()

    doc = parse_async(docx_data)

    assert doc.extended_props_part is not None, "extendedPropsPart should exist"  # noqa: S101
    assert doc.extended_props_part.appVersion == "16.0000"  # noqa: S101
    assert doc.extended_props_part.application == "Microsoft Office Word"  # noqa: S101
    assert doc.extended_props_part.characters == EXPECTED_CHARACTERS  # noqa: S101
    assert doc.extended_props_part.company == ""  # noqa: S101
    assert doc.extended_props_part.lines == EXPECTED_LINES  # noqa: S101
    assert doc.extended_props_part.pages == EXPECTED_PAGES  # noqa: S101
    assert doc.extended_props_part.paragraphs == 1  # noqa: S101
    assert doc.extended_props_part.template == "Normal.dotm"  # noqa: S101
    assert doc.extended_props_part.words == EXPECTED_WORDS  # noqa: S101
