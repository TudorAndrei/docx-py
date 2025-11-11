"""Tests for document rendering."""

import re
from pathlib import Path

import pytest
from bs4 import BeautifulSoup

from docxjs import render_async


def format_html(text: str) -> str:
    """Format HTML text for comparison."""
    text = re.sub(r"\t+|\s+", " ", text, flags=re.IGNORECASE)
    return re.sub(r"><", ">\n<", text)


def get_test_cases() -> list[dict[str, Path]]:
    """Get test cases from render-test directory."""
    test_dir = Path(__file__).parent / "render-test"
    test_cases = []

    for test_dir_path in test_dir.iterdir():
        if test_dir_path.is_dir():
            docx_path = test_dir_path / "document.docx"
            result_path = test_dir_path / "result.html"

            if docx_path.exists() and result_path.exists():
                test_cases.append(
                    {
                        "name": test_dir_path.name,
                        "docx_path": docx_path,
                        "expected_html_path": result_path,
                    }
                )

    return test_cases


@pytest.mark.parametrize("test_case", get_test_cases())
def test_render_document(test_case: dict[str, Path]) -> None:
    """Test document rendering."""
    with test_case["docx_path"].open("rb") as f:
        docx_data = f.read()

    with test_case["expected_html_path"].open(encoding="utf-8") as f:
        expected_html = f.read()

    soup = BeautifulSoup("<html><body></body></html>", "html.parser")
    body = soup.body

    render_async(docx_data, body, body)

    actual_html = body.decode_contents()
    actual_formatted = format_html(actual_html)
    expected_formatted = format_html(expected_html)

    if actual_formatted != expected_formatted:
        pass

    error_msg = f"HTML output does not match expected output for {test_case['name']}"
    assert actual_formatted == expected_formatted, error_msg  # noqa: S101
