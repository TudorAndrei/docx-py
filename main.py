"""Main entry point for the docxjs Python port."""

import sys


def main() -> None:
    """Display example usage information."""
    sys.stdout.write("docx.js Python port - Ready to use!\n")
    sys.stdout.write("Example usage:\n")
    sys.stdout.write("  from docxjs import render_async\n")
    sys.stdout.write("  from bs4 import BeautifulSoup\n")
    sys.stdout.write("  \n")
    sys.stdout.write("  with open('document.docx', 'rb') as f:\n")
    sys.stdout.write("      docx_data = f.read()\n")
    sys.stdout.write("  \n")
    soup_line = (
        "  soup = BeautifulSoup('<html><body></body></html>', "
        "'html.parser')\n"
    )
    sys.stdout.write(soup_line)
    sys.stdout.write("  body = soup.body\n")
    sys.stdout.write("  \n")
    sys.stdout.write("  doc = render_async(docx_data, body, body)\n")
    sys.stdout.write("  print(soup.prettify())\n")


if __name__ == "__main__":
    main()
