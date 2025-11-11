"""Streamlit app for DOCX to HTML rendering."""
import streamlit as st
from bs4 import BeautifulSoup

from docxjs import render_async

st.set_page_config(page_title="DOCX Renderer", page_icon="📄", layout="wide")

st.title("📄 DOCX to HTML Renderer")
st.markdown("Upload a DOCX file to render it as HTML using the docx.js Python port")

uploaded_file = st.file_uploader("Choose a DOCX file", type=["docx"])

if uploaded_file is not None:
    try:
        with st.spinner("Parsing and rendering DOCX file..."):
            docx_data = uploaded_file.read()

            soup = BeautifulSoup(
                "<html><head></head><body></body></html>", "html.parser",
            )
            body = soup.body
            head = soup.head if soup.head else soup.new_tag("head")
            if not soup.head:
                soup.html.insert(0, head)

            doc = render_async(docx_data, body, head, {
                "in_wrapper": True,
                "class_name": "docx",
                "ignore_fonts": False,
                "render_headers": True,
                "render_footers": True,
                "render_footnotes": True,
                "render_endnotes": True,
            })

            html_output = str(soup.prettify())

            st.success("Document rendered successfully!")

            col1, col2 = st.columns([2, 1])

            with col1:
                st.subheader("Rendered Output")
                st.components.v1.html(html_output, height=800, scrolling=True)

            with col2:
                st.subheader("HTML Source")
                with st.expander("View HTML", expanded=False):
                    st.code(html_output, language="html")

                st.subheader("Download")
                st.download_button(
                    label="Download HTML",
                    data=html_output,
                    file_name=(
                        f"{uploaded_file.name.replace('.docx', '')}_rendered.html"
                    ),
                    mime="text/html",
                )

                st.subheader("Document Info")
                if doc.document_part and doc.document_part.body:
                    st.info("Document loaded successfully")
                    if doc.styles_part:
                        st.write(f"Styles: {len(doc.styles_part.styles)}")
                    if doc.numbering_part:
                        st.write(
                            f"Numbering definitions: "
                            f"{len(doc.numbering_part.dom_numberings)}",
                        )
                    if doc.footnotes_part:
                        st.write(f"Footnotes: {len(doc.footnotes_part.notes)}")
                    if doc.endnotes_part:
                        st.write(f"Endnotes: {len(doc.endnotes_part.notes)}")

    except (ValueError, KeyError, AttributeError, OSError) as e:
        st.error(f"Error rendering document: {e!s}")
        st.exception(e)
        if st.checkbox("Show debug info"):
            st.code(str(e), language="text")

else:
    st.info("👆 Please upload a DOCX file to get started")

    with st.expander("Example Usage Code"):
        st.code("""
from docx_preview import render_async
from bs4 import BeautifulSoup

with open('document.docx', 'rb') as f:
    docx_data = f.read()

soup = BeautifulSoup('<html><body></body></html>', 'html.parser')
body = soup.body

doc = render_async(docx_data, body, body)
print(soup.prettify())
        """, language="python")

