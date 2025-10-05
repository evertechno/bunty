"""
Streamlit Multi-format Converter
- PDF → Text / Word / HTML
- HTML → Word / Text
- Bulk processing supported
- Upload + Download buttons with previews

Requirements:
    pip install streamlit pdfminer.six pymupdf pillow python-docx beautifulsoup4
Run:
    streamlit run streamlit_converter.py
"""

import streamlit as st
import io, base64, os
from pdfminer.high_level import extract_text_to_fp, extract_text
from pdfminer.layout import LAParams
import fitz  # PyMuPDF
from docx import Document
from bs4 import BeautifulSoup

# --------------------- Streamlit Setup ---------------------
st.set_page_config(page_title="Universal Converter", layout="wide")
st.title("📄 Universal Converter (PDF ↔ Text/Word/HTML)")
st.markdown("Upload PDFs or HTML files and convert them into different formats with precision. Supports single & bulk files.")

uploaded_files = st.file_uploader("Upload files", type=["pdf", "html"], accept_multiple_files=True)
method = st.selectbox(
    "Choose conversion type",
    [
        "PDF → Text",
        "PDF → Word",
        "PDF → HTML (text-preserving)",
        "PDF → HTML (image fallback)",
        "HTML → Text",
        "HTML → Word",
    ]
)
quality = st.slider("Image DPI (for PDF → HTML image fallback)", 100, 400, 150, 25)

# --------------------- Conversion Functions ---------------------
def pdf_to_html_text(pdf_bytes: bytes) -> bytes:
    outfp = io.BytesIO()
    laparams = LAParams()
    extract_text_to_fp(io.BytesIO(pdf_bytes), outfp, output_type="html", laparams=laparams)
    return outfp.getvalue()

def pdf_to_html_images(pdf_bytes: bytes, dpi: int = 150) -> bytes:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_html = []
    for page_number in range(len(doc)):
        page = doc.load_page(page_number)
        mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_bytes = pix.tobytes("png")
        b64 = base64.b64encode(img_bytes).decode("ascii")
        pages_html.append(
            f'<div style="page-break-after: always; text-align:center;">'
            f'<img src="data:image/png;base64,{b64}" alt="page-{page_number+1}" style="max-width:100%; height:auto;"/></div>'
        )
    html = "<html><body>" + "\n".join(pages_html) + "</body></html>"
    return html.encode("utf-8")

def pdf_to_text_bytes(pdf_bytes: bytes) -> bytes:
    text = extract_text(io.BytesIO(pdf_bytes))
    return text.encode("utf-8")

def pdf_to_word_bytes(pdf_bytes: bytes) -> bytes:
    text = extract_text(io.BytesIO(pdf_bytes))
    doc = Document()
    for line in text.splitlines():
        doc.add_paragraph(line)
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

def html_to_text_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    text = soup.get_text(separator="\n")
    return text.encode("utf-8")

def html_to_word_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    doc = Document()
    text = soup.get_text(separator="\n")
    for line in text.splitlines():
        doc.add_paragraph(line)
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

# --------------------- Processing ---------------------
if uploaded_files:
    convert = st.button("🚀 Convert All Files")
    if convert:
        for uploaded_file in uploaded_files:
            file_bytes = uploaded_file.read()
            file_name, ext = os.path.splitext(uploaded_file.name)
            output_bytes, mime_type, out_ext = None, None, None

            try:
                if method == "PDF → Text" and ext.lower() == ".pdf":
                    output_bytes = pdf_to_text_bytes(file_bytes)
                    mime_type, out_ext = "text/plain", ".txt"
                elif method == "PDF → Word" and ext.lower() == ".pdf":
                    output_bytes = pdf_to_word_bytes(file_bytes)
                    mime_type, out_ext = "application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx"
                elif method == "PDF → HTML (text-preserving)" and ext.lower() == ".pdf":
                    output_bytes = pdf_to_html_text(file_bytes)
                    mime_type, out_ext = "text/html", ".html"
                elif method == "PDF → HTML (image fallback)" and ext.lower() == ".pdf":
                    output_bytes = pdf_to_html_images(file_bytes, dpi=quality)
                    mime_type, out_ext = "text/html", ".html"
                elif method == "HTML → Text" and ext.lower() == ".html":
                    output_bytes = html_to_text_bytes(file_bytes)
                    mime_type, out_ext = "text/plain", ".txt"
                elif method == "HTML → Word" and ext.lower() == ".html":
                    output_bytes = html_to_word_bytes(file_bytes)
                    mime_type, out_ext = "application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx"

                if output_bytes:
                    st.success(f"✅ {uploaded_file.name} converted successfully!")
                    if mime_type.startswith("text/"):
                        preview_text = output_bytes.decode("utf-8", errors="replace")[:2000]
                        st.text_area(f"Preview of {uploaded_file.name}", preview_text, height=200)
                    elif mime_type == "text/html":
                        st.components.v1.html(output_bytes.decode("utf-8", errors="replace")[:50000], height=400, scrolling=True)

                    st.download_button(
                        f"⬇ Download {file_name+out_ext}",
                        data=output_bytes,
                        file_name=file_name + out_ext,
                        mime=mime_type
                    )
                else:
                    st.error(f"❌ Conversion not supported for {uploaded_file.name} with {method}")

            except Exception as e:
                st.error(f"⚠️ Failed to convert {uploaded_file.name}: {e}")
else:
    st.info("Upload one or more PDF/HTML files to start.")
