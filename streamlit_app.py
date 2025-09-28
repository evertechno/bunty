"""
Streamlit PDF → HTML & Text converter
- Text-preserving HTML conversion using pdfminer.six
- Image-based fallback HTML conversion using PyMuPDF (pixel-perfect)
- PDF → Text conversion option
- Upload + Download buttons with previews

Requirements:
    pip install streamlit pdfminer.six pymupdf pillow

Run:
    streamlit run streamlit_pdf_to_html_converter.py
"""

import streamlit as st
import io
import base64
from pdfminer.high_level import extract_text_to_fp, extract_text
from pdfminer.layout import LAParams
import fitz  # PyMuPDF

st.set_page_config(page_title="PDF Converter", layout="wide")
st.title("PDF Converter (PDF → HTML & Text)")
st.markdown("Upload a PDF and convert it to HTML or plain text. Data is preserved — choose the conversion method you prefer.")

uploaded_file = st.file_uploader("Choose a PDF file", type=["pdf"])  
method = st.radio("Conversion method", ["Text-preserving HTML (pdfminer)", "Image fallback HTML (pixel-perfect)", "Plain Text"], index=0)
embed_pdf = st.checkbox("Embed original PDF inside the HTML (only for HTML conversion)", value=False)
quality = st.slider("Image DPI for fallback (higher = larger file)", min_value=100, max_value=400, value=150, step=25)


def pdf_to_html_text(pdf_bytes: bytes) -> bytes:
    """Convert PDF bytes to HTML using pdfminer.six. Returns HTML bytes."""
    outfp = io.BytesIO()
    laparams = LAParams()
    try:
        extract_text_to_fp(io.BytesIO(pdf_bytes), outfp, output_type='html', laparams=laparams)
        html_bytes = outfp.getvalue()
        if not html_bytes.strip():
            raise ValueError("pdfminer produced empty output")
        return html_bytes
    finally:
        outfp.close()


def pdf_to_html_images(pdf_bytes: bytes, dpi: int = 150) -> bytes:
    """Render each PDF page to a PNG using PyMuPDF and embed as base64 images in an HTML file."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_html = []
    for page_number in range(len(doc)):
        page = doc.load_page(page_number)
        mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_bytes = pix.tobytes(output="png")
        b64 = base64.b64encode(img_bytes).decode('ascii')
        pages_html.append(f'<div style="page-break-after: always; text-align:center;">\n'
                          f'<img src="data:image/png;base64,{b64}" alt="page-{page_number+1}" style="max-width:100%; height:auto;"/>\n'
                          f'</div>')
    title = "PDF_as_images"
    html = f'<!doctype html><html><head><meta charset="utf-8"><title>{title}</title></head><body>' + "\n".join(pages_html) + '</body></html>'
    return html.encode('utf-8')


def pdf_to_text_bytes(pdf_bytes: bytes) -> bytes:
    """Extract plain text from PDF using pdfminer."""
    text = extract_text(io.BytesIO(pdf_bytes))
    if not text.strip():
        raise ValueError("No text extracted (possibly a scanned PDF)")
    return text.encode("utf-8")


def embed_original_pdf_in_html(html_bytes: bytes, pdf_bytes: bytes) -> bytes:
    """Append an <embed> of the original PDF at the end of the HTML, encoded as base64."""
    b64 = base64.b64encode(pdf_bytes).decode('ascii')
    embed_snippet = f"\n<hr/>\n<h2>Original PDF (embedded)</h2>\n<embed src=\"data:application/pdf;base64,{b64}\" width=\"100%\" height=\"800px\"></embed>\n"
    try:
        s = html_bytes.decode('utf-8')
        if "</body>" in s:
            s = s.replace("</body>", embed_snippet + "</body>")
        else:
            s = s + embed_snippet
        return s.encode('utf-8')
    except Exception:
        return html_bytes + embed_snippet.encode('utf-8')


if uploaded_file is not None:
    pdf_bytes = uploaded_file.read()
    st.markdown(f"**File:** {uploaded_file.name} — {len(pdf_bytes):,} bytes")

    convert = st.button("Convert")

    if convert:
        with st.spinner("Converting..."):
            output_bytes = None
            mime_type = None
            ext = None
            conversion_error = None

            if method.startswith("Text-preserving HTML"):
                try:
                    output_bytes = pdf_to_html_text(pdf_bytes)
                    mime_type = "text/html"
                    ext = ".html"
                except Exception as e:
                    conversion_error = str(e)
                    st.warning(f"Text conversion failed: {conversion_error}. Falling back to image-based conversion.")
                    output_bytes = pdf_to_html_images(pdf_bytes, dpi=quality)
                    mime_type = "text/html"
                    ext = ".html"
            elif method.startswith("Image fallback"):
                try:
                    output_bytes = pdf_to_html_images(pdf_bytes, dpi=quality)
                    mime_type = "text/html"
                    ext = ".html"
                except Exception as e:
                    st.error(f"Image conversion failed: {e}")
            elif method == "Plain Text":
                try:
                    output_bytes = pdf_to_text_bytes(pdf_bytes)
                    mime_type = "text/plain"
                    ext = ".txt"
                except Exception as e:
                    st.error(f"Text extraction failed: {e}")

            # Optionally embed the original PDF in HTML output
            if output_bytes and mime_type == "text/html" and embed_pdf:
                try:
                    output_bytes = embed_original_pdf_in_html(output_bytes, pdf_bytes)
                except Exception:
                    st.warning("Failed to embed original PDF; continuing without embedding.")

            if output_bytes is None:
                st.error("Conversion failed. See messages above.")
            else:
                if mime_type == "text/html":
                    st.subheader("Preview (first 10 MB)")
                    preview_text = output_bytes[:10 * 1024 * 1024]
                    try:
                        st.components.v1.html(preview_text.decode('utf-8', errors='replace'), height=600, scrolling=True)
                    except Exception:
                        st.warning("Could not render the HTML preview in the browser. You can still download the file.")
                elif mime_type == "text/plain":
                    st.subheader("Preview (first 5,000 characters)")
                    preview_text = output_bytes.decode("utf-8", errors="replace")[:5000]
                    st.text_area("Extracted Text", preview_text, height=400)

                default_name = uploaded_file.name.rsplit('.', 1)[0] + ext
                st.download_button(f"Download {ext.upper()[1:]} file", data=output_bytes, file_name=default_name, mime=mime_type)

    st.markdown("---")
    st.markdown("**Tips:**\n- If the text-preserving HTML misses graphics/tables, try the Image fallback.\n- For scanned PDFs, text extraction may fail (consider OCR separately).\n- Large PDFs will produce large HTML files, especially with the image fallback.")
else:
    st.info("Upload a PDF file to get started.")
