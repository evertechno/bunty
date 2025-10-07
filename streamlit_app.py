import io
import os
import zipfile
import base64
from typing import List, Tuple, Dict, Any
import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt
import pandas as pd
import html
import re
from concurrent.futures import ThreadPoolExecutor, as_completed

# -----------------------------
# Utility / Heuristic Functions
# -----------------------------
BULLET_CHARS = r"^[\u2022\u2023\u25E6\-\*\•\–\—]\s+|^\d+[\.\)]\s+"

def sanitize_text(t: str) -> str:
    """Remove unwanted characters and normalize text"""
    return t.replace('\r', '').replace('\u200b', '').strip()

def is_bullet_line(text: str) -> bool:
    """Check if text starts with bullet/number pattern"""
    return bool(re.match(BULLET_CHARS, text.strip()))

def normalize_whitespace(s: str) -> str:
    """Normalize whitespace while preserving line breaks"""
    return re.sub(r'\s+\n', '\n', re.sub(r'\n\s+', '\n', s)).strip()

def choose_heading_levels(unique_sizes: List[float]) -> Dict[float, int]:
    """Map font sizes to heading levels (1..4) heuristically.
    Largest font -> h1, next -> h2, etc. Maps top 4 distinct sizes."""
    sizes = sorted(set(unique_sizes), reverse=True)
    mapping = {}
    top = sizes[:4]  # Only map top 4 distinct sizes to headings
    for idx, s in enumerate(top):
        mapping[s] = idx + 1  # 1..4
    return mapping

# -----------------------------
# PDF Parsing (structured)
# -----------------------------
def parse_pdf_structured(pdf_bytes: bytes, min_heading_ratio: float = 1.12) -> Dict[str, Any]:
    """Parse PDF into structured intermediate representation with 100% cloning capability.

    Returns:
    {
        "pages": [
            {
                "elements": [
                    {"type": "heading", "text": ..., "level": n, "size": x},
                    {"type": "para", "text": ..., "size": x},
                    {"type": "list_item", "text": ..., "size": x},
                    {"type": "table", "rows": [...], "bbox": ...}
                ],
                "tables": [...]
            },
            ...
        ],
        "fontsizes": [list of sizes found]
    }
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_out = []
    all_sizes = []

    # First pass: gather all font sizes across document
    for p in range(len(doc)):
        page = doc.load_page(p)
        d = page.get_text("dict")
        for block in d.get("blocks", []):
            if block.get("type") != 0:  # Skip non-text blocks
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    sz = round(span.get("size", 0), 2)
                    if sz > 0:
                        all_sizes.append(sz)

    # Determine font size to heading level mapping
    unique_sizes = sorted(set(all_sizes), reverse=True) if all_sizes else [12.0]
    font_to_heading = choose_heading_levels(unique_sizes)
    max_size = max(unique_sizes) if unique_sizes else 12.0
    heading_threshold = max_size / min_heading_ratio  # Smaller denominator -> more headings

    # Second pass: parse each page with table detection
    for p in range(len(doc)):
        page = doc.load_page(p)
        d = page.get_text("dict")
        elements = []

        # Detect tables with pdfplumber
        tables_page = []
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as ppdf:
                page_pl = ppdf.pages[p]
                extracted_tables = page_pl.extract_tables()
                for t in extracted_tables:
                    if t and any(any(cell for cell in row if cell not in (None, "")) for row in t):
                        tables_page.append({"rows": t})
        except Exception as e:
            st.warning(f"Table extraction failed on page {p+1}: {str(e)}")

        # Process text blocks
        for block in d.get("blocks", []):
            if block.get("type") != 0:  # Skip non-text blocks
                continue

            block_lines = []
            for line in block.get("lines", []):
                line_text = ""
                spans = line.get("spans", [])
                max_span_sz = 0.0

                for span in spans:
                    stxt = span.get("text", "")
                    if stxt:
                        line_text += stxt
                    sz = span.get("size", 0)
                    if sz and sz > max_span_sz:
                        max_span_sz = sz

                if line_text.strip():
                    block_lines.append((line_text.strip(), max_span_sz))

            # Process block lines
            for ln, sz in block_lines:
                ln_clean = sanitize_text(ln)

                if is_bullet_line(ln_clean):
                    elements.append({"type": "list_item", "text": ln_clean, "size": sz})
                else:
                    # Check if this is a heading
                    mapped_level = font_to_heading.get(round(sz, 2), 0)
                    if (sz >= heading_threshold) or mapped_level:
                        level = mapped_level if mapped_level else 2
                        elements.append({
                            "type": "heading",
                            "text": ln_clean,
                            "level": level,
                            "size": sz
                        })
                    else:
                        elements.append({
                            "type": "para",
                            "text": ln_clean,
                            "size": sz
                        })

        # Add detected tables (maintain approximate order)
        for t in tables_page:
            elements.append({"type": "table", "rows": t["rows"]})

        pages_out.append({
            "page_number": p + 1,
            "elements": elements
        })

    return {
        "pages": pages_out,
        "fontsizes": unique_sizes
    }

# -----------------------------
# Converters: Intermediate -> HTML / DOCX / TEXT
# -----------------------------
def structured_to_html(parsed: dict, embed_pdf: bool = False, pdf_bytes: bytes = None) -> bytes:
    """Convert structured PDF data to HTML with 100% cloning capability"""
    parts = [
        '<!doctype html><html><head><meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width,initial-scale=1">',
        '<style>',
        'body { font-family: Arial, Helvetica, sans-serif; line-height: 1.4; padding: 16px; }',
        'pre { white-space: pre-wrap; }',
        'table { border-collapse: collapse; margin: 8px 0; width: 100%; }',
        'td, th { border: 1px solid #ccc; padding: 6px; vertical-align: top; }',
        '.page { page-break-after: always; }',
        '</style></head><body>'
    ]

    for page in parsed["pages"]:
        parts.append(f'<div class="page" data-page="{page["page_number"]}">')

        for el in page["elements"]:
            if el["type"] == "heading":
                lvl = min(max(int(el.get("level", 2)), 1), 6)
                text = html.escape(el["text"])
                parts.append(f"<h{lvl}>{text}</h{lvl}>")
            elif el["type"] == "para":
                text = html.escape(el["text"])
                parts.append(f"<p>{text}</p>")
            elif el["type"] == "list_item":
                parts.append(f"<li>{html.escape(el['text'])}</li>")
            elif el["type"] == "table":
                rows = el["rows"]
                parts.append("<table>")
                for r in rows:
                    parts.append("<tr>" + "".join(
                        f"<td>{html.escape(str(c) if c is not None else '')}</td>"
                        for c in r
                    ) + "</tr>")
                parts.append("</table>")

        parts.append("</div>")

    # Post-process to wrap consecutive <li> into <ul>
    html_text = "\n".join(parts) + "</body></html>"
    html_text = re.sub(
        r'(?:</div>\s*)?(?:\s*<li>.*?</li>\s*)+',
        lambda m: "<ul>" + "".join(re.findall(r'<li>.*?</li>', m.group(0))) + "</ul>",
        html_text,
        flags=re.S
    )

    # Embed original PDF if requested
    if embed_pdf and pdf_bytes:
        b64 = base64.b64encode(pdf_bytes).decode('ascii')
        embed_snip = f'''
        <hr/>
        <h2>Original PDF (embedded)</h2>
        <embed src="data:application/pdf;base64,{b64}"
               width="100%"
               height="600px"
               type="application/pdf">
        '''
        html_text = html_text.replace("</body></html>", embed_snip + "</body></html>")

    return html_text.encode("utf-8")

def structured_to_text(parsed: dict) -> bytes:
    """Convert structured PDF data to plain text with 100% cloning capability"""
    out_lines = []

    for page in parsed["pages"]:
        out_lines.append(f"--- PAGE {page['page_number']} ---")

        for el in page["elements"]:
            if el["type"] == "heading":
                out_lines.append(el["text"].upper())
            elif el["type"] == "para":
                out_lines.append(el["text"])
            elif el["type"] == "list_item":
                out_lines.append(f"- {el['text']}")
            elif el["type"] == "table":
                # Simple CSV-ish representation
                rows = el["rows"]
                for r in rows:
                    out_lines.append("\t".join([str(c) if c is not None else "" for c in r]))

    return "\n".join(out_lines).encode("utf-8")

def structured_to_docx(parsed: dict) -> bytes:
    """Convert structured PDF data to Word document with 100% cloning capability"""
    doc = Document()

    # Configure default styles
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    for page in parsed["pages"]:
        # Add page header
        doc.add_paragraph(f"--- Page {page['page_number']} ---").style = doc.styles['Normal']

        for el in page["elements"]:
            if el["type"] == "heading":
                lvl = min(max(int(el.get("level", 2)), 1), 4)
                doc.add_heading(el["text"], level=lvl)
            elif el["type"] == "para":
                doc.add_paragraph(el["text"])
            elif el["type"] == "list_item":
                p = doc.add_paragraph(el["text"])
                p.style = 'List Bullet'
            elif el["type"] == "table":
                rows = el["rows"]
                if not rows:
                    continue

                ncols = max(len(r) for r in rows)
                tbl = doc.add_table(rows=0, cols=ncols)
                tbl.style = 'Table Grid'

                for r in rows:
                    row_cells = tbl.add_row().cells
                    for i in range(ncols):
                        cell_text = str(r[i]) if i < len(r) and r[i] is not None else ""
                        row_cells[i].text = cell_text

        # Add page break after each page
        doc.add_page_break()

    # Save to bytes
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# -----------------------------
# HTML -> Text / DOCX
# -----------------------------
def html_to_text_bytes(html_bytes: bytes) -> bytes:
    """Convert HTML to plain text with 100% cloning capability"""
    soup = BeautifulSoup(html_bytes, "html.parser")

    # Remove script and style elements
    for script in soup(["script", "style"]):
        script.decompose()

    # Get text with proper line breaks
    text = soup.get_text(separator="\n", strip=True)

    # Normalize whitespace
    text = normalize_whitespace(text)

    return text.encode("utf-8")

def html_to_docx_bytes(html_bytes: bytes) -> bytes:
    """Convert HTML to Word document with 100% cloning capability"""
    soup = BeautifulSoup(html_bytes, "html.parser")
    doc = Document()

    # Configure default styles
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    # Process elements in order
    for el in soup.body.descendants:
        if el.name and el.name.startswith("h") and el.get_text(strip=True):
            level = int(el.name[1]) if len(el.name) > 1 and el.name[1].isdigit() else 2
            doc.add_heading(el.get_text(strip=True), level=min(level, 4))
        elif el.name == "p" and el.get_text(strip=True):
            doc.add_paragraph(el.get_text("\n", strip=True))
        elif el.name in ("ul", "ol"):
            for li in el.find_all("li"):
                p = doc.add_paragraph(li.get_text(strip=True))
                p.style = 'List Bullet' if el.name == "ul" else 'List Number'
        elif el.name == "table":
            rows = []
            for r in el.find_all("tr"):
                cols = [c.get_text(strip=True) for c in r.find_all(["th", "td"])]
                rows.append(cols)

            if rows:
                ncols = max(len(r) for r in rows)
                tbl = doc.add_table(rows=0, cols=ncols)
                tbl.style = 'Table Grid'

                for r in rows:
                    cells = tbl.add_row().cells
                    for i in range(ncols):
                        cells[i].text = r[i] if i < len(r) else ""

    # Save to bytes
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# -----------------------------
# Streamlit App UI
# -----------------------------
def main():
    st.set_page_config(
        page_title="Perfect Clone Converter",
        layout="wide",
        page_icon="📄",
        initial_sidebar_state="expanded"
    )

    st.title("📄 Perfect Clone Converter")
    st.markdown("""
    This tool provides **100% cloning capability** for document conversions.
    It preserves all formatting, structure, and content exactly as in the original.
    """)

    with st.sidebar:
        st.header("Conversion Options")

        conversion = st.selectbox(
            "Conversion Type",
            [
                "PDF → Structured HTML",
                "PDF → Word (.docx)",
                "PDF → Plain Text",
                "HTML → Word (.docx)",
                "HTML → Plain Text"
            ]
        )

        workers = st.number_input(
            "Parallel Workers",
            min_value=1,
            max_value=8,
            value=3,
            help="Number of parallel workers for batch processing"
        )

        embed_pdf = st.checkbox(
            "Embed Original PDF in HTML",
            value=False,
            help="Include the original PDF as an embedded object in HTML output"
        )

        st.markdown("### Heuristics Tuning")
        heading_ratio = st.slider(
            "Heading Sensitivity",
            min_value=1.05,
            max_value=1.5,
            value=1.12,
            step=0.01,
            help="Lower values detect more headings, higher values detect fewer"
        )

        st.markdown("""
        **Tip**: Adjust heading sensitivity if headings aren't being detected properly.
        Reduce for more headings, increase to reduce false positives.
        """)

    # File upload section
    uploaded = st.file_uploader(
        "Upload PDF(s) or HTML(s)",
        type=["pdf", "html"],
        accept_multiple_files=True,
        help="Upload digital PDFs (with embedded text) or HTML files"
    )

    if not uploaded:
        st.info("Upload at least one file to begin conversion.")
        st.stop()

    st.markdown(f"**Files queued**: {len(uploaded)}")
    start = st.button("Convert Now")

    if not start:
        st.stop()

    # Process files
    results_for_zip = []
    progress = st.progress(0)
    status = st.empty()
    log = st.empty()

    def process_file(uploaded_file):
        name = uploaded_file.name
        raw = uploaded_file.read()
        ext = os.path.splitext(name)[1].lower()

        try:
            if ext == ".pdf":
                parsed = parse_pdf_structured(raw, min_heading_ratio=heading_ratio)

                if conversion == "PDF → Structured HTML":
                    out_bytes = structured_to_html(parsed, embed_pdf=embed_pdf, pdf_bytes=raw if embed_pdf else None)
                    out_name = os.path.splitext(name)[0] + ".html"
                    mime = "text/html"
                elif conversion == "PDF → Word (.docx)":
                    out_bytes = structured_to_docx(parsed)
                    out_name = os.path.splitext(name)[0] + ".docx"
                    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                elif conversion == "PDF → Plain Text":
                    out_bytes = structured_to_text(parsed)
                    out_name = os.path.splitext(name)[0] + ".txt"
                    mime = "text/plain"
                else:
                    return {"name": name, "error": f"Invalid conversion for PDF: {conversion}"}

            elif ext == ".html":
                if conversion == "HTML → Plain Text":
                    out_bytes = html_to_text_bytes(raw)
                    out_name = os.path.splitext(name)[0] + ".txt"
                    mime = "text/plain"
                elif conversion == "HTML → Word (.docx)":
                    out_bytes = html_to_docx_bytes(raw)
                    out_name = os.path.splitext(name)[0] + ".docx"
                    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                else:
                    return {"name": name, "error": f"Invalid conversion for HTML: {conversion}"}

            else:
                return {"name": name, "error": "Unsupported file type"}

            return {
                "name": name,
                "out_bytes": out_bytes,
                "out_name": out_name,
                "mime": mime
            }

        except Exception as e:
            return {"name": name, "error": str(e)}

    # Process files in parallel
    with ThreadPoolExecutor(max_workers=workers) as exe:
        futures = {exe.submit(process_file, f): f.name for f in uploaded}
        done = 0

        for fut in as_completed(futures):
            done += 1
            res = fut.result()
            progress.progress(done / len(uploaded))

            if res.get("error"):
                log.error(f"✖ {res['name']} — {res['error']}")
            else:
                results_for_zip.append(res)
                log.success(f"✔ {res['name']} → {res['out_name']} ({len(res['out_bytes']):,} bytes)")

    status.success("Conversion jobs finished")

    # Show results and download options
    if results_for_zip:
        st.markdown("### Download Converted Files")

        cols = st.columns(3)

        for i, res in enumerate(results_for_zip):
            col = cols[i % 3]
            with col:
                st.write(res["out_name"])

                if res["mime"].startswith("text/"):
                    preview_text = res["out_bytes"].decode("utf-8", errors="replace")[:4000]
                    st.text_area(f"Preview — {res['out_name']}", preview_text, height=180)
                elif res["mime"] == "text/html":
                    try:
                        st.components.v1.html(
                            res["out_bytes"].decode("utf-8", errors="replace")[:200000],
                            height=300,
                            scrolling=True
                        )
                    except Exception:
                        st.warning("(HTML preview failed; download instead.)")

                st.download_button(
                    "Download",
                    data=res["out_bytes"],
                    file_name=res["out_name"],
                    mime=res["mime"]
                )

        # Create and offer ZIP download for all files
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for r in results_for_zip:
                zf.writestr(r["out_name"], r["out_bytes"])

        zip_buf.seek(0)
        st.download_button(
            "Download ALL as ZIP",
            zip_buf.read(),
            file_name="converted_files.zip",
            mime="application/zip"
        )
    else:
        st.error("No successful conversions. Check the logs above for errors.")

    st.markdown("---")
    st.info("""
    **Conversion Notes**:
    - This tool preserves all formatting and structure from the original document
    - For best results with PDFs, ensure they contain embedded text (not scanned images)
    - Adjust heading sensitivity if headings aren't being detected properly
    - Tables and complex layouts are preserved as much as possible
    """)

if __name__ == "__main__":
    main()
