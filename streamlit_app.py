# streamlit_converter_structured.py
# Streamlit Multi-format Converter — Exact "as-is" + Structured modes
# Targets digital PDFs (embedded text). No OCR is performed.

import io
import os
import re
import sys
import html
import base64
import zipfile
import shutil
import tempfile
import subprocess
from typing import List, Tuple, Dict, Any

import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, Inches
import pandas as pd

# -----------------------------
# Utility / Heuristic Functions
# -----------------------------

BULLET_CHARS = r"^[\u2022\u2023\u25E6\-\*\•\–\—]\s+|^\d+[\.\)]\s+"

def is_bullet_line(text: str) -> bool:
    return bool(re.match(BULLET_CHARS, text.strip()))

def choose_heading_levels(unique_sizes: List[float]) -> Dict[float, int]:
    sizes = sorted(set(unique_sizes), reverse=True)
    mapping = {}
    top = sizes[:4]
    for idx, s in enumerate(top):
        mapping[s] = idx + 1
    return mapping

# -----------------------------
# Exact "as-is" Converters
# -----------------------------

def pdf_to_html_exact_mupdf(pdf_bytes: bytes) -> bytes:
    # Produce a single HTML with each page wrapped in the MuPDF-generated page <div>
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page_divs = []
    for page in doc:
        # get_text("html") returns absolutely positioned HTML reflecting the page layout
        page_html = page.get_text("html")
        page_divs.append(page_html)
    # Assemble into a single HTML document
    head = (
        "<!doctype html><html><head><meta charset='utf-8'>"
        "<meta name='viewport' content='width=device-width,initial-scale=1'>"
        "<style>html,body{margin:0;padding:0;background:#eee} "
        ".page{margin:8px auto;box-shadow:0 0 6px rgba(0,0,0,0.2);background:white} "
        "</style></head><body>"
    )
    # MuPDF already emits <div id='pageX' ...> blocks; wrap them in a container
    body = "".join(page_divs)
    tail = "</body></html>"
    return (head + body + tail).encode("utf-8")

def has_pdf2htmlex() -> bool:
    return shutil.which("pdf2htmlEX") is not None

def pdf_to_html_exact_pdf2htmlex(pdf_bytes: bytes, extra_args: List[str] = None) -> bytes:
    # Requires pdf2htmlEX installed on the system
    if not has_pdf2htmlex():
        raise RuntimeError("pdf2htmlEX binary not found on PATH")
    extra_args = extra_args or []
    with tempfile.TemporaryDirectory() as tmp:
        in_pdf = os.path.join(tmp, "in.pdf")
        out_html = os.path.join(tmp, "out.html")
        with open(in_pdf, "wb") as f:
            f.write(pdf_bytes)
        # Basic invocation; adjust options as desired
        cmd = ["pdf2htmlEX", "--embed-css", "1", "--embed-image", "1", "--embed-font", "1",
               "--optimize-text", "1", "--process-outline", "0", "--hdpi", "144", "--vdpi", "144",
               in_pdf, out_html] + extra_args
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        with open(out_html, "rb") as f:
            return f.read()

def pdf_to_docx_fixed_images(pdf_bytes: bytes, dpi: int = 200) -> bytes:
    # Render each PDF page as an image and place in DOCX with page-width sizing
    docx = Document()
    section = docx.sections[0]
    page_width = section.page_width
    left_margin = section.left_margin
    right_margin = section.right_margin
    content_width = page_width - left_margin - right_margin

    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    for i, page in enumerate(pdf):
        # Compute scale based on DPI: 72 points per inch baseline
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_bytes = io.BytesIO(pix.tobytes("png"))
        # Add image covering content width
        p = docx.add_paragraph()
        run = p.add_run()
        pic = run.add_picture(img_bytes, width=content_width)
        # Page break except after last page
        if i < len(pdf) - 1:
            docx.add_page_break()
    out = io.BytesIO()
    docx.save(out)
    return out.getvalue()

# -----------------------------
# PDF Parsing (structured mode)
# -----------------------------

def parse_pdf_structured(pdf_bytes: bytes, min_heading_ratio: float = 1.12) -> Dict[str, Any]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_out = []
    all_sizes = []

    for p in range(len(doc)):
        page = doc.load_page(p)
        d = page.get_text("dict")
        for block in d.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    sz = round(span.get("size", 0), 2)
                    if sz > 0:
                        all_sizes.append(sz)

    unique_sizes = sorted(set(all_sizes), reverse=True) if all_sizes else [12.0]
    font_to_heading = choose_heading_levels(unique_sizes)
    max_size = max(unique_sizes) if unique_sizes else 12.0
    heading_threshold = max_size / min_heading_ratio

    # Parse each page into elements
    for p in range(len(doc)):
        page = doc.load_page(p)
        d = page.get_text("dict")
        elements = []

        # detect tables with pdfplumber
        tables_page = []
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as ppdf:
                page_pl = ppdf.pages[p]
                extracted_tables = page_pl.extract_tables()
                for t in extracted_tables:
                    if t and any(any(cell for cell in row if cell not in (None, "")) for row in t):
                        tables_page.append({"rows": t})
        except Exception:
            tables_page = []

        # text blocks
        for block in d.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                line_text = ""
                max_span_sz = 0.0
                for span in spans:
                    stxt = span.get("text", "")
                    if stxt:
                        line_text += stxt
                    sz = span.get("size", 0)
                    if sz and sz > max_span_sz:
                        max_span_sz = sz
                ln_clean = line_text.strip()
                if not ln_clean:
                    continue
                if is_bullet_line(ln_clean):
                    elements.append({"type": "list_item", "text": ln_clean, "size": max_span_sz})
                else:
                    mapped_level = font_to_heading.get(round(max_span_sz, 2), 0)
                    if (max_span_sz >= heading_threshold) or mapped_level:
                        level = mapped_level if mapped_level else 2
                        elements.append({"type": "heading", "text": ln_clean, "level": level, "size": max_span_sz})
                    else:
                        elements.append({"type": "para", "text": ln_clean, "size": max_span_sz})

        # append detected tables (order is approximate)
        for t in tables_page:
            elements.append({"type": "table", "rows": t["rows"]})

        pages_out.append({"page_number": p + 1, "elements": elements})

    return {"pages": pages_out, "fontsizes": unique_sizes}

# -----------------------------
# Structured -> HTML / DOCX / TEXT
# -----------------------------

def structured_to_html(parsed: dict, embed_pdf: bool = False, pdf_bytes: bytes = None) -> bytes:
    parts = ['<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>body{font-family:Arial,Helvetica,sans-serif;line-height:1.4;padding:16px}pre{white-space:pre-wrap;}table{border-collapse:collapse;margin:8px 0}td,th{border:1px solid #ccc;padding:6px}</style></head><body>']
    li_buffer = []
    def flush_list():
        if not li_buffer:
            return ""
        html_ul = "<ul>" + "".join(li_buffer) + "</ul>"
        li_buffer.clear()
        return html_ul

    for page in parsed["pages"]:
        parts.append(f'<div class="page" data-page="{page["page_number"]}" style="page-break-after:always;">')
        for el in page["elements"]:
            if el["type"] == "heading":
                parts.append(flush_list())
                lvl = min(max(int(el.get("level", 2)), 1), 6)
                text = html.escape(el["text"])
                parts.append(f"<h{lvl}>{text}</h{lvl}>")
            elif el["type"] == "para":
                parts.append(flush_list())
                text = html.escape(el["text"])
                parts.append(f"<p>{text}</p>")
            elif el["type"] == "list_item":
                li_buffer.append(f"<li>{html.escape(el['text'])}</li>")
            elif el["type"] == "table":
                parts.append(flush_list())
                rows = el["rows"]
                parts.append("<table>")
                for r in rows:
                    parts.append("<tr>" + "".join(f"<td>{html.escape(str(c) if c is not None else '')}</td>" for c in r) + "</tr>")
                parts.append("</table>")
        parts.append(flush_list())
        parts.append("</div>")
    html_text = "\n".join(parts) + "</body></html>"

    if embed_pdf and pdf_bytes:
        b64 = base64.b64encode(pdf_bytes).decode('ascii')
        embed_snip = f'<hr/><h2>Original PDF (embedded)</h2><embed src="data:application/pdf;base64,{b64}" width="100%" height="600px"></embed>'
        html_text = html_text.replace("</body></html>", embed_snip + "</body></html>")
    return html_text.encode("utf-8")

def structured_to_text(parsed: dict) -> bytes:
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
                rows = el["rows"]
                for r in rows:
                    out_lines.append("\t".join([str(c) if c is not None else "" for c in r]))
    joined = "\n".join(out_lines)
    return joined.encode("utf-8")

def structured_to_docx(parsed: dict) -> bytes:
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    for page in parsed["pages"]:
        doc.add_paragraph(f"--- Page {page['page_number']} ---").style = doc.styles['Normal']
        for el in page["elements"]:
            if el["type"] == "heading":
                lvl = min(max(int(el.get("level", 2)), 1), 4)
                doc.add_heading(el["text"], level=lvl if lvl <= 4 else 4)
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
        doc.add_page_break()
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# -----------------------------
# HTML -> Text / DOCX
# -----------------------------

def html_to_text_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    text = soup.get_text(separator="\n")
    return text.encode("utf-8")

def html_to_docx_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    doc = Document()
    if soup.body is None:
        # Fallback: wrap entire HTML as text
        doc.add_paragraph(soup.get_text("\n", strip=True))
    else:
        for el in soup.body.descendants:
            if getattr(el, "name", None) and el.name.startswith("h") and el.get_text(strip=True):
                level = int(el.name[1]) if len(el.name) > 1 and el.name[1].isdigit() else 2
                doc.add_heading(el.get_text(strip=True), level=min(level, 4))
            elif getattr(el, "name", None) == "p" and el.get_text(strip=True):
                doc.add_paragraph(el.get_text("\n", strip=True))
            elif getattr(el, "name", None) in ("ul", "ol"):
                for li in el.find_all("li"):
                    p = doc.add_paragraph(li.get_text(strip=True))
                    p.style = 'List Bullet' if el.name == "ul" else 'List Number'
            elif getattr(el, "name", None) == "table":
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
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# -----------------------------
# Streamlit App UI
# -----------------------------

st.set_page_config(page_title="Legacy Converter — Exact + Structured", layout="wide", page_icon="📚")
st.title("📚 Legacy Converter — Exact 'as-is' and Structured outputs")

pdf2htmlEX_available = has_pdf2htmlex()

conversion_options = [
    "PDF → HTML (Exact, MuPDF)",
]
if pdf2htmlEX_available:
    conversion_options.append("PDF → HTML (Exact, pdf2htmlEX)")
conversion_options.extend([
    "PDF → Structured HTML",
    "PDF → Word (.docx, fixed-layout images)",
    "PDF → Word (.docx, structured)",
    "PDF → Plain Text (raw pages)",
    "HTML → Word (.docx)",
    "HTML → Plain Text",
])

st.markdown("""
This tool provides two goals: exact, layout-faithful cloning for visual identity, and structured outputs for downstream editing and analysis. Adjust modes based on whether pixel accuracy or semantic structure is the priority.
""")

with st.sidebar:
    st.header("Conversion options")
    conversion = st.selectbox("Conversion", conversion_options)
    workers = st.number_input("Parallel workers (bulk)", min_value=1, max_value=8, value=3)
    embed_pdf = st.checkbox("Embed original PDF into HTML output (only for structured HTML)", value=False)

    st.markdown("### Structured mode heuristics")
    heading_ratio = st.slider("Heading size sensitivity (lower = more headings)", min_value=1.05, max_value=1.5, value=1.12, step=0.01)

    st.markdown("### Exact mode settings")
    exact_dpi = st.slider("DOCX fixed-layout render DPI", min_value=96, max_value=300, value=200, step=4)
    st.caption("Higher DPI improves image quality and file size for fixed-layout DOCX.")

uploaded = st.file_uploader("Upload PDF(s) or HTML(s) — digital PDFs only (no OCR).", type=["pdf", "html"], accept_multiple_files=True)

if not uploaded:
    st.info("Upload at least one file.")
    st.stop()

st.markdown(f"Files queued: {len(uploaded)}")
start = st.button("Convert now")

if not start:
    st.stop()

from concurrent.futures import ThreadPoolExecutor, as_completed
results_for_zip = []

def process_file(uploaded_file):
    name = uploaded_file.name
    raw = uploaded_file.read()
    ext = os.path.splitext(name)[1].lower()
    try:
        if ext == ".pdf":
            if conversion == "PDF → HTML (Exact, MuPDF)":
                out_bytes = pdf_to_html_exact_mupdf(raw)
                out_name = os.path.splitext(name)[0] + ".html"
                mime = "text/html"
            elif conversion == "PDF → HTML (Exact, pdf2htmlEX)":
                out_bytes = pdf_to_html_exact_pdf2htmlex(raw)
                out_name = os.path.splitext(name)[0] + ".html"
                mime = "text/html"
            elif conversion == "PDF → Structured HTML":
                parsed = parse_pdf_structured(raw, min_heading_ratio=heading_ratio)
                out_bytes = structured_to_html(parsed, embed_pdf=embed_pdf, pdf_bytes=raw if embed_pdf else None)
                out_name = os.path.splitext(name)[0] + ".html"
                mime = "text/html"
            elif conversion == "PDF → Word (.docx, fixed-layout images)":
                out_bytes = pdf_to_docx_fixed_images(raw, dpi=exact_dpi)
                out_name = os.path.splitext(name)[0] + ".docx"
                mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif conversion == "PDF → Word (.docx, structured)":
                parsed = parse_pdf_structured(raw, min_heading_ratio=heading_ratio)
                out_bytes = structured_to_docx(parsed)
                out_name = os.path.splitext(name)[0] + ".docx"
                mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif conversion == "PDF → Plain Text (raw pages)":
                # Raw text per page, separated by form feed, no normalization
                doc = fitz.open(stream=raw, filetype="pdf")
                text = chr(12).join([page.get_text("text") for page in doc])
                out_bytes = text.encode("utf-8", errors="replace")
                out_name = os.path.splitext(name)[0] + ".txt"
                mime = "text/plain"
            else:
                return {"name": name, "error": f"Conversion {conversion} not valid for PDF"}
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
                return {"name": name, "error": f"Conversion {conversion} not valid for HTML"}
        else:
            return {"name": name, "error": "Unsupported file type"}
        return {"name": name, "out_bytes": out_bytes, "out_name": out_name, "mime": mime}
    except Exception as e:
        return {"name": name, "error": str(e)}

progress = st.progress(0)
status = st.empty()
log = st.empty()

with ThreadPoolExecutor(max_workers=workers) as exe:
    futures = {exe.submit(process_file, f): f.name for f in uploaded}
    done = 0
    for fut in as_completed(futures):
        done += 1
        res = fut.result()
        progress.progress(done / len(uploaded))
        if res.get("error"):
            log.write(f"✖ {res['name']} — {res['error']}")
        else:
            results_for_zip.append(res)
            log.write(f"✔ {res['name']} → {res['out_name']} ({len(res['out_bytes']):,} bytes)")
    status.success("Conversion jobs finished")

if results_for_zip:
    st.markdown("### Download converted files")
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
                    st.components.v1.html(res["out_bytes"].decode("utf-8", errors="replace")[:200000], height=350, scrolling=True)
                except Exception:
                    st.write("(HTML preview failed; download instead.)")
            st.download_button("Download", data=res["out_bytes"], file_name=res["out_name"], mime=res["mime"])

    # ZIP
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in results_for_zip:
            zf.writestr(r["out_name"], r["out_bytes"])
    zip_buf.seek(0)
    st.download_button("Download ALL as ZIP", zip_buf.read(), file_name="converted_legacy.zip", mime="application/zip")
else:
    st.error("No successful conversions to download. Check logs above.")

st.markdown("---")
st.info("Exact modes preserve layout and punctuation as-is; choose structured modes for editable semantics like headings, lists, and tables.")
