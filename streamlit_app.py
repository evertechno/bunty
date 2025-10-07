"""
Production-ready Streamlit converter — structured, "as‑is" output
- PDF -> HTML / DOCX / TXT (tries to preserve visual structure)
- HTML -> DOCX / TXT

Heuristics & improvements in this version:
- Single pdfplumber session for table extraction (faster + accurate)
- Better list detection including ordered vs unordered
- Proper grouping of list items into <ul>/<ol> in HTML output
- Better error handling and logging
- Configurable heuristics in sidebar and per-document tuning
- Optional embedding of the original PDF as an (visually "invisible") embed
- Bulk processing with ThreadPoolExecutor and progress

Requirements:
    pip install streamlit pdfminer.six pymupdf pdfplumber beautifulsoup4 python-docx pandas

Run:
    streamlit run streamlit_converter_structured.py

Notes:
- This targets digital PDFs (embedded text). No OCR is performed.
- If you need bit-exact visual cloning (pixel-perfect), consider exporting to a PDF viewer or keeping original PDF alongside HTML.
"""

import io
import os
import zipfile
import base64
import logging
import re
from typing import List, Tuple, Dict, Any, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt

# -----------------------------
# Logging
# -----------------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("converter")

# -----------------------------
# Heuristics / Utilities
# -----------------------------
BULLET_UNORDERED = r"^[\u2022\u2023\u25E6\-\*\•\–\—]\s+(.*)"
BULLET_ORDERED = r"^(\d{1,3})([\.|\)])\s+(.*)"


def sanitize_text(t: str) -> str:
    return t.replace('\r', '').rstrip()


def is_bullet_line(text: str) -> Tuple[bool, Optional[str]]:
    """Return (is_bullet, 'ul'|'ol'|None)"""
    s = text.strip()
    if re.match(BULLET_UNORDERED, s):
        return True, "ul"
    m = re.match(BULLET_ORDERED, s)
    if m:
        return True, "ol"
    return False, None


def normalize_whitespace(s: str) -> str:
    return re.sub(r"\s+\n", "\n", re.sub(r"\n\s+", "\n", s)).strip()


def choose_heading_levels(unique_sizes: List[float]) -> Dict[float, int]:
    """Map font sizes -> heading levels heuristically. Biggest -> h1, next -> h2..."""
    sizes = sorted(set(unique_sizes), reverse=True)
    mapping: Dict[float, int] = {}
    top = sizes[:4]
    for idx, s in enumerate(top):
        mapping[round(s, 2)] = idx + 1
    return mapping

# -----------------------------
# PDF Parsing (structured)
# -----------------------------

def parse_pdf_structured(pdf_bytes: bytes, min_heading_ratio: float = 1.12) -> Dict[str, Any]:
    """Parse PDF into a structured intermediate representation.

    Output format:
      {"pages": [ {"page_number":int, "elements":[{...}, ...]}, ...], "fontsizes": [...] }

    Elements have types: heading, para, list_item (with list_type), table
    """
    # open fitz doc for text blocks and pdfplumber for tables
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        logger.exception("Failed to open PDF with PyMuPDF")
        raise

    # pre-scan font sizes
    all_sizes: List[float] = []
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
    # lower divisor => more headings. We invert to produce a threshold: sizes >= (max_size / min_heading_ratio)
    heading_threshold = max_size / (min_heading_ratio if min_heading_ratio > 0 else 1.12)

    pages_out: List[Dict[str, Any]] = []

    # Use pdfplumber once for tables (better performance)
    tables_by_page: Dict[int, List[Dict[str, Any]]] = {}
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as ppdf:
            for i, plpage in enumerate(ppdf.pages):
                try:
                    extracted = plpage.extract_tables()
                except Exception:
                    extracted = []
                tables = []
                for t in extracted:
                    # filter out empty-looking tables
                    if t and any(any(cell for cell in row if cell not in (None, "")) for row in t):
                        tables.append({"rows": t})
                tables_by_page[i] = tables
    except Exception:
        logger.exception("pdfplumber table extraction failed")
        tables_by_page = {}

    # parse each page with fitz and inject tables with approximate order (tables appended at end of page)
    for p in range(len(doc)):
        page = doc.load_page(p)
        d = page.get_text("dict")
        elements: List[Dict[str, Any]] = []

        # iterate blocks to capture reading order
        for block in d.get("blocks", []):
            if block.get("type") != 0:
                continue
            # gather block lines and their max span size
            block_lines: List[Tuple[str, float]] = []
            for line in block.get("lines", []):
                line_text = ""
                max_span_sz = 0.0
                for span in line.get("spans", []):
                    stxt = span.get("text", "")
                    if stxt:
                        line_text += stxt
                    sz = span.get("size", 0)
                    if sz and sz > max_span_sz:
                        max_span_sz = sz
                if line_text.strip():
                    block_lines.append((line_text.strip(), max_span_sz))

            # if block_lines empty continue
            if not block_lines:
                continue

            # If block looks like a single-line heading or multiple lines treat individually
            for ln, sz in block_lines:
                ln_clean = sanitize_text(ln)
                is_bullet, list_type = is_bullet_line(ln_clean)
                if is_bullet:
                    # store list items with type
                    # strip leading bullet chars for cleaner content
                    if list_type == "ul":
                        text = re.sub(BULLET_UNORDERED, r"\1", ln_clean).strip()
                    else:
                        text = re.sub(BULLET_ORDERED, r"\3", ln_clean).strip()
                    elements.append({"type": "list_item", "text": text, "size": sz, "list_type": list_type})
                else:
                    mapped_level = font_to_heading.get(round(sz, 2), 0)
                    if (sz >= heading_threshold) or mapped_level:
                        level = mapped_level if mapped_level else 2
                        elements.append({"type": "heading", "text": ln_clean, "level": level, "size": sz})
                    else:
                        # heuristics: short lines in uppercase or titlecase with larger-than-paragraph fonts -> heading
                        if len(ln_clean) < 120 and (ln_clean.isupper() and sz > (max_size * 0.9)):
                            elements.append({"type": "heading", "text": ln_clean, "level": 2, "size": sz})
                        else:
                            elements.append({"type": "para", "text": ln_clean, "size": sz})

        # append page-detected tables (best-effort)
        for t in tables_by_page.get(p, []):
            elements.append({"type": "table", "rows": t["rows"]})

        pages_out.append({"page_number": p + 1, "elements": elements})

    return {"pages": pages_out, "fontsizes": unique_sizes}

# -----------------------------
# Converters: Intermediate -> HTML / DOCX / TEXT
# -----------------------------

def _wrap_list_items_html(elements: List[Dict[str, Any]]) -> str:
    """Emit HTML for a page elements list, grouping consecutive list_items into <ul> or <ol>."""
    out: List[str] = []
    i = 0
    n = len(elements)
    while i < n:
        el = elements[i]
        if el["type"] == "list_item":
            # gather consecutive list items of same type
            list_type = el.get("list_type", "ul")
            tag = "ol" if list_type == "ol" else "ul"
            items: List[str] = []
            while i < n and elements[i]["type"] == "list_item" and elements[i].get("list_type") == list_type:
                items.append(f"<li>{html_escape(elements[i]['text'])}</li>")
                i += 1
            out.append(f"<{tag}>" + "\n".join(items) + f"</{tag}>")
            continue
        elif el["type"] == "heading":
            lvl = min(max(int(el.get("level", 2)), 1), 6)
            out.append(f"<h{lvl}>{html_escape(el['text'])}</h{lvl}>")
        elif el["type"] == "para":
            out.append(f"<p>{html_escape(el['text'])}</p>")
        elif el["type"] == "table":
            rows = el["rows"]
            out.append("<div class=\"table-wrap\">\n<table>")
            for r in rows:
                out.append("<tr>" + "".join(f"<td>{html_escape(str(c) if c is not None else '')}</td>" for c in r) + "</tr>")
            out.append("</table>\n</div>")
        else:
            out.append(f"<p>{html_escape(str(el))}</p>")
        i += 1
    return "\n".join(out)


def html_escape(s: str) -> str:
    return (s.replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('\"', '&quot;'))


def structured_to_html(parsed: dict, embed_pdf: bool = False, pdf_bytes: bytes = None, invisible_embed: bool = False) -> bytes:
    css = (
        "body{font-family:Arial,Helvetica,sans-serif;line-height:1.4;padding:16px;margin:0}"
        "pre{white-space:pre-wrap;}"
        "table{border-collapse:collapse;margin:8px 0;width:100%}"
        "td,th{border:1px solid #ddd;padding:6px}"
        "h1,h2,h3,h4{margin-top:1em;margin-bottom:0.5em}"
        ".page{padding:18px 20px}" 
        ".table-wrap{overflow:auto;margin:8px 0}"
    )

    parts: List[str] = [
        '<!doctype html>', '<html>', '<head>', '<meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width,initial-scale=1">', f'<style>{css}</style>', '</head>', '<body>'
    ]

    for page in parsed["pages"]:
        parts.append(f'<div class="page" data-page="{page["page_number"]}" style="page-break-after:always;">')
        parts.append(f'<div class="page-header" style="display:none">Page {page["page_number"]}</div>')
        page_html = _wrap_list_items_html(page["elements"])
        parts.append(page_html)
        parts.append('</div>')

    # optionally embed original PDF. invisible_embed tries to hide visible chrome (no border, no toolbar in some browsers)
    if embed_pdf and pdf_bytes:
        b64 = base64.b64encode(pdf_bytes).decode('ascii')
        if invisible_embed:
            embed_snip = (
                '<div style="margin-top:8px;">'
                f'<object data="data:application/pdf;base64,{b64}" type="application/pdf" width="100%" height="600px" style="border:none;outline:none;">'
                'Original PDF (fallback link below)'
                '</object>'
                '</div>'
            )
        else:
            embed_snip = (
                '<hr/>'
                '<h2>Original PDF (embedded)</h2>'
                f'<embed src="data:application/pdf;base64,{b64}" width="100%" height="600px"></embed>'
            )
        parts.append(embed_snip)

    parts.append('</body></html>')
    return "\n".join(parts).encode('utf-8')


def structured_to_text(parsed: dict) -> bytes:
    out_lines: List[str] = []
    for page in parsed["pages"]:
        out_lines.append(f"--- PAGE {page['page_number']} ---")
        for el in page["elements"]:
            if el["type"] == "heading":
                out_lines.append(el["text"].upper())
            elif el["type"] == "para":
                out_lines.append(el["text"])
            elif el["type"] == "list_item":
                prefix = "1." if el.get("list_type") == "ol" else "-"
                out_lines.append(f"{prefix} {el['text']}")
            elif el["type"] == "table":
                rows = el["rows"]
                for r in rows:
                    out_lines.append('\t'.join([str(c) if c is not None else "" for c in r]))
            else:
                out_lines.append(str(el))
    joined = "\n".join(out_lines)
    return joined.encode('utf-8')


def structured_to_docx(parsed: dict) -> bytes:
    doc = Document()
    style = doc.styles['Normal']
    try:
        style.font.name = 'Arial'
        style.font.size = Pt(11)
    except Exception:
        # in some python-docx installations setting font name may not work for default style; ignore
        pass

    for page in parsed["pages"]:
        # page header (kept small & optional)
        doc.add_paragraph(f"--- Page {page['page_number']} ---").style = doc.styles['Normal']
        last_was_list = False
        current_list_paras = []

        for el in page["elements"]:
            if el["type"] == "heading":
                lvl = min(max(int(el.get("level", 2)), 1), 4)
                doc.add_heading(el["text"], level=lvl if lvl <= 4 else 4)
                last_was_list = False
            elif el["type"] == "para":
                doc.add_paragraph(el["text"])
                last_was_list = False
            elif el["type"] == "list_item":
                # emulate lists using paragraph styles
                p = doc.add_paragraph(el["text"])
                if el.get("list_type") == "ol":
                    try:
                        p.style = 'List Number'
                    except Exception:
                        p.style = 'List Bullet'
                else:
                    try:
                        p.style = 'List Bullet'
                    except Exception:
                        pass
                last_was_list = True
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
                last_was_list = False
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
    return text.encode('utf-8')


def html_to_docx_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    doc = Document()
    for el in soup.body.descendants:
        if getattr(el, 'name', None) is None:
            continue
        name = el.name.lower()
        if name.startswith('h') and el.get_text(strip=True):
            try:
                level = int(name[1]) if len(name) > 1 and name[1].isdigit() else 2
            except Exception:
                level = 2
            doc.add_heading(el.get_text(strip=True), level=min(level, 4))
        elif name == 'p' and el.get_text(strip=True):
            doc.add_paragraph(el.get_text("\n", strip=True))
        elif name in ('ul', 'ol'):
            for li in el.find_all('li'):
                p = doc.add_paragraph(li.get_text(strip=True))
                if name == 'ul':
                    try:
                        p.style = 'List Bullet'
                    except Exception:
                        pass
                else:
                    try:
                        p.style = 'List Number'
                    except Exception:
                        pass
        elif name == 'table':
            rows = []
            for r in el.find_all('tr'):
                cols = [c.get_text(strip=True) for c in r.find_all(['th', 'td'])]
                rows.append(cols)
            if rows:
                ncols = max(len(r) for r in rows)
                tbl = doc.add_table(rows=0, cols=ncols)
                tbl.style = 'Table Grid'
                for r in rows:
                    row_cells = tbl.add_row().cells
                    for i in range(ncols):
                        row_cells[i].text = r[i] if i < len(r) else ""
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# -----------------------------
# Streamlit App UI
# -----------------------------

st.set_page_config(page_title="Legacy Converter — Preserve Structure", layout="wide", page_icon="📚")
st.title("📚 Legacy Converter — Preserve structure, create a legacy")

st.markdown("""
This app tries to **preserve structure** during conversions:
- Headings (detected via font size)
- Paragraphs (text blocks)
- Lists (bullets / numbered)
- Tables (via pdfplumber)

It uses heuristics — tune thresholds from the sidebar for better results on specific documents.
""")

# Sidebar controls
with st.sidebar:
    st.header("Conversion options")
    conversion = st.selectbox("Conversion", [
        "PDF → Structured HTML",
        "PDF → Word (.docx)",
        "PDF → Plain Text",
        "HTML → Word (.docx)",
        "HTML → Plain Text",
    ])
    workers = st.number_input("Parallel workers (bulk)", min_value=1, max_value=8, value=3)
    embed_pdf = st.checkbox("Embed original PDF into HTML output (only for HTML outputs)", value=False)
    invisible_embed = st.checkbox("Make embedded PDF visually minimal/invisible (no border)", value=True)

    st.markdown("### Heuristics tuning")
    heading_ratio = st.slider("Heading size sensitivity (lower = more headings)", min_value=1.05, max_value=1.5, value=1.12, step=0.01)
    st.markdown("Tip: reduce heading sensitivity if your headings are not being detected; increase to reduce false headings.")

uploaded = st.file_uploader("Upload PDF(s) or HTML(s) — digital PDFs only (no OCR).", type=["pdf", "html"], accept_multiple_files=True)

if not uploaded:
    st.info("Upload at least one file. This converter targets digital PDFs (embedded text) and HTML files.")
    st.stop()

st.markdown(f"Files queued: {len(uploaded)}")
start = st.button("Create legacy — Convert now")
if not start:
    st.stop()

# Process files
results_for_zip: List[Dict[str, Any]] = []


def process_file(uploaded_file) -> Dict[str, Any]:
    name = uploaded_file.name
    raw = uploaded_file.read()
    ext = os.path.splitext(name)[1].lower()
    try:
        if ext == ".pdf":
            parsed = parse_pdf_structured(raw, min_heading_ratio=heading_ratio)
            if conversion == "PDF → Structured HTML":
                out_bytes = structured_to_html(parsed, embed_pdf=embed_pdf, pdf_bytes=raw if embed_pdf else None, invisible_embed=invisible_embed)
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
        logger.exception("Failed processing %s", name)
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
            try:
                log.write(f"✔ {res['name']} → {res['out_name']} ({len(res['out_bytes']):,} bytes)")
            except Exception:
                log.write(f"✔ {res['name']} → {res.get('out_name')} ")

status.success("Conversion jobs finished")

# Show results and allow downloads; also create ZIP for bulk download
if results_for_zip:
    st.markdown("### Download converted files")
    cols = st.columns(3)
    for i, res in enumerate(results_for_zip):
        col = cols[i % 3]
        with col:
            st.write(res["out_name"])
            if res["mime"].startswith("text/"):
                try:
                    preview_text = res["out_bytes"].decode("utf-8", errors="replace")[:4000]
                except Exception:
                    preview_text = "(Preview not available)"
                st.text_area(f"Preview — {res['out_name']}", preview_text, height=180)
            elif res["mime"] == "text/html":
                try:
                    st.components.v1.html(res["out_bytes"].decode("utf-8", errors="replace"), height=360, scrolling=True)
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
st.info("This tool preserves structure using heuristics (font sizes, bullets, table detection). If a document has odd layout or complex multi-column formatting, tweak the 'Heading size sensitivity' slider or ask me to add per-document tuning rules.")
