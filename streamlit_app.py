# streamlit_converter_structured_prod.py
"""
Production-ready Streamlit converter:
- Structured PDF -> HTML / DOCX / TXT (semantic)
- HTML -> DOCX / TXT
- Pixel-perfect options:
    * Embed original PDF (data URI)
    * Rasterize pages to PNG and generate image-based HTML
Notes:
- Targets digital PDFs (embedded text). No OCR included.
- For OCR workflows, integrate Tesseract or an OCR API before parsing.
"""

import io
import os
import re
import base64
import logging
import zipfile
from typing import List, Dict, Any, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt
import html
import traceback

# -------------------------
# Logging
# -------------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("converter")

# -------------------------
# Heuristics & Utilities
# -------------------------
BULLET_CHARS = r"^\s*([\u2022\u2023\u25E6\-\*\•\–\—]|\d+[\.\)])\s+"

def sanitize_text(t: str) -> str:
    # preserve punctuation and as-is characters, remove only CR and trailing spaces
    return t.replace('\r', '').rstrip()

def is_bullet_line(text: str) -> Tuple[bool, str]:
    m = re.match(BULLET_CHARS, text)
    if m:
        marker = m.group(1)
        return True, marker
    return False, ""

def choose_heading_levels(unique_sizes: List[float]) -> Dict[float, int]:
    sizes = sorted(set(unique_sizes), reverse=True)
    mapping = {}
    top = sizes[:6]  # allow mapping up to top 6 sizes
    # map top sizes to heading levels (range 1..4); others map to 0 (paragraph)
    for idx, s in enumerate(top):
        if idx < 4:
            mapping[round(s, 2)] = idx + 1
        else:
            mapping[round(s, 2)] = 0
    return mapping

def sort_blocks_reading_order(blocks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    # Sort blocks primarily by top (y0), secondarily by left (x0)
    return sorted(blocks, key=lambda b: (round(b.get("y0", 0), 2), round(b.get("x0", 0), 2)))

# -------------------------
# PDF Parsing (Improved)
# -------------------------
def parse_pdf_structured(pdf_bytes: bytes, min_heading_ratio: float = 1.12, detect_tables: bool = True) -> Dict[str, Any]:
    """
    Parse PDF into structured intermediate representation with bbox ordering and table placement.
    Returns:
      { "pages": [ {"page_number": n, "elements": [ {"type": "...", "text":..., "bbox": (x0,y0,x1,y1), ...}, ... ] }, ... ],
        "fontsizes": [...]
      }
    """
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        logger.exception("Failed to open PDF with PyMuPDF")
        raise

    # collect sizes
    all_sizes = []
    per_page_blocks: List[List[Dict[str, Any]]] = []

    for p_idx in range(len(doc)):
        page = doc.load_page(p_idx)
        d = page.get_text("dict")
        page_blocks = []
        for block in d.get("blocks", []):
            if block.get("type") != 0:
                continue
            bbox = block.get("bbox", [0, 0, 0, 0])
            x0, y0, x1, y1 = bbox
            # collect each line's text and max span size for heading heuristics
            for line in block.get("lines", []):
                line_text = ""
                max_span_sz = 0.0
                for span in line.get("spans", []):
                    txt = span.get("text", "")
                    if txt:
                        line_text += txt
                    sz = span.get("size", 0)
                    if sz and sz > max_span_sz:
                        max_span_sz = sz
                if line_text.strip():
                    pt = {
                        "text": sanitize_text(line_text),
                        "size": round(max_span_sz, 2),
                        "bbox": tuple(line.get("bbox", bbox)),
                        "x0": float(line.get("bbox", bbox)[0]),
                        "y0": float(line.get("bbox", bbox)[1]),
                        "x1": float(line.get("bbox", bbox)[2]),
                        "y1": float(line.get("bbox", bbox)[3]),
                    }
                    page_blocks.append(pt)
                    if max_span_sz > 0:
                        all_sizes.append(max_span_sz)
        per_page_blocks.append(page_blocks)

    unique_sizes = sorted(set([round(s, 2) for s in all_sizes]), reverse=True) if all_sizes else [12.0]
    font_to_heading = choose_heading_levels(unique_sizes)
    max_size = max(unique_sizes) if unique_sizes else 12.0
    heading_threshold = max_size / min_heading_ratio

    pages_out = []

    # use pdfplumber to extract page tables (if enabled)
    pdfplumber_tables_by_page = {}
    if detect_tables:
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as ppdf:
                for p_idx in range(len(ppdf.pages)):
                    page_pl = ppdf.pages[p_idx]
                    extracted_tables = page_pl.extract_tables()
                    # keep only non-empty tables
                    tables_clean = []
                    for t in extracted_tables or []:
                        if t and any(any(cell for cell in row if cell not in (None, "")) for row in t):
                            tables_clean.append({"rows": t, "y0": page_pl.bbox[1] if hasattr(page_pl, "bbox") else 0})
                    pdfplumber_tables_by_page[p_idx] = tables_clean
        except Exception:
            logger.exception("pdfplumber table extraction failed, skipping tables.")
            pdfplumber_tables_by_page = {}

    # Build page elements in reading order and merge tables near their vertical position
    for p_idx, blocks in enumerate(per_page_blocks):
        page_elements = []
        # sort lines by y0 then x0
        blocks_sorted = sorted(blocks, key=lambda b: (round(b["y0"], 2), round(b["x0"], 2)))
        # group consecutive bullets into lists while reading
        for b in blocks_sorted:
            txt = b["text"].strip()
            sz = b["size"] if b.get("size") else 0.0
            is_bullet, marker = is_bullet_line(txt)
            if is_bullet:
                # remove the bullet characters only from start so inner punctuation preserved
                cleaned = re.sub(BULLET_CHARS, "", txt, count=1)
                page_elements.append({"type": "list_item", "text": cleaned, "marker": marker, "size": sz, "bbox": b["bbox"], "x0": b["x0"], "y0": b["y0"]})
            else:
                mapped_level = font_to_heading.get(round(sz, 2), 0)
                if (sz >= heading_threshold) or mapped_level:
                    lvl = mapped_level if mapped_level else 2
                    page_elements.append({"type": "heading", "text": txt, "level": lvl, "size": sz, "bbox": b["bbox"], "x0": b["x0"], "y0": b["y0"]})
                else:
                    page_elements.append({"type": "para", "text": txt, "size": sz, "bbox": b["bbox"], "x0": b["x0"], "y0": b["y0"]})

        # Insert tables into page_elements based on approximate vertical position (best-effort)
        tables_here = pdfplumber_tables_by_page.get(p_idx, [])
        for t in tables_here:
            # we don't have exact bbox from pdfplumber easily, so append at end but try to place near last element
            page_elements.append({"type": "table", "rows": t["rows"], "bbox": None})

        pages_out.append({"page_number": p_idx + 1, "elements": page_elements})

    return {"pages": pages_out, "fontsizes": unique_sizes}

# -------------------------
# Converters: Intermediate -> HTML / DOCX / TEXT
# -------------------------

def structured_to_html(parsed: dict, embed_pdf: bool = False, pdf_bytes: bytes = None, rasterize_images: bool = False, image_dpi: int = 150) -> bytes:
    """
    Build HTML with semantic tags in reading-order.
    If rasterize_images is True, generate PNG for each page and insert <img> for pixel-perfect pages.
    """
    parts = [
        '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">',
        '<style>body{font-family:Arial,Helvetica,sans-serif;line-height:1.45;padding:18px;max-width:1000px;margin:auto}pre{white-space:pre-wrap}table{border-collapse:collapse;margin:8px 0}td,th{border:1px solid #ccc;padding:6px;text-align:left}hr{border:none;border-top:1px solid #eee;margin:18px 0}</style>',
        '</head><body>'
    ]

    # pixel-perfect HTML via PNG images
    if rasterize_images and pdf_bytes:
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            for p_idx in range(len(doc)):
                page = doc.load_page(p_idx)
                mat = fitz.Matrix(image_dpi / 72.0, image_dpi / 72.0)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img_bytes = pix.tobytes("png")
                b64 = base64.b64encode(img_bytes).decode("ascii")
                parts.append(f'<div class="page" data-page="{p_idx+1}" style="page-break-after:always;margin-bottom:24px;">')
                parts.append(f'<img src="data:image/png;base64,{b64}" alt="page-{p_idx+1}" style="width:100%;height:auto;display:block;border:1px solid #ddd;box-shadow:0 1px 2px rgba(0,0,0,0.05)"/>')
                parts.append("</div>")
            # if we rasterized, optionally embed original pdf below
            if embed_pdf:
                b64pdf = base64.b64encode(pdf_bytes).decode('ascii')
                parts.append('<hr/><h3>Original PDF (embedded)</h3>')
                parts.append(f'<embed src="data:application/pdf;base64,{b64pdf}" width="100%" height="600px"></embed>')
            parts.append("</body></html>")
            return "\n".join(parts).encode("utf-8")
        except Exception:
            logger.exception("Rasterization failed; falling back to semantic HTML.")

    # Semantic HTML: build lists, headings, paragraphs and tables preserving order
    for page in parsed["pages"]:
        parts.append(f'<div class="page" data-page="{page["page_number"]}" style="page-break-after:always;">')
        elems = page.get("elements", [])
        i = 0
        while i < len(elems):
            el = elems[i]
            if el["type"] == "heading":
                lvl = min(max(int(el.get("level", 2)), 1), 6)
                parts.append(f"<h{lvl}>{html.escape(el['text'])}</h{lvl}>")
                i += 1
            elif el["type"] == "para":
                parts.append(f"<p>{html.escape(el['text'])}</p>")
                i += 1
            elif el["type"] == "list_item":
                # gather the run of consecutive list_items and check markers (ordered vs unordered)
                run = []
                markers = []
                while i < len(elems) and elems[i]["type"] == "list_item":
                    run.append(elems[i])
                    markers.append(elems[i].get("marker", ""))
                    i += 1
                # determine ordered vs unordered
                ordered = any(re.match(r"^\d", m) for m in markers)
                tag = "ol" if ordered else "ul"
                parts.append(f"<{tag}>")
                for li in run:
                    parts.append(f"<li>{html.escape(li['text'])}</li>")
                parts.append(f"</{tag}>")
            elif el["type"] == "table":
                rows = el.get("rows", [])
                parts.append("<div class='table-wrap'><table>")
                for r in rows:
                    parts.append("<tr>" + "".join(f"<td>{html.escape(str(c) if c is not None else '')}</td>" for c in r) + "</tr>")
                parts.append("</table></div>")
                i += 1
            else:
                # unknown type -> output as paragraph
                parts.append(f"<p>{html.escape(el.get('text', ''))}</p>")
                i += 1
        parts.append("</div>")

    if embed_pdf and pdf_bytes:
        b64 = base64.b64encode(pdf_bytes).decode('ascii')
        parts.append('<hr/><h3>Original PDF (embedded)</h3>')
        parts.append(f'<embed src="data:application/pdf;base64,{b64}" width="100%" height="600px"></embed>')
    parts.append("</body></html>")
    return "\n".join(parts).encode("utf-8")


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
                rows = el.get("rows", [])
                for r in rows:
                    out_lines.append("\t".join([str(c) if c is not None else "" for c in r]))
    return "\n".join(out_lines).encode("utf-8")


def structured_to_docx(parsed: dict) -> bytes:
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    for page in parsed["pages"]:
        doc.add_paragraph(f"--- Page {page['page_number']} ---").style = doc.styles['Normal']
        elems = page.get("elements", [])
        i = 0
        while i < len(elems):
            el = elems[i]
            if el["type"] == "heading":
                lvl = min(max(int(el.get("level", 2)), 1), 4)
                doc.add_heading(el["text"], level=lvl if lvl <= 4 else 4)
                i += 1
            elif el["type"] == "para":
                doc.add_paragraph(el["text"])
                i += 1
            elif el["type"] == "list_item":
                # collect run and use 'List Bullet' or 'List Number'
                run = []
                markers = []
                while i < len(elems) and elems[i]["type"] == "list_item":
                    run.append(elems[i])
                    markers.append(elems[i].get("marker", ""))
                    i += 1
                is_ordered = any(re.match(r"^\d", m) for m in markers)
                for li in run:
                    p = doc.add_paragraph(li["text"])
                    p.style = 'List Number' if is_ordered else 'List Bullet'
            elif el["type"] == "table":
                rows = el.get("rows", [])
                if not rows:
                    i += 1
                    continue
                ncols = max(len(r) for r in rows)
                tbl = doc.add_table(rows=0, cols=ncols)
                tbl.style = 'Table Grid'
                for r in rows:
                    row_cells = tbl.add_row().cells
                    for ci in range(ncols):
                        row_cells[ci].text = str(r[ci]) if ci < len(r) and r[ci] is not None else ""
                i += 1
            else:
                doc.add_paragraph(el.get("text", ""))
                i += 1
        doc.add_page_break()
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# -------------------------
# HTML -> DOCX / TEXT
# -------------------------
def html_to_text_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    text = soup.get_text(separator="\n")
    return text.encode("utf-8")

def html_to_docx_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    doc = Document()
    for el in soup.body.descendants:
        if getattr(el, "name", None):
            if el.name and el.name.startswith("h") and el.get_text(strip=True):
                try:
                    level = int(el.name[1]) if len(el.name) > 1 and el.name[1].isdigit() else 2
                except Exception:
                    level = 2
                doc.add_heading(el.get_text(strip=True), level=min(level, 4))
            elif el.name == "p" and el.get_text(strip=True):
                doc.add_paragraph(el.get_text("\n", strip=True))
            elif el.name in ("ul", "ol"):
                is_ordered = el.name == "ol"
                for li in el.find_all("li"):
                    p = doc.add_paragraph(li.get_text(strip=True))
                    p.style = 'List Number' if is_ordered else 'List Bullet'
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
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# -------------------------
# Streamlit UI
# -------------------------
st.set_page_config(page_title="Legacy Converter — Prod", layout="wide", page_icon="📚")
st.title("📚 Legacy Converter — Production-ready (Structured & Pixel options)")

st.markdown("""
**What this does**
- Converts digital PDFs into structured HTML / DOCX / TXT preserving headings, lists, tables.
- Offers pixel-perfect options: embed original PDF, or rasterize pages to PNG images and produce image-based HTML.
- Tune heading sensitivity and enable table detection.
""")

with st.sidebar:
    st.header("Options")
    conversion = st.selectbox("Conversion", [
        "PDF → Structured HTML",
        "PDF → Rasterized HTML (PNG pages)",
        "PDF → Word (.docx)",
        "PDF → Plain Text",
        "HTML → Word (.docx)",
        "HTML → Plain Text"
    ])
    workers = st.number_input("Parallel workers (bulk)", min_value=1, max_value=8, value=3)
    embed_pdf = st.checkbox("Embed original PDF into HTML output", value=False)
    rasterize_images = st.checkbox("Rasterize PDF pages to PNG (pixel-perfect HTML)", value=False)
    image_dpi = st.number_input("Rasterize DPI (higher = larger images)", min_value=72, max_value=300, value=150)
    st.markdown("### Heuristics")
    heading_ratio = st.slider("Heading size sensitivity (lower = more headings)", min_value=1.05, max_value=1.5, value=1.12, step=0.01)
    detect_tables = st.checkbox("Detect tables using pdfplumber (best-effort)", value=True)
    st.markdown("Tip: reduce heading sensitivity if headings are not detected; enable rasterize for exact visual clones.")

uploaded = st.file_uploader("Upload PDF(s) or HTML(s) — digital PDFs only (no OCR).", type=["pdf", "html"], accept_multiple_files=True)

if not uploaded:
    st.info("Upload at least one file.")
    st.stop()

st.markdown(f"Files queued: {len(uploaded)}")
start = st.button("Create legacy — Convert now")

if not start:
    st.stop()

results_for_zip = []
progress = st.progress(0)
status = st.empty()
log = st.empty()

def process_file(uploaded_file) -> Dict[str, Any]:
    name = uploaded_file.name
    raw = uploaded_file.read()
    ext = os.path.splitext(name)[1].lower()
    try:
        if ext == ".pdf":
            parsed = parse_pdf_structured(raw, min_heading_ratio=heading_ratio, detect_tables=detect_tables)
            if conversion == "PDF → Structured HTML":
                out_bytes = structured_to_html(parsed, embed_pdf=embed_pdf, pdf_bytes=raw, rasterize_images=False)
                out_name = os.path.splitext(name)[0] + ".html"
                mime = "text/html"
            elif conversion == "PDF → Rasterized HTML (PNG pages)":
                out_bytes = structured_to_html(parsed, embed_pdf=embed_pdf, pdf_bytes=raw, rasterize_images=True, image_dpi=image_dpi)
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
        logger.exception("Processing failed for %s", name)
        tb = traceback.format_exc()
        return {"name": name, "error": f"{str(e)}\n{tb}"}

# process concurrently
with ThreadPoolExecutor(max_workers=workers) as exe:
    futures = {exe.submit(process_file, f): f.name for f in uploaded}
    done = 0
    for fut in as_completed(futures):
        done += 1
        res = fut.result()
        progress.progress(done / len(uploaded))
        if res.get("error"):
            log.write(f"✖ {res['name']} — {res['error'][:4000]}")
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
                    st.components.v1.html(res["out_bytes"].decode("utf-8", errors="replace")[:200000], height=300, scrolling=True)
                except Exception:
                    st.write("(HTML preview failed; download instead.)")
            st.download_button("Download", data=res["out_bytes"], file_name=res["out_name"], mime=res["mime"])

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in results_for_zip:
            zf.writestr(r["out_name"], r["out_bytes"])
    zip_buf.seek(0)
    st.download_button("Download ALL as ZIP", zip_buf.read(), file_name="converted_legacy.zip", mime="application/zip")
else:
    st.error("No successful conversions to download. Check logs above.")

st.markdown("---")
st.info("If your main goal is *pixel-perfect visual cloning*, enable 'Rasterize PDF pages to PNG' or leave the original PDF embedded. For semantic conversions (editable Word/HTML), tune 'Heading size sensitivity' and enable 'Detect tables' for better table extraction.")
