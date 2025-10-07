# streamlit_embed_structured.py
"""
Streamlit app: Invisible Embedded PDF + Code Inset & Structured Conversion
=========================================================================
Features:
1. Upload a PDF (digital). The app embeds the **entire PDF as-is** (data URI) into the Viewer tab.
   - The embedded viewer is styled to be visually clean / invisible (no border, no grid, minimal UI).
   - We append PDF URL parameters to hide toolbar/navigation when supported by the browser (#toolbar=0&navpanes=0).
   - If you want pixel-perfect visual clones, use the Rasterize option in the "Structure" tab.
2. Code Editor tab:
   - Paste/enter transformation or structuring rules (text or Python snippet).
   - Save code snippets to disk for later use (no execution of user code by the app).
   - Maintain snippet list and allow download.
3. Structure tab:
   - Run a safe, built-in structured PDF parser (heuristics) that extracts headings, paragraphs,
     lists, and tables (no user code executed).
   - Produce semantic HTML, DOCX and TXT outputs from parsed structure.
   - Option to embed the original PDF in the output (invisible embed or normal).
   - Option to rasterize pages to images for pixel-perfect HTML (if you need exact visual copy).
Important security note:
- The "Code Editor" saves snippets but DOES NOT execute them. This prevents remote code execution risk.
- The app uses PyMuPDF (fitz), pdfplumber, python-docx, BeautifulSoup and Streamlit.

Save this file as: streamlit_embed_structured.py
Run:
    pip install streamlit PyMuPDF pdfplumber python-docx beautifulsoup4
    streamlit run streamlit_embed_structured.py
"""

import os
import io
import re
import base64
import zipfile
import traceback
from typing import List, Dict, Any, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt
import html
import logging

# ---------------------------
# Logging
# ---------------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("embed_structured")

# ---------------------------
# Utilities and heuristics
# ---------------------------
BULLET_CHARS = r"^\s*([\u2022\u2023\u25E6\-\*\•\–\—]|\d+[\.\)])\s+"

def sanitize_text(t: str) -> str:
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
    top = sizes[:6]
    for idx, s in enumerate(top):
        if idx < 4:
            mapping[round(s, 2)] = idx + 1
        else:
            mapping[round(s, 2)] = 0
    return mapping

# ---------------------------
# Parsing PDF into structured intermediate representation
# ---------------------------
def parse_pdf_structured(pdf_bytes: bytes, min_heading_ratio: float = 1.12, detect_tables: bool = True) -> Dict[str, Any]:
    """
    Returns:
      {
        "pages": [
            {"page_number": n,
             "elements": [
                 {"type":"heading"/"para"/"list_item"/"table", "text":..., "level":n, "rows":[...], "bbox":(...)}
             ]
            }, ...
        ],
        "fontsizes": [...]
      }
    """
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        logger.exception("PyMuPDF open failed")
        raise

    all_sizes = []
    per_page_blocks: List[List[Dict[str, Any]]] = []

    for p_idx in range(len(doc)):
        page = doc.load_page(p_idx)
        d = page.get_text("dict")
        page_blocks = []
        for block in d.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                line_text = ""
                max_span_sz = 0.0
                bbox = line.get("bbox", block.get("bbox", [0,0,0,0]))
                for span in line.get("spans", []):
                    txt = span.get("text", "")
                    if txt:
                        line_text += txt
                    sz = span.get("size", 0)
                    if sz and sz > max_span_sz:
                        max_span_sz = sz
                if line_text.strip():
                    b = line.get("bbox", bbox)
                    pt = {
                        "text": sanitize_text(line_text),
                        "size": round(max_span_sz, 2),
                        "bbox": tuple(b),
                        "x0": float(b[0]), "y0": float(b[1]),
                        "x1": float(b[2]), "y1": float(b[3]),
                    }
                    page_blocks.append(pt)
                    if max_span_sz > 0:
                        all_sizes.append(max_span_sz)
        per_page_blocks.append(page_blocks)

    unique_sizes = sorted(set([round(s,2) for s in all_sizes]), reverse=True) if all_sizes else [12.0]
    font_to_heading = choose_heading_levels(unique_sizes)
    max_size = max(unique_sizes) if unique_sizes else 12.0
    heading_threshold = max_size / min_heading_ratio

    pages_out = []

    pdfplumber_tables_by_page = {}
    if detect_tables:
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as ppdf:
                for p_idx in range(len(ppdf.pages)):
                    page_pl = ppdf.pages[p_idx]
                    extracted_tables = page_pl.extract_tables()
                    tables_clean = []
                    for t in extracted_tables or []:
                        if t and any(any(cell for cell in row if cell not in (None, "")) for row in t):
                            tables_clean.append({"rows": t})
                    pdfplumber_tables_by_page[p_idx] = tables_clean
        except Exception:
            logger.exception("pdfplumber extraction failed")
            pdfplumber_tables_by_page = {}

    for p_idx, blocks in enumerate(per_page_blocks):
        page_elements = []
        blocks_sorted = sorted(blocks, key=lambda b: (round(b["y0"],2), round(b["x0"],2)))
        for b in blocks_sorted:
            txt = b["text"].strip()
            sz = b.get("size", 0.0)
            is_bullet, marker = is_bullet_line(txt)
            if is_bullet:
                cleaned = re.sub(BULLET_CHARS, "", txt, count=1)
                page_elements.append({"type":"list_item","text":cleaned,"marker":marker,"size":sz,"bbox":b["bbox"],"x0":b["x0"],"y0":b["y0"]})
            else:
                mapped = font_to_heading.get(round(sz,2), 0)
                if (sz >= heading_threshold) or mapped:
                    lvl = mapped if mapped else 2
                    page_elements.append({"type":"heading","text":txt,"level":lvl,"size":sz,"bbox":b["bbox"],"x0":b["x0"],"y0":b["y0"]})
                else:
                    page_elements.append({"type":"para","text":txt,"size":sz,"bbox":b["bbox"],"x0":b["x0"],"y0":b["y0"]})
        # append tables (best-effort)
        for t in pdfplumber_tables_by_page.get(p_idx, []):
            page_elements.append({"type":"table","rows": t["rows"], "bbox": None})
        pages_out.append({"page_number": p_idx+1, "elements": page_elements})

    return {"pages": pages_out, "fontsizes": unique_sizes}

# ---------------------------
# Intermediate -> HTML / DOCX / TEXT
# ---------------------------
def structured_to_html(parsed: dict, embed_pdf: bool = False, pdf_bytes: bytes = None, rasterize_images: bool = False, image_dpi: int = 150) -> bytes:
    parts = [
        '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">',
        '<style>body{font-family:Arial,Helvetica,sans-serif;line-height:1.45;padding:18px;max-width:1000px;margin:auto}pre{white-space:pre-wrap}table{border-collapse:collapse;margin:8px 0}td,th{border:1px solid #ccc;padding:6px;text-align:left}hr{border:none;border-top:1px solid #eee;margin:18px 0}</style>',
        '</head><body>'
    ]

    # Rasterize pages into images if requested (pixel-perfect)
    if rasterize_images and pdf_bytes:
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            for p_idx in range(len(doc)):
                page = doc.load_page(p_idx)
                mat = fitz.Matrix(image_dpi/72.0, image_dpi/72.0)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img_bytes = pix.tobytes("png")
                b64 = base64.b64encode(img_bytes).decode("ascii")
                parts.append(f'<div class="page" data-page="{p_idx+1}" style="page-break-after:always;margin-bottom:24px;">')
                parts.append(f'<img src="data:image/png;base64,{b64}" alt="page-{p_idx+1}" style="width:100%;height:auto;display:block;border:0;box-shadow:none;"/>')
                parts.append("</div>")
            if embed_pdf:
                b64pdf = base64.b64encode(pdf_bytes).decode('ascii')
                parts.append('<hr/><h3>Original PDF (embedded)</h3>')
                parts.append(f'<embed src="data:application/pdf;base64,{b64pdf}#toolbar=0&navpanes=0" width="100%" height="600px" style="border:none;"></embed>')
            parts.append("</body></html>")
            return "\n".join(parts).encode("utf-8")
        except Exception:
            logger.exception("Rasterization failed - falling back to semantic HTML")

    # Semantic HTML
    for page in parsed["pages"]:
        parts.append(f'<div class="page" data-page="{page["page_number"]}" style="page-break-after:always;">')
        elems = page.get("elements", [])
        i = 0
        while i < len(elems):
            el = elems[i]
            if el["type"] == "heading":
                lvl = min(max(int(el.get("level",2)),1),6)
                parts.append(f"<h{lvl}>{html.escape(el['text'])}</h{lvl}>")
                i += 1
            elif el["type"] == "para":
                parts.append(f"<p>{html.escape(el['text'])}</p>")
                i += 1
            elif el["type"] == "list_item":
                run = []
                markers = []
                while i < len(elems) and elems[i]["type"] == "list_item":
                    run.append(elems[i])
                    markers.append(elems[i].get("marker",""))
                    i += 1
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
                parts.append(f"<p>{html.escape(el.get('text',''))}</p>")
                i += 1
        parts.append("</div>")

    if embed_pdf and pdf_bytes:
        b64 = base64.b64encode(pdf_bytes).decode('ascii')
        parts.append('<hr/><h3>Original PDF (embedded)</h3>')
        parts.append(f'<embed src="data:application/pdf;base64,{b64}#toolbar=0&navpanes=0" width="100%" height="600px" style="border:none;"></embed>')
    parts.append("</body></html>")
    return "\n".join(parts).encode("utf-8")


def structured_to_text(parsed: dict) -> bytes:
    out = []
    for page in parsed["pages"]:
        out.append(f"--- PAGE {page['page_number']} ---")
        for el in page["elements"]:
            if el["type"] == "heading":
                out.append(el["text"].upper())
            elif el["type"] == "para":
                out.append(el["text"])
            elif el["type"] == "list_item":
                out.append(f"- {el['text']}")
            elif el["type"] == "table":
                rows = el.get("rows", [])
                for r in rows:
                    out.append("\t".join([str(c) if c is not None else "" for c in r]))
    return "\n".join(out).encode("utf-8")


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
                lvl = min(max(int(el.get("level",2)),1),4)
                doc.add_heading(el["text"], level=lvl if lvl <=4 else 4)
                i += 1
            elif el["type"] == "para":
                doc.add_paragraph(el["text"])
                i += 1
            elif el["type"] == "list_item":
                run = []
                markers = []
                while i < len(elems) and elems[i]["type"] == "list_item":
                    run.append(elems[i])
                    markers.append(elems[i].get("marker",""))
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
                doc.add_paragraph(el.get("text",""))
                i += 1
        doc.add_page_break()
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# ---------------------------
# Helper: embed PDF as data URI (invisible/styled)
# ---------------------------
def pdf_to_data_uri(pdf_bytes: bytes, hide_toolbar: bool = True) -> str:
    b64 = base64.b64encode(pdf_bytes).decode("ascii")
    suffix = "#toolbar=0&navpanes=0" if hide_toolbar else ""
    return f"data:application/pdf;base64,{b64}{suffix}"

# ---------------------------
# App UI
# ---------------------------
st.set_page_config(page_title="Embed & Structure — Invisible PDF + Code Editor", layout="wide")
st.title("Embed & Structure — Invisible PDF + Code Editor")

st.markdown("""
**Overview**
- Upload a PDF and view it embedded (styled to be visually clean/invisible — no border or grid).
- Use the Code Editor tab to paste transformation rules or snippets and save them (the app will NOT execute your code).
- Use the Structure tab to run built-in heuristics that extract headings, lists, paragraphs, and tables into editable outputs.
""")

# Sidebar controls
with st.sidebar:
    st.header("App options")
    hide_pdf_toolbar = st.checkbox("Hide PDF toolbar/navigation (when supported)", value=True)
    default_dpi = st.slider("Rasterize DPI (for pixel-perfect HTML)", 72, 300, 150)
    parallel_workers = st.number_input("Parallel workers (bulk)", min_value=1, max_value=8, value=3)
    st.markdown("Saved code snippets are stored locally under ./snippets/")

# Ensure snippets dir
os.makedirs("snippets", exist_ok=True)

# Tabs: Viewer, Code Editor, Structure
tab_viewer, tab_editor, tab_structure = st.tabs(["Viewer (Embedded PDF)", "Code Editor (save snippets)", "Structure & Export"])

# File uploader at the top to persist selection across tabs
uploaded_files = st.file_uploader("Upload PDF(s) — digital PDFs only (no OCR).", type=["pdf"], accept_multiple_files=True, key="uploader")

if not uploaded_files:
    st.info("Upload one or more PDF files to proceed.")
    st.stop()

# Helper to pick one active PDF file for viewer/editor operations
file_names = [f.name for f in uploaded_files]
active_index = st.selectbox("Active file (viewer/editor/structure)", options=list(range(len(file_names))), format_func=lambda i: file_names[i])

active_file = uploaded_files[active_index]
active_name = active_file.name
active_bytes = active_file.read()

# ---- Viewer Tab ----
with tab_viewer:
    st.subheader("Viewer — embedded PDF (clean/invisible style)")
    st.markdown("This embeds the full original PDF (data URI). The embed is styled to remove borders and minimize visible UI.")
    # build embed HTML
    data_uri = pdf_to_data_uri(active_bytes, hide_toolbar=hide_pdf_toolbar)
    # Use an <iframe> or <embed>. Some browsers show built-in UI; we use #toolbar=0 parameter as best-effort.
    embed_html = f"""
    <div style="position:relative;width:100%;height:85vh;overflow:hidden;border:none;background:transparent;">
      <iframe src="{data_uri}" style="position:absolute;left:0;top:0;width:100%;height:100%;border:none;box-shadow:none;" sandbox="allow-same-origin allow-scripts" frameborder="0"></iframe>
    </div>
    """
    st.components.v1.html(embed_html, height=700)

    st.markdown("---")
    st.write("Download original PDF:")
    st.download_button("Download original PDF", data=active_bytes, file_name=active_name, mime="application/pdf")

# ---- Code Editor Tab ----
with tab_editor:
    st.subheader("Code Editor — paste rules / snippets and save (no execution)")
    st.markdown("Paste any transformation rules, JSON mapping, or notes. These are saved to ./snippets/ and not executed by this app (for safety).")

    # list existing snippets
    snippet_files = sorted([fn for fn in os.listdir("snippets") if fn.endswith(".txt") or fn.endswith(".py") or fn.endswith(".md")])
    col1, col2 = st.columns([2,1])
    with col1:
        new_snippet_name = st.text_input("Snippet filename (e.g. rules.py or transform.txt)", value=f"{os.path.splitext(active_name)[0]}_rules.txt")
        snippet_code = st.text_area("Snippet content", height=300, placeholder="# Paste rules or transformations here. This WILL NOT be executed by the app.")
        save_btn = st.button("Save snippet")
    with col2:
        st.markdown("Saved snippets")
        for s in snippet_files:
            st.write(s)
        selected_snip = st.selectbox("Open snippet", options=["-- choose --"] + snippet_files)
        if st.button("Load selected"):
            if selected_snip != "-- choose --":
                with open(os.path.join("snippets", selected_snip), "r", encoding="utf-8") as fh:
                    code = fh.read()
                st.experimental_set_query_params()  # no-op used to cause re-render in some streamlit versions
                st.info(f"Loaded snippet: {selected_snip}")
                # show in a new text_area
                st.text_area("Loaded snippet content (read-only)", value=code, height=300)

    if save_btn:
        if not new_snippet_name.strip():
            st.error("Provide a valid filename.")
        else:
            safe_name = re.sub(r"[^A-Za-z0-9_\-\.]", "_", new_snippet_name)
            path = os.path.join("snippets", safe_name)
            try:
                with open(path, "w", encoding="utf-8") as fh:
                    fh.write(snippet_code)
                st.success(f"Saved snippet to {path}")
            except Exception as e:
                st.error(f"Failed to save: {e}")

    st.markdown("---")
    st.markdown("Download snippets as ZIP:")
    if st.button("Bundle snippets as ZIP"):
        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
            for s in snippet_files:
                zf.write(os.path.join("snippets", s), arcname=s)
        zbuf.seek(0)
        st.download_button("Download snippets.zip", zbuf.read(), file_name="snippets.zip", mime="application/zip")

# ---- Structure & Export Tab ----
with tab_structure:
    st.subheader("Structure & Export — built-in heuristics (safe)")
    st.markdown("This runs the app's built-in parser to extract structure (headings, paragraphs, lists, tables). It **does not** execute user snippets saved in the Code Editor.")

    cols_opts = st.columns(4)
    with cols_opts[0]:
        heading_ratio = st.slider("Heading sensitivity (lower = more headings)", min_value=1.05, max_value=1.5, value=1.12, step=0.01)
    with cols_opts[1]:
        detect_tables = st.checkbox("Detect tables with pdfplumber", value=True)
    with cols_opts[2]:
        embed_original = st.checkbox("Embed original PDF (in output HTML)", value=False)
    with cols_opts[3]:
        rasterize_pages = st.checkbox("Rasterize pages to PNG for pixel-perfect HTML", value=False)

    if st.button("Run structuring on active PDF"):
        try:
            st.info("Parsing PDF (this may take a few seconds)...")
            parsed = parse_pdf_structured(active_bytes, min_heading_ratio=heading_ratio, detect_tables=detect_tables)
            st.success("Parsing finished — preview below.")
            # show a short preview of structure
            preview_count = 0
            for page in parsed["pages"]:
                st.markdown(f"**Page {page['page_number']}**")
                for el in page["elements"][:60]:
                    t = el["type"]
                    if t == "heading":
                        st.markdown(f"- **H{el.get('level',2)}**: {el.get('text')[:200]}")
                    elif t == "para":
                        st.write(f"- {el.get('text')[:200]}")
                    elif t == "list_item":
                        st.write(f"- (list) {el.get('text')[:200]}")
                    elif t == "table":
                        st.write(f"- (table) {len(el.get('rows',[]))} rows")
                    preview_count += 1
                    if preview_count > 80:
                        st.write("... (preview truncated)")
                        break
                if preview_count > 80:
                    break

            # produce outputs (HTML, DOCX, TXT)
            st.markdown("### Generate outputs")
            out_html = structured_to_html(parsed, embed_pdf=embed_original, pdf_bytes=active_bytes if embed_original else None, rasterize_images=rasterize_pages, image_dpi=default_dpi)
            out_txt = structured_to_text(parsed)
            out_docx = structured_to_docx(parsed)

            st.markdown("Download outputs (or preview HTML below):")
            st.download_button("Download structured HTML", data=out_html, file_name=f"{os.path.splitext(active_name)[0]}_structured.html", mime="text/html")
            st.download_button("Download structured TXT", data=out_txt, file_name=f"{os.path.splitext(active_name)[0]}_structured.txt", mime="text/plain")
            st.download_button("Download structured DOCX", data=out_docx, file_name=f"{os.path.splitext(active_name)[0]}_structured.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

            # HTML preview (safe)
            try:
                st.markdown("#### HTML Preview")
                st.components.v1.html(out_html.decode("utf-8", errors="replace"), height=600, scrolling=True)
            except Exception:
                st.write("HTML preview failed; download the file to inspect fully.")
        except Exception as e:
            tb = traceback.format_exc()
            st.error(f"Structuring failed: {e}")
            st.text(tb)

    st.markdown("---")
    st.markdown("Advanced: Bundle outputs for all uploaded PDFs")
    if st.button("Run batch structuring for all uploaded PDFs"):
        results = []
        with st.spinner("Processing files..."):
            with ThreadPoolExecutor(max_workers=parallel_workers) as exe:
                futures = {}
                for f in uploaded_files:
                    raw = f.read()
                    futures[exe.submit(parse_pdf_structured, raw, min_heading_ratio=heading_ratio, detect_tables=detect_tables)] = (f.name, raw)
                for fut in as_completed(futures):
                    name, raw = futures[fut]
                    try:
                        parsed = fut.result()
                        html_bytes = structured_to_html(parsed, embed_pdf=embed_original, pdf_bytes=raw if embed_original else None, rasterize_images=rasterize_pages, image_dpi=default_dpi)
                        txt_bytes = structured_to_text(parsed)
                        docx_bytes = structured_to_docx(parsed)
                        results.append({"name": name, "html": html_bytes, "txt": txt_bytes, "docx": docx_bytes})
                        st.write(f"✔ {name} processed")
                    except Exception as e:
                        st.write(f"✖ {name} failed: {e}")

        if results:
            # create zip
            zbuf = io.BytesIO()
            with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
                for r in results:
                    base = os.path.splitext(r["name"])[0]
                    zf.writestr(f"{base}_structured.html", r["html"])
                    zf.writestr(f"{base}_structured.txt", r["txt"])
                    zf.writestr(f"{base}_structured.docx", r["docx"])
            zbuf.seek(0)
            st.download_button("Download ALL structured outputs (ZIP)", zbuf.read(), file_name="structured_outputs.zip", mime="application/zip")
        else:
            st.info("No successful results to bundle.")

st.markdown("---")
st.info("Notes:\n- The Viewer embeds the original PDF as a data URI; the 'invisible' / no-border look is achieved via iframe styling and PDF URL parameters (#toolbar=0). Some browsers may still show limited UI depending on their PDF plugin.\n- The Code Editor stores snippets locally in ./snippets (they are NOT executed). Use them to save transformation rules, mapping notes, or Python scripts you'll execute offline.\n- The Structuring uses safe internal heuristics only. If you want the app to execute a saved snippet to transform parsed output, we can add a sandboxed runner — but that is more complex and requires careful security design.")
