import io
import os
import zipfile
from typing import List, Tuple, Dict, Any
import fitz  # PyMuPDF
import pdfplumber
from bs4 import BeautifulSoup
from docx.shared import Document as DOCX
import pandas as pd
import html
import re

# Utility / Heuristic Functions
BULLET_CHARS = r"^[\u2022\u2023\u25E6\-\*\•\–\—]\s+|^\d+[\.$]\s+"

def sanitize_text(t: str) -> str:
    return t.replace('\r', '').rstrip()

def is_bullet_line(text: str) -> bool:
    return bool(re.match(BULLET_CHARS, text.strip()))

def normalize_whitespace(s: str) -> str:
    return re.sub(r'\s+\n', '\n', re.sub(r'\n\s+', '\n', s)).strip()

def choose_heading_levels(unique_sizes: List[float]) -> Dict[float, int]:
    """ Map font sizes -> heading levels (1..4) heuristically. Biggest -> h1, next -> h2, etc. If many sizes, map top ones to headings. """
    sizes = sorted(set(unique_sizes), reverse=True)
    mapping = {}
    
    # Map up to 4 distinct sizes to heading levels top = 1, 2, 3, 4
    for idx, size in enumerate(sizes[:4]):
        mapping[size] = idx + 1
    
    return mapping

# PDF Parsing (structured)
def parse_pdf_structured(pdf_bytes: bytes, min_heading_ratio: float = 1.12) -> Dict[str, Any]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_out = []
    all_sizes = []

    for page in doc:
        page_text = fitz.get_text(document=page)
        elements = []

        # First detect tables with pdfplumber on this page
        tables_page = []
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as ppdf:
                page_pl = ppdf.pages[p]
                for t in ppdf.extract_tables():
                    tables_page.append({"rows": t})
        except Exception:
            tables_page = []

        # Add tables as elements (this keeps order approximate — we append tables at end of page parsing)
        for table in tables_page:
            rows = table.get_text("rows")
            if rows:
                rows = '\n'.join(rows)
                elements.append({"type": "table", "rows": rows})

        # Now process text blocks
        for block in page.get_text("blocks").split("\n"):
            span = block.get("spans", [])
            sz = round(span.get("size", 0), 2)
            if sz > 0:
                line_text = span.get_text()
                if is_bullet_line(line_text):
                    ln = line_text.strip()
                    elements.append({"type": "list_item", "text": ln})
                else:
                    ln = line_text
                    elements.append({"type": "heading", "text": ln, "level": min(max(int(ln.split(" ", 1)[0]), 1), 6)})
            elif sz > 0:
                ln = line_text.strip()
                elements.append({"type": "para", "text": ln})

        pages_out.append({"page_number": p + 1, "elements": elements})

    return {"pages": pages_out, "fontsizes": normalize_whitespace(all_sizes)}

def structured_to_html(parsed: Dict[str, Any], embed_pdf: bool = False, pdf_bytes: bytes = None) -> bytes:
    parts = ['<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>body{font-family:Arial,Helvetica,sans-serif;line-height:1.4;padding:16px}pre{white-space:pre-wrap;}</style></head><body>']
    for page in parsed["pages"]:
        parts.append(f'<div class="page" data-page="{page["page_number"]}" style="page-break-after:always;">')
        for el in page["elements"]:
            if el["type"] == "heading":
                lvl = min(max(int(el.get("level", 2)), 1), 6)
                text = html.escape(el["text"])
                parts.append(f'<h{lvl}>{text}</h{lvl>>')
            elif el["type"] == "para":
                text = html.escape(el["text"])
                parts.append(f'<p>{text}</p>')
            elif el["type"] == "list_item":
                text = ' '.join(el['text'] for el in el.get('list_items', []))
                parts.append(f'<li>{text}</li>')
            elif el["type"] == "table":
                rows = el["rows"]
                for r in rows:
                    cells = [cell.get_text() if cell else "" for cell in r]
                    parts.append(f"<td>{', '.join(cells)}</td>")
        parts.append("</div>")
    html_text = "\n".join(parts) + "</body></html>"
    return html_text.encode("utf-8")

def structured_to_text(parsed: Dict[str, Any]) -> bytes:
    out_lines = []
    for page in parsed["pages"]:
        out_lines.append(f"--- PAGE {page['page_number']} ---")
        for el in page["elements"]:
            if el["type"] == "heading":
                text = html.escape(el["text"].upper())
                out_lines.append(text)
            elif el["type"] == "para":
                text = html.escape(el["text"])
                out_lines.append(text)
            elif el["type"] == "list_item":
                text = ' '.join(el["text"] for el in el.get("list_items", []))
                out_lines.append(f"- {text}")
        html_text = "\n".join(out_lines)
        return html_text.encode("utf-8")

def structured_to_docx(parsed: Dict[str, Any]) -> bytes:
    doc = Document()
    doc.add_paragraph(f"--- Page {parsed['page_number']} ---")
    last_was_list = False
    for el in parsed["elements"]:
        if el["type"] == "heading":
            lvl = min(max(int(el.get("level", 2)), 1), 4)
            doc.add_heading(el["text"].upper(), level=lvl)
        elif el["type"] == "para":
            doc.add_paragraph(el["text"])
        elif el["type"] == "list_item":
            rows = el["rows"]
            ncols = max(len(r) for r in rows)
            tbl = doc.add_table(rows=0, cols=ncols)
            for r in rows:
                row_cells = tbl.add_row().cells
                for i in range(ncols):
                    cells = [cell.get_text() if cell else "" for cell in row_cells[i]]
                    doc.add_paragraph(f"{', '.join(cells)}")
    doc.save(io.BytesIO())
    return io.BytesIO(doc.get_content())

# HTML -> Text / DOCX ##############################################################################################
def html_to_text_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    text = soup.get_text(separator="\n")
    return text.encode("utf-8")

def html_to_docx_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    doc = Document()
    for el in soup.body.descendants:
        if el.name and el.name.startswith("h") and el.get_text(strip=True):
            level = int(el.name[1])
            doc.add_heading(el.get_text(strip=True).upper(), level=level)
        elif el.name == "p" and el.get_text(strip=True):
            doc.add_paragraph(el.get_text("\n", strip=True))
        elif el.name in ("ul", "ol"):
            for li in el.find_all("li"):
                p = doc.add_paragraph(li.get_text(strip=True))
                p.style = 'List Bullet' if el.name == "ul" else 'List Number'
    doc.save(io.BytesIO())
    return io.BytesIO(doc.get_content())

# Streamlit App UI
def create_converter_interface():
    st.set_page_config(page_title="Legacy Converter — Preserve structure", layout="wide", page_icon="📚")
    st.title("Legacy Converter — Preserve structure, create a legacy")

    st.markdown("""This app tries to **preserve structure** during conversions: - Headings (detected via font size) - Paragraphs (text blocks) - Lists (bullets / numbered) - Tables (via pdfplumber) It uses heuristics — tune thresholds from the sidebar for better results on specific documents."""
    st.sidebarheader("Conversion options")
    conversion_options = ["PDF → Structured HTML", "PDF → Word (.docx)", "PDF → Plain Text", "HTML → Word (.docx)", "HTML → Plain Text"]
    conversion_type = st.selectbox("Conversion", conversion_options)
    workers = st.number_input("Parallel workers (bulk)", min_value=1, max_value=8, value=3)
    embed_pdf = st.checkbox("Embed original PDF into HTML output (only for HTML outputs)")
    heading_ratio = st.slider("Heading size sensitivity (lower = more headings)", 0.5, 2, 1.5)

    uploaded_file = st.file_uploader("Upload PDF(s) or HTML(s) — digital PDFs only (no OCR).", type=["pdf", "html"], accept_multiple_files=True)

    if not uploaded_file:
        st.info("Upload at least one file. This converter targets digital PDFs (embedded text) and HTML files.")

    start = st.button("Create legacy — Convert now")

    def on_create_conversion(page_number):
        with ThreadPoolExecutor(max_workers=workers) as executor:
            future = executor.submit(process_file, uploaded_file.name)
            result = future.result()
            log.f.write(f"✔ {result['name']} → {result['out_name']} ({len(result['out_bytes']):,})")
        st.info(f"Conversion {result['name']} completed successfully.")

    if not uploaded_file:
        return st.exptitle("Uploaded Files")

    st.subheader("Files Queued for Conversion")
    st.markdown("Files to be converted:")
    for i, file in enumerate(uploaded_file):
        st.write(f"{i + 1}. {file.name}")

    create_conversion_button = st.button("Process Files (Parallel)", on_click=lambda: process_files(uploaded_file))
    create_conversion_button.click()

    def process_files(file_name):
        try:
            raw_bytes = uploaded_file.read()
            parsed = parse_pdf_structured(raw_bytes, min_heading_ratio=heading_ratio)
            out_bytes = structured_to_html(parsed, embed_pdf=embed_pdf, pdf_bytes=raw_bytes)
            return {"name": file_name, "out_bytes": out_bytes.decode("utf-8"), "out_name": file_name.lower() + ".html", "mime": "text/html"}
        except Exception as e:
            log.write(f"✖ {file_name} — {str(e)}")

    progress = st.progress(0, total=len(uploaded_file))
    res = {}
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(process_file, file): file.name for file in uploaded_file}
        done = 0
        for future in futures:
            res[future.name] = future.result()
            progress.update(future.progress())
    
    if res:
        log.writelines(f"\n{' '.join(f'{k}: {v["out_name"]}' for k, v in res.items())}")
        st.markdown("--- Conversion Results")
        for file, content in res.items():
            preview = content[:4000].decode("utf-8") + "\n" + "".join(f"{line.strip()}" for line in content[4000:])
            st.write(preview)
            st.download_button("Download", data=content, file_name=file.lower() + ".html", mime="text/html")

    return res

def process_file(uploaded_file):
    name = uploaded_file.name
    raw_bytes = uploaded_file.read()
    ext = os.path.splitext(name)[1].lower()
    if ext == ".pdf":
        return structured_to_html(parse_pdf_structured(raw_bytes, min_heading_ratio=heading_ratio))
    elif ext == ".html":
        return structured_to_docx(parse_pdf_structured(raw_bytes, min_heading_ratio=heading_ratio))
    else:
        return {"name": name, "error": f"Conversion {ext} not valid for PDF"}

if __name__ == "__main__":
    converter = create_converter_interface()
