import io
import os
import zipfile
from typing import List, Dict, Tuple, Any
import fitz  # PyMuPDF
import pdfplumber
from bs4 import BeautifulSoup
from docx.shared import Document as DOCX
import pandas as pd
import html
import re

# Utility Functions
BULLET_CHARS = r"^[\u2022\u2023\u25E6\-\*\•\–\—]\s+|^\d+[\.$]\s+"

def sanitize_text(text: str) -> str:
    return text.replace('\r', '').rstrip()

def is_bullet_line(text: str) -> bool:
    return bool(re.match(BULLET_CHARS, text.strip()))

def normalize_whitespace(s: str) -> str:
    return re.sub(r'\s+\n', '\n', re.sub(r'\n\s+', '\n', s)).strip()

def choose_heading_levels(unique_sizes: List[float]) -> Dict[float, int]:
    sizes = sorted(set(unique_sizes), reverse=True)
    mapping = {size: idx + 1 for idx, size in enumerate(sizes) if size > 0}
    return mapping

# PDF Parsing (Structured)
def parse_pdf_structured(pdf_bytes: bytes, min_heading_ratio: float = 1.12) -> Dict[str, Any]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_out = []
    all_sizes = []

    for page in doc:
        page_text = fitz.get_text(document=page)
        elements = extract_elements(page_text)
        tables_page = extract_tables(page_text)
        
        for element in elements:
            if element["type"] == "heading":
                level = min(max(int(element.get("level", 2)), 1), 6)
                text = html.escape(element["text"])
                elements.append({"type": "heading", "text": text, "level": level})
            elif element["type"] == "para":
                text = html.escape(element["text"])
                elements.append({"type": "para", "text": text})
            elif element["type"] == "list_item":
                text = html.escape(element["text"])
                elements.append({"type": "list_item", "text": text})
            elif element["type"] == "table":
                rows = extract_table_elements(page_text)
                tables_page.append({"rows": rows})

    return {"pages": pages_out, "fontsizes": normalize_whitespace(all_sizes)}

def extract_elements(text: str) -> List[Dict[str, Any]]:
    elements = []
    for line in text.split("\n"):
        span = line.get("spans", [])
        if span:
            stxt = span.get_text()
            size = round(span.get("size", 0), 2)
            if is_bullet_line(stxt):
                ln = stxt.strip()
                elements.append({"type": "list_item", "text": ln})
            else:
                ln = stxt
                elements.append({"type": "heading", "text": ln, "level": min(max(int(ln.split(" ", 1)[0]), 1), 6)})
            elif size > 0:
                elements.append({"type": "para", "text": ln})
    return elements

def extract_tables(text: str) -> List[List[str]]:
    tables_page = []
    try:
        with pdfplumber.open(io.BytesIO(text)) as ppdf:
            for page_pl in ppdf.pages:
                if isinstance(page_pl, pdfplumber.pages.Table):
                    tables_page.append([cell.get_text() for cell in page_pl.rows[0].text])
    except Exception:
        pass
    return tables_page

# Converters
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
                rows = '\n'.join([f'<td>{str(c)}</td>' for c in el.get('rows', [])])
                parts.append(f'<table><tr>{rows}</tr></table>')
    parts.append("</div>")
    html_text = "\n".join(parts) + "</body></html>"
    return html_text.encode("utf-8")

def structured_to_text(parsed: Dict[str, Any]) -> bytes:
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
                rows = '\n'.join([f'<td>{str(c)}</td>' for c in row if c is not None])
                out_lines.append(f"\t{rows}")
    return "\n".join(out_lines).encode("utf-8")

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
            doc.add_paragraph(el["text"])
        elif el["type"] == "table":
            rows = '\n'.join([f'<td>{str(c)}</td>' for c in row if c is not None])
            ncols = max(len(r) for r in rows)
            tbl = doc.add_table(rows=0, cols=ncols)
            for r in rows:
                row_cells = tbl.add_row().cells
                for i in range(ncols):
                    cells = [cell.get_text() if cell else "" for cell in row_cells[i]]
                    out_lines.append(f"<li>{', '.join(cells)}</li>")
    doc.save(io.BytesIO())
    return io.BytesIO(doc.get_content())

# HTML to Text Conversion
def html_to_text_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    text = soup.get_text(separator="\n")
    return text.encode("utf-8")

def html_to_docx_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    doc = Document()
    for el in soup.body.descendants:
        if el.name.startswith("h") and len(el.name) > 1 and el.name[1].isdigit():
            level = int(el.name[1])
            doc.add_heading(el.get_text(), level=level)
        elif el.name in ("p", "ul", "ol"):
            for li in el.find_all("li"):
                p = doc.add_paragraph(li.get_text())
                p.style = 'List Bullet' if el.name == "ul" else 'List Number'
    rows = '\n'.join([f"<tr>{str(c)}</tr>" for r in el.find_all("tr") for c in r.find_all(["th", "td"])])
    tbl = doc.add_table(rows=0, cols=max(len(r) for r in rows))
    for r in rows:
        cells = tbl.add_row().cells
        for i in range(len(cells)):
            cells[i].text = cells[i][0] if cells[i] else ""
    doc.save(io.BytesIO())
    return io.BytesIO(doc.get_content())

# Streamlit App UI
def create_converter_interface():
    st.set_page_config(page_title="Legacy Converter", layout="wide", page_icon="📚")
    
    st.title("Legacy Converter — Preserve Structure")
    st.markdown("""
    This app tries to preserve structure during conversions:
    - Headings (detected via font size)
    - Paragraphs (text blocks)
    - Lists (bullets/numbered)
    - Tables (via pdfplumber)
    It uses heuristics — tune thresholds from the sidebar for better results on specific documents.
    """)
    
    st.sidebarheader("Conversion Options")
    conversion_choices = ["PDF → Structured HTML", "PDF → Word (.docx)", "PDF → Plain Text", "HTML → Word (.docx)", "HTML → Plain Text"]
    conversion_type = st.selectbox("Conversion", conversion_choices)
    
    workers = st.number_input("Parallel workers (bulk)", min_value=1, max_value=8, value=3)
    embed_pdf = st.checkbox("Embed original PDF into HTML output (only for HTML outputs)")
    heading_ratio = st.slider("Heading size sensitivity (lower = more headings)", 0.5, 2, 1.5)
    
    uploaded_file = st.file_uploader("Upload PDF(s) or HTML(s) — digital PDFs only (no OCR)", type=["pdf", "html"], accept_multiple_files=True)
    
    if not uploaded_file:
        st.info("Upload at least one file. This converter targets digital PDFs (embedded text).")
    
    start = st.button("Create legacy — Convert now")

    def on_create_conversion(page_number):
        with ThreadPoolExecutor(max_workers=workers) as executor:
            future = executor.submit(process_file, uploaded_file.name)
            res = future.result()
            log.f.write(f"✔ {res['name']} → {res['out_name']} ({len(res['out_bytes']):,})")
            st.info(f"Conversion {res['name']} completed successfully.")

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
            preview = content[:4000].decode("utf-8") + "\n" + "".join(f"{line.strip()}" for line in content[4000:]) if content else ""
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
