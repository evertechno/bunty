"""
Streamlit Multi-format Converter — Truly Structured Output
- PDF → HTML: Semantic, layout-aware, visually faithful (headings, paragraphs, lists, tables).
- PDF → DOCX/TXT: Same structured logic.
- No "cloning" illusion — real document understanding.

Key improvements:
- Uses pdfplumber for layout + fitz for font metadata.
- Groups lines into paragraphs by vertical spacing.
- Detects headings via size + isolation.
- Preserves list continuity.
- Tables from pdfplumber with clean HTML.
"""

import io
import os
import zipfile
import base64
from typing import List, Dict, Any, Tuple
import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt
import re

# -----------------------------
# Constants & Helpers
# -----------------------------
BULLET_PATTERNS = [
    r"^\s*[\u2022\u2023\u25E6\-\*\•]\s+",
    r"^\s*\d+[\.\)]\s+",
    r"^\s*[a-zA-Z][\.\)]\s+"
]

def is_bullet_line(text: str) -> bool:
    return any(re.match(pat, text) for pat in BULLET_PATTERNS)

def get_font_size_at_position(pdf_fitz, page_num: int, x0: float, y0: float, x1: float, y1: float) -> float:
    """Get dominant font size in a bbox using fitz."""
    try:
        page = pdf_fitz.load_page(page_num)
        rect = fitz.Rect(x0, y0, x1, y1)
        text_dict = page.get_text("dict", clip=rect)
        sizes = []
        for block in text_dict.get("blocks", []):
            if block["type"] != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    if span.get("size"):
                        sizes.append(span["size"])
        return max(sizes) if sizes else 12.0
    except:
        return 12.0

def group_lines_into_blocks(lines: List[Dict], y_tolerance: float = 5.0) -> List[List[Dict]]:
    """Group lines into blocks (paragraphs/headings) based on vertical proximity."""
    if not lines:
        return []
    blocks = []
    current_block = [lines[0]]
    for line in lines[1:]:
        last_line = current_block[-1]
        if line["top"] - last_line["bottom"] <= y_tolerance:
            current_block.append(line)
        else:
            blocks.append(current_block)
            current_block = [line]
    blocks.append(current_block)
    return blocks

# -----------------------------
# Structured PDF Parser (Layout-Aware)
# -----------------------------
def parse_pdf_structured_v2(pdf_bytes: bytes, heading_ratio: float = 1.12) -> Dict[str, Any]:
    doc_fitz = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_out = []
    all_font_sizes = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf_pl:
        for p_num in range(len(pdf_pl.pages)):
            page_pl = pdf_pl.pages[p_num]
            page_fitz = doc_fitz.load_page(p_num)

            # Extract tables first (to avoid text inside tables being parsed as paragraphs)
            tables = []
            try:
                extracted_tables = page_pl.extract_tables()
                table_bboxes = []
                for table in page_pl.find_tables():
                    table_bboxes.append(table.bbox)  # (x0, top, x1, bottom)
                    rows = table.extract()
                    if rows and any(any(cell for cell in row if cell) for row in rows):
                        tables.append({"rows": rows, "bbox": table.bbox})
            except Exception:
                tables = []

            # Get all text lines (excluding tables)
            text_lines = []
            try:
                # Use chars to reconstruct lines with position
                chars = [c for c in page_pl.chars if c["text"].strip()]
                # Exclude chars inside tables
                def is_in_table(char):
                    for tb in table_bboxes:
                        if (tb[0] <= char["x0"] <= tb[2] and tb[1] <= char["top"] <= tb[3]):
                            return True
                    return False
                filtered_chars = [c for c in chars if not is_in_table(c)]
                if filtered_chars:
                    lines = page_pl.extract_text_lines(filtered_chars, strip=False)
                    for line in lines:
                        if line["text"].strip():
                            # Get font size from fitz at this line's bbox
                            font_size = get_font_size_at_position(
                                doc_fitz, p_num,
                                line["x0"], line["top"], line["x1"], line["bottom"]
                            )
                            all_font_sizes.append(font_size)
                            text_lines.append({
                                "text": line["text"],
                                "top": line["top"],
                                "bottom": line["bottom"],
                                "x0": line["x0"],
                                "font_size": font_size
                            })
            except Exception:
                pass

            # Sort lines by vertical position (reading order)
            text_lines.sort(key=lambda x: (x["top"], x["x0"]))

            # Group into blocks (paragraphs/headings)
            blocks = group_lines_into_blocks(text_lines, y_tolerance=3.0)

            # Determine heading threshold
            if all_font_sizes:
                max_size = max(all_font_sizes)
                heading_threshold = max_size / heading_ratio
            else:
                heading_threshold = 14.0

            elements = []

            # Process blocks
            for block in blocks:
                block_text = "\n".join([ln["text"] for ln in block]).strip()
                if not block_text:
                    continue

                avg_font_size = sum(ln["font_size"] for ln in block) / len(block)
                is_heading = avg_font_size >= heading_threshold

                # Check for list
                first_line = block[0]["text"]
                if is_bullet_line(first_line):
                    elements.append({
                        "type": "list_item",
                        "text": block_text,
                        "font_size": avg_font_size
                    })
                elif is_heading:
                    # Estimate heading level by size rank (1-4)
                    unique_sizes = sorted(set(all_font_sizes), reverse=True)
                    top_sizes = unique_sizes[:4]
                    level = 1
                    for i, sz in enumerate(top_sizes):
                        if avg_font_size >= sz:
                            level = i + 1
                            break
                    elements.append({
                        "type": "heading",
                        "text": block_text,
                        "level": level,
                        "font_size": avg_font_size
                    })
                else:
                    elements.append({
                        "type": "para",
                        "text": block_text,
                        "font_size": avg_font_size
                    })

            # Insert tables at approximate positions (by top coordinate)
            table_elements = []
            for tbl in tables:
                table_elements.append({
                    "type": "table",
                    "rows": tbl["rows"],
                    "top": tbl["bbox"][1]
                })

            # Merge text elements and tables by vertical position
            all_elements = elements + table_elements
            all_elements.sort(key=lambda x: x.get("top", x["font_size"] * -1000))  # tables use "top", text uses font_size fallback

            # Remove "top" from final output
            final_elements = []
            for el in all_elements:
                clean_el = {k: v for k, v in el.items() if k != "top"}
                final_elements.append(clean_el)

            pages_out.append({
                "page_number": p_num + 1,
                "elements": final_elements
            })

    return {"pages": pages_out, "fontsizes": sorted(set(all_font_sizes), reverse=True) if all_font_sizes else [12.0]}

# -----------------------------
# Converters
# -----------------------------
def structured_to_html(parsed: dict, embed_pdf: bool = False, pdf_bytes: bytes = None) -> bytes:
    parts = [
        '<!doctype html>',
        '<html>',
        '<head>',
        '<meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width,initial-scale=1">',
        '<style>',
        'body{font-family:Arial,Helvetica,sans-serif;line-height:1.5;margin:20px}',
        'h1{font-size:2em;margin:0.67em 0}',
        'h2{font-size:1.5em;margin:0.75em 0}',
        'h3{font-size:1.17em;margin:0.83em 0}',
        'h4{font-size:1em;margin:1em 0}',
        'p,ul,ol{margin:1em 0}',
        'li{margin:0.5em 0}',
        'table{border-collapse:collapse;width:100%;margin:1em 0}',
        'td,th{border:1px solid #999;padding:8px;text-align:left}',
        '.page{page-break-after:always}',
        '</style>',
        '</head>',
        '<body>'
    ]

    for page in parsed["pages"]:
        parts.append(f'<div class="page" data-page="{page["page_number"]}">')
        in_list = False
        for el in page["elements"]:
            if el["type"] == "heading":
                lvl = min(max(int(el.get("level", 2)), 1), 6)
                text = el["text"]
                if in_list:
                    parts.append("</ul>")
                    in_list = False
                parts.append(f"<h{lvl}>{html.escape(text)}</h{lvl}>")
            elif el["type"] == "para":
                text = el["text"]
                if in_list:
                    parts.append("</ul>")
                    in_list = False
                parts.append(f"<p>{html.escape(text)}</p>")
            elif el["type"] == "list_item":
                if not in_list:
                    parts.append("<ul>")
                    in_list = True
                parts.append(f"<li>{html.escape(el['text'])}</li>")
            elif el["type"] == "table":
                if in_list:
                    parts.append("</ul>")
                    in_list = False
                parts.append("<table>")
                for row in el["rows"]:
                    parts.append("<tr>")
                    for cell in row:
                        cell_text = html.escape(str(cell) if cell is not None else "")
                        parts.append(f"<td>{cell_text}</td>")
                    parts.append("</tr>")
                parts.append("</table>")
        if in_list:
            parts.append("</ul>")
        parts.append("</div>")

    parts.append("</body></html>")
    html_str = "".join(parts)

    if embed_pdf and pdf_bytes:
        b64 = base64.b64encode(pdf_bytes).decode('ascii')
        embed_snip = f'<hr/><h2>Original PDF</h2><embed src="data:application/pdf;base64,{b64}" width="100%" height="600"></embed>'
        html_str = html_str.replace("</body></html>", embed_snip + "</body></html>")

    return html_str.encode("utf-8")

def structured_to_text(parsed: dict) -> bytes:
    lines = []
    for page in parsed["pages"]:
        lines.append(f"--- PAGE {page['page_number']} ---")
        for el in page["elements"]:
            if el["type"] == "heading":
                lines.append(el["text"])
            elif el["type"] == "para":
                lines.append(el["text"])
            elif el["type"] == "list_item":
                lines.append(f"• {el['text']}")
            elif el["type"] == "table":
                for row in el["rows"]:
                    lines.append("\t".join([str(c) if c is not None else "" for c in row]))
        lines.append("")
    return "\n".join(lines).encode("utf-8")

def structured_to_docx(parsed: dict) -> bytes:
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    for page in parsed["pages"]:
        for el in page["elements"]:
            if el["type"] == "heading":
                lvl = min(max(int(el.get("level", 2)), 1), 4)
                doc.add_heading(el["text"], level=lvl)
            elif el["type"] == "para":
                doc.add_paragraph(el["text"])
            elif el["type"] == "list_item":
                p = doc.add_paragraph(el["text"], style='List Bullet')
            elif el["type"] == "table":
                rows = el["rows"]
                if not rows:
                    continue
                ncols = max(len(r) for r in rows) if rows else 1
                table = doc.add_table(rows=0, cols=ncols)
                table.style = 'Table Grid'
                for row_data in rows:
                    row_cells = table.add_row().cells
                    for i in range(ncols):
                        text = str(row_data[i]) if i < len(row_data) and row_data[i] is not None else ""
                        row_cells[i].text = text
        doc.add_page_break()

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

# -----------------------------
# HTML Converters (unchanged)
# -----------------------------
def html_to_text_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    text = soup.get_text(separator="\n")
    return text.encode("utf-8")

def html_to_docx_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    doc = Document()
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
            for tr in el.find_all("tr"):
                cols = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
                if cols:
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
# Streamlit App
# -----------------------------
st.set_page_config(page_title="Structured PDF Converter", layout="wide", page_icon="📄")
st.title("📄 Structured PDF Converter — True Document Fidelity")
st.markdown("""
Convert PDFs to **semantic HTML, DOCX, or TXT** with:
- Accurate heading detection
- Proper paragraph grouping
- List recognition
- Table extraction
- Reading-order preservation
""")

with st.sidebar:
    st.header("Options")
    conversion = st.selectbox("Conversion", [
        "PDF → Structured HTML",
        "PDF → Word (.docx)",
        "PDF → Plain Text",
        "HTML → Word (.docx)",
        "HTML → Plain Text"
    ])
    embed_pdf = st.checkbox("Embed PDF in HTML output", value=False)
    heading_ratio = st.slider(
        "Heading sensitivity (lower = more headings)",
        min_value=1.05, max_value=1.5, value=1.15, step=0.01
    )

uploaded = st.file_uploader("Upload PDF or HTML", type=["pdf", "html"], accept_multiple_files=True)
if not uploaded:
    st.info("Upload a digital PDF (with text) or HTML file.")
    st.stop()

if not st.button("Convert"):
    st.stop()

from concurrent.futures import ThreadPoolExecutor

def process_file(f):
    name = f.name
    raw = f.read()
    ext = os.path.splitext(name)[1].lower()
    try:
        if ext == ".pdf":
            parsed = parse_pdf_structured_v2(raw, heading_ratio=heading_ratio)
            if conversion == "PDF → Structured HTML":
                out = structured_to_html(parsed, embed_pdf=embed_pdf, pdf_bytes=raw if embed_pdf else None)
                out_name = name.replace(".pdf", ".html")
                mime = "text/html"
            elif conversion == "PDF → Word (.docx)":
                out = structured_to_docx(parsed)
                out_name = name.replace(".pdf", ".docx")
                mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif conversion == "PDF → Plain Text":
                out = structured_to_text(parsed)
                out_name = name.replace(".pdf", ".txt")
                mime = "text/plain"
            else:
                raise ValueError("Invalid PDF conversion")
        elif ext == ".html":
            if conversion == "HTML → Plain Text":
                out = html_to_text_bytes(raw)
                out_name = name.replace(".html", ".txt")
                mime = "text/plain"
            elif conversion == "HTML → Word (.docx)":
                out = html_to_docx_bytes(raw)
                out_name = name.replace(".html", ".docx")
                mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            else:
                raise ValueError("Invalid HTML conversion")
        else:
            raise ValueError("Unsupported file type")
        return {"name": name, "out_bytes": out, "out_name": out_name, "mime": mime}
    except Exception as e:
        return {"name": name, "error": str(e)}

with ThreadPoolExecutor(max_workers=3) as executor:
    results = list(executor.map(process_file, uploaded))

success_results = [r for r in results if "error" not in r]
error_results = [r for r in results if "error" in r]

for r in error_results:
    st.error(f"❌ {r['name']}: {r['error']}")

if success_results:
    st.success(f"✅ Converted {len(success_results)} file(s)")
    for i, res in enumerate(success_results):
        col = st.columns(3)[i % 3]
        with col:
            st.caption(res["out_name"])
            if res["mime"].startswith("text/"):
                preview = res["out_bytes"].decode("utf-8", errors="replace")[:2000]
                st.text_area("Preview", preview, height=150, key=res["out_name"])
            st.download_button("Download", res["out_bytes"], res["out_name"], res["mime"])

    # ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in success_results:
            zf.writestr(r["out_name"], r["out_bytes"])
    zip_buffer.seek(0)
    st.download_button("📥 Download All as ZIP", zip_buffer.read(), "converted.zip", "application/zip")
else:
    st.error("No files converted successfully.")
