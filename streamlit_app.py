"""
streamlit_converter_structured.py

Streamlit Multi-format Converter — Structured, "as-is" output.
Heuristics used:
 - PyMuPDF (fitz) to read text blocks / spans with font sizes -> detect headings vs paragraphs
 - pdfplumber to detect tabular data -> produce <table> and docx tables
 - BeautifulSoup for HTML->Text/Word conversions
 - python-docx to create Word (.docx) with headings, paragraphs, lists, tables

Notes:
 - This targets digital PDFs (embedded text). No OCR is performed.
 - Heuristics (font-size thresholds, list detection) can be tuned in the UI.
 - The goal is to preserve structured *content* as faithfully as possible,
   not necessarily pixel-perfect visual layout, especially for PDF conversions.
"""
import io
import os
import zipfile
import base64
from typing import List, Tuple, Dict, Any, Optional

import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
from bs4 import BeautifulSoup, Tag
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
import pandas as pd
import html
import re

# -----------------------------
# Utility / Heuristic Functions
# -----------------------------

BULLET_CHARS = r"^[\u2022\u2023\u25E6\-\*\•\–\—]\s*" # Added * to end, removed digit pattern
NUMBER_CHARS = r"^\d+[\.\)]\s*"

def sanitize_text(t: str) -> str:
    """Removes common control characters and leading/trailing whitespace."""
    return t.replace('\r', '').strip()

def is_bullet_line(text: str) -> bool:
    """Checks if a line starts with common bullet characters."""
    return bool(re.match(BULLET_CHARS, text))

def is_numbered_line(text: str) -> bool:
    """Checks if a line starts with a number pattern."""
    return bool(re.match(NUMBER_CHARS, text))

def normalize_whitespace_for_output(s: str) -> str:
    """
    Simplifies whitespace for cleaner output, keeping essential line breaks.
    This is less aggressive than `re.sub(r'\s+', ' ', s).strip()` to preserve
    some natural line breaks within paragraphs.
    """
    s = re.sub(r'[ \t]+', ' ', s) # Multiple spaces to single space
    s = re.sub(r'\n{2,}', '\n\n', s) # Multiple newlines to at most two
    return s.strip()


def choose_heading_levels(unique_sizes: List[float], max_levels: int = 4) -> Dict[float, int]:
    """
    Map font sizes -> heading levels (1..max_levels) heuristically.
    Biggest -> h1, next -> h2, etc. If many sizes, map top ones to headings.
    Smaller sizes not in top `max_levels` get mapped to 0 (paragraph).
    """
    if not unique_sizes:
        return {}
    
    # Sort distinct sizes from largest to smallest
    sorted_sizes = sorted(list(set(unique_sizes)), reverse=True)
    
    mapping = {}
    # Assign heading levels to the largest unique font sizes
    for idx, size in enumerate(sorted_sizes[:max_levels]):
        mapping[round(size, 2)] = idx + 1 # Levels 1 to max_levels
    
    return mapping

# -----------------------------
# PDF Parsing (structured)
# -----------------------------

def parse_pdf_structured(pdf_bytes: bytes, min_heading_ratio: float = 1.12, min_para_size: float = 8.0) -> Dict[str, Any]:
    """
    Parse PDF into a structured intermediate representation.
    Heuristics:
     - Uses PyMuPDF blocks & spans for text and font sizes.
     - Uses pdfplumber for table detection.
     - Identify bullets/numbers via regex.
     - Infers headings based on font size.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_out = []
    all_sizes = []

    # First pass: Collect all non-zero font sizes to determine document's font hierarchy
    for p_idx in range(len(doc)):
        page = doc.load_page(p_idx)
        d = page.get_text("dict")
        for block in d.get("blocks", []):
            if block.get("type") != 0: # Skip images
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    sz = round(span.get("size", 2), 2) # Default to 2 to avoid zero size issues if span has no size
                    if sz > 0:
                        all_sizes.append(sz)

    if not all_sizes:
        unique_sizes = [12.0] # Default if no font sizes found (e.g., empty PDF)
    else:
        # Filter out very small, likely decorative or artefact sizes
        unique_sizes = sorted(set(s for s in all_sizes if s >= min_para_size), reverse=True)

    # Determine heading levels based on unique font sizes
    font_to_heading = choose_heading_levels(unique_sizes)
    
    # Find the most common (likely body text) font size
    body_text_size = 0.0
    if all_sizes:
        # Use Counter to find the most frequent size
        from collections import Counter
        size_counts = Counter(round(s, 2) for s in all_sizes if s >= min_para_size)
        if size_counts:
            body_text_size = size_counts.most_common(1)[0][0]
            # Ensure body text size is not accidentally mapped as a heading
            if body_text_size in font_to_heading and font_to_heading[body_text_size] > 0:
                # If the most common size is a heading, it implies the document
                # has very large body text or unusual hierarchy.
                # We can adjust by saying if it's the smallest mapped heading, treat as para.
                if font_to_heading[body_text_size] == max(font_to_heading.values()):
                    del font_to_heading[body_text_size] # Demote the smallest heading if it's the body size

    # Default heading threshold if unique_sizes is empty or too few elements
    max_doc_size = max(unique_sizes) if unique_sizes else 12.0
    heading_threshold = max_doc_size / min_heading_ratio 

    # Second pass: Parse each page into elements
    for p_idx in range(len(doc)):
        page = doc.load_page(p_idx)
        elements = []

        # Table detection with pdfplumber
        tables_on_page = []
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as ppdf:
                page_pl = ppdf.pages[p_idx]
                extracted_tables = page_pl.extract_tables()
                for t in extracted_tables:
                    if t and any(any(cell for cell in row if cell not in (None, "")) for row in t):
                        # Clean up cells: replace None with empty string
                        cleaned_table = [[str(c) if c is not None else "" for c in row] for row in t]
                        tables_on_page.append({"rows": cleaned_table})
        except Exception as e:
            st.warning(f"Error extracting tables from PDF page {p_idx+1}: {e}. Skipping table extraction for this page.")
            tables_on_page = []

        # Text block processing with PyMuPDF
        current_list_items = [] # To group consecutive list items
        for block in d.get("blocks", []):
            if block.get("type") != 0: # Skip images and other non-text blocks
                continue

            block_text_lines = []
            max_block_span_sz = 0.0
            
            # Combine lines within a block, track max font size for heading inference
            for line in block.get("lines", []):
                line_spans_text = []
                for span in line.get("spans", []):
                    stxt = span.get("text", "")
                    if stxt:
                        line_spans_text.append(stxt)
                    sz = span.get("size", 0)
                    if sz > max_block_span_sz:
                        max_block_span_sz = sz
                
                # Join spans into a line, then add to block lines if not empty
                if line_spans_text:
                    block_text_lines.append("".join(line_spans_text).strip())

            # Process block_text_lines
            if not block_text_lines:
                continue
            
            # Flush existing list items before processing new block
            if current_list_items:
                elements.append({"type": "list", "items": current_list_items, "list_type": "bullet"}) # Default to bullet
                current_list_items = []

            # Determine if this block is a list or normal text
            is_block_list = all(is_bullet_line(line) or is_numbered_line(line) for line in block_text_lines if line)
            
            if is_block_list and block_text_lines:
                list_type = "numbered" if any(is_numbered_line(line) for line in block_text_lines) else "bullet"
                for ln in block_text_lines:
                    cleaned_ln = sanitize_text(ln)
                    if cleaned_ln:
                        current_list_items.append(cleaned_ln)
                # Group all lines from this block as a single list if it's purely list items
                if current_list_items:
                    elements.append({"type": "list", "items": current_list_items, "list_type": list_type})
                    current_list_items = [] # Reset after adding
            else:
                # It's not a list, process as paragraphs or headings
                block_content = "\n".join(block_text_lines)
                cleaned_block_content = sanitize_text(block_content)
                if not cleaned_block_content:
                    continue

                # Heading detection
                inferred_level = font_to_heading.get(round(max_block_span_sz, 2), 0)
                if inferred_level > 0 or (max_block_span_sz >= heading_threshold and max_block_span_sz > body_text_size):
                    elements.append({"type": "heading", "text": cleaned_block_content, "level": inferred_level if inferred_level > 0 else 2, "size": max_block_span_sz})
                else:
                    elements.append({"type": "para", "text": cleaned_block_content, "size": max_block_span_sz})

        # Ensure any leftover list items are added
        if current_list_items:
            elements.append({"type": "list", "items": current_list_items, "list_type": "bullet"})


        # Integrate detected tables into the elements list.
        # This is a simplification; precise positioning would require
        # comparing bounding boxes of text blocks and tables.
        # For 'cloning', tables are often distinct visual elements.
        for t in tables_on_page:
            elements.append({"type": "table", "rows": t["rows"]})

        pages_out.append({"page_number": p_idx + 1, "elements": elements})

    return {"pages": pages_out, "fontsizes": unique_sizes}


# -----------------------------
# Converters: Intermediate -> HTML / DOCX / TEXT
# -----------------------------

def structured_to_html(parsed: dict, embed_pdf: bool = False, pdf_bytes: bytes = None) -> bytes:
    parts = [
        '<!doctype html>',
        '<html>',
        '<head>',
        '<meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width,initial-scale=1">',
        '<title>Converted Document</title>',
        '<style>',
        'body{font-family:Arial,Helvetica,sans-serif;line-height:1.6;padding:16px;margin:0;}',
        'pre{white-space:pre-wrap;font-family:monospace;}',
        'table{width:100%;border-collapse:collapse;margin:1em 0;table-layout:fixed;}',
        'td,th{border:1px solid #ccc;padding:8px;text-align:left;vertical-align:top;word-wrap:break-word;}',
        'th{background-color:#f0f0f0;font-weight:bold;}',
        '.page{page-break-after:always; margin-bottom: 2em; padding: 1em; border: 1px dashed #eee;}',
        '.page-number{text-align:center; color:#888; font-size:0.9em; margin-bottom:1em;}',
        'h1, h2, h3, h4, h5, h6 { margin-top: 1em; margin-bottom: 0.5em; }',
        'ul, ol { margin-left: 1.5em; margin-bottom: 1em; }',
        'p { margin-bottom: 1em; }',
        '</style>',
        '</head>',
        '<body>'
    ]
    
    for page in parsed["pages"]:
        parts.append(f'<div class="page" data-page="{page["page_number"]}">')
        parts.append(f'<p class="page-number">--- Page {page["page_number"]} ---</p>')
        
        # Track current list type to properly open/close ul/ol tags
        current_list_type = None # "bullet" or "numbered"
        
        for el in page["elements"]:
            if el["type"] == "heading":
                # Close any open list before a heading
                if current_list_type:
                    parts.append(f'</{"ul" if current_list_type == "bullet" else "ol"}>')
                    current_list_type = None
                
                lvl = min(max(int(el.get("level", 2)), 1), 6) # Ensure level is between 1 and 6
                text = html.escape(normalize_whitespace_for_output(el["text"]))
                parts.append(f"<h{lvl}>{text}</h{lvl}>")
            
            elif el["type"] == "para":
                # Close any open list before a paragraph
                if current_list_type:
                    parts.append(f'</{"ul" if current_list_type == "bullet" else "ol"}>')
                    current_list_type = None
                
                text = html.escape(normalize_whitespace_for_output(el["text"]))
                parts.append(f"<p>{text}</p>")
            
            elif el["type"] == "list":
                list_type_html = "ul" if el["list_type"] == "bullet" else "ol"
                
                # If list type changes or no list is open, close current and open new
                if current_list_type != el["list_type"]:
                    if current_list_type: # Close previous list if open
                        parts.append(f'</{"ul" if current_list_type == "bullet" else "ol"}>')
                    parts.append(f'<{list_type_html}>')
                    current_list_type = el["list_type"]
                
                for item_text in el["items"]:
                    parts.append(f"<li>{html.escape(normalize_whitespace_for_output(item_text))}</li>")
            
            elif el["type"] == "table":
                # Close any open list before a table
                if current_list_type:
                    parts.append(f'</{"ul" if current_list_type == "bullet" else "ol"}>')
                    current_list_type = None
                
                rows = el["rows"]
                parts.append("<table>")
                for r_idx, r in enumerate(rows):
                    parts.append("<tr>")
                    # Assume first row is header if table has more than one row and it seems like a header
                    tag = "th" if r_idx == 0 and len(rows) > 1 else "td"
                    parts.append("".join(f"<{tag}>{html.escape(str(c) if c is not None else '')}</{tag}>" for c in r))
                    parts.append("</tr>")
                parts.append("</table>")
        
        # After processing all elements on a page, close any open list
        if current_list_type:
            parts.append(f'</{"ul" if current_list_type == "bullet" else "ol"}>')
            current_list_type = None
            
        parts.append("</div>") # Close page div

    # Embed original PDF optionally
    if embed_pdf and pdf_bytes:
        b64 = base64.b64encode(pdf_bytes).decode('ascii')
        embed_snip = f'<hr/><h2>Original PDF (embedded)</h2><p><i>Note: The embedded PDF is for reference and may have a different visual layout than the converted HTML.</i></p><embed src="data:application/pdf;base64,{b64}" width="100%" height="600px" type="application/pdf"></embed>'
        parts.append(embed_snip)
        
    parts.append("</body></html>")
    return "\n".join(parts).encode("utf-8")


def structured_to_text(parsed: dict) -> bytes:
    out_lines = []
    for page in parsed["pages"]:
        out_lines.append(f"\n--- PAGE {page['page_number']} ---\n")
        
        for el in page["elements"]:
            if el["type"] == "heading":
                out_lines.append(normalize_whitespace_for_output(el["text"]).upper())
                out_lines.append("-" * len(el["text"])) # Simple underline for headings
                out_lines.append("")
            elif el["type"] == "para":
                out_lines.append(normalize_whitespace_for_output(el["text"]))
                out_lines.append("")
            elif el["type"] == "list":
                for item_text in el["items"]:
                    prefix = "- " if el["list_type"] == "bullet" else "1. " # Simplified numbering for text
                    out_lines.append(f"{prefix}{normalize_whitespace_for_output(item_text)}")
                out_lines.append("")
            elif el["type"] == "table":
                rows = el["rows"]
                # A simple text-based table approximation, tab-separated or fixed width
                for r_idx, r in enumerate(rows):
                    # Pad each cell for alignment, or just join with tabs
                    cells_formatted = [str(c) if c is not None else "" for c in r]
                    out_lines.append("\t".join(cells_formatted))
                    if r_idx == 0 and len(rows) > 1: # Header separator
                        out_lines.append("\t".join(["-" * len(cell) for cell in cells_formatted]))
                out_lines.append("") # Add a blank line after table
    joined = "\n".join(out_lines).strip()
    return joined.encode("utf-8")


def structured_to_docx(parsed: dict) -> bytes:
    doc = Document()
    
    # Define default styles for better consistency
    styles = doc.styles
    if 'Normal' not in styles:
        styles.add_style('Normal', WD_STYLE_TYPE.PARAGRAPH)
    styles['Normal'].font.name = 'Arial'
    styles['Normal'].font.size = Pt(11)
    
    # Ensure heading styles exist, or create them
    for i in range(1, 5): # H1 to H4
        if f'Heading {i}' not in styles:
            styles.add_style(f'Heading {i}', WD_STYLE_TYPE.PARAGRAPH)
        styles[f'Heading {i}'].font.name = 'Arial'
        styles[f'Heading {i}'].font.size = Pt(16 - i * 2) # Example sizing
        styles[f'Heading {i}'].paragraph_format.space_before = Pt(12)
        styles[f'Heading {i}'].paragraph_format.space_after = Pt(6)

    # List styles
    if 'List Bullet' not in styles:
        styles.add_style('List Bullet', WD_STYLE_TYPE.PARAGRAPH)
        styles['List Bullet'].font.name = 'Arial'
        styles['List Bullet'].font.size = Pt(11)
    if 'List Number' not in styles:
        styles.add_style('List Number', WD_STYLE_TYPE.PARAGRAPH)
        styles['List Number'].font.name = 'Arial'
        styles['List Number'].font.size = Pt(11)

    for page in parsed["pages"]:
        # Add a "Page X" marker
        p_page_num = doc.add_paragraph(f"--- Page {page['page_number']} ---")
        p_page_num.style = 'Normal'
        p_page_num.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        for el in page["elements"]:
            if el["type"] == "heading":
                lvl = min(max(int(el.get("level", 2)), 1), 4) # DOCX supports Heading 1-9, but 1-4 are common
                p = doc.add_heading(normalize_whitespace_for_output(el["text"]), level=lvl)
            elif el["type"] == "para":
                p = doc.add_paragraph(normalize_whitespace_for_output(el["text"]))
                p.style = 'Normal'
            elif el["type"] == "list":
                list_style = 'List Bullet' if el["list_type"] == "bullet" else 'List Number'
                for item_text in el["items"]:
                    p = doc.add_paragraph(normalize_whitespace_for_output(item_text))
                    p.style = list_style
            elif el["type"] == "table":
                rows = el["rows"]
                if not rows:
                    continue

                ncols = max(len(r) for r in rows)
                # Create table with header and data rows
                tbl = doc.add_table(rows=0, cols=ncols)
                tbl.style = 'Table Grid' # Apply a basic grid style

                for r_idx, r in enumerate(rows):
                    row_cells = tbl.add_row().cells
                    for i in range(ncols):
                        cell_text = str(r[i]) if i < len(r) and r[i] is not None else ""
                        row_cells[i].text = cell_text
                        if r_idx == 0: # Apply bold to header row
                            row_cells[i].paragraphs[0].runs[0].bold = True
                
                # Add some spacing after the table
                doc.add_paragraph("")
                
        # Add a page break after each page's content
        doc.add_page_break()
        
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# -----------------------------
# HTML -> Text / DOCX
# -----------------------------

def html_to_text_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    # Use different separator to distinguish block elements better
    text = soup.get_text(separator="\n\n")
    return normalize_whitespace_for_output(text).encode("utf-8")

def html_to_docx_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    doc = Document()
    
    # Default styles
    styles = doc.styles
    if 'Normal' not in styles:
        styles.add_style('Normal', WD_STYLE_TYPE.PARAGRAPH)
    styles['Normal'].font.name = 'Arial'
    styles['Normal'].font.size = Pt(11)

    # Iterate through body children to respect block element order
    for element in soup.body.children:
        if isinstance(element, Tag):
            if element.name and element.name.startswith("h") and element.get_text(strip=True):
                level = int(element.name[1]) if len(element.name) > 1 and element.name[1].isdigit() else 2
                doc.add_heading(normalize_whitespace_for_output(element.get_text(strip=True)), level=min(level, 4))
            elif element.name == "p" and element.get_text(strip=True):
                doc.add_paragraph(normalize_whitespace_for_output(element.get_text("\n", strip=True))).style = 'Normal'
            elif element.name in ("ul", "ol"):
                list_style = 'List Bullet' if element.name == "ul" else 'List Number'
                for li in element.find_all("li", recursive=False): # Only direct children li
                    if li.get_text(strip=True):
                        p = doc.add_paragraph(normalize_whitespace_for_output(li.get_text(strip=True)))
                        p.style = list_style
            elif element.name == "table":
                rows_data = []
                for r in element.find_all("tr", recursive=False):
                    cols = [normalize_whitespace_for_output(c.get_text(strip=True)) for c in r.find_all(["th", "td"], recursive=False)]
                    rows_data.append(cols)
                
                if rows_data:
                    ncols = max(len(r) for r in rows_data)
                    tbl = doc.add_table(rows=0, cols=ncols)
                    tbl.style = 'Table Grid'
                    for r_idx, r in enumerate(rows_data):
                        cells = tbl.add_row().cells
                        for i in range(ncols):
                            cells[i].text = r[i] if i < len(r) else ""
                            if r_idx == 0 and element.find("th"): # If table has a th, assume first row is header
                                cells[i].paragraphs[0].runs[0].bold = True
                    doc.add_paragraph("") # Add a blank line after table
            elif element.name == "hr":
                doc.add_paragraph("---") # Simple representation of a horizontal rule
            
            # TODO: Handle img, pre, blockquote, etc. for more comprehensive HTML conversion

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# -----------------------------
# Streamlit App UI
# -----------------------------

st.set_page_config(page_title="Legacy Converter — Structured Cloning", layout="wide", page_icon="📚")
st.title("📚 Legacy Converter — Preserve structure, create a legacy")

st.markdown("""
This app aims to **clone the *structured content*** of your documents with high fidelity:
-   **Headings:** Detected via font size hierarchy.
-   **Paragraphs:** Main text blocks.
-   **Lists:** Bulleted and numbered items.
-   **Tables:** Extracted using advanced table detection.

**Goal:** To achieve a conversion where the *textual and semantic content* is replicated as precisely as possible ("not even a comma" difference in *content*).
**Limitations:**
-   **Digital PDFs only:** No OCR is performed on scanned PDFs or images.
-   **Visual layout:** Exact pixel-perfect rendering (e.g., specific font faces, colors, precise spacing, complex multi-column layouts, graphics) is challenging and not the primary focus for *structured text cloning*.
-   **File Size:** Very large files might take time or hit Streamlit's memory limits.
""")

with st.sidebar:
    st.header("Conversion Options")
    conversion = st.selectbox("Select Conversion Type", [
        "PDF → Structured HTML",
        "PDF → Word (.docx)",
        "PDF → Plain Text",
        "HTML → Word (.docx)",
        "HTML → Plain Text"
    ])
    
    st.markdown("### Advanced Settings")
    workers = st.number_input("Parallel processing limit", min_value=1, max_value=8, value=3, help="Number of files to process simultaneously. Adjust based on your machine's CPU/memory.")
    embed_pdf = st.checkbox("Embed original PDF into HTML output", value=False, help="For PDF to HTML, includes the original PDF content base64 encoded within the HTML output for reference.")
    
    st.markdown("### Heuristics Tuning (for PDF conversion)")
    heading_ratio = st.slider("Heading size sensitivity (lower = more headings)", min_value=1.05, max_value=1.5, value=1.12, step=0.01, help="A lower value will make the converter more likely to identify text as a heading (e.g., if a font size is only slightly larger than body text). Increase to be more strict.")
    min_para_size = st.slider("Minimum text font size (pixels)", min_value=4.0, max_value=12.0, value=8.0, step=0.5, help="Text smaller than this might be considered noise or ignored for structure detection. Default is 8pt.")

uploaded_files = st.file_uploader("Upload PDF(s) or HTML(s) — Digital PDFs for best results.", type=["pdf", "html"], accept_multiple_files=True)

if not uploaded_files:
    st.info("Upload at least one file to begin conversion.")
    st.stop()

st.markdown(f"**Files queued for conversion:** {len(uploaded_files)} file(s)")
start_conversion = st.button("🚀 Start Conversion Now")

if not start_conversion:
    st.stop()

# --- Conversion Logic ---
results_for_zip = []
errors_occurred = []

def process_single_file(uploaded_file_obj):
    file_name = uploaded_file_obj.name
    raw_bytes = uploaded_file_obj.read()
    file_ext = os.path.splitext(file_name)[1].lower()
    
    result_entry = {"name": file_name}
    
    try:
        if file_ext == ".pdf":
            parsed_content = parse_pdf_structured(raw_bytes, min_heading_ratio=heading_ratio, min_para_size=min_para_size)
            if conversion == "PDF → Structured HTML":
                output_bytes = structured_to_html(parsed_content, embed_pdf=embed_pdf, pdf_bytes=raw_bytes if embed_pdf else None)
                output_name = os.path.splitext(file_name)[0] + ".html"
                mime_type = "text/html"
            elif conversion == "PDF → Word (.docx)":
                output_bytes = structured_to_docx(parsed_content)
                output_name = os.path.splitext(file_name)[0] + ".docx"
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif conversion == "PDF → Plain Text":
                output_bytes = structured_to_text(parsed_content)
                output_name = os.path.splitext(file_name)[0] + ".txt"
                mime_type = "text/plain"
            else:
                raise ValueError(f"Conversion '{conversion}' is not valid for PDF files.")
                
        elif file_ext == ".html":
            if conversion == "HTML → Plain Text":
                output_bytes = html_to_text_bytes(raw_bytes)
                output_name = os.path.splitext(file_name)[0] + ".txt"
                mime_type = "text/plain"
            elif conversion == "HTML → Word (.docx)":
                output_bytes = html_to_docx_bytes(raw_bytes)
                output_name = os.path.splitext(file_name)[0] + ".docx"
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            else:
                raise ValueError(f"Conversion '{conversion}' is not valid for HTML files.")
        else:
            raise ValueError(f"Unsupported file type: {file_ext}")
            
        result_entry.update({"out_bytes": output_bytes, "out_name": output_name, "mime": mime_type})
        return result_entry
        
    except Exception as e:
        result_entry["error"] = str(e)
        return result_entry

# UI for progress and status
progress_bar = st.progress(0)
status_message = st.empty()
conversion_log = st.empty()

# Use ThreadPoolExecutor for concurrent processing
with ThreadPoolExecutor(max_workers=workers) as executor:
    future_to_file = {executor.submit(process_single_file, f): f.name for f in uploaded_files}
    
    completed_count = 0
    for future in as_completed(future_to_file):
        completed_count += 1
        file_name_processed = future_to_file[future]
        progress_percent = completed_count / len(uploaded_files)
        progress_bar.progress(progress_percent)
        
        try:
            result = future.result()
            if "error" in result:
                errors_occurred.append(result)
                conversion_log.write(f"❌ Failed to convert '{result['name']}': {result['error']}")
            else:
                results_for_zip.append(result)
                conversion_log.write(f"✅ Converted '{result['name']}' to '{result['out_name']}' ({len(result['out_bytes']):,} bytes)")
        except Exception as exc:
            errors_occurred.append({"name": file_name_processed, "error": f"Unhandled exception during processing: {exc}"})
            conversion_log.write(f"❌ Failed to convert '{file_name_processed}': Unhandled error: {exc}")

if errors_occurred:
    status_message.error(f"Conversion completed with {len(errors_occurred)} error(s). Please check the log above.")
else:
    status_message.success("All files converted successfully!")

# Display results and download options
if results_for_zip:
    st.markdown("---")
    st.markdown("### Download Converted Files")
    
    # Create columns for individual downloads
    num_cols = min(len(results_for_zip), 3) # Max 3 columns
    cols = st.columns(num_cols)
    
    for i, res in enumerate(results_for_zip):
        with cols[i % num_cols]:
            st.markdown(f"**{res['out_name']}**")
            
            # Preview for text-based outputs
            if res["mime"].startswith("text/plain"):
                preview_text = res["out_bytes"].decode("utf-8", errors="replace")
                st.text_area(f"Preview (first 2000 chars)", preview_text[:2000], height=180, key=f"preview_txt_{i}")
            elif res["mime"].startswith("text/html"):
                # HTML preview, limited to avoid browser/memory issues with huge HTML
                # Streamlit's html component is sandboxed, so external links won't work.
                try:
                    st.components.v1.html(res["out_bytes"].decode("utf-8", errors="replace")[:200000], height=300, scrolling=True)
                except Exception:
                    st.write("*(HTML preview failed or is too large. Download instead.)*")
            else:
                st.write(f"*(No preview for {res['mime']})*")
                
            st.download_button(
                label="Download",
                data=res["out_bytes"],
                file_name=res["out_name"],
                mime=res["mime"],
                key=f"download_btn_{i}"
            )

    # Offer a single ZIP download for all successfully converted files
    st.markdown("---")
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for res in results_for_zip:
            zf.writestr(res["out_name"], res["out_bytes"])
    zip_buffer.seek(0)
    
    st.download_button(
        label="⬇️ Download ALL Converted Files as ZIP",
        data=zip_buffer.read(),
        file_name="converted_structured_documents.zip",
        mime="application/zip"
    )

elif not errors_occurred:
    st.info("No files were successfully converted.")

st.markdown("---")
st.caption("Developed with PyMuPDF, pdfplumber, BeautifulSoup, and python-docx. Version 1.1")
