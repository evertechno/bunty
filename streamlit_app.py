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
import re
import html
from typing import List, Tuple, Dict, Any, Optional
# --- Added missing imports for concurrency ---
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
from bs4 import BeautifulSoup, Tag
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
import pandas as pd

# -----------------------------
# Utility / Heuristic Functions
# -----------------------------

BULLET_CHARS = r"^[\u2022\u2023\u25E6\-\*\•\–\—]\s*" 
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
            # Log warning to console/logs, don't interrupt user flow
            print(f"Warning: Error extracting tables from PDF page {p_idx+1}: {e}")
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
                    # Strip existing bullets/numbers from text if present, as HTML handles them
                    clean_item_text = re.sub(BULLET_CHARS if el["list_type"]=="bullet" else NUMBER_CHARS, "", item_text).strip()
                    parts.append(f"<li>{html.escape(normalize_whitespace_for_output(clean_item_text if clean_item_text else item_text))}</li>")
            
            elif el["type"] == "table":
                # Close any open list before a table
                if current_list_type:
                    parts.append(f'</{"ul" if current_list_type == "bullet" else "ol"}>')
                    current_list_type = None
                
                rows = el["rows"]
                parts.append("<table>")
                for r_idx, r in enumerate(rows):
                    parts.append("<tr>")
                    # Assume first row is header if table has more than one row
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
        embed_snip = f'<hr/><h2>Original PDF (embedded)</h2><p><i>Note: The embedded PDF is for reference.</i></p><embed src="data:application/pdf;base64,{b64}" width="100%" height="800px" type="application/pdf"></embed>'
        parts.append(embed_snip)
        
    parts.append("</body></html>")
    return "\n".join(parts).encode("utf-8")


def structured_to_text(parsed: dict) -> bytes:
    out_lines = []
    for page in parsed["pages"]:
        out_lines.append(f"\n--- PAGE {page['page_number']} ---\n")
        
        for el in page["elements"]:
            norm_text = normalize_whitespace_for_output(el["text"]) if "text" in el else ""
            if el["type"] == "heading":
                out_lines.append(norm_text.upper())
                out_lines.append("=" * len(norm_text)) # Stronger underline for headings
                out_lines.append("")
            elif el["type"] == "para":
                out_lines.append(norm_text)
                out_lines.append("")
            elif el["type"] == "list":
                for item_text in el["items"]:
                    # Keep original bullets/numbers for plain text if they exist in source
                    out_lines.append(normalize_whitespace_for_output(item_text))
                out_lines.append("")
            elif el["type"] == "table":
                rows = el["rows"]
                # A simple text-based table approximation, tab-separated
                for r_idx, r in enumerate(rows):
                    cells_formatted = [normalize_whitespace_for_output(str(c)) if c is not None else "" for c in r]
                    out_lines.append("\t".join(cells_formatted))
                    if r_idx == 0 and len(rows) > 1: # Header separator
                        out_lines.append("\t".join(["-" * len(cell) if cell else "---" for cell in cells_formatted]))
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
        # Heuristic sizing relative to Normal
        styles[f'Heading {i}'].font.size = Pt(11 + (5-i)*2) 
        styles[f'Heading {i}'].font.bold = True
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
        doc.add_paragraph("") # Spacing
        
        for el in page["elements"]:
            if el["type"] == "heading":
                lvl = min(max(int(el.get("level", 2)), 1), 4)
                doc.add_heading(normalize_whitespace_for_output(el["text"]), level=lvl)
            elif el["type"] == "para":
                p = doc.add_paragraph(normalize_whitespace_for_output(el["text"]))
                p.style = 'Normal'
            elif el["type"] == "list":
                list_style = 'List Bullet' if el["list_type"] == "bullet" else 'List Number'
                for item_text in el["items"]:
                    # Strip existing bullets/numbers as docx style handles them
                    clean_item_text = re.sub(BULLET_CHARS if el["list_type"]=="bullet" else NUMBER_CHARS, "", item_text).strip()
                    p = doc.add_paragraph(normalize_whitespace_for_output(clean_item_text if clean_item_text else item_text))
                    p.style = list_style
            elif el["type"] == "table":
                rows = el["rows"]
                if not rows:
                    continue

                ncols = max(len(r) for r in rows)
                # Create table
                tbl = doc.add_table(rows=0, cols=ncols)
                tbl.style = 'Table Grid'

                for r_idx, r in enumerate(rows):
                    row_cells = tbl.add_row().cells
                    for i in range(ncols):
                        cell_text = normalize_whitespace_for_output(str(r[i])) if i < len(r) and r[i] is not None else ""
                        row_cells[i].text = cell_text
                        if r_idx == 0 and len(rows) > 1: # Apply bold to header row
                            for run in row_cells[i].paragraphs[0].runs:
                                run.font.bold = True
                
                doc.add_paragraph("") # Spacing after table
                
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
    # Use newlines to distinguish block elements
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

    # Basic recursive parser to handle structure
    def parse_element(element, doc_obj):
        if isinstance(element, Tag):
            # Headings
            if element.name and element.name.startswith("h") and len(element.name) == 2 and element.name[1].isdigit():
                level = int(element.name[1])
                text = normalize_whitespace_for_output(element.get_text(strip=True))
                if text:
                    doc_obj.add_heading(text, level=min(level, 4))
            # Paragraphs
            elif element.name == "p":
                text = normalize_whitespace_for_output(element.get_text(" ", strip=True))
                if text:
                    doc_obj.add_paragraph(text).style = 'Normal'
            # Lists
            elif element.name in ("ul", "ol"):
                list_style = 'List Bullet' if element.name == "ul" else 'List Number'
                for li in element.find_all("li", recursive=False):
                    text = normalize_whitespace_for_output(li.get_text(" ", strip=True))
                    if text:
                        doc_obj.add_paragraph(text).style = list_style
            # Tables
            elif element.name == "table":
                rows_data = []
                for r in element.find_all("tr"):
                    cols = [normalize_whitespace_for_output(c.get_text(" ", strip=True)) for c in r.find_all(["th", "td"])]
                    rows_data.append(cols)
                
                if rows_data:
                    ncols = max(len(r) for r in rows_data)
                    tbl = doc_obj.add_table(rows=0, cols=ncols)
                    tbl.style = 'Table Grid'
                    has_header = element.find("th") is not None
                    for r_idx, r in enumerate(rows_data):
                        cells = tbl.add_row().cells
                        for i in range(ncols):
                            cells[i].text = r[i] if i < len(r) else ""
                            if r_idx == 0 and has_header and cells[i].paragraphs: 
                                for run in cells[i].paragraphs[0].runs:
                                    run.font.bold = True
                    doc_obj.add_paragraph("")
            # Horizontal Rule
            elif element.name == "hr":
                doc_obj.add_paragraph("_________________________________").alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Recurse for container tags like div, section, body if not handled above
            elif element.name in ("body", "div", "section", "article", "main"):
                for child in element.children:
                    parse_element(child, doc_obj)

    # Start parsing from body
    if soup.body:
        for child in soup.body.children:
            parse_element(child, doc)
    else:
        # Fallback if no body tag
        for child in soup.children:
             parse_element(child, doc)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# -----------------------------
# Streamlit App UI
# -----------------------------

st.set_page_config(page_title="Legacy Converter — Structured Cloning", layout="wide", page_icon="📚")
st.title("📚 Legacy Converter — Preserve structure, create a legacy")

st.markdown("""
This app aims to **clone the *structured content*** of your documents with high fidelity, prioritizing semantic structure (headings, lists, tables, paragraphs) over exact visual layout.

**Goal:** "Not even a comma" difference in the *textual content* and its logical organization.

**Supported Conversions & Features:**
*   **PDF → HTML/DOCX/TXT:** Uses heuristics to reconstruct structure from digital PDFs.
*   **HTML → DOCX/TXT:** Preserves existing HTML structure.
*   **Table Detection:** Extracts tables into proper HTML/DOCX table formats.
*   **List & Heading Recovery:** Rebuilds lists and heading hierarchies.

*Note: Only supports digital PDFs (no OCR). Large files may take time.*
""")

with st.sidebar:
    st.header("⚙️ Conversion Options")
    conversion = st.selectbox("Select Conversion Type", [
        "PDF → Structured HTML",
        "PDF → Word (.docx)",
        "PDF → Plain Text",
        "HTML → Word (.docx)",
        "HTML → Plain Text"
    ])
    
    st.markdown("---")
    st.markdown("### 🔧 Tuning (PDF Inputs)")
    st.info("Adjust these if structure isn't detected correctly.")
    heading_ratio = st.slider(
        "Heading Detection Sensitivity", 
        min_value=1.01, max_value=1.5, value=1.10, step=0.01, 
        help="Lower value = Smaller font differences count as headings. Increase if paragraphs are wrongly detected as headings."
    )
    min_para_size = st.number_input(
        "Minimum Text Size (pt)", 
        min_value=4.0, max_value=14.0, value=7.0, step=0.5,
        help="Ignore text smaller than this (e.g., page numbers, tiny footnotes) to clean up structure."
    )

    st.markdown("---")
    st.markdown("### ⚙️ System")
    workers = st.number_input("Parallel Workers", min_value=1, max_value=8, value=4, help="Process multiple files at once.")
    embed_pdf = st.checkbox("Embed Source PDF in HTML", value=True, help="Adds original PDF to HTML output for side-by-side reference.")

uploaded_files = st.file_uploader("Drop PDF or HTML files here", type=["pdf", "html"], accept_multiple_files=True)

if not uploaded_files:
    st.info("👋 Upload files to start. This tool does not store your data.")
    st.stop()

start_conversion = st.button(f"🚀 Convert {len(uploaded_files)} File(s)")

if not start_conversion:
    st.stop()

# --- Conversion Logic ---
results_for_zip = []
errors_occurred = []

def process_single_file(uploaded_file_obj):
    file_name = uploaded_file_obj.name
    # Read once
    raw_bytes = uploaded_file_obj.read()
    file_ext = os.path.splitext(file_name)[1].lower()
    
    result_entry = {"name": file_name}
    
    try:
        # Determine input type and process
        if file_ext == ".pdf":
            # Parse PDF structure once
            parsed_content = parse_pdf_structured(raw_bytes, min_heading_ratio=heading_ratio, min_para_size=min_para_size)
            
            # Route to appropriate converter
            if conversion == "PDF → Structured HTML":
                output_bytes = structured_to_html(parsed_content, embed_pdf=embed_pdf, pdf_bytes=raw_bytes if embed_pdf else None)
                ext, mime = ".html", "text/html"
            elif conversion == "PDF → Word (.docx)":
                output_bytes = structured_to_docx(parsed_content)
                ext, mime = ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif conversion == "PDF → Plain Text":
                output_bytes = structured_to_text(parsed_content)
                ext, mime = ".txt", "text/plain"
            else:
                raise ValueError("Invalid conversion path for PDF.")
                
        elif file_ext == ".html":
            # Route HTML inputs
            if conversion == "HTML → Plain Text":
                output_bytes = html_to_text_bytes(raw_bytes)
                ext, mime = ".txt", "text/plain"
            elif conversion == "HTML → Word (.docx)":
                output_bytes = html_to_docx_bytes(raw_bytes)
                ext, mime = ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            else:
                raise ValueError("Invalid conversion path for HTML.")
        else:
            raise ValueError(f"Unsupported file type: {file_ext}")
            
        # Success: Prepare result package
        output_name = os.path.splitext(file_name)[0] + "_converted" + ext
        result_entry.update({"out_bytes": output_bytes, "out_name": output_name, "mime": mime})
        return result_entry
        
    except Exception as e:
        # Failure: Capture error
        result_entry["error"] = str(e)
        return result_entry

# Use columns for cleaner layout during processing
stat_col, log_col = st.columns([1, 2])
with stat_col:
    st.write("Processing Status")
    progress_bar = st.progress(0)
    status_message = st.empty()

with log_col:
    st.write("Activity Log")
    log_area = st.container(height=200)

# Concurrent Processing
with ThreadPoolExecutor(max_workers=workers) as executor:
    future_to_file = {executor.submit(process_single_file, f): f.name for f in uploaded_files}
    
    completed_count = 0
    for future in as_completed(future_to_file):
        completed_count += 1
        file_name_processed = future_to_file[future]
        progress_bar.progress(completed_count / len(uploaded_files))
        
        try:
            result = future.result()
            if "error" in result:
                errors_occurred.append(result)
                log_area.error(f"❌ **{result['name']}**: {result['error']}")
            else:
                results_for_zip.append(result)
                out_size_kb = len(result['out_bytes']) / 1024
                log_area.success(f"✅ **{result['name']}** → {result['out_name']} ({out_size_kb:.1f} KB)")
        except Exception as exc:
            errors_occurred.append({"name": file_name_processed, "error": str(exc)})
            log_area.error(f"❌ **{file_name_processed}**: Critical error - {exc}")

# Final Status Update
if errors_occurred:
    status_message.warning(f"Completed with {len(errors_occurred)} errors.")
else:
    status_message.success("All files converted successfully!")

st.markdown("---")

# Results & Download Section
if results_for_zip:
    st.header("📥 Download Results")
    
    # Tabbed interface for previews if multiple files, otherwise direct view
    if len(results_for_zip) > 1:
        tabs = st.tabs([res['name'] for res in results_for_zip])
        iterable = zip(tabs, results_for_zip)
    else:
        iterable = [(st.container(), results_for_zip[0])]

    for container, res in iterable:
        with container:
            col1, col2 = st.columns([3, 1])
            with col1:
                st.subheader(res['out_name'])
            with col2:
                # Prominent download button for each file
                st.download_button(
                    label=f"⬇️ Download {res['out_name']}",
                    data=res["out_bytes"],
                    file_name=res["out_name"],
                    mime=res["mime"],
                    key=f"dl_{res['out_name']}",
                    use_container_width=True,
                    type="primary"
                )

            # Smart Preview
            if res["mime"].startswith("text/"):
                try:
                    preview_text = res["out_bytes"].decode("utf-8", errors="replace")
                    if res["mime"] == "text/html":
                        with st.expander("👁️ Preview HTML (Rendered)", expanded=False):
                            # Sandboxed HTML preview
                            st.components.v1.html(preview_text[:500000], height=400, scrolling=True)
                        with st.expander("📄 View HTML Source Code", expanded=False):
                             st.code(preview_text[:10000], language="html")
                    else: # Plain text
                        with st.expander("👁️ Preview Text", expanded=True):
                            st.text_area("Content", preview_text[:10000], height=300, label_visibility="collapsed")
                except Exception as e:
                     st.warning(f"Could not generate preview: {e}")
            elif "wordprocessingml" in res["mime"]:
                st.info("📝 DOCX file created. Please download to view.")
            
            st.divider()

    # Bulk ZIP Download
    if len(results_for_zip) > 1:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for res in results_for_zip:
                zf.writestr(res["out_name"], res["out_bytes"])
        zip_buffer.seek(0)
        
        st.download_button(
            label=f"📦 Download All {len(results_for_zip)} Files (ZIP)",
            data=zip_buffer.read(),
            file_name="converted_legacy_docs.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

elif not errors_occurred:
    # Should only happen if list was empty initially but passed check
    st.warning("No results to display.")
