"""
streamlit_converter_structured.py

Streamlit Multi-format Converter — Structured, "as-is" output.
Heuristics used:
 - PyMuPDF (fitz) to read ALL text spans with bounding boxes -> reconstruct lines/paragraphs/headings
 - PyMuPDF to extract images
 - pdfplumber to detect tabular data -> produce <table> and docx tables
 - BeautifulSoup for HTML->Text/Word conversions
 - python-docx to create Word (.docx) with headings, paragraphs, lists, tables

Notes:
 - This targets digital PDFs (embedded text and images). No OCR is performed.
 - Heuristics (font-size thresholds, list detection, repeated text filtering) can be tuned in the UI.
 - **Goal:** To convert the *entire content* of the PDF (text, images, tables) into structured HTML/DOCX/TXT,
   prioritizing *semantic content and reasonable visual flow* over pixel-perfect layout replication (which is extremely hard).
   For absolute visual fidelity on complex pages, converting the page to an image might be a better but less semantic approach.
"""
import io
import os
import zipfile
import base64
import re
import html
from typing import List, Tuple, Dict, Any, Optional
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
from bs4 import BeautifulSoup, Tag
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
# pandas is imported but not explicitly used in the final version, could be removed if not needed for future features
# import pandas as pd 

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
    Reduces multiple spaces and newlines for consistent output.
    """
    s = re.sub(r'[ \t]+', ' ', s) # Multiple spaces to single space
    s = re.sub(r'\n{3,}', '\n\n', s) # Reduce excessive newlines to at most two
    return s.strip()

def choose_heading_levels(unique_sizes: List[float], max_levels: int = 4) -> Dict[float, int]:
    """
    Map font sizes -> heading levels (1..max_levels) heuristically.
    Biggest -> h1, next -> h2, etc. If many sizes, map top ones to headings.
    Smaller sizes not in top `max_levels` get mapped to 0 (paragraph).
    """
    if not unique_sizes:
        return {}
    
    sorted_sizes = sorted(list(set(unique_sizes)), reverse=True)
    mapping = {}
    for idx, size in enumerate(sorted_sizes[:max_levels]):
        mapping[round(size, 2)] = idx + 1 # Levels 1 to max_levels
    
    return mapping

# -----------------------------
# Advanced PDF Content Extraction (Text + Images + Tables)
# -----------------------------

def extract_page_elements_detailed(
    doc: fitz.Document, 
    page_idx: int, 
    min_para_size: float = 8.0, 
    line_gap_threshold: float = 0.5, # Factor of font_height to consider a new line
    para_gap_threshold: float = 1.5, # Factor of font_height to consider a new paragraph
    min_font_change_for_heading: float = 0.15 # % change for new heading heuristic
) -> Tuple[List[Dict], List[float]]:
    """
    Extracts all text spans, images, and tables from a page, and then reconstructs
    logical elements (headings, paragraphs, lists, images, tables) with their bboxes.
    Returns: (list_of_combined_elements, all_span_sizes_on_page)
    """
    page = doc.load_page(page_idx)
    page_width = page.rect.width
    page_height = page.rect.height
    combined_elements = []
    all_span_sizes_on_page = []

    # 1. Extract raw text spans
    raw_spans = []
    text_blocks_raw = page.get_text("rawdict").get("blocks", [])
    for block in text_blocks_raw:
        if block.get("type") == 0:  # Text block
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    if span.get("text", "").strip():
                        s_bbox = fitz.Rect(span["bbox"])
                        # Filter out very small, likely noise spans
                        if s_bbox.height < 2 or s_bbox.width < 2: continue

                        raw_spans.append({
                            "text": span["text"],
                            "size": round(span["size"], 2),
                            "font": span["font"],
                            "flags": span["flags"],
                            "color": span["color"],
                            "bbox": s_bbox # Store as fitz.Rect
                        })
                        all_span_sizes_on_page.append(span["size"])

    # Sort raw spans primarily by vertical position, then horizontal for reading order
    raw_spans.sort(key=lambda s: (s["bbox"].y0, s["bbox"].x0))

    # 2. Extract images
    img_list = page.get_images(full=True)
    for img_index, img_info in enumerate(img_list):
        xref = img_info[0]
        base_image = doc.extract_image(xref)
        image_bytes = base_image["image"]
        image_ext = base_image["ext"]
        image_mime = f"image/{image_ext}"
        
        # Get image bbox directly from display list if possible for accuracy
        image_bbox = None
        for item in page.get_displaylist()._image_info:
            if item.get("xref") == xref:
                image_bbox = fitz.Rect(item.get("bbox"))
                break
        
        if not image_bbox: # Fallback if bbox not found
            # A more robust fallback would be to try to match img_info with page.get_pixmap()._image_info
            # For simplicity, use a generic size if bbox cannot be precisely determined.
            image_bbox = fitz.Rect(0, 0, page_width / 4, page_height / 4) 

        b64_image = base64.b64encode(image_bytes).decode('utf-8')
        combined_elements.append({
            "type": "image",
            "data": f"data:{image_mime};base64,{b64_image}",
            "bbox": image_bbox,
            "y0_sort": image_bbox.y0 # For sorting
        })
    
    # 3. Extract tables
    tables_on_page = []
    try:
        with pdfplumber.open(io.BytesIO(doc.tobytes())) as ppdf:
            plumber_page = ppdf.pages[page_idx]
            extracted_tables = plumber_page.extract_tables()
            for t_idx, t_rows in enumerate(extracted_tables):
                if t_rows and any(any(cell for cell in row if cell not in (None, "")) for row in t_rows):
                    cleaned_table = [[str(c) if c is not None else "" for c in row] for row in t_rows]
                    table_plumber_bbox = plumber_page.find_tables()[t_idx].bbox
                    table_bbox = fitz.Rect(table_plumber_bbox)
                    combined_elements.append({
                        "type": "table",
                        "rows": cleaned_table,
                        "bbox": table_bbox,
                        "y0_sort": table_bbox.y0 # For sorting
                    })
    except Exception as e:
        print(f"Warning: Error extracting tables from PDF page {page_idx+1} with pdfplumber: {e}")

    # 4. Reconstruct logical text elements from sorted spans
    text_elements_from_spans = []
    if raw_spans:
        current_paragraph_spans = []
        current_paragraph_bbox = None
        
        for i, span in enumerate(raw_spans):
            is_new_line = False
            is_new_paragraph = False
            
            # Heuristic for new line (if not first span)
            if current_paragraph_spans:
                last_span = current_paragraph_spans[-1]
                # Check if vertical gap is significant (new line)
                if span["bbox"].y0 > last_span["bbox"].y1 + line_gap_threshold * span["bbox"].height:
                    is_new_line = True
                    # Check if gap is even larger (new paragraph)
                    if span["bbox"].y0 > last_span["bbox"].y1 + para_gap_threshold * span["bbox"].height:
                        is_new_paragraph = True
                # Check horizontal gap if on same "line" (could indicate column break or explicit new line)
                elif span["bbox"].x0 > last_span["bbox"].x1 + char_gap_threshold * span["bbox"].width * 5: # Large horizontal gap
                     is_new_line = True # Consider it a new line
                     
                # Also consider significant font size change as a potential new paragraph/heading
                if abs(span["size"] - last_span["size"]) > min_font_change_for_heading * span["size"] and \
                   span["size"] >= min_para_size: # Only if above min readable font size
                    is_new_paragraph = True

            if is_new_paragraph or not current_paragraph_spans:
                if current_paragraph_spans:
                    # Flush previous paragraph
                    text_elements_from_spans.append({"type": "temp_text", "spans": current_paragraph_spans})
                current_paragraph_spans = [span]
            elif is_new_line:
                # Add a newline marker if it's a new line within the same logical paragraph
                current_paragraph_spans.append({"text": "\n", "size": span["size"], "bbox": span["bbox"]}) # Placeholder for newline
                current_paragraph_spans.append(span)
            else:
                current_paragraph_spans.append(span)

        if current_paragraph_spans:
            text_elements_from_spans.append({"type": "temp_text", "spans": current_paragraph_spans})

    # Now, convert temp_text elements into proper paragraphs/headings/lists
    final_text_elements = []
    for temp_el in text_elements_from_spans:
        full_text = "".join([s["text"] for s in temp_el["spans"] if s["text"] != "\n"]).strip()
        if not full_text: continue

        # Calculate combined bbox for the element
        min_x, min_y, max_x, max_y = page_width, page_height, 0, 0
        for span in temp_el["spans"]:
            if span["text"] != "\n": # Don't use newline markers for bbox calc
                min_x = min(min_x, span["bbox"].x0)
                min_y = min(min_y, span["bbox"].y0)
                max_x = max(max_x, span["bbox"].x1)
                max_y = max(max_y, span["bbox"].y1)
        element_bbox = fitz.Rect(min_x, min_y, max_x, max_y)

        # Determine type (list, para, potential heading)
        first_line = full_text.split('\n')[0]
        if is_bullet_line(first_line) or is_numbered_line(first_line):
            list_type = "numbered" if is_numbered_line(first_line) else "bullet"
            # Split into individual list items for better processing
            items = [sanitize_text(line) for line in full_text.split('\n') if sanitize_text(line)]
            final_text_elements.append({
                "type": "list",
                "items": items,
                "list_type": list_type,
                "bbox": element_bbox,
                "y0_sort": element_bbox.y0,
                "size": max(s["size"] for s in temp_el["spans"] if s["text"] != "\n") # Representative size
            })
        else:
            final_text_elements.append({
                "type": "para", # Default to paragraph, will be refined to heading later
                "text": normalize_whitespace_for_output(full_text),
                "bbox": element_bbox,
                "y0_sort": element_bbox.y0,
                "size": max(s["size"] for s in temp_el["spans"] if s["text"] != "\n") # Representative size
            })
    
    combined_elements.extend(final_text_elements)
    
    # Sort all elements (text, images, tables) by their y0_sort property
    combined_elements.sort(key=lambda x: x["y0_sort"])

    return combined_elements, all_span_sizes_on_page


def parse_pdf_structured(pdf_bytes: bytes, min_heading_ratio: float = 1.12, 
                         min_para_size: float = 7.0, 
                         filter_repeated_text: bool = True,
                         line_gap_threshold: float = 0.5,
                         para_gap_threshold: float = 1.5,
                         min_font_change_for_heading: float = 0.15
                         ) -> Dict[str, Any]:
    """
    Parse PDF into a structured intermediate representation, incorporating images and advanced text parsing.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_out = []
    all_doc_sizes = []
    
    # --- First pass: Collect font sizes and identify repeated text elements ---
    repeated_text_candidates = Counter()
    temp_all_page_elements_raw = [] # To store elements before filtering/final typing

    for p_idx in range(len(doc)):
        # Use a preliminary extraction to get lines and their approximate bboxes
        # This is a lighter pass than `extract_page_elements_detailed`
        page = doc.load_page(p_idx)
        raw_text_dict = page.get_text("dict")
        page_lines_for_filter = []

        for block in raw_text_dict.get("blocks", []):
            if block.get("type") == 0:
                for line in block.get("lines", []):
                    line_text = "".join([span["text"] for span in line.get("spans", [])]).strip()
                    if line_text:
                        # Normalize text and use a rounded bbox center for matching repeated elements
                        bbox_center = (fitz.Rect(line["bbox"]).x0 // 20, fitz.Rect(line["bbox"]).y0 // 20)
                        normalized_line_text = normalize_whitespace_for_output(line_text)
                        page_lines_for_filter.append((normalized_line_text, bbox_center, fitz.Rect(line["bbox"])))
                        repeated_text_candidates[(normalized_line_text, bbox_center)] += 1
        
        temp_all_page_elements_raw.append(page_lines_for_filter)

    # Determine truly repeated elements (appear on > 50% of pages in same position)
    num_pages = len(doc)
    threshold = num_pages * 0.5
    if num_pages < 2: # Don't filter if only one page or very few pages
        filter_repeated_text = False

    filtered_out_elements_set = set()
    if filter_repeated_text:
        for (text, bbox_center), count in repeated_text_candidates.items():
            if count >= threshold: # Use >= for exact match
                filtered_out_elements_set.add((text, bbox_center))

    # --- Second pass: Detailed extraction and filtering ---
    for p_idx in range(len(doc)):
        page_elements, current_page_sizes = extract_page_elements_detailed(
            doc, p_idx, min_para_size, line_gap_threshold, para_gap_threshold, min_font_change_for_heading
        )
        all_doc_sizes.extend(current_page_sizes)

        final_page_elements = []
        for el in page_elements:
            if el["type"] in ["para", "heading", "list"]:
                # For filtering, use the text of the element and its approximated bbox center
                text_for_filter = el["text"].split('\n')[0] if el["type"] != "list" else el["items"][0]
                text_for_filter = normalize_whitespace_for_output(text_for_filter)
                
                # Use bbox of the element directly
                bbox_center = (el["bbox"].x0 // 20, el["bbox"].y0 // 20)
                
                if filter_repeated_text and (text_for_filter, bbox_center) in filtered_out_elements_set:
                    continue # Skip this element as it's a repeated header/footer

            final_page_elements.append(el)
        
        pages_out.append({"page_number": p_idx + 1, "elements": final_page_elements})

    # --- Final pass: Apply global heading styles based on document-wide font analysis ---
    unique_doc_sizes = sorted(set(s for s in all_doc_sizes if s >= min_para_size), reverse=True)
    font_to_heading = choose_heading_levels(unique_doc_sizes)

    # Get body text size (most common size) for better heading inference
    body_text_size = 0.0
    if all_doc_sizes:
        size_counts = Counter(round(s, 2) for s in all_doc_sizes if s >= min_para_size)
        if size_counts:
            body_text_size = size_counts.most_common(1)[0][0]
            # If body_text_size was incorrectly assigned a heading level, demote it
            if body_text_size in font_to_heading and font_to_heading[body_text_size] == max(font_to_heading.values()):
                del font_to_heading[body_text_size]
    
    max_doc_size = max(unique_doc_sizes) if unique_doc_sizes else 12.0
    heading_threshold = max_doc_size / min_heading_ratio # Heuristic threshold

    for page_data in pages_out:
        for el in page_data["elements"]:
            if el["type"] == "para" and "size" in el: # Only refine paragraphs
                mapped_level = font_to_heading.get(round(el["size"], 2), 0)
                # If mapped or significantly larger than body text and threshold
                if mapped_level > 0 or (el["size"] >= heading_threshold and el["size"] > body_text_size):
                    el["type"] = "heading"
                    el["level"] = mapped_level if mapped_level > 0 else 2 # Default to H2 if just threshold

    return {"pages": pages_out, "fontsizes": unique_doc_sizes}


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
        'img { max-width: 100%; height: auto; display: block; margin: 1em auto; }', 
        '</style>',
        '</head>',
        '<body>'
    ]
    
    for page in parsed["pages"]:
        parts.append(f'<div class="page" data-page="{page["page_number"]}">')
        parts.append(f'<p class="page-number">--- Page {page["page_number"]} ---</p>')
        
        current_list_type = None 
        
        for el in page["elements"]:
            if el["type"] == "heading":
                if current_list_type:
                    parts.append(f'</{"ul" if current_list_type == "bullet" else "ol"}>')
                    current_list_type = None
                
                lvl = min(max(int(el.get("level", 2)), 1), 6) 
                text = html.escape(normalize_whitespace_for_output(el["text"]))
                parts.append(f"<h{lvl}>{text}</h{lvl}>")
            
            elif el["type"] == "para":
                if current_list_type:
                    parts.append(f'</{"ul" if current_list_type == "bullet" else "ol"}>')
                    current_list_type = None
                
                text = html.escape(normalize_whitespace_for_output(el["text"]))
                parts.append(f"<p>{text}</p>")
            
            elif el["type"] == "list":
                list_type_html = "ul" if el["list_type"] == "bullet" else "ol"
                
                if current_list_type != el["list_type"]:
                    if current_list_type: 
                        parts.append(f'</{"ul" if current_list_type == "bullet" else "ol"}>')
                    parts.append(f'<{list_type_html}>')
                    current_list_type = el["list_type"]
                
                for item_text in el["items"]:
                    clean_item_text = re.sub(BULLET_CHARS if el["list_type"]=="bullet" else NUMBER_CHARS, "", item_text).strip()
                    parts.append(f"<li>{html.escape(normalize_whitespace_for_output(clean_item_text if clean_item_text else item_text))}</li>")
            
            elif el["type"] == "table":
                if current_list_type:
                    parts.append(f'</{"ul" if current_list_type == "bullet" else "ol"}>')
                    current_list_type = None
                
                rows = el["rows"]
                parts.append("<table>")
                for r_idx, r in enumerate(rows):
                    parts.append("<tr>")
                    tag = "th" if r_idx == 0 and len(rows) > 1 else "td"
                    parts.append("".join(f"<{tag}>{html.escape(str(c) if c is not None else '')}</{tag}>" for c in r))
                    parts.append("</tr>")
                parts.append("</table>")
            
            elif el["type"] == "image":
                if current_list_type:
                    parts.append(f'</{"ul" if current_list_type == "bullet" else "ol"}>')
                    current_list_type = None
                
                # HTML Image width/height. Use max-width:100% for responsiveness,
                # but set original dimensions as attributes for reference.
                width_attr = f'width="{el["bbox"].width}"' if el["bbox"].width > 0 else ''
                height_attr = f'height="{el["bbox"].height}"' if el["bbox"].height > 0 else ''

                parts.append(f'<img src="{el["data"]}" alt="Image from PDF" {width_attr} {height_attr}/>')


        if current_list_type:
            parts.append(f'</{"ul" if current_list_type == "bullet" else "ol"}>')
            current_list_type = None
            
        parts.append("</div>") 

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
            if el["type"] == "heading":
                norm_text = normalize_whitespace_for_output(el["text"])
                out_lines.append(norm_text.upper())
                out_lines.append("=" * min(len(norm_text), 80)) # Limit underline length
                out_lines.append("")
            elif el["type"] == "para":
                out_lines.append(normalize_whitespace_for_output(el["text"]))
                out_lines.append("")
            elif el["type"] == "list":
                for item_text in el["items"]:
                    # Keep original bullets/numbers for plain text if they exist in source
                    out_lines.append(normalize_whitespace_for_output(item_text))
                out_lines.append("")
            elif el["type"] == "table":
                rows = el["rows"]
                for r_idx, r in enumerate(rows):
                    cells_formatted = [normalize_whitespace_for_output(str(c)) if c is not None else "" for c in r]
                    out_lines.append("\t".join(cells_formatted))
                    if r_idx == 0 and len(rows) > 1:
                        out_lines.append("\t".join(["-" * len(cell) if cell else "---" for cell in cells_formatted]))
                out_lines.append("") 
            elif el["type"] == "image":
                out_lines.append(f"[IMAGE: Embedded, size {el['bbox'].width:.0f}x{el['bbox'].height:.0f}px]")
                out_lines.append("")
    joined = "\n".join(out_lines).strip()
    return joined.encode("utf-8")


def structured_to_docx(parsed: dict) -> bytes:
    doc = Document()
    
    styles = doc.styles
    if 'Normal' not in styles:
        styles.add_style('Normal', WD_STYLE_TYPE.PARAGRAPH)
    styles['Normal'].font.name = 'Arial'
    styles['Normal'].font.size = Pt(11)
    
    for i in range(1, 5): 
        if f'Heading {i}' not in styles:
            styles.add_style(f'Heading {i}', WD_STYLE_TYPE.PARAGRAPH)
        styles[f'Heading {i}'].font.name = 'Arial'
        styles[f'Heading {i}'].font.size = Pt(11 + (5-i)*2) 
        styles[f'Heading {i}'].font.bold = True
        styles[f'Heading {i}'].paragraph_format.space_before = Pt(12)
        styles[f'Heading {i}'].paragraph_format.space_after = Pt(6)

    if 'List Bullet' not in styles:
        styles.add_style('List Bullet', WD_STYLE_TYPE.PARAGRAPH)
    styles['List Bullet'].font.name = 'Arial'
    styles['List Bullet'].font.size = Pt(11)
    
    if 'List Number' not in styles:
        styles.add_style('List Number', WD_STYLE_TYPE.PARAGRAPH)
    styles['List Number'].font.name = 'Arial'
    styles['List Number'].font.size = Pt(11)

    for page in parsed["pages"]:
        p_page_num = doc.add_paragraph(f"--- Page {page['page_number']} ---")
        p_page_num.style = 'Normal'
        p_page_num.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("") 
        
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
                    clean_item_text = re.sub(BULLET_CHARS if el["list_type"]=="bullet" else NUMBER_CHARS, "", item_text).strip()
                    p = doc.add_paragraph(normalize_whitespace_for_output(clean_item_text if clean_item_text else item_text))
                    p.style = list_style
            elif el["type"] == "table":
                rows = el["rows"]
                if not rows:
                    continue

                ncols = max(len(r) for r in rows)
                tbl = doc.add_table(rows=0, cols=ncols)
                tbl.style = 'Table Grid'

                for r_idx, r in enumerate(rows):
                    row_cells = tbl.add_row().cells
                    for i in range(ncols):
                        cell_text = normalize_whitespace_for_output(str(r[i])) if i < len(r) and r[i] is not None else ""
                        row_cells[i].text = cell_text
                        if r_idx == 0 and len(rows) > 1: 
                            for run in row_cells[i].paragraphs[0].runs:
                                run.font.bold = True
                
                doc.add_paragraph("") 
            elif el["type"] == "image":
                try:
                    img_data_b64 = el['data'].split(';base64,')[1]
                    img_bytes = base64.b64decode(img_data_b64)
                    
                    img_stream = io.BytesIO(img_bytes)
                    
                    # Convert pixel width/height to Inches for docx
                    # Assuming 96 DPI for screen pixels to print inches
                    width_inches = el["bbox"].width / 96.0
                    height_inches = el["bbox"].height / 96.0

                    doc.add_picture(img_stream, 
                                    width=Inches(width_inches) if width_inches > 0 else None, 
                                    height=Inches(height_inches) if height_inches > 0 else None)
                    doc.add_paragraph("") 
                except Exception as img_e:
                    doc.add_paragraph(f"[Failed to embed image: {img_e}]")
                    print(f"Error embedding image in DOCX: {img_e}")

        doc.add_page_break()
        
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# -----------------------------
# HTML -> Text / DOCX
# -----------------------------

def html_to_text_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    text = soup.get_text(separator="\n\n")
    return normalize_whitespace_for_output(text).encode("utf-8")

def html_to_docx_bytes(html_bytes: bytes) -> bytes:
    soup = BeautifulSoup(html_bytes, "html.parser")
    doc = Document()
    
    styles = doc.styles
    if 'Normal' not in styles:
        styles.add_style('Normal', WD_STYLE_TYPE.PARAGRAPH)
    styles['Normal'].font.name = 'Arial'
    styles['Normal'].font.size = Pt(11)

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
            # Images
            elif element.name == "img" and element.has_attr('src'):
                img_src = element['src']
                if img_src.startswith('data:image'):
                    try:
                        mime_type, img_data_b64 = img_src.split(';base64,')
                        img_bytes = base64.b64decode(img_data_b64)
                        img_stream = io.BytesIO(img_bytes)

                        # Attempt to get width/height from HTML attributes
                        width = None
                        height = None
                        if element.has_attr('width'):
                            try: width = Inches(float(element['width']) / 96) 
                            except ValueError: pass
                        if element.has_attr('height'):
                            try: height = Inches(float(element['height']) / 96)
                            except ValueError: pass

                        doc_obj.add_picture(img_stream, width=width, height=height)
                        doc_obj.add_paragraph("") 
                    except Exception as img_e:
                        doc_obj.add_paragraph(f"[Failed to embed image from HTML: {img_e}]")
                elif img_src: # External image links
                    doc_obj.add_paragraph(f"[Image: {img_src}]")
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

**Goal:** "Not even a comma" difference in the *textual content* and its logical organization, including all embedded images and tables.

**Supported Conversions & Features:**
*   **PDF → HTML/DOCX/TXT:** Uses advanced heuristics to reconstruct text structure (headings, paragraphs, lists), extracts all **images**, and detects **tables**.
*   **HTML → DOCX/TXT:** Preserves existing HTML structure, including images and tables.
*   **Table Detection:** Extracts tables into proper HTML/DOCX table formats.
*   **List & Heading Recovery:** Rebuilds lists and heading hierarchies.
*   **Repeated Content Filtering:** (Optional) Can filter out headers/footers that repeat across many pages.

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
    filter_repeated_text_sidebar = st.checkbox( # Renamed to avoid conflict with function param
        "Filter Repeated Headers/Footers", 
        value=True, 
        help="Attempts to remove text that appears in the same position on most pages (e.g., 'Thank You', page numbers). Useful for presentations."
    )


    st.markdown("---")
    st.markdown("### ⚙️ System")
    workers = st.number_input("Parallel Workers", min_value=1, max_value=8, value=4, help="Process multiple files at once.")
    embed_pdf = st.checkbox("Embed Source PDF in HTML", value=True, help="Adds original PDF to HTML output for side-by-side reference. Increases HTML file size.")

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
    raw_bytes = uploaded_file_obj.read()
    file_ext = os.path.splitext(file_name)[1].lower()
    
    result_entry = {"name": file_name}
    
    try:
        if file_ext == ".pdf":
            parsed_content = parse_pdf_structured(
                raw_bytes, 
                min_heading_ratio=heading_ratio, 
                min_para_size=min_para_size,
                filter_repeated_text=filter_repeated_text_sidebar, # Use sidebar value
                line_gap_threshold=0.5, # Pass tuning parameters
                para_gap_threshold=1.5,
                min_font_change_for_heading=0.15
            )
            
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
            
        output_name = os.path.splitext(file_name)[0] + "_converted" + ext
        result_entry.update({"out_bytes": output_bytes, "out_name": output_name, "mime": mime})
        return result_entry
        
    except Exception as e:
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
            errors_occurred.append({"name": file_name_processed, "error": f"Unhandled exception during conversion: {exc}"})
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
                st.download_button(
                    label=f"⬇️ Download {res['out_name']}",
                    data=res["out_bytes"],
                    file_name=res["out_name"],
                    mime=res["mime"],
                    key=f"dl_{res['out_name']}",
                    use_container_width=True,
                    type="primary"
                )

            if res["mime"].startswith("text/"):
                try:
                    preview_text = res["out_bytes"].decode("utf-8", errors="replace")
                    if res["mime"] == "text/html":
                        with st.expander("👁️ Preview HTML (Rendered)", expanded=False):
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
    st.warning("No results to display.")

st.markdown("---")
st.caption("Developed with PyMuPDF, pdfplumber, BeautifulSoup, and python-docx. Version 1.1")
