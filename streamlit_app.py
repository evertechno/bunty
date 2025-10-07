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
import pandas as pd # pandas is imported but not explicitly used in the final version, could be removed if not needed for future features

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
    
    # Sort distinct sizes from largest to smallest
    sorted_sizes = sorted(list(set(unique_sizes)), reverse=True)
    
    mapping = {}
    # Assign heading levels to the largest unique font sizes
    for idx, size in enumerate(sorted_sizes[:max_levels]):
        mapping[round(size, 2)] = idx + 1 # Levels 1 to max_levels
    
    return mapping

# -----------------------------
# Advanced PDF Content Extraction (Text + Images)
# -----------------------------

def extract_page_content_advanced(page: fitz.Page, min_para_size: float = 8.0, 
                                 line_gap_threshold: float = 0.5, char_gap_threshold: float = 0.1) -> Tuple[List[Dict], List[Dict], List[float]]:
    """
    Extracts all text spans with their properties, groups them into logical lines and elements,
    and extracts images.
    Returns: (text_elements, image_elements, all_span_sizes)
    """
    text_elements = []
    image_elements = []
    all_span_sizes = []

    # 1. Extract and process all text spans
    raw_text_dict = page.get_text("rawdict")
    spans = []
    for block in raw_text_dict.get("blocks", []):
        if block.get("type") == 0:  # Text block
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    if span.get("text", "").strip(): # Only consider non-empty text
                        spans.append({
                            "text": span["text"],
                            "size": round(span["size"], 2),
                            "font": span["font"],
                            "bbox": fitz.Rect(span["bbox"]) # Convert to fitz.Rect object
                        })
                        all_span_sizes.append(span["size"])

    # Sort spans by vertical position, then horizontal for reading order
    spans.sort(key=lambda s: (s["bbox"].y0, s["bbox"].x0))

    if not spans:
        return [], [], all_span_sizes # No text on page

    current_line_spans = []
    current_line_y = -1

    # Group spans into lines first
    lines = []
    for span in spans:
        if not current_line_spans or (span["bbox"].y0 - current_line_y > line_gap_threshold * span["bbox"].height):
            # Start a new line
            if current_line_spans:
                lines.append(current_line_spans)
            current_line_spans = [span]
            current_line_y = span["bbox"].y0
        else:
            # Add to current line, update Y (average or min)
            current_line_spans.append(span)
            current_line_y = min(current_line_y, span["bbox"].y0) # Adjust line Y to lowest span in line

    if current_line_spans:
        lines.append(current_line_spans)

    # Now group lines into elements (paragraphs, headings, lists)
    current_element_lines = []
    current_element_type = "para" # Default
    current_element_size = 0.0 # Max size in current element
    
    for line_spans in lines:
        line_text = "".join([s["text"] for s in line_spans]).strip()
        if not line_text:
            continue

        line_max_size = max(s["size"] for s in line_spans)
        
        # Check for list items
        is_list = is_bullet_line(line_text) or is_numbered_line(line_text)
        
        # Heuristic for new element: significant font size change, empty line, or list type change
        new_element_needed = False
        if not current_element_lines: # First line of content
            new_element_needed = True
        elif abs(line_max_size - current_element_size) > min_para_size * 0.15: # Significant font change
            new_element_needed = True
        elif is_list and current_element_type != "list": # Transition to list
            new_element_needed = True
        elif not is_list and current_element_type == "list": # Transition from list
            new_element_needed = True
        elif not line_text.strip(): # Explicit blank line
             new_element_needed = True

        if new_element_needed and current_element_lines:
            # Flush previous element
            combined_text = "\n".join(sanitize_text("".join([s["text"] for s in ls])) for ls in current_element_lines if ls)
            if combined_text:
                if current_element_type == "list":
                    # This logic should group list items properly when flush
                    text_elements.append({"type": "list", "items": [sanitize_text("".join([s["text"] for s in ls])) for ls in current_element_lines], "list_type": "bullet"}) # Simplified to bullet for now
                else: # para or heading
                    text_elements.append({"type": current_element_type, "text": combined_text, "size": current_element_size})
            current_element_lines = []
            current_element_size = 0.0

        # Add current line to new element
        current_element_lines.append(line_spans)
        current_element_size = max(current_element_size, line_max_size)
        current_element_type = "list" if is_list else "para" # Re-evaluate type for current line

    # Flush the last element after loop
    if current_element_lines:
        combined_text = "\n".join(sanitize_text("".join([s["text"] for s in ls])) for ls in current_element_lines if ls)
        if combined_text:
            if current_element_type == "list":
                text_elements.append({"type": "list", "items": [sanitize_text("".join([s["text"] for s in ls])) for ls in current_element_lines], "list_type": "bullet"})
            else:
                text_elements.append({"type": current_element_type, "text": combined_text, "size": current_element_size})


    # 2. Extract images
    img_list = page.get_images(full=True)
    for img_index, img in enumerate(img_list):
        xref = img[0]
        base_image = doc.extract_image(xref)
        image_bytes = base_image["image"]
        image_ext = base_image["ext"]
        image_mime = f"image/{image_ext}"
        
        # Get image bbox (approximated, as PyMuPDF's get_images doesn't provide it directly in this form easily)
        # We need to look for image objects on the page with their actual bounding boxes
        # This is a bit more involved: iterate through page objects
        found_bbox = None
        for item in page.get_displaylist()._image_info: # Accessing internal structure, might break
            if item.get("xref") == xref:
                found_bbox = item.get("bbox")
                break
        
        if not found_bbox: # Fallback: if bbox not found, estimate or ignore
            found_bbox = (0, 0, page.rect.width, page.rect.height) # Placeholder

        # Encode image to base64 for embedding in HTML
        b64_image = base64.b64encode(image_bytes).decode('utf-8')
        image_elements.append({
            "type": "image",
            "data": f"data:{image_mime};base64,{b64_image}",
            "bbox": fitz.Rect(found_bbox) # Store bbox for potential sorting
        })
    
    # Sort images by vertical position for better flow
    image_elements.sort(key=lambda img: img["bbox"].y0)

    return text_elements, image_elements, all_span_sizes


def parse_pdf_structured(pdf_bytes: bytes, min_heading_ratio: float = 1.12, min_para_size: float = 8.0, 
                         filter_repeated_text: bool = True) -> Dict[str, Any]:
    """
    Parse PDF into a structured intermediate representation, incorporating images and advanced text parsing.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_out = []
    all_doc_sizes = []
    
    # Collect potential repeated text for filtering
    repeated_text_candidates = Counter()
    per_page_text_lines = [] # Store lines with their approximate bbox for filtering

    # First pass to collect all text and potential repeated elements
    for p_idx in range(len(doc)):
        page = doc.load_page(p_idx)
        page_lines_with_bbox = []
        for block in page.get_text("dict").get("blocks", []):
            if block.get("type") == 0:
                for line in block.get("lines", []):
                    line_text = "".join([span["text"] for span in line.get("spans", [])]).strip()
                    if line_text:
                        # Use a rounded bbox center for matching repeated elements
                        bbox_center = (fitz.Rect(line["bbox"]).x0 // 20, fitz.Rect(line["bbox"]).y0 // 20)
                        page_lines_with_bbox.append((line_text, bbox_center))
                        repeated_text_candidates[(line_text, bbox_center)] += 1
        per_page_text_lines.append(page_lines_with_bbox)

    # Determine truly repeated elements (appear on > 50% of pages in same position)
    num_pages = len(doc)
    threshold = num_pages * 0.5
    filtered_out_elements = set()
    if filter_repeated_text:
        for (text, bbox_center), count in repeated_text_candidates.items():
            if count > threshold:
                filtered_out_elements.add((text, bbox_center))


    # Second pass: Process each page for final elements
    for p_idx in range(len(doc)):
        page = doc.load_page(p_idx)
        page_elements = []

        # Extract images and text elements with font sizes
        raw_text_elements, image_elements, current_page_sizes = \
            extract_page_content_advanced(page, min_para_size=min_para_size)
        all_doc_sizes.extend(current_page_sizes)

        # Apply filtering for repeated text
        filtered_text_elements = []
        for el in raw_text_elements:
            if el["type"] in ["para", "heading", "list"]:
                # Flatten list items for checking, use first item's bbox for list entry
                check_text = el["text"] if el["type"] != "list" else el["items"][0]
                
                # Approximate bbox center for comparison (assumes element has bbox, which rawdict helps with)
                # This is a bit of a hack without direct bbox for combined elements, so re-extract a span-level bbox
                first_span_bbox = None
                if el["type"] == "para" or el["type"] == "heading":
                    # For para/heading, get bbox of the first line of text for comparison
                    raw_text_dict = page.get_text("rawdict")
                    for block in raw_text_dict.get("blocks", []):
                        if block.get("type") == 0:
                            for line in block.get("lines", []):
                                line_content = "".join([s["text"] for s in line["spans"]]).strip()
                                if line_content == el["text"].split('\n')[0]: # Match first line
                                    first_span_bbox = fitz.Rect(line["bbox"])
                                    break
                            if first_span_bbox: break
                elif el["type"] == "list" and el["items"]:
                     raw_text_dict = page.get_text("rawdict")
                     for block in raw_text_dict.get("blocks", []):
                        if block.get("type") == 0:
                            for line in block.get("lines", []):
                                line_content = "".join([s["text"] for s in line["spans"]]).strip()
                                if line_content == el["items"][0]: # Match first item
                                    first_span_bbox = fitz.Rect(line["bbox"])
                                    break
                            if first_span_bbox: break
                
                if first_span_bbox:
                    bbox_center = (first_span_bbox.x0 // 20, first_span_bbox.y0 // 20)
                    if (sanitize_text(check_text.split('\n')[0]), bbox_center) in filtered_out_elements:
                         # Skip this element if it's a repeated header/footer
                        continue
            filtered_text_elements.append(el)

        # Combine text and images by vertical position (approximated)
        all_content_elements = []
        for el in filtered_text_elements:
            # Need an approximate y-position for text elements
            if el["type"] in ["para", "heading", "list"] and el.get("text") or el.get("items"):
                # Use max_block_span_sz from original parse_pdf_structured for rough estimate
                # For `extract_page_content_advanced`, we would need to get bbox for the whole combined element
                # For now, let's just use first line/item's bbox y0
                y0_approx = None
                if el["type"] in ["para", "heading"]:
                    first_line = el["text"].split('\n')[0]
                    for text_line_data, bbox_center in per_page_text_lines[p_idx]:
                        if sanitize_text(text_line_data) == sanitize_text(first_line):
                            # This needs a more direct way to get y0 from `extract_page_content_advanced`
                            # For now, we'll sort the output after table insertion.
                            pass
                elif el["type"] == "list" and el["items"]:
                    first_item = el["items"][0]
                    for text_line_data, bbox_center in per_page_text_lines[p_idx]:
                        if sanitize_text(text_line_data) == sanitize_text(first_item):
                            pass
                
                all_content_elements.append({
                    "type": el["type"],
                    "content": el,
                    "y0": el["text_bbox"].y0 if "text_bbox" in el else (0 if not text_elements else text_elements[0]["bbox"].y0) # Placeholder
                })
        
        # Simpler approach: Just append images and tables after text elements for now, then sort.
        for el in filtered_text_elements:
            page_elements.append(el)
        for img_el in image_elements:
            page_elements.append(img_el)

        # Table detection with pdfplumber
        tables_on_page = []
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as ppdf:
                page_pl = ppdf.pages[p_idx]
                extracted_tables = page_pl.extract_tables()
                for t in extracted_tables:
                    if t and any(any(cell for cell in row if cell not in (None, "")) for row in t):
                        cleaned_table = [[str(c) if c is not None else "" for c in row] for row in t]
                        # Store bbox for tables for sorting
                        table_bbox = page_pl.find_tables()[extracted_tables.index(t)].bbox
                        tables_on_page.append({"type": "table", "rows": cleaned_table, "bbox": fitz.Rect(table_bbox)})
        except Exception as e:
            print(f"Warning: Error extracting tables from PDF page {p_idx+1} with pdfplumber: {e}")
            tables_on_page = []

        # Add tables to page elements and sort everything by approximate vertical position
        for tbl_el in tables_on_page:
            page_elements.append(tbl_el)
        
        # Sort all elements (text, images, tables) by their top-most coordinate
        # This requires `y0` or similar to be present on all elements
        def get_y0_for_sorting(element):
            if element["type"] == "image":
                return element["bbox"].y0
            elif element["type"] == "table":
                return element["bbox"].y0
            elif "bbox" in element: # Text elements with explicit bbox
                return element["bbox"].y0
            # For `raw_text_elements`, the lines already sorted, so taking first line's y0 is complex here.
            # We need to ensure `extract_page_content_advanced` returns elements with `bbox` or a `y0` property.
            # For simplicity for now, this part needs careful definition of `element` structure from `extract_page_content_advanced`.
            # Let's assume for combined text elements, we can get a representative y0.
            # Revisit this: `extract_page_content_advanced` should yield "elements" not just "spans"
            # and these "elements" should have a `bbox` for sorting.

            # Temporary fallback for text elements (this is imperfect)
            if element["type"] in ["para", "heading", "list"]:
                # Try to get y0 from one of its component spans if available
                if "content" in element and isinstance(element["content"], dict):
                    if element["content"].get("items") and isinstance(element["content"]["items"], list):
                        # For list items, first item's line top
                        line_text_raw = element["content"]["items"][0]
                        line_text_clean = sanitize_text(line_text_raw)
                        # This needs to be improved to get precise y0 for combined elements
                        # For now, a rough estimate is better than unsorted.
                        # We will make extract_page_content_advanced return combined elements with bbox
                        pass
                return 0 # Put unsortable elements at top
            return 0 # Default if bbox not found


        # The `extract_page_content_advanced` returns individual span properties, not combined element bboxes directly.
        # This makes sorting text elements with images/tables problematic.
        # Let's simplify: `extract_page_content_advanced` will produce elements (heading/para/list/image),
        # each with a `bbox` property.
        
        # Re-parse: to make sorting work, we need all elements (text, image, table) to have a .y0
        # The `raw_text_elements` returned from `extract_page_content_advanced` need a `bbox` property.
        
        final_elements_on_page = []
        
        # Re-iterate `raw_text_elements` to assign representative bbox
        for el in raw_text_elements:
            if el["type"] in ["para", "heading"]:
                # To assign bbox for a combined element: get the bbox of all its constituent lines
                # This would require modifying `extract_page_content_advanced` to return combined_bbox
                # For now, let's use the bbox of the first line/span within the element as a proxy for sorting
                first_line_text = el["text"].split('\n')[0]
                temp_bbox = None
                for block in page.get_text("dict").get("blocks", []): # This is inefficient, but necessary to get bbox
                    if block.get("type") == 0:
                        for line in block.get("lines", []):
                            line_content = "".join([span["text"] for span in line.get("spans", [])]).strip()
                            if sanitize_text(line_content) == sanitize_text(first_line_text):
                                temp_bbox = fitz.Rect(line["bbox"])
                                break
                        if temp_bbox: break
                el["bbox"] = temp_bbox if temp_bbox else fitz.Rect(0,0,page.rect.width, 10) # Fallback bbox
                final_elements_on_page.append(el)

            elif el["type"] == "list" and el["items"]:
                first_item_text = el["items"][0]
                temp_bbox = None
                for block in page.get_text("dict").get("blocks", []):
                    if block.get("type") == 0:
                        for line in block.get("lines", []):
                            line_content = "".join([span["text"] for span in line.get("spans", [])]).strip()
                            if sanitize_text(line_content) == sanitize_text(first_item_text):
                                temp_bbox = fitz.Rect(line["bbox"])
                                break
                        if temp_bbox: break
                el["bbox"] = temp_bbox if temp_bbox else fitz.Rect(0,0,page.rect.width, 10) # Fallback bbox
                final_elements_on_page.append(el)
            else: # Images (already have bbox)
                final_elements_on_page.append(el)
        
        for tbl_el in tables_on_page:
            final_elements_on_page.append(tbl_el) # Tables already have bbox

        # Final sort of all elements on the page
        final_elements_on_page.sort(key=lambda x: x["bbox"].y0)

        pages_out.append({"page_number": p_idx + 1, "elements": final_elements_on_page})

    # Recalculate global font hierarchy from all non-filtered text
    unique_doc_sizes = sorted(set(s for s in all_doc_sizes if s >= min_para_size), reverse=True)
    font_to_heading = choose_heading_levels(unique_doc_sizes)

    # Apply heading levels to parsed text elements in pages_out
    for page_data in pages_out:
        for el in page_data["elements"]:
            if el["type"] == "para" and "size" in el: # Only for "para" initially
                mapped_level = font_to_heading.get(round(el["size"], 2), 0)
                if mapped_level > 0 or (el["size"] >= (max(unique_doc_sizes) / min_heading_ratio) and el["size"] > (Counter(round(s, 2) for s in all_doc_sizes if s >= min_para_size).most_common(1)[0][0] if all_doc_sizes else 0)):
                    el["type"] = "heading"
                    el["level"] = mapped_level if mapped_level > 0 else 2 # Default to H2 if heuristic
    
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
        'img { max-width: 100%; height: auto; display: block; margin: 1em auto; }', # Added img style
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
                    # Clean bullets/numbers as HTML handles them visually
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
                parts.append(f'<img src="{el["data"]}" alt="Image from PDF" style="width:{el["bbox"].width}px; height:{el["bbox"].height}px;"/>') # Embed images with their original dimensions


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
            norm_text = normalize_whitespace_for_output(el["text"]) if "text" in el else ""
            if el["type"] == "heading":
                out_lines.append(norm_text.upper())
                out_lines.append("=" * len(norm_text)) 
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
                for r_idx, r in enumerate(rows):
                    cells_formatted = [normalize_whitespace_for_output(str(c)) if c is not None else "" for c in r]
                    out_lines.append("\t".join(cells_formatted))
                    if r_idx == 0 and len(rows) > 1:
                        out_lines.append("\t".join(["-" * len(cell) if cell else "---" for cell in cells_formatted]))
                out_lines.append("") 
            elif el["type"] == "image":
                out_lines.append(f"[IMAGE: {el['data'][:50]}... (base64 data)]")
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
                # docx needs raw bytes, not base64
                try:
                    # Extract image bytes from base64 data string
                    img_data_b64 = el['data'].split(';base64,')[1]
                    img_bytes = base64.b64decode(img_data_b64)
                    
                    # Add image to docx, convert pixel dimensions to EMUs (used by docx)
                    # 1 inch = 914400 EMUs, 1px = 914400/96 EMUs (assuming 96 DPI)
                    # For a more precise conversion, you'd need the PDF's DPI
                    # Here, using 96 DPI as a common web standard
                    width_emus = int(el["bbox"].width * 914400 / 96)
                    height_emus = int(el["bbox"].height * 914400 / 96)

                    # Ensure image bytes are valid for python-docx
                    img_stream = io.BytesIO(img_bytes)
                    doc.add_picture(img_stream, width=Inches(el["bbox"].width / 96) if el["bbox"].width > 0 else None, 
                                    height=Inches(el["bbox"].height / 96) if el["bbox"].height > 0 else None)
                    doc.add_paragraph("") # Spacing after image
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
                        # Extract base64 data
                        mime_type, img_data_b64 = img_src.split(';base64,')
                        img_bytes = base64.b64decode(img_data_b64)
                        img_stream = io.BytesIO(img_bytes)

                        # Attempt to get width/height from HTML attributes
                        width = None
                        height = None
                        if element.has_attr('width'):
                            try: width = Inches(float(element['width']) / 96) # Assume 96 DPI
                            except ValueError: pass
                        if element.has_attr('height'):
                            try: height = Inches(float(element['height']) / 96)
                            except ValueError: pass

                        doc_obj.add_picture(img_stream, width=width, height=height)
                        doc_obj.add_paragraph("") # Spacing
                    except Exception as img_e:
                        doc_obj.add_paragraph(f"[Failed to embed image from HTML: {img_e}]")
                # Handle external image links (not directly embedding, but marking)
                elif img_src:
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
    filter_repeated_text = st.checkbox(
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
    # Read once
    raw_bytes = uploaded_file_obj.read()
    file_ext = os.path.splitext(file_name)[1].lower()
    
    result_entry = {"name": file_name}
    
    try:
        # Determine input type and process
        if file_ext == ".pdf":
            # Parse PDF structure once
            parsed_content = parse_pdf_structured(raw_bytes, 
                                                min_heading_ratio=heading_ratio, 
                                                min_para_size=min_para_size,
                                                filter_repeated_text=filter_repeated_text)
            
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
                errors_occurred.append(result) # Append the dict for full context
                log_area.error(f"❌ **{result['name']}**: {result['error']}")
            else:
                results_for_zip.append(result)
                out_size_kb = len(result['out_bytes']) / 1024
                log_area.success(f"✅ **{result['name']}** → {result['out_name']} ({out_size_kb:.1f} KB)")
        except Exception as exc: # FIX: Changed to catch specific Exception and assign it
            errors_occurred.append({"name": file_name_processed, "error": f"Unhandled exception: {exc}"})
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
        # Single item, use a dummy container to make loop generic
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

st.markdown("---")
st.caption("Developed with PyMuPDF, pdfplumber, BeautifulSoup, and python-docx. Version 1.1")
