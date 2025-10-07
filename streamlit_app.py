""" streamlit_converter_near_original.py
Streamlit Multi-format Converter — Near-Original Output
Direct PDF rendering + Structured extraction + Custom code insertion
"""

import io
import os
import zipfile
import base64
import json
from typing import List, Dict, Tuple, Any, Optional
import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd
import html
import re
from concurrent.futures import ThreadPoolExecutor, as_completed

# -----------------------------
# Enhanced PDF Text Extraction
# -----------------------------

def extract_text_near_original(pdf_bytes: bytes) -> Dict[str, Any]:
    """Extract text with near-original formatting and structure"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_data = []
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        
        # Extract text with layout preservation
        text_blocks = page.get_text("dict")
        
        # Get tables using pdfplumber
        tables_data = extract_tables_from_page(pdf_bytes, page_num)
        
        # Extract images (for potential embedding)
        images = []
        image_list = page.get_images()
        for img_index, img in enumerate(image_list):
            try:
                xref = img[0]
                pix = fitz.Pixmap(doc, xref)
                if pix.n - pix.alpha < 4:  # RGB or CMYK
                    img_data = pix.tobytes("png")
                    images.append({
                        "index": img_index,
                        "data": base64.b64encode(img_data).decode(),
                        "bbox": img[1:5] if len(img) > 4 else (0, 0, 100, 100)
                    })
                pix = None
            except:
                pass
        
        page_data = {
            "number": page_num + 1,
            "width": page.rect.width,
            "height": page.rect.height,
            "text_blocks": text_blocks.get("blocks", []),
            "tables": tables_data,
            "images": images
        }
        
        pages_data.append(page_data)
    
    doc.close()
    
    return {
        "pages": pages_data,
        "metadata": doc.metadata,
        "total_pages": len(pages_data)
    }

def extract_tables_from_page(pdf_bytes: bytes, page_num: int) -> List[Dict]:
    """Extract tables with proper structure"""
    tables = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            page = pdf.pages[page_num]
            extracted_tables = page.extract_tables()
            
            for i, table in enumerate(extracted_tables):
                if table and any(any(cell for cell in row) for row in table):
                    # Get table bounding box
                    table_obj = page.find_tables()
                    bbox = table_obj[i].bbox if i < len(table_obj) else None
                    
                    # Clean table data
                    cleaned_table = []
                    for row in table:
                        cleaned_row = [clean_cell_text(cell) for cell in row]
                        if any(cleaned_row):  # Only add non-empty rows
                            cleaned_table.append(cleaned_row)
                    
                    if cleaned_table:
                        tables.append({
                            "rows": cleaned_table,
                            "bbox": bbox,
                            "page": page_num + 1
                        })
    except Exception as e:
        st.warning(f"Table extraction issue on page {page_num + 1}: {e}")
    
    return tables

def clean_cell_text(text: Any) -> str:
    """Clean table cell text"""
    if text is None:
        return ""
    text = str(text).strip()
    # Remove excessive whitespace but preserve meaningful spaces
    text = re.sub(r'\s+', ' ', text)
    return text

def is_heading(text: str, font_size: float, position: float, page_height: float) -> bool:
    """Determine if text is a heading"""
    # Criteria for heading detection
    if font_size >= 14:
        return True
    if font_size >= 12 and position < page_height * 0.2:  # Top of page
        return True
    if text.isupper() and len(text) < 100:  # Short uppercase text
        return True
    if re.match(r'^(chapter|section|part)\s+[0-9IVX]+', text, re.IGNORECASE):
        return True
    return False

def get_heading_level(font_size: float) -> int:
    """Determine heading level based on font size"""
    if font_size >= 18:
        return 1
    elif font_size >= 16:
        return 2
    elif font_size >= 14:
        return 3
    elif font_size >= 12:
        return 4
    else:
        return 0

# -----------------------------
# Direct PDF Rendering in iframe
# -----------------------------

def create_pdf_iframe_embed(pdf_bytes: bytes, width: str = "100%", height: str = "600px") -> str:
    """Create iframe embed code for PDF"""
    pdf_b64 = base64.b64encode(pdf_bytes).decode('ascii')
    return f'''
    <iframe src="data:application/pdf;base64,{pdf_b64}" 
            width="{width}" 
            height="{height}" 
            style="border: none; margin: 20px 0;">
    </iframe>
    '''

def create_pdf_viewer_html(pdf_bytes: bytes, show_grid: bool = False) -> str:
    """Create HTML with PDF viewer"""
    pdf_b64 = base64.b64encode(pdf_bytes).decode('ascii')
    
    grid_style = ""
    if show_grid:
        grid_style = """
        .grid-overlay {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-image: 
                linear-gradient(rgba(0,0,0,0.1) 1px, transparent 1px),
                linear-gradient(90deg, rgba(0,0,0,0.1) 1px, transparent 1px);
            background-size: 20px 20px;
            pointer-events: none;
            z-index: 10;
        }
        """
    
    return f'''
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body {{ margin: 0; padding: 20px; background: #f5f5f5; }}
            .pdf-container {{ 
                position: relative; 
                max-width: 100%; 
                margin: 0 auto;
                box-shadow: 0 4px 8px rgba(0,0,0,0.1);
                background: white;
            }}
            {grid_style}
        </style>
    </head>
    <body>
        <div class="pdf-container">
            <embed src="data:application/pdf;base64,{pdf_b64}" 
                   width="100%" 
                   height="800px" 
                   type="application/pdf">
            {'<div class="grid-overlay"></div>' if show_grid else ''}
        </div>
    </body>
    </html>
    '''

# -----------------------------
# Enhanced HTML Conversion with Custom Code Insertion
# -----------------------------

def convert_to_html_with_custom_code(parsed_data: Dict, pdf_bytes: bytes, 
                                   custom_css: str = "", custom_js: str = "", 
                                   embed_pdf: bool = False, show_grid: bool = False) -> bytes:
    """Convert to HTML with custom code insertion options"""
    
    base_css = """
    <style>
    /* Near-original PDF styling */
    body {
        font-family: "Times New Roman", Georgia, serif;
        font-size: 12pt;
        line-height: 1.6;
        margin: 0;
        padding: 40px;
        background: white;
        color: #000000;
        max-width: 8.5in;
        margin: 0 auto;
    }
    
    .page-section {
        margin-bottom: 40px;
        page-break-inside: avoid;
    }
    
    .pdf-viewer-section {
        margin: 40px 0;
        border: 2px solid #e0e0e0;
        border-radius: 8px;
        overflow: hidden;
    }
    
    h1, h2, h3, h4 {
        font-family: Arial, Helvetica, sans-serif;
        font-weight: bold;
        margin: 24px 0 12px 0;
        color: #2c3e50;
    }
    
    h1 { font-size: 24pt; border-bottom: 2px solid #3498db; padding-bottom: 8px; }
    h2 { font-size: 20pt; }
    h3 { font-size: 16pt; }
    h4 { font-size: 14pt; }
    
    p {
        margin: 12px 0;
        text-align: justify;
        font-size: 12pt;
        line-height: 1.6;
    }
    
    .text-block {
        margin: 8px 0;
        padding: 4px 0;
    }
    
    ul, ol {
        margin: 16px 0 16px 30px;
    }
    
    li {
        margin: 6px 0;
        font-size: 12pt;
    }
    
    table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
        font-size: 11pt;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    
    th {
        background: #34495e;
        color: white;
        font-weight: bold;
        padding: 12px 15px;
        text-align: left;
        border: 1px solid #2c3e50;
    }
    
    td {
        padding: 10px 15px;
        border: 1px solid #bdc3c7;
        background: white;
    }
    
    tr:nth-child(even) td {
        background: #f8f9fa;
    }
    
    .content-section {
        background: white;
        padding: 30px;
        margin: 20px 0;
        border-radius: 8px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    
    .metadata {
        background: #ecf0f1;
        padding: 15px;
        border-radius: 5px;
        margin: 20px 0;
        font-size: 10pt;
        color: #7f8c8d;
    }
    </style>
    """
    
    html_parts = [
        '<!DOCTYPE html>',
        '<html lang="en">',
        '<head>',
        '<meta charset="UTF-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
        '<title>Near-Original PDF Conversion</title>',
        base_css,
        f'<style>{custom_css}</style>',
        '</head>',
        '<body>',
        '<div class="content-section">',
        '<h1>📄 Document Conversion</h1>',
        '<div class="metadata">',
        f'<strong>Total Pages:</strong> {parsed_data["total_pages"]}<br>',
        f'<strong>Creator:</strong> {parsed_data.get("metadata", {}).get("creator", "N/A")}<br>',
        f'<strong>Producer:</strong> {parsed_data.get("metadata", {}).get("producer", "N/A")}',
        '</div>'
    ]
    
    # Add PDF viewer if requested
    if embed_pdf:
        html_parts.extend([
            '<div class="pdf-viewer-section">',
            '<h2>📊 Original PDF Viewer</h2>',
            create_pdf_iframe_embed(pdf_bytes),
            '</div>'
        ])
    
    # Process each page
    for page in parsed_data["pages"]:
        html_parts.append(f'<div class="page-section" data-page="{page["number"]}">')
        html_parts.append(f'<h3>📖 Page {page["number"]}</h3>')
        
        # Process text blocks
        text_content = []
        for block in page.get("text_blocks", []):
            if block.get("type") == 0:  # Text block
                block_text = extract_text_from_block(block)
                if block_text.strip():
                    # Determine if this is a heading
                    font_size = get_block_font_size(block)
                    if is_heading(block_text, font_size, block.get("bbox", [0,0,0,0])[1], page["height"]):
                        level = get_heading_level(font_size)
                        html_parts.append(f'<h{level}>{html.escape(block_text)}</h{level}>')
                    else:
                        html_parts.append(f'<p class="text-block">{html.escape(block_text)}</p>')
        
        # Process tables
        for table in page.get("tables", []):
            html_parts.append(convert_table_to_html(table))
        
        html_parts.append('</div>')  # Close page-section
    
    html_parts.extend([
        '</div>',  # Close content-section
        f'<script>{custom_js}</script>',
        '</body>',
        '</html>'
    ])
    
    return "\n".join(html_parts).encode('utf-8')

def extract_text_from_block(block: Dict) -> str:
    """Extract text from a block with proper spacing"""
    text_lines = []
    for line in block.get("lines", []):
        line_text = ""
        for span in line.get("spans", []):
            span_text = span.get("text", "").strip()
            if span_text:
                line_text += span_text + " "
        if line_text.strip():
            text_lines.append(line_text.strip())
    return "\n".join(text_lines)

def get_block_font_size(block: Dict) -> float:
    """Get the predominant font size from a block"""
    sizes = []
    for line in block.get("lines", []):
        for span in line.get("spans", []):
            sizes.append(span.get("size", 12))
    return max(sizes) if sizes else 12

def convert_table_to_html(table: Dict) -> str:
    """Convert table data to HTML table"""
    rows = table.get("rows", [])
    if not rows:
        return ""
    
    html_parts = ['<table>']
    
    # Check if first row looks like headers
    first_row = rows[0]
    is_header_row = any(cell and any(c.isupper() for c in str(cell)) for cell in first_row)
    
    if is_header_row:
        html_parts.append('<thead><tr>')
        for cell in first_row:
            html_parts.append(f'<th>{html.escape(str(cell))}</th>')
        html_parts.append('</tr></thead><tbody>')
        rows = rows[1:]
    else:
        html_parts.append('<tbody>')
    
    for row in rows:
        html_parts.append('<tr>')
        for cell in row:
            html_parts.append(f'<td>{html.escape(str(cell))}</td>')
        html_parts.append('</tr>')
    
    html_parts.append('</tbody></table>')
    return "".join(html_parts)

# -----------------------------
# Enhanced DOCX Conversion
# -----------------------------

def convert_to_docx_enhanced(parsed_data: Dict) -> bytes:
    """Convert to DOCX with near-original structure"""
    doc = Document()
    
    # Set document properties
    doc.core_properties.title = "PDF Conversion"
    doc.core_properties.author = "Streamlit Converter"
    
    # Add title
    title = doc.add_heading('PDF Document Conversion', 0)
    
    # Add metadata
    meta_para = doc.add_paragraph()
    meta_para.add_run(f"Pages: {parsed_data['total_pages']} | ")
    meta_para.add_run(f"Creator: {parsed_data.get('metadata', {}).get('creator', 'N/A')}")
    
    doc.add_paragraph()  # Spacing
    
    # Process pages
    for page in parsed_data["pages"]:
        if page["number"] > 1:
            doc.add_page_break()
        
        # Add page header
        page_header = doc.add_heading(f'Page {page["number"]}', level=2)
        
        # Process text blocks
        for block in page.get("text_blocks", []):
            if block.get("type") == 0:
                block_text = extract_text_from_block(block)
                if block_text.strip():
                    font_size = get_block_font_size(block)
                    if is_heading(block_text, font_size, block.get("bbox", [0,0,0,0])[1], page["height"]):
                        level = get_heading_level(font_size)
                        doc.add_heading(block_text, level=min(level, 4))
                    else:
                        para = doc.add_paragraph(block_text)
        
        # Process tables
        for table in page.get("tables", []):
            rows = table.get("rows", [])
            if rows:
                num_cols = max(len(row) for row in rows)
                doc_table = doc.add_table(rows=len(rows), cols=num_cols)
                doc_table.style = 'Table Grid'
                
                for i, row in enumerate(rows):
                    for j, cell in enumerate(row):
                        if j < num_cols:
                            doc_table.cell(i, j).text = str(cell) if cell else ""
                
                doc.add_paragraph()  # Spacing after table
    
    # Save to buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

# -----------------------------
# Enhanced Text Conversion
# -----------------------------

def convert_to_text_enhanced(parsed_data: Dict) -> bytes:
    """Convert to text with structure preservation"""
    text_parts = []
    
    text_parts.append("PDF DOCUMENT CONVERSION")
    text_parts.append("=" * 50)
    text_parts.append(f"Total Pages: {parsed_data['total_pages']}")
    text_parts.append(f"Creator: {parsed_data.get('metadata', {}).get('creator', 'N/A')}")
    text_parts.append("")
    
    for page in parsed_data["pages"]:
        text_parts.append(f"PAGE {page['number']}")
        text_parts.append("-" * 30)
        
        # Text blocks
        for block in page.get("text_blocks", []):
            if block.get("type") == 0:
                block_text = extract_text_from_block(block)
                if block_text.strip():
                    text_parts.append(block_text)
                    text_parts.append("")
        
        # Tables
        for table in page.get("tables", []):
            rows = table.get("rows", [])
            if rows:
                text_parts.append("TABLE:")
                # Find column widths for alignment
                col_widths = [0] * max(len(row) for row in rows)
                for row in rows:
                    for j, cell in enumerate(row):
                        if j < len(col_widths):
                            col_widths[j] = max(col_widths[j], len(str(cell)) if cell else 0)
                
                for row in rows:
                    row_text = "| "
                    for j, cell in enumerate(row):
                        if j < len(col_widths):
                            cell_text = str(cell) if cell else ""
                            row_text += cell_text.ljust(col_widths[j]) + " | "
                    text_parts.append(row_text)
                
                text_parts.append("")
    
    return "\n".join(text_parts).encode('utf-8')

# -----------------------------
# HTML to Other Formats
# -----------------------------

def html_to_docx_enhanced(html_bytes: bytes) -> bytes:
    """Convert HTML to enhanced DOCX"""
    soup = BeautifulSoup(html_bytes, 'html.parser')
    doc = Document()
    
    # Remove unwanted elements
    for element in soup(["script", "style"]):
        element.decompose()
    
    # Process content
    body = soup.find('body') or soup
    
    def process_element(element):
        if element.name == 'h1':
            doc.add_heading(element.get_text(strip=True), level=1)
        elif element.name == 'h2':
            doc.add_heading(element.get_text(strip=True), level=2)
        elif element.name == 'h3':
            doc.add_heading(element.get_text(strip=True), level=3)
        elif element.name == 'h4':
            doc.add_heading(element.get_text(strip=True), level=4)
        elif element.name == 'p':
            doc.add_paragraph(element.get_text(strip=True))
        elif element.name in ['ul', 'ol']:
            for li in element.find_all('li'):
                if element.name == 'ul':
                    doc.add_paragraph(li.get_text(strip=True), style='List Bullet')
                else:
                    doc.add_paragraph(li.get_text(strip=True), style='List Number')
        elif element.name == 'table':
            rows = element.find_all('tr')
            if rows:
                num_cols = max(len(row.find_all(['td', 'th'])) for row in rows)
                table = doc.add_table(rows=len(rows), cols=num_cols)
                table.style = 'Table Grid'
                
                for i, row in enumerate(rows):
                    cells = row.find_all(['td', 'th'])
                    for j, cell in enumerate(cells):
                        if j < num_cols:
                            table.cell(i, j).text = cell.get_text(strip=True)
        
        # Process children
        if hasattr(element, 'children'):
            for child in element.children:
                if hasattr(child, 'name'):
                    process_element(child)
    
    process_element(body)
    
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

def html_to_text_enhanced(html_bytes: bytes) -> bytes:
    """Convert HTML to enhanced text"""
    soup = BeautifulSoup(html_bytes, 'html.parser')
    
    # Remove unwanted elements
    for element in soup(["script", "style"]):
        element.decompose()
    
    # Get structured text
    text = soup.get_text(separator='\n')
    
    # Clean up
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return '\n'.join(lines).encode('utf-8')

# -----------------------------
# Streamlit UI with Enhanced Options
# -----------------------------

def main():
    st.set_page_config(
        page_title="Near-Original PDF Converter",
        layout="wide",
        page_icon="🎯"
    )
    
    st.title("🎯 Near-Original PDF Converter")
    st.markdown("""
    ### Direct PDF Rendering + Structured Extraction + Custom Code Insertion
    
    **Features:**
    - 📊 **Direct PDF rendering** in iframe (no grids by default)
    - 🏗️ **Proper structure extraction** (headings, tables, paragraphs)
    - 💻 **Custom code insertion** (CSS/JavaScript)
    - 📄 **Multiple output formats** (HTML, DOCX, Text)
    - ⚡ **Bulk processing** with parallel conversion
    """)
    
    with st.sidebar:
        st.header("⚙️ Conversion Settings")
        
        conversion_type = st.selectbox(
            "Conversion Type",
            [
                "PDF → Enhanced HTML",
                "PDF → Enhanced DOCX", 
                "PDF → Enhanced Text",
                "HTML → Enhanced DOCX",
                "HTML → Enhanced Text"
            ]
        )
        
        st.subheader("🎨 Output Options")
        
        embed_pdf = st.checkbox(
            "Embed PDF Viewer in HTML",
            value=True,
            help="Include original PDF as embedded viewer in HTML output"
        )
        
        show_grid = st.checkbox(
            "Show Grid Overlay (HTML only)",
            value=False,
            help="Add grid overlay to PDF viewer"
        )
        
        preserve_structure = st.checkbox(
            "Preserve Original Structure",
            value=True,
            help="Maintain headings, tables, and formatting"
        )
        
        st.subheader("💻 Custom Code")
        
        custom_css = st.text_area(
            "Custom CSS",
            height=100,
            help="Add custom CSS styles to HTML output"
        )
        
        custom_js = st.text_area(
            "Custom JavaScript", 
            height=100,
            help="Add custom JavaScript to HTML output"
        )
        
        workers = st.number_input(
            "Parallel Workers",
            min_value=1,
            max_value=8,
            value=4
        )
    
    # File upload section
    uploaded_files = st.file_uploader(
        "📁 Upload PDF or HTML Files",
        type=["pdf", "html", "htm"],
        accept_multiple_files=True,
        help="Supported: Digital PDFs (text-based), HTML files"
    )
    
    if not uploaded_files:
        st.info("👆 Upload files to begin near-original conversion")
        st.stop()
    
    # File information display
    st.subheader("📋 Files Ready for Conversion")
    file_data = []
    for file in uploaded_files:
        file_type = "PDF" if file.name.lower().endswith('.pdf') else "HTML"
        file_data.append({
            "Name": file.name,
            "Type": file_type, 
            "Size (KB)": f"{len(file.getvalue()) / 1024:.1f}"
        })
    
    df_files = pd.DataFrame(file_data)
    st.dataframe(df_files, use_container_width=True)
    
    # Quick preview for PDF files
    pdf_files = [f for f in uploaded_files if f.name.lower().endswith('.pdf')]
    if pdf_files:
        with st.expander("🔍 Quick PDF Preview"):
            selected_pdf = st.selectbox("Select PDF for preview", [f.name for f in pdf_files])
            pdf_file = next(f for f in pdf_files if f.name == selected_pdf)
            
            # Display PDF using iframe
            pdf_bytes = pdf_file.getvalue()
            pdf_b64 = base64.b64encode(pdf_bytes).decode()
            
            st.components.v1.html(
                create_pdf_iframe_embed(pdf_bytes, "100%", "400px"),
                height=420
            )
    
    # Conversion button
    if st.button("🚀 Start Near-Original Conversion", type="primary"):
        progress_bar = st.progress(0)
        status_area = st.empty()
        results = []
        
        def process_single_file(file):
            try:
                filename = file.name
                file_bytes = file.getvalue()
                file_ext = os.path.splitext(filename)[1].lower()
                
                if file_ext == '.pdf':
                    # Parse PDF with near-original structure
                    parsed_data = extract_text_near_original(file_bytes)
                    
                    if "PDF → Enhanced HTML" in conversion_type:
                        output_bytes = convert_to_html_with_custom_code(
                            parsed_data, file_bytes, custom_css, custom_js, embed_pdf, show_grid
                        )
                        output_name = filename.replace('.pdf', '_enhanced.html')
                        mime_type = "text/html"
                    
                    elif "PDF → Enhanced DOCX" in conversion_type:
                        output_bytes = convert_to_docx_enhanced(parsed_data)
                        output_name = filename.replace('.pdf', '_enhanced.docx')
                        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    
                    elif "PDF → Enhanced Text" in conversion_type:
                        output_bytes = convert_to_text_enhanced(parsed_data)
                        output_name = filename.replace('.pdf', '_enhanced.txt')
                        mime_type = "text/plain"
                
                elif file_ext in ['.html', '.htm']:
                    if "HTML → Enhanced DOCX" in conversion_type:
                        output_bytes = html_to_docx_enhanced(file_bytes)
                        output_name = filename.replace('.html', '_enhanced.docx').replace('.htm', '_enhanced.docx')
                        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    
                    elif "HTML → Enhanced Text" in conversion_type:
                        output_bytes = html_to_text_enhanced(file_bytes)
                        output_name = filename.replace('.html', '_enhanced.txt').replace('.htm', '_enhanced.txt')
                        mime_type = "text/plain"
                
                return {
                    "success": True,
                    "original_name": filename,
                    "output_name": output_name,
                    "output_bytes": output_bytes,
                    "mime_type": mime_type,
                    "size": len(output_bytes)
                }
                
            except Exception as e:
                return {
                    "success": False,
                    "original_name": filename,
                    "error": str(e)
                }
        
        # Process files
        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures = {executor.submit(process_single_file, file): file for file in uploaded_files}
            
            completed = 0
            for future in as_completed(futures):
                completed += 1
                progress_bar.progress(completed / len(uploaded_files))
                
                result = future.result()
                if result["success"]:
                    results.append(result)
                    status_area.success(f"✅ {result['original_name']} → {result['output_name']}")
                else:
                    status_area.error(f"❌ {result['original_name']} failed: {result['error']}")
        
        # Display results
        if results:
            st.success(f"🎉 Conversion completed! {len(results)} files converted successfully.")
            
            # Individual downloads
            st.subheader("📥 Download Converted Files")
            cols = st.columns(3)
            
            for i, result in enumerate(results):
                with cols[i % 3]:
                    st.download_button(
                        label=f"⬇️ {result['output_name']}",
                        data=result['output_bytes'],
                        file_name=result['output_name'],
                        mime=result['mime_type'],
                        key=f"dl_{i}"
                    )
                    
                    # Preview for text-based files
                    if result['mime_type'].startswith('text/'):
                        with st.expander(f"Preview: {result['output_name']}"):
                            preview_text = result['output_bytes'][:1000].decode('utf-8', errors='replace')
                            st.text_area("", preview_text, height=150, key=f"prev_{i}")
            
            # Bulk download as ZIP
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for result in results:
                    zip_file.writestr(result['output_name'], result['output_bytes'])
            
            zip_buffer.seek(0)
            
            st.download_button(
                label="📦 Download All as ZIP",
                data=zip_buffer.getvalue(),
                file_name="near_original_conversion.zip",
                mime="application/zip"
            )
            
            # Display sample of HTML output
            html_results = [r for r in results if r['mime_type'] == 'text/html']
            if html_results and st.checkbox("Show HTML Output Preview"):
                sample_html = html_results[0]['output_bytes'].decode('utf-8', errors='replace')
                st.components.v1.html(sample_html, height=600, scrolling=True)
        
        else:
            st.error("❌ No files were successfully converted. Please check the errors above.")

if __name__ == "__main__":
    main()
