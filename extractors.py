"""
extractors.py — Local text extraction from DOCX and PDF files.
No LLM involvement. Pure local processing.
"""

import io
from docx import Document
import pdfplumber


def extract_from_docx(file_bytes: bytes) -> str:
    """Extract text from a DOCX file.
    
    Reads paragraphs and table cells, preserving section structure.
    """
    doc = Document(io.BytesIO(file_bytes))
    parts = []

    for element in doc.element.body:
        tag = element.tag.split('}')[-1]  # Strip namespace

        if tag == 'p':
            # It's a paragraph
            from docx.text.paragraph import Paragraph
            para = Paragraph(element, doc)
            text = para.text.strip()
            if text:
                parts.append(text)

        elif tag == 'tbl':
            # It's a table — extract row by row
            from docx.table import Table
            table = Table(element, doc)
            for row in table.rows:
                row_texts = []
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    if cell_text:
                        row_texts.append(cell_text)
                if row_texts:
                    parts.append(" : ".join(row_texts))

    return "\n".join(parts)


def extract_from_pdf(file_bytes: bytes) -> str:
    """Extract text from a PDF file using pdfplumber.
    
    Layout-aware extraction that preserves reading order.
    """
    text_parts = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text_parts.append(page_text.strip())

            # Also extract tables if present
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    row_texts = [cell.strip() for cell in row if cell and cell.strip()]
                    if row_texts:
                        text_parts.append(" : ".join(row_texts))

    return "\n".join(text_parts)


def extract_text(file_bytes: bytes, filename: str) -> str:
    """Dispatch to the correct extractor based on file extension.
    
    Args:
        file_bytes: Raw file content as bytes.
        filename: Original filename (used to detect extension).
    
    Returns:
        Extracted plain text string.
    
    Raises:
        ValueError: If the file format is not supported.
    """
    ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''

    if ext == 'docx':
        return extract_from_docx(file_bytes)
    elif ext == 'pdf':
        return extract_from_pdf(file_bytes)
    elif ext == 'txt':
        return file_bytes.decode('utf-8', errors='replace')
    else:
        raise ValueError(
            f"Unsupported file format: .{ext}. "
            f"Supported formats: .docx, .pdf, .txt"
        )
