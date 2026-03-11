#!/usr/bin/env python3
"""
PDF text extractor for NanoClaw medical PPT generation.
Extracts text from PDF files using pdfplumber.
Usage: python3 pdf_reader.py <path_to_pdf>
"""

import sys
import pdfplumber


def extract_text(path: str) -> str:
    """Extract all text from a PDF file."""
    try:
        with pdfplumber.open(path) as pdf:
            pages_text = []
            for i, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                if text:
                    pages_text.append(f"--- Page {i} ---\n{text}")
            return "\n\n".join(pages_text)
    except Exception as e:
        return f"ERROR: Failed to extract text from PDF: {e}"


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 pdf_reader.py <path_to_pdf>")
        sys.exit(1)

    result = extract_text(sys.argv[1])
    print(result)
