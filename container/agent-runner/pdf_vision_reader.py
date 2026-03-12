#!/usr/bin/env python3
"""
Enhanced PDF reader with Vision AI support.
Converts PDF pages to images and uses Gemini Vision to understand
complex layouts (tables, diagrams, charts) that text extraction misses.

Usage:
    python3 pdf_vision_reader.py <path_to_pdf>                    # text + vision analysis
    python3 pdf_vision_reader.py <path_to_pdf> --text-only        # text extraction only
    python3 pdf_vision_reader.py <path_to_pdf> --extract-images   # also save page images
"""

import json
import os
import sys
from pathlib import Path

import pdfplumber

try:
    from pdf2image import convert_from_path
    HAS_PDF2IMAGE = True
except ImportError:
    HAS_PDF2IMAGE = False

try:
    from google import genai
    from google.genai import types
    HAS_GENAI = True
except ImportError:
    HAS_GENAI = False


def extract_text(path: str) -> list[dict]:
    """Extract text from each page of a PDF."""
    pages = []
    try:
        with pdfplumber.open(path) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                text = page.extract_text() or ""
                tables = page.extract_tables() or []
                pages.append({
                    "page": i,
                    "text": text,
                    "tables": tables,
                    "has_images": len(page.images) > 0 if hasattr(page, 'images') else False,
                })
    except Exception as e:
        pages.append({"page": 1, "text": f"ERROR: {e}", "tables": [], "has_images": False})
    return pages


def pdf_to_images(path: str, output_dir: str | None = None, dpi: int = 200) -> list[str]:
    """Convert PDF pages to images. Returns list of image paths."""
    if not HAS_PDF2IMAGE:
        print("WARNING: pdf2image not installed, skipping image conversion", file=sys.stderr)
        return []

    images = convert_from_path(path, dpi=dpi)
    if output_dir is None:
        output_dir = str(Path(path).parent / f"{Path(path).stem}_pages")
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    paths = []
    for i, img in enumerate(images, 1):
        img_path = str(Path(output_dir) / f"page_{i}.png")
        img.save(img_path, "PNG")
        paths.append(img_path)
    return paths


def analyze_page_with_vision(image_path: str, page_num: int, text_content: str) -> str:
    """Use Gemini Vision to analyze a PDF page image."""
    if not HAS_GENAI:
        return ""

    api_key = os.environ.get("ZENMUX_API_KEY", "")
    if not api_key:
        return ""

    client = genai.Client(
        api_key=api_key,
        vertexai=True,
        http_options=types.HttpOptions(
            api_version="v1",
            base_url="https://zenmux.ai/api/vertex-ai",
        ),
    )

    image_data = Path(image_path).read_bytes()

    prompt = (
        f"This is page {page_num} of a medical/scientific PDF document. "
        "Please analyze the visual content that text extraction might miss:\n"
        "1. Describe any diagrams, flowcharts, or illustrations\n"
        "2. Describe any tables with their data\n"
        "3. Identify any charts/graphs and their key data points\n"
        "4. Note any important formatting or layout elements\n"
        "5. If there are images, describe what they show\n\n"
        f"The extracted text from this page is:\n{text_content[:500]}\n\n"
        "Focus on visual content NOT captured in the text above. "
        "Reply in the same language as the document."
    )

    try:
        response = client.models.generate_content(
            model="google/gemini-3.1-flash-image-preview",
            contents=[
                types.Part.from_bytes(data=image_data, mime_type="image/png"),
                prompt,
            ],
        )
        return response.text or ""
    except Exception as e:
        print(f"WARNING: Vision analysis failed for page {page_num}: {e}", file=sys.stderr)
        return ""


def process_pdf(path: str, text_only: bool = False, extract_images: bool = False) -> str:
    """Process a PDF with text extraction and optional vision analysis."""
    pages = extract_text(path)

    # If text-only or vision not available, return text extraction results
    if text_only or not HAS_PDF2IMAGE or not HAS_GENAI:
        output_parts = []
        for page in pages:
            output_parts.append(f"--- Page {page['page']} ---")
            output_parts.append(page["text"])
            if page["tables"]:
                output_parts.append(f"\n[Tables detected: {len(page['tables'])}]")
                for ti, table in enumerate(page["tables"], 1):
                    output_parts.append(f"Table {ti}:")
                    for row in table:
                        output_parts.append("  | " + " | ".join(str(c or "") for c in row) + " |")
            if page["has_images"]:
                output_parts.append("[Page contains images/figures]")
        return "\n\n".join(output_parts)

    # Vision-enhanced mode
    image_dir = str(Path(path).parent / f"{Path(path).stem}_pages") if extract_images else None
    page_images = pdf_to_images(path, output_dir=image_dir)

    output_parts = []
    for page in pages:
        output_parts.append(f"--- Page {page['page']} ---")
        output_parts.append(page["text"])

        if page["tables"]:
            output_parts.append(f"\n[Tables detected: {len(page['tables'])}]")
            for ti, table in enumerate(page["tables"], 1):
                output_parts.append(f"Table {ti}:")
                for row in table:
                    output_parts.append("  | " + " | ".join(str(c or "") for c in row) + " |")

        # Use vision for pages with images or complex layouts
        page_idx = page["page"] - 1
        if page_idx < len(page_images) and (page["has_images"] or page["tables"]):
            vision_result = analyze_page_with_vision(
                page_images[page_idx], page["page"], page["text"]
            )
            if vision_result:
                output_parts.append(f"\n[Vision Analysis]\n{vision_result}")

    # Clean up temp images if not extracting
    if not extract_images:
        import shutil
        temp_dir = str(Path(path).parent / f"{Path(path).stem}_pages")
        if Path(temp_dir).exists():
            shutil.rmtree(temp_dir, ignore_errors=True)

    return "\n\n".join(output_parts)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 pdf_vision_reader.py <path_to_pdf> [--text-only] [--extract-images]")
        sys.exit(1)

    pdf_path = sys.argv[1]
    text_only = "--text-only" in sys.argv
    extract_images = "--extract-images" in sys.argv

    result = process_pdf(pdf_path, text_only=text_only, extract_images=extract_images)
    print(result)
