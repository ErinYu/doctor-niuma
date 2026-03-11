#!/usr/bin/env python3
"""
Medical PPT Generator for NanoClaw / DoctorNiuMa
Accepts JSON from stdin or first argument, outputs a .pptx file.

JSON schema:
{
  "title": str,
  "subtitle": str (optional),
  "author": str (optional),
  "date": str (optional),
  "style": "academic" | "clinical" | "patient" | "concise",
  "output_path": str,  # e.g. "/workspace/group/report.pptx"
  "slides": [
    {
      "layout": "title" | "section" | "content" | "two_col" | "table" | "blank",
      "title": str,
      "content": [str, ...],     # bullet points (supports nested: "  - sub-bullet")
      "left": [str, ...],        # two_col left column
      "right": [str, ...],       # two_col right column
      "table": [[str, ...], ...],# first row = header
      "notes": str               # speaker notes
    }
  ]
}
"""

import json
import sys
import os
from datetime import datetime

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Inches, Pt
    import pptx.oxml.ns as nsmap
    from lxml import etree
except ImportError:
    print("ERROR: python-pptx not installed. Run: pip3 install python-pptx", file=sys.stderr)
    sys.exit(1)

# ─── Color Palette ────────────────────────────────────────────────────────────
PALETTES = {
    "academic": {
        "primary":   RGBColor(0x1B, 0x4F, 0x72),  # dark navy blue
        "secondary": RGBColor(0x21, 0x8F, 0xBE),  # steel blue
        "accent":    RGBColor(0xE6, 0x7E, 0x22),  # amber
        "bg":        RGBColor(0xFF, 0xFF, 0xFF),
        "text":      RGBColor(0x1C, 0x2B, 0x3A),
        "subtext":   RGBColor(0x5D, 0x6D, 0x7E),
        "table_hdr": RGBColor(0x1B, 0x4F, 0x72),
        "table_alt": RGBColor(0xEB, 0xF5, 0xFB),
    },
    "clinical": {
        "primary":   RGBColor(0x0D, 0x47, 0xA1),  # clinical blue
        "secondary": RGBColor(0x19, 0x76, 0xD2),
        "accent":    RGBColor(0xC6, 0x28, 0x28),  # alert red
        "bg":        RGBColor(0xFF, 0xFF, 0xFF),
        "text":      RGBColor(0x21, 0x21, 0x21),
        "subtext":   RGBColor(0x61, 0x61, 0x61),
        "table_hdr": RGBColor(0x0D, 0x47, 0xA1),
        "table_alt": RGBColor(0xE3, 0xF2, 0xFD),
    },
    "patient": {
        "primary":   RGBColor(0x00, 0x69, 0x5C),  # teal
        "secondary": RGBColor(0x00, 0x96, 0x88),
        "accent":    RGBColor(0xFF, 0x8F, 0x00),  # warm amber
        "bg":        RGBColor(0xFF, 0xFF, 0xFF),
        "text":      RGBColor(0x21, 0x21, 0x21),
        "subtext":   RGBColor(0x54, 0x7F, 0x7A),
        "table_hdr": RGBColor(0x00, 0x69, 0x5C),
        "table_alt": RGBColor(0xE0, 0xF2, 0xF1),
    },
    "concise": {
        "primary":   RGBColor(0x37, 0x47, 0x4F),  # slate
        "secondary": RGBColor(0x54, 0x6E, 0x7A),
        "accent":    RGBColor(0x00, 0x89, 0x7B),  # teal accent
        "bg":        RGBColor(0xFF, 0xFF, 0xFF),
        "text":      RGBColor(0x21, 0x21, 0x21),
        "subtext":   RGBColor(0x60, 0x7D, 0x8B),
        "table_hdr": RGBColor(0x37, 0x47, 0x4F),
        "table_alt": RGBColor(0xEC, 0xEF, 0xF1),
    },
}

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

def get_palette(style: str) -> dict:
    return PALETTES.get(style, PALETTES["academic"])


def set_bg(slide, color: RGBColor):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_text_box(slide, text: str, left, top, width, height,
                 font_size=18, bold=False, color=None, align=PP_ALIGN.LEFT,
                 wrap=True):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    return txBox


def add_rect(slide, left, top, width, height, fill_color: RGBColor):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape


def parse_bullets(lines: list[str]) -> list[tuple[int, str]]:
    """Parse bullet lines into (level, text) tuples. Indent = 2 spaces per level."""
    result = []
    for line in lines:
        stripped = line.lstrip()
        indent = len(line) - len(stripped)
        level = indent // 2
        # Strip leading "- " or "• "
        text = stripped.lstrip('-•').strip()
        result.append((min(level, 2), text))
    return result


def add_bullets_to_tf(tf, bullets: list[tuple[int, str]], palette: dict,
                       base_size=18):
    sizes = [base_size, base_size - 2, base_size - 4]
    colors = [palette["text"], palette["subtext"], palette["subtext"]]

    first = True
    for level, text in bullets:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.level = level
        p.space_before = Pt(2)
        run = p.add_run()
        run.text = text
        run.font.size = Pt(sizes[level])
        run.font.color.rgb = colors[level]


# ─── Slide Builders ───────────────────────────────────────────────────────────

def build_title_slide(prs, data: dict, palette: dict):
    slide_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(slide_layout)
    set_bg(slide, palette["bg"])

    # Left accent bar
    add_rect(slide, Inches(0), Inches(0), Inches(0.4), SLIDE_H, palette["primary"])
    # Bottom accent strip
    add_rect(slide, Inches(0.4), Inches(6.2), SLIDE_W - Inches(0.4), Inches(0.08), palette["secondary"])

    # Title
    title_text = data.get("title", "")
    add_text_box(slide, title_text,
                 Inches(0.8), Inches(2.0), Inches(11.5), Inches(2.0),
                 font_size=36, bold=True, color=palette["primary"])

    # Subtitle
    subtitle = data.get("subtitle", "")
    if subtitle:
        add_text_box(slide, subtitle,
                     Inches(0.8), Inches(3.9), Inches(11.5), Inches(0.8),
                     font_size=22, color=palette["subtext"])

    # Author + date bottom line
    meta_parts = []
    if data.get("author"):
        meta_parts.append(data["author"])
    if data.get("date"):
        meta_parts.append(data["date"])
    else:
        meta_parts.append(datetime.now().strftime("%Y-%m-%d"))
    meta = "   |   ".join(meta_parts)
    add_text_box(slide, meta,
                 Inches(0.8), Inches(6.5), Inches(11.5), Inches(0.5),
                 font_size=14, color=palette["subtext"])
    return slide


def build_section_slide(prs, slide_data: dict, palette: dict):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_bg(slide, palette["primary"])

    title = slide_data.get("title", "")
    add_text_box(slide, title,
                 Inches(1.5), Inches(2.8), Inches(10), Inches(1.8),
                 font_size=40, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF),
                 align=PP_ALIGN.CENTER)

    subtitle = slide_data.get("content", [])
    if subtitle:
        sub_text = subtitle[0] if isinstance(subtitle, list) else subtitle
        add_text_box(slide, sub_text,
                     Inches(1.5), Inches(4.5), Inches(10), Inches(0.8),
                     font_size=20, color=RGBColor(0xCC, 0xDD, 0xEE),
                     align=PP_ALIGN.CENTER)
    return slide


def build_content_slide(prs, slide_data: dict, palette: dict):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_bg(slide, palette["bg"])

    # Title bar
    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.2), palette["primary"])
    title = slide_data.get("title", "")
    add_text_box(slide, title,
                 Inches(0.4), Inches(0.1), Inches(12.5), Inches(1.0),
                 font_size=28, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

    # Accent line below title
    add_rect(slide, Inches(0), Inches(1.2), SLIDE_W, Inches(0.05), palette["secondary"])

    # Bullets
    content = slide_data.get("content", [])
    if content:
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(1.4), Inches(12.3), Inches(5.8))
        tf = txBox.text_frame
        tf.word_wrap = True
        bullets = parse_bullets(content)
        add_bullets_to_tf(tf, bullets, palette, base_size=20)

    return slide


def build_two_col_slide(prs, slide_data: dict, palette: dict):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_bg(slide, palette["bg"])

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.2), palette["primary"])
    title = slide_data.get("title", "")
    add_text_box(slide, title,
                 Inches(0.4), Inches(0.1), Inches(12.5), Inches(1.0),
                 font_size=28, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    add_rect(slide, Inches(0), Inches(1.2), SLIDE_W, Inches(0.05), palette["secondary"])

    # Divider
    add_rect(slide, Inches(6.55), Inches(1.4), Inches(0.05), Inches(5.8), palette["table_alt"])

    left = slide_data.get("left", [])
    right = slide_data.get("right", [])

    for col_content, left_pos in [(left, Inches(0.4)), (right, Inches(6.8))]:
        if col_content:
            txBox = slide.shapes.add_textbox(left_pos, Inches(1.4), Inches(5.9), Inches(5.8))
            tf = txBox.text_frame
            tf.word_wrap = True
            bullets = parse_bullets(col_content)
            add_bullets_to_tf(tf, bullets, palette, base_size=18)

    return slide


def build_table_slide(prs, slide_data: dict, palette: dict):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_bg(slide, palette["bg"])

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.2), palette["primary"])
    title = slide_data.get("title", "")
    add_text_box(slide, title,
                 Inches(0.4), Inches(0.1), Inches(12.5), Inches(1.0),
                 font_size=28, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    add_rect(slide, Inches(0), Inches(1.2), SLIDE_W, Inches(0.05), palette["secondary"])

    table_data = slide_data.get("table", [])
    if not table_data:
        return slide

    rows = len(table_data)
    cols = max(len(r) for r in table_data)
    if rows == 0 or cols == 0:
        return slide

    # Calculate sizes
    tbl_width = Inches(12.3)
    tbl_height = min(Inches(5.5), Inches(0.5 * rows))
    left = Inches(0.5)
    top = Inches(1.4)

    table = slide.shapes.add_table(rows, cols, left, top, tbl_width, tbl_height).table

    for r_idx, row in enumerate(table_data):
        for c_idx in range(cols):
            cell = table.cell(r_idx, c_idx)
            text = row[c_idx] if c_idx < len(row) else ""
            cell.text = str(text)
            tf = cell.text_frame
            tf.paragraphs[0].font.size = Pt(14 if r_idx > 0 else 15)
            tf.paragraphs[0].font.bold = (r_idx == 0)

            # Header row styling
            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = palette["table_hdr"]
                tf.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            elif r_idx % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = palette["table_alt"]

    return slide


BUILDERS = {
    "title":   build_title_slide,
    "section": build_section_slide,
    "content": build_content_slide,
    "two_col": build_two_col_slide,
    "table":   build_table_slide,
}


def add_speaker_notes(slide, notes_text: str):
    if not notes_text:
        return
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text


def generate(data: dict) -> str:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    style = data.get("style", "academic")
    palette = get_palette(style)
    output_path = data.get("output_path", "/workspace/group/presentation.pptx")

    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # Always add title slide first
    title_slide = build_title_slide(prs, data, palette)

    slides = data.get("slides", [])
    for slide_data in slides:
        layout = slide_data.get("layout", "content")
        if layout == "title":
            # Explicit title slide in slides list (skip, already added)
            continue
        builder = BUILDERS.get(layout, build_content_slide)
        slide = builder(prs, slide_data, palette)
        notes = slide_data.get("notes", "")
        if notes:
            add_speaker_notes(slide, notes)

    prs.save(output_path)
    return output_path


def main():
    if len(sys.argv) > 1:
        raw = sys.argv[1]
    else:
        raw = sys.stdin.read()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        print(f"ERROR: Invalid JSON input: {e}", file=sys.stderr)
        sys.exit(1)

    path = generate(data)
    print(path)


if __name__ == "__main__":
    main()
