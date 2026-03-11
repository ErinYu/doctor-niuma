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
      "layout": "title" | "section" | "content" | "two_col" | "table" | "three_col" |
                "left_sidebar" | "right_sidebar" | "quote" | "center_focus" | "comparison" | "blank",
      "title": str,
      "content": [str, ...],     # bullet points (supports nested: "  - sub-bullet")
      "left": [str, ...],        # left column content
      "right": [str, ...],       # right column content
      "col1/col2/col3": [str],   # for three_col layout
      "sidebar": [str, ...],     # for sidebar layouts
      "table": [[str, ...], ...],# first row = header
      "attribution": str,        # for quote layout
      "left_label": str,         # for comparison layout
      "right_label": str,        # for comparison layout
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
    from PIL import Image as PILImage
    import os
    import sys
    from io import BytesIO
except ImportError as e:
    print(f"ERROR: Missing dependency: {e}", file=sys.stderr)
    sys.exit(1)


def create_chart_from_spec(chart_spec: dict, palette: dict):
    """Create chart image from specification (integrates chart_generator)."""
    try:
        # Import chart generator
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from chart_generator import create_chart

        return create_chart(chart_spec, palette)
    except ImportError:
        print("WARNING: Chart generator not available, skipping chart", file=sys.stderr)
        return None

# ─── Color Palette ────────────────────────────────────────────────────────────
PALETTES = {
    "academic": {
        "primary":        RGBColor(0x1B, 0x4F, 0x72),  # dark navy blue
        "secondary":      RGBColor(0x21, 0x8F, 0xBE),  # steel blue
        "accent":         RGBColor(0xE6, 0x7E, 0x22),  # amber
        "bg":             RGBColor(0xFF, 0xFF, 0xFF),
        "text":           RGBColor(0x1C, 0x2B, 0x3A),
        "subtext":        RGBColor(0x5D, 0x6D, 0x7E),
        "table_hdr":      RGBColor(0x1B, 0x4F, 0x72),
        "table_alt":      RGBColor(0xEB, 0xF5, 0xFB),
        # Gradient definitions
        "gradient_start": RGBColor(0x1B, 0x4F, 0x72),
        "gradient_end":   RGBColor(0x2C, 0x6F, 0x92),
        "background_tint": RGBColor(0xF0, 0xF5, 0xFA),
    },
    "clinical": {
        "primary":        RGBColor(0x0D, 0x47, 0xA1),  # clinical blue
        "secondary":      RGBColor(0x19, 0x76, 0xD2),
        "accent":         RGBColor(0xC6, 0x28, 0x28),  # alert red
        "bg":             RGBColor(0xFF, 0xFF, 0xFF),
        "text":           RGBColor(0x21, 0x21, 0x21),
        "subtext":        RGBColor(0x61, 0x61, 0x61),
        "table_hdr":      RGBColor(0x0D, 0x47, 0xA1),
        "table_alt":      RGBColor(0xE3, 0xF2, 0xFD),
        # Gradient definitions
        "gradient_start": RGBColor(0x0D, 0x47, 0xA1),
        "gradient_end":   RGBColor(0x1A, 0x60, 0xC0),
        "background_tint": RGBColor(0xE3, 0xF2, 0xFD),
    },
    "patient": {
        "primary":        RGBColor(0x00, 0x69, 0x5C),  # teal
        "secondary":      RGBColor(0x00, 0x96, 0x88),
        "accent":         RGBColor(0xFF, 0x8F, 0x00),  # warm amber
        "bg":             RGBColor(0xFF, 0xFF, 0xFF),
        "text":           RGBColor(0x21, 0x21, 0x21),
        "subtext":        RGBColor(0x54, 0x7F, 0x7A),
        "table_hdr":      RGBColor(0x00, 0x69, 0x5C),
        "table_alt":      RGBColor(0xE0, 0xF2, 0xF1),
        # Gradient definitions
        "gradient_start": RGBColor(0x00, 0x69, 0x5C),
        "gradient_end":   RGBColor(0x00, 0x80, 0x70),
        "background_tint": RGBColor(0xE0, 0xF2, 0xF1),
    },
    "concise": {
        "primary":        RGBColor(0x37, 0x47, 0x4F),  # slate
        "secondary":      RGBColor(0x54, 0x6E, 0x7A),
        "accent":         RGBColor(0x00, 0x89, 0x7B),  # teal accent
        "bg":             RGBColor(0xFF, 0xFF, 0xFF),
        "text":           RGBColor(0x21, 0x21, 0x21),
        "subtext":        RGBColor(0x60, 0x7D, 0x8B),
        "table_hdr":      RGBColor(0x37, 0x47, 0x4F),
        "table_alt":      RGBColor(0xEC, 0xEF, 0xF1),
        # Gradient definitions
        "gradient_start": RGBColor(0x37, 0x47, 0x4F),
        "gradient_end":   RGBColor(0x4A, 0x5A, 0x62),
        "background_tint": RGBColor(0xEC, 0xEF, 0xF1),
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


def add_gradient_fill(slide, start_color: RGBColor, end_color: RGBColor):
    """Add gradient background to slide (simulated with solid for compatibility)."""
    # Note: python-pptx doesn't support gradient fills directly
    # Using solid fill with start_color as fallback
    set_bg(slide, start_color)


def add_pattern_overlay(slide, base_color: RGBColor, opacity=0.05):
    """Add subtle pattern overlay to slide background."""
    # Set base background
    set_bg(slide, base_color)
    # Pattern overlay would require custom XML - simplified for now


def add_side_bar(slide, color: RGBColor, width=Inches(0.3)):
    """Add colored side bar to slide."""
    add_rect(slide, Inches(0), Inches(0), width, SLIDE_H, color)


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


# ─── Typography ───────────────────────────────────────────────────────────────

FONTS = {
    "academic": {
        "heading": "Calibri",
        "body": "Calibri",
        "code": "Consolas",
        "size_h1": 36,
        "size_h2": 28,
        "size_h3": 24,
        "size_body": 20,
        "size_small": 16,
    },
    "clinical": {
        "heading": "Arial",
        "body": "Calibri",
        "code": "Consolas",
        "size_h1": 36,
        "size_h2": 28,
        "size_h3": 24,
        "size_body": 20,
        "size_small": 16,
    },
    "patient": {
        "heading": "Verdana",
        "body": "Calibri",
        "code": "Consolas",
        "size_h1": 36,
        "size_h2": 28,
        "size_h3": 24,
        "size_body": 20,
        "size_small": 16,
    },
    "concise": {
        "heading": "Calibri",
        "body": "Calibri",
        "code": "Consolas",
        "size_h1": 36,
        "size_h2": 28,
        "size_h3": 24,
        "size_body": 20,
        "size_small": 16,
    },
}


def get_font_spec(style: str, element_type="body") -> tuple:
    """Get (font_name, font_size) for given style and element type."""
    fonts = FONTS.get(style, FONTS["academic"])
    element_type = element_type.lower()

    if element_type == "h1":
        return (fonts["heading"], fonts["size_h1"])
    elif element_type == "h2":
        return (fonts["heading"], fonts["size_h2"])
    elif element_type == "h3":
        return (fonts["heading"], fonts["size_h3"])
    elif element_type == "small":
        return (fonts["body"], fonts["size_small"])
    else:
        return (fonts["body"], fonts["size_body"])


# ─── Decorative Elements ─────────────────────────────────────────────────────

DECORATIVE_ELEMENTS = {
    "accent_bar": "Colored bar at slide edge",
    "corner_accent": "Triangle accent in corner",
    "circle_accent": "Partial circle border",
    "divider": "Horizontal line with style options",
    "icon_bullets": "Custom bullet icons (check, arrow, dot)",
    "number_badge": "Stylized numbers for lists",
}


def add_accent_bar(slide, position="left", color=None, palette=None):
    """Add accent bar at slide edge."""
    if palette and color is None:
        color = palette["primary"]
    if position == "left":
        add_rect(slide, Inches(0), Inches(0), Inches(0.3), SLIDE_H, color)
    elif position == "right":
        add_rect(slide, SLIDE_W - Inches(0.3), Inches(0), Inches(0.3), SLIDE_H, color)
    elif position == "top":
        add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(0.08), color)


def add_corner_accent(slide, corner="top-left", color=None, palette=None):
    """Add triangular accent in corner."""
    if palette and color is None:
        color = palette["accent"]
    size = Inches(1.5)
    if corner == "top-left":
        add_rect(slide, Inches(0), Inches(0), size, Inches(0.15), color)
        add_rect(slide, Inches(0), Inches(0), Inches(0.15), size, color)
    elif corner == "top-right":
        add_rect(slide, SLIDE_W - size, Inches(0), size, Inches(0.15), color)
        add_rect(slide, SLIDE_W - Inches(0.15), Inches(0), Inches(0.15), size, color)


def add_divider(slide, y_pos, width=Inches(10), color=None, palette=None, style="solid"):
    """Add horizontal divider line."""
    if palette and color is None:
        color = palette["subtext"]
    left = (SLIDE_W - width) / 2
    add_rect(slide, left, y_pos, width, Inches(0.05), color)


def add_decorations(slide, layout_type, palette):
    """Apply decorative elements based on layout type."""
    # Accent bars for content slides
    if layout_type in ["content", "two_col", "three_col", "left_sidebar", "right_sidebar", "comparison"]:
        add_accent_bar(slide, position="left", palette=palette)

    # Corner accents for special slides
    if layout_type in ["quote", "center_focus"]:
        add_corner_accent(slide, corner="top-right", palette=palette)


def add_image_to_slide(slide, image_path: str, left, top, width, style="plain", palette=None):
    """Add an image to a slide with optional styling."""
    if not os.path.exists(image_path):
        print(f"WARNING: Image not found: {image_path}", file=sys.stderr)
        return None

    try:
        pic = slide.shapes.add_picture(image_path, left, top, width=width)
        pic = pic._element  # Get underlying element for advanced styling

        # Apply style effects
        if style == "rounded":
            # Rounded corners effect (simplified - creates border)
            pic.get_or_add_ln().get_or_add_spPr().get_or_add_ln().get_or_add_w()
        elif style == "shadow":
            # Shadow effect
            pass  # Would require XML manipulation
        elif style == "border" and palette:
            # Border effect
            pass  # Would require XML manipulation

        return pic
    except Exception as e:
        print(f"WARNING: Failed to add image {image_path}: {e}", file=sys.stderr)
        return None


def build_image_left_slide(prs, slide_data: dict, palette: dict):
    """Image on left (40%), text on right (60%)."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_bg(slide, palette["bg"])

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.2), palette["primary"])
    title = slide_data.get("title", "")
    add_text_box(slide, title,
                 Inches(0.4), Inches(0.1), Inches(12.5), Inches(1.0),
                 font_size=28, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    add_rect(slide, Inches(0), Inches(1.2), SLIDE_W, Inches(0.05), palette["secondary"])

    # Image (left 40%)
    image_spec = slide_data.get("image", {})
    if isinstance(image_spec, str):
        image_spec = {"path": image_spec}
    image_path = image_spec.get("path", "")
    image_style = image_spec.get("style", "plain")

    if image_path and os.path.exists(image_path):
        img_width = Inches(4.8)
        add_image_to_slide(slide, image_path, Inches(0.5), Inches(1.4), img_width, image_style, palette)

    # Content (right 60%)
    content = slide_data.get("content", [])
    if content:
        txBox = slide.shapes.add_textbox(Inches(5.5), Inches(1.4), Inches(7.3), Inches(5.6))
        tf = txBox.text_frame
        tf.word_wrap = True
        bullets = parse_bullets(content)
        add_bullets_to_tf(tf, bullets, palette, base_size=18)

    return slide


def build_image_right_slide(prs, slide_data: dict, palette: dict):
    """Text on left (60%), image on right (40%)."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_bg(slide, palette["bg"])

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.2), palette["primary"])
    title = slide_data.get("title", "")
    add_text_box(slide, title,
                 Inches(0.4), Inches(0.1), Inches(12.5), Inches(1.0),
                 font_size=28, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    add_rect(slide, Inches(0), Inches(1.2), SLIDE_W, Inches(0.05), palette["secondary"])

    # Content (left 60%)
    content = slide_data.get("content", [])
    if content:
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(1.4), Inches(7.3), Inches(5.6))
        tf = txBox.text_frame
        tf.word_wrap = True
        bullets = parse_bullets(content)
        add_bullets_to_tf(tf, bullets, palette, base_size=18)

    # Image (right 40%)
    image_spec = slide_data.get("image", {})
    if isinstance(image_spec, str):
        image_spec = {"path": image_spec}
    image_path = image_spec.get("path", "")
    image_style = image_spec.get("style", "plain")

    if image_path and os.path.exists(image_path):
        img_width = Inches(4.8)
        add_image_to_slide(slide, image_path, Inches(8.0), Inches(1.4), img_width, image_style, palette)

    return slide


def build_image_top_slide(prs, slide_data: dict, palette: dict):
    """Image on top (50%), text on bottom (50%)."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_bg(slide, palette["bg"])

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.2), palette["primary"])
    title = slide_data.get("title", "")
    add_text_box(slide, title,
                 Inches(0.4), Inches(0.1), Inches(12.5), Inches(1.0),
                 font_size=28, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    add_rect(slide, Inches(0), Inches(1.2), SLIDE_W, Inches(0.05), palette["secondary"])

    # Image (top 50% of content area)
    image_spec = slide_data.get("image", {})
    if isinstance(image_spec, str):
        image_spec = {"path": image_spec}
    image_path = image_spec.get("path", "")
    image_style = image_spec.get("style", "plain")

    if image_path and os.path.exists(image_path):
        img_width = Inches(10)
        add_image_to_slide(slide, image_path, Inches(1.5), Inches(1.4), img_width, image_style, palette)

    # Content (bottom 50% of content area)
    content = slide_data.get("content", [])
    if content:
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        bullets = parse_bullets(content)
        add_bullets_to_tf(tf, bullets, palette, base_size=16)

    return slide


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

    # Apply enhanced table styling
    for r_idx, row in enumerate(table_data):
        for c_idx in range(cols):
            cell = table.cell(r_idx, c_idx)
            text = row[c_idx] if c_idx < len(row) else ""
            cell.text = str(text)
            tf = cell.text_frame

            # Enhanced header row styling
            if r_idx == 0:
                # Gradient-like effect with solid fill using accent
                cell.fill.solid()
                cell.fill.fore_color.rgb = palette["primary"]
                tf.paragraphs[0].font.size = Pt(16)
                tf.paragraphs[0].font.bold = True
                tf.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                # Center align header text
                tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            else:
                # Data row styling
                tf.paragraphs[0].font.size = Pt(14)
                tf.paragraphs[0].font.color.rgb = palette["text"]

            # Enhanced zebra striping with background_tint color
            if r_idx > 0 and r_idx % 2 == 1:
                cell.fill.solid()
                cell.fill.fore_color.rgb = palette["background_tint"]

            # Enhanced border styling
            border = cell.border
            border.top.color.rgb = palette["subtext"]
            border.bottom.color.rgb = palette["subtext"]
            border.left.color.rgb = palette["subtext"]
            border.right.color.rgb = palette["subtext"]
            border.top.width = Pt(1)
            border.bottom.width = Pt(1)
            border.left.width = Pt(1)
            border.right.width = Pt(1)

    return slide


def build_three_col_slide(prs, slide_data: dict, palette: dict):
    """Three equal columns for comparisons."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_bg(slide, palette["bg"])

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.2), palette["primary"])
    title = slide_data.get("title", "")
    add_text_box(slide, title,
                 Inches(0.4), Inches(0.1), Inches(12.5), Inches(1.0),
                 font_size=28, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    add_rect(slide, Inches(0), Inches(1.2), SLIDE_W, Inches(0.05), palette["secondary"])

    # Three equal columns
    col_width = Inches(4.0)
    positions = [Inches(0.4), Inches(4.5), Inches(8.6)]
    dividers = [Inches(4.4), Inches(8.5)]

    # Vertical dividers
    for div_pos in dividers:
        add_rect(slide, div_pos, Inches(1.4), Inches(0.05), Inches(5.6), palette["table_alt"])

    for i in range(3):
        col_key = f"col{i+1}"
        col_content = slide_data.get(col_key, slide_data.get("content", []))
        if col_content:
            txBox = slide.shapes.add_textbox(positions[i], Inches(1.4), col_width, Inches(5.6))
            tf = txBox.text_frame
            tf.word_wrap = True
            bullets = parse_bullets(col_content if isinstance(col_content, list) else [col_content])
            add_bullets_to_tf(tf, bullets, palette, base_size=16)

    return slide


def build_left_sidebar_slide(prs, slide_data: dict, palette: dict):
    """Content area with left sidebar for notes/definitions."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_bg(slide, palette["bg"])

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.2), palette["primary"])
    title = slide_data.get("title", "")
    add_text_box(slide, title,
                 Inches(0.4), Inches(0.1), Inches(12.5), Inches(1.0),
                 font_size=28, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    add_rect(slide, Inches(0), Inches(1.2), SLIDE_W, Inches(0.05), palette["secondary"])

    # Sidebar background (left 30%)
    add_rect(slide, Inches(0), Inches(1.25), Inches(3.9), Inches(6.0), palette["table_alt"])

    # Divider
    add_rect(slide, Inches(3.9), Inches(1.4), Inches(0.05), Inches(5.6), palette["secondary"])

    # Sidebar content (notes/definitions)
    sidebar = slide_data.get("sidebar", slide_data.get("left", []))
    if sidebar:
        txBox = slide.shapes.add_textbox(Inches(0.3), Inches(1.5), Inches(3.5), Inches(5.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        bullets = parse_bullets(sidebar if isinstance(sidebar, list) else [sidebar])
        add_bullets_to_tf(tf, bullets, palette, base_size=14)

    # Main content
    content = slide_data.get("content", [])
    if content:
        txBox = slide.shapes.add_textbox(Inches(4.1), Inches(1.5), Inches(8.7), Inches(5.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        bullets = parse_bullets(content)
        add_bullets_to_tf(tf, bullets, palette, base_size=18)

    return slide


def build_right_sidebar_slide(prs, slide_data: dict, palette: dict):
    """Content area with right sidebar for stats/highlights."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_bg(slide, palette["bg"])

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.2), palette["primary"])
    title = slide_data.get("title", "")
    add_text_box(slide, title,
                 Inches(0.4), Inches(0.1), Inches(12.5), Inches(1.0),
                 font_size=28, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    add_rect(slide, Inches(0), Inches(1.2), SLIDE_W, Inches(0.05), palette["secondary"])

    # Sidebar background (right 30%)
    add_rect(slide, Inches(9.4), Inches(1.25), Inches(3.9), Inches(6.0), palette["table_alt"])

    # Divider
    add_rect(slide, Inches(9.35), Inches(1.4), Inches(0.05), Inches(5.6), palette["secondary"])

    # Main content
    content = slide_data.get("content", [])
    if content:
        txBox = slide.shapes.add_textbox(Inches(0.4), Inches(1.5), Inches(8.7), Inches(5.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        bullets = parse_bullets(content)
        add_bullets_to_tf(tf, bullets, palette, base_size=18)

    # Sidebar content (stats/highlights)
    sidebar = slide_data.get("sidebar", slide_data.get("right", []))
    if sidebar:
        txBox = slide.shapes.add_textbox(Inches(9.5), Inches(1.5), Inches(3.7), Inches(5.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        bullets = parse_bullets(sidebar if isinstance(sidebar, list) else [sidebar])
        add_bullets_to_tf(tf, bullets, palette, base_size=14)

    return slide


def build_quote_slide(prs, slide_data: dict, palette: dict):
    """Large quote with attribution."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_bg(slide, palette["bg"])

    # Decorative accent bar on left
    add_rect(slide, Inches(0), Inches(0), Inches(0.6), SLIDE_H, palette["accent"])

    # Quote content
    quote_text = slide_data.get("content", [""])[0] if slide_data.get("content") else ""
    add_text_box(slide, f'"{quote_text}"',
                 Inches(1.2), Inches(2.0), Inches(11.5), Inches(3.5),
                 font_size=32, color=palette["primary"], align=PP_ALIGN.CENTER)

    # Attribution
    attribution = slide_data.get("attribution", slide_data.get("title", ""))
    if attribution:
        add_text_box(slide, f"— {attribution}",
                     Inches(1.2), Inches(5.5), Inches(11.5), Inches(0.8),
                     font_size=20, bold=True, color=palette["subtext"], align=PP_ALIGN.CENTER)

    # Accent line below
    add_rect(slide, Inches(4), Inches(6.2), Inches(5.3), Inches(0.08), palette["accent"])

    return slide


def build_center_focus_slide(prs, slide_data: dict, palette: dict):
    """Centered content with accent border."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_bg(slide, palette["bg"])

    # Outer accent border
    border_thickness = Inches(0.15)
    add_rect(slide, Inches(0), Inches(0), SLIDE_W, border_thickness, palette["primary"])  # Top
    add_rect(slide, Inches(0), SLIDE_H - border_thickness, SLIDE_W, border_thickness, palette["primary"])  # Bottom
    add_rect(slide, Inches(0), Inches(0), border_thickness, SLIDE_H, palette["primary"])  # Left
    add_rect(slide, SLIDE_W - border_thickness, Inches(0), border_thickness, SLIDE_H, palette["primary"])  # Right

    # Title
    title = slide_data.get("title", "")
    add_text_box(slide, title,
                 Inches(1.5), Inches(1.5), Inches(10.3), Inches(1.2),
                 font_size=32, bold=True, color=palette["primary"], align=PP_ALIGN.CENTER)

    # Centered content
    content = slide_data.get("content", [])
    if content:
        txBox = slide.shapes.add_textbox(Inches(1.5), Inches(2.8), Inches(10.3), Inches(4.2))
        tf = txBox.text_frame
        tf.word_wrap = True
        bullets = parse_bullets(content)
        add_bullets_to_tf(tf, bullets, palette, base_size=20)
        # Center align all paragraphs
        for p in tf.paragraphs:
            p.alignment = PP_ALIGN.CENTER

    return slide


def build_comparison_slide(prs, slide_data: dict, palette: dict):
    """Before/after or Option A/B comparison matrix."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_bg(slide, palette["bg"])

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.2), palette["primary"])
    title = slide_data.get("title", "")
    add_text_box(slide, title,
                 Inches(0.4), Inches(0.1), Inches(12.5), Inches(1.0),
                 font_size=28, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    add_rect(slide, Inches(0), Inches(1.2), SLIDE_W, Inches(0.05), palette["secondary"])

    # Get column labels (default: "Option A", "Option B")
    left_label = slide_data.get("left_label", "Option A")
    right_label = slide_data.get("right_label", "Option B")

    # Column headers with different backgrounds
    add_rect(slide, Inches(0.4), Inches(1.4), Inches(6.2), Inches(0.6), palette["secondary"])
    add_rect(slide, Inches(6.7), Inches(1.4), Inches(6.2), Inches(0.6), palette["primary"])

    add_text_box(slide, left_label,
                 Inches(0.4), Inches(1.45), Inches(6.2), Inches(0.5),
                 font_size=20, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)
    add_text_box(slide, right_label,
                 Inches(6.7), Inches(1.45), Inches(6.2), Inches(0.5),
                 font_size=20, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), align=PP_ALIGN.CENTER)

    # Content columns
    left = slide_data.get("left", slide_data.get("content", []))
    right = slide_data.get("right", [])

    for col_content, left_pos in [(left, Inches(0.4)), (right, Inches(6.7))]:
        if col_content:
            txBox = slide.shapes.add_textbox(left_pos, Inches(2.1), Inches(6.2), Inches(5.0))
            tf = txBox.text_frame
            tf.word_wrap = True
            bullets = parse_bullets(col_content if isinstance(col_content, list) else [col_content])
            add_bullets_to_tf(tf, bullets, palette, base_size=16)

    return slide


def build_process_slide(prs, slide_data: dict, palette: dict):
    """Step-by-step process flow with numbered steps."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_bg(slide, palette["bg"])

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.2), palette["primary"])
    title = slide_data.get("title", "")
    add_text_box(slide, title,
                 Inches(0.4), Inches(0.1), Inches(12.5), Inches(1.0),
                 font_size=28, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    add_rect(slide, Inches(0), Inches(1.2), SLIDE_W, Inches(0.05), palette["secondary"])

    # Process steps - horizontal layout
    steps = slide_data.get("steps", slide_data.get("content", []))
    if not steps:
        return slide

    num_steps = min(len(steps), 5)  # Max 5 steps horizontally
    step_width = Inches(12.0 / num_steps)

    for i, step in enumerate(steps[:5]):
        x_pos = Inches(0.5 + i * (12.0 / num_steps))

        # Step circle with number
        circle = slide.shapes.add_shape(
            9,  # MSO_SHAPE_TYPE.OVAL
            x_pos, Inches(1.5), Inches(0.7), Inches(0.7)
        )
        circle.fill.solid()
        circle.fill.fore_color.rgb = palette["accent"]
        circle.line.color.rgb = palette["primary"]
        circle.line.width = Pt(2)

        # Step number
        txBox = slide.shapes.add_textbox(x_pos, Inches(1.6), Inches(0.7), Inches(0.5))
        tf = txBox.text_frame
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        run = tf.paragraphs[0].add_run()
        run.text = str(i + 1)
        run.font.size = Pt(24)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        # Step text below
        txBox = slide.shapes.add_textbox(x_pos, Inches(2.4), step_width, Inches(3.5))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = step if isinstance(step, str) else step[0] if step else ""
        run.font.size = Pt(14)
        run.font.color.rgb = palette["text"]

        # Arrow to next step (except last)
        if i < num_steps - 1:
            add_rect(slide, x_pos + Inches(0.7), Inches(1.8), Inches(0.3), Inches(0.05), palette["subtext"])

    return slide


def build_chart_slide(prs, slide_data: dict, palette: dict):
    """Slide with generated chart image."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_bg(slide, palette["bg"])

    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.2), palette["primary"])
    title = slide_data.get("title", "")
    add_text_box(slide, title,
                 Inches(0.4), Inches(0.1), Inches(12.5), Inches(1.0),
                 font_size=28, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    add_rect(slide, Inches(0), Inches(1.2), SLIDE_W, Inches(0.05), palette["secondary"])

    # Get chart specification
    chart_spec = slide_data.get("chart", {})
    if not chart_spec:
        return slide

    # Generate chart image
    chart_buf = create_chart_from_spec(chart_spec, palette)
    if chart_buf:
        # Save chart to temp file and insert
        import tempfile
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            tmp.write(chart_buf.getvalue())
            tmp_path = tmp.name

        try:
            # Add chart image centered
            img_width = Inches(10)
            img_height = Inches(5)
            left = (SLIDE_W - img_width) / 2
            top = Inches(1.5)
            slide.shapes.add_picture(tmp_path, left, top, width=img_width, height=img_height)
        finally:
            os.unlink(tmp_path)

    # Add caption/notes if provided
    caption = slide_data.get("caption", slide_data.get("content", [""])[0] if slide_data.get("content") else "")
    if caption:
        add_text_box(slide, caption,
                     Inches(0.5), Inches(6.6), Inches(12.3), Inches(1.0),
                     font_size=14, color=palette["subtext"])

    return slide


BUILDERS = {
    "title":          build_title_slide,
    "section":        build_section_slide,
    "content":        build_content_slide,
    "two_col":        build_two_col_slide,
    "table":          build_table_slide,
    "three_col":      build_three_col_slide,
    "left_sidebar":   build_left_sidebar_slide,
    "right_sidebar":  build_right_sidebar_slide,
    "quote":          build_quote_slide,
    "center_focus":   build_center_focus_slide,
    "comparison":     build_comparison_slide,
    "process":        build_process_slide,
    "image_left":     build_image_left_slide,
    "image_right":    build_image_right_slide,
    "image_top":      build_image_top_slide,
    "chart":          build_chart_slide,
}


def add_speaker_notes(slide, notes_text: str):
    if not notes_text:
        return
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text


def generate_style_previews(output_dir="/workspace/group/previews"):
    """Generate preview images for each style."""
    try:
        os.makedirs(output_dir, exist_ok=True)

        for style_name in PALETTES.keys():
            # Create a sample presentation
            prs = Presentation()
            prs.slide_width = SLIDE_W
            prs.slide_height = SLIDE_H

            palette = get_palette(style_name)

            # Create a sample slide with title and content
            slide_layout = prs.slide_layouts[6]  # blank
            slide = prs.slides.add_slide(slide_layout)
            set_bg(slide, palette["bg"])

            # Add decorative elements
            add_accent_bar(slide, position="left", palette=palette)

            # Title
            add_text_box(slide, f"{style_name.title()} Style",
                         Inches(1.5), Inches(2.0), Inches(10), Inches(0.8),
                         font_size=36, bold=True, color=palette["primary"])

            # Sample content
            add_text_box(slide, f"This is a sample {style_name} presentation style.",
                         Inches(1.5), Inches(3.0), Inches(10), Inches(0.6),
                         font_size=18, color=palette["text"])

            # Add decorative divider
            add_rect(slide, Inches(3), Inches(3.8), Inches(7), Inches(0.08), palette["secondary"])

            # Sample bullets
            txBox = slide.shapes.add_textbox(Inches(1.5), Inches(4.2), Inches(10), Inches(2.5))
            tf = txBox.text_frame
            tf.word_wrap = True
            bullets = [
                (0, "Feature 1: Professional typography"),
                (0, "Feature 2: Coordinated color palette"),
                (0, "Feature 3: Visual depth and accents"),
            ]
            add_bullets_to_tf(tf, bullets, palette, base_size=16)

            # Save preview
            preview_path = os.path.join(output_dir, f"{style_name}_preview.png")
            prs.save(preview_path)
            print(f"Generated preview: {preview_path}")

    except Exception as e:
        print(f"WARNING: Could not generate previews: {e}", file=sys.stderr)


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
