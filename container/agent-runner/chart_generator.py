#!/usr/bin/env python3
"""
Chart Generator for NanoClaw PPT System
Generates chart images using matplotlib for insertion into PowerPoint slides.

Chart data schema:
{
  "type": "bar" | "line" | "pie" | "scatter" | "combo",
  "title": str,
  "data": {
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series": [
      {
        "name": "Revenue",
        "values": [100, 120, 140, 160],
        "color": "primary" | "secondary" | "accent" | "#RRGGBB"
      }
    ]
  },
  "options": {
    "show_legend": bool,
    "show_values": bool,
    "style": "clustered" | "stacked" | "percentage"
  }
}
"""

import sys
import os
from io import BytesIO
from typing import Dict, List, Any

try:
    import matplotlib
    matplotlib.use('Agg')  # Non-interactive backend
    import matplotlib.pyplot as plt
    import numpy as np
except ImportError:
    print("ERROR: matplotlib not installed. Run: pip install matplotlib", file=sys.stderr)
    sys.exit(1)

from pptx.dml.color import RGBColor


def hex_to_rgb(hex_color: str) -> RGBColor:
    """Convert hex color string to RGBColor."""
    hex_color = hex_color.lstrip('#')
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))


def get_chart_color(color_spec: str, palette: Dict[str, RGBColor]) -> RGBColor:
    """Get color from palette or hex string."""
    if color_spec in palette:
        return palette[color_spec]
    elif color_spec.startswith('#'):
        return hex_to_rgb(color_spec)
    else:
        return palette["primary"]


def create_bar_chart(ax, chart_data: Dict, palette: Dict[str, RGBColor]):
    """Create a bar chart."""
    categories = chart_data["categories"]
    series_list = chart_data["series"]

    x = np.arange(len(categories))
    width = 0.8 / len(series_list)

    for i, series in enumerate(series_list):
        values = series["values"]
        color = get_chart_color(series.get("color", "primary"), palette)
        offset = (i - len(series_list) / 2) * width
        ax.bar(x + offset, values, width, label=series["name"], color=color.rgb[:3])

    ax.set_xticks(x)
    ax.set_xticklabels(categories)
    ax.set_title(chart_data.get("title", ""))


def create_line_chart(ax, chart_data: Dict, palette: Dict[str, RGBColor]):
    """Create a line chart."""
    categories = chart_data["categories"]
    series_list = chart_data["series"]

    for series in series_list:
        values = series["values"]
        color = get_chart_color(series.get("color", "primary"), palette)
        ax.plot(categories, values, marker='o', label=series["name"], color=color.rgb[:3])

    ax.set_title(chart_data.get("title", ""))


def create_pie_chart(ax, chart_data: Dict, palette: Dict[str, RGBColor]):
    """Create a pie chart."""
    series = chart_data["series"][0]
    values = series["values"]
    labels = chart_data["categories"]

    # Use palette colors or cycle through colors
    colors = []
    for i, val in enumerate(values):
        if i < len(series_list):
            color_spec = series_list[i].get("color", "primary")
            colors.append(get_chart_color(color_spec, palette).rgb[:3])
        else:
            # Fallback colors
            fallback_colors = [
                (0x1B, 0x4F, 0x72), (0x21, 0x8F, 0xBE), (0xE6, 0x7E, 0x22),
                (0x00, 0x69, 0x5C), (0x37, 0x47, 0x4F)
            ]
            colors.append(fallback_colors[i % len(fallback_colors)])

    ax.pie(values, labels=labels, autopct='%1.1f%%', colors=colors)
    ax.set_title(chart_data.get("title", ""))


def create_scatter_chart(ax, chart_data: Dict, palette: Dict[str, RGBColor]):
    """Create a scatter plot."""
    series = chart_data["series"][0]
    x_values = series.get("x", series.get("values", []))
    y_values = series.get("y", [])
    color = get_chart_color(series.get("color", "primary"), palette)

    if len(x_values) == len(y_values):
        ax.scatter(x_values, y_values, color=color.rgb[:3], alpha=0.7)

    ax.set_title(chart_data.get("title", ""))
    ax.set_xlabel(chart_data.get("x_label", "X"))
    ax.set_ylabel(chart_data.get("y_label", "Y"))


def create_combo_chart(ax, chart_data: Dict, palette: Dict[str, RGBColor]):
    """Create a combo chart (bar + line)."""
    categories = chart_data["categories"]
    series_list = chart_data["series"]

    if len(series_list) < 2:
        # Fall back to bar chart
        create_bar_chart(ax, chart_data, palette)
        return

    x = np.arange(len(categories))
    width = 0.35

    # First series as bars
    bar_series = series_list[0]
    bar_color = get_chart_color(bar_series.get("color", "primary"), palette)
    ax.bar(x - width/2, bar_series["values"], width, label=bar_series["name"], color=bar_color.rgb[:3])

    # Second series as line
    line_series = series_list[1]
    line_color = get_chart_color(line_series.get("color", "accent"), palette)
    ax.plot(x, line_series["values"], marker='o', label=line_series["name"], color=line_color.rgb[:3])

    ax.set_xticks(x)
    ax.set_xticklabels(categories)
    ax.set_title(chart_data.get("title", ""))


def create_chart(chart_spec: Dict, palette: Dict[str, RGBColor]) -> BytesIO:
    """
    Create a chart from specification and return as BytesIO buffer.

    Args:
        chart_spec: Chart data specification
        palette: Color palette for the chart

    Returns:
        BytesIO buffer containing PNG image data
    """
    chart_type = chart_spec.get("type", "bar")
    data = chart_spec.get("data", {})
    options = chart_spec.get("options", {})

    # Create figure
    fig, ax = plt.subplots(figsize=(10, 6))

    # Create chart based on type
    if chart_type == "bar":
        create_bar_chart(ax, data, palette)
    elif chart_type == "line":
        create_line_chart(ax, data, palette)
    elif chart_type == "pie":
        create_pie_chart(ax, data, palette)
    elif chart_type == "scatter":
        create_scatter_chart(ax, data, palette)
    elif chart_type == "combo":
        create_combo_chart(ax, data, palette)
    else:
        raise ValueError(f"Unknown chart type: {chart_type}")

    # Add legend if requested
    if options.get("show_legend", True):
        ax.legend()

    # Add values if requested
    if options.get("show_values", False) and chart_type in ["bar", "line"]:
        # Add value labels on bars/points
        pass  # Simplified for now

    plt.tight_layout()

    # Save to BytesIO
    buf = BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight',
                transparent=True, facecolor='none')
    buf.seek(0)
    plt.close(fig)

    return buf


def main():
    """CLI entry point for chart generation."""
    if len(sys.argv) < 2:
        print("Usage: chart_generator.py '<chart_json>'")
        sys.exit(1)

    import json
    from pptx.dml.color import RGBColor

    raw = sys.argv[1]
    try:
        chart_spec = json.loads(raw)
    except json.JSONDecodeError as e:
        print(f"ERROR: Invalid JSON: {e}", file=sys.stderr)
        sys.exit(1)

    # Default palette
    palette = {
        "primary": RGBColor(0x1B, 0x4F, 0x72),
        "secondary": RGBColor(0x21, 0x8F, 0xBE),
        "accent": RGBColor(0xE6, 0x7E, 0x22),
    }

    buf = create_chart(chart_spec, palette)
    # Return buffer info - caller writes to file
    print(f"CHART Generated: {len(buf.getvalue())} bytes")


if __name__ == "__main__":
    main()
