#!/usr/bin/env python3
"""Generate and analyze images using Gemini 3.1 Flash Image model via Zenmux proxy.

Usage:
    python3 gemini_image_gen.py generate '{"prompt": "...", "output_path": "...", "style": "medical"}'
    python3 gemini_image_gen.py generate  (reads JSON from stdin)
    python3 gemini_image_gen.py analyze /path/to/image.png "What is shown?"
"""

import json
import os
import sys
import mimetypes
from pathlib import Path

from google import genai
from google.genai import types


def get_client() -> genai.Client:
    api_key = os.environ.get("ZENMUX_API_KEY", "")
    if not api_key:
        # Try reading from secrets file written by agent runner
        try:
            with open("/tmp/.zenmux_key", "r") as f:
                api_key = f.read().strip()
        except FileNotFoundError:
            pass
    if not api_key:
        raise RuntimeError("ZENMUX_API_KEY not available (env or /tmp/.zenmux_key)")
    return genai.Client(
        api_key=api_key,
        vertexai=True,
        http_options=types.HttpOptions(
            api_version="v1",
            base_url="https://zenmux.ai/api/vertex-ai",
        ),
    )


MODEL = "google/gemini-3-pro-image-preview"

STYLE_PREFIXES = {
    "medical": (
        "Create a clean, professional medical illustration. "
        "Use precise anatomical detail with soft, clinical colors. "
    ),
    "diagram": (
        "Create a clear, well-labeled technical diagram. "
        "Use clean lines, distinct shapes, and a structured layout. "
    ),
    "infographic": (
        "Create a modern, visually engaging infographic. "
        "Use bold colors, icons, and clear data hierarchy. "
    ),
    "comparison": (
        "Create a side-by-side comparison illustration. "
        "Use consistent framing and clear visual contrast between the two sides. "
    ),
}


def build_prompt(raw_prompt: str, style: str | None = None) -> str:
    prefix = STYLE_PREFIXES.get(style or "", "")
    return f"{prefix}{raw_prompt}"


def generate_image(params: dict) -> str:
    """Generate an image from a text prompt and save it to disk.

    Args:
        params: dict with keys ``prompt``, ``output_path``, and optional ``style``.

    Returns:
        The absolute path of the saved image.
    """
    raw_prompt = params.get("prompt")
    output_path = params.get("output_path")
    style = params.get("style")

    if not raw_prompt:
        raise ValueError("Missing required field: prompt")
    if not output_path:
        raise ValueError("Missing required field: output_path")

    prompt = build_prompt(raw_prompt, style)
    client = get_client()

    response = client.models.generate_content(
        model=MODEL,
        contents=[prompt],
        config=types.GenerateContentConfig(
            response_modalities=["TEXT", "IMAGE"],
        ),
    )

    # Walk through response parts and find the first inline image.
    image_saved = False
    for part in response.candidates[0].content.parts:
        if part.inline_data is not None:
            out = Path(output_path)
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_bytes(part.inline_data.data)
            image_saved = True
            break

    if not image_saved:
        raise RuntimeError(
            "Gemini did not return an image. Model response text: "
            + (response.text or "(empty)")
        )

    return str(Path(output_path).resolve())


def analyze_image(
    image_path: str,
    question: str = "Describe this image in detail",
) -> str:
    """Use Gemini Vision to analyze an existing image.

    Args:
        image_path: Path to a local image file.
        question: The question or instruction to send alongside the image.

    Returns:
        The model's textual response.
    """
    path = Path(image_path)
    if not path.is_file():
        raise FileNotFoundError(f"Image not found: {image_path}")

    mime_type, _ = mimetypes.guess_type(str(path))
    if mime_type is None:
        mime_type = "image/png"

    image_data = path.read_bytes()
    client = get_client()

    response = client.models.generate_content(
        model=MODEL,
        contents=[
            types.Part.from_bytes(data=image_data, mime_type=mime_type),
            question,
        ],
    )
    return response.text


def _parse_generate_input(args: list[str]) -> dict:
    """Parse JSON input from argv or stdin."""
    if len(args) >= 1 and args[0].strip().startswith("{"):
        return json.loads(args[0])
    # Fall back to stdin
    data = sys.stdin.read().strip()
    if not data:
        raise ValueError("No JSON input provided via argument or stdin")
    return json.loads(data)


def main() -> None:
    if len(sys.argv) < 2:
        print(
            "Usage:\n"
            "  gemini_image_gen.py generate '{...}'\n"
            "  gemini_image_gen.py analyze <image_path> [question]",
            file=sys.stderr,
        )
        sys.exit(1)

    command = sys.argv[1]

    try:
        if command == "generate":
            params = _parse_generate_input(sys.argv[2:])
            result_path = generate_image(params)
            print(result_path)

        elif command == "analyze":
            if len(sys.argv) < 3:
                print("ERROR: analyze requires an image path", file=sys.stderr)
                sys.exit(1)
            image_path = sys.argv[2]
            question = sys.argv[3] if len(sys.argv) >= 4 else "Describe this image in detail"
            result = analyze_image(image_path, question)
            print(result)

        else:
            print(f"ERROR: Unknown command '{command}'", file=sys.stderr)
            sys.exit(1)

    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
