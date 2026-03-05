#!/usr/bin/env python3
"""
Convert files to Markdown using Microsoft's MarkItDown.
Supports: PDF, DOCX, PPTX, XLSX, images, HTML, CSV, JSON, XML, ZIP, YouTube URLs, EPub
"""

import sys
from markitdown import MarkItDown


def convert(source: str, output: str = None):
    md = MarkItDown()
    result = md.convert(source)
    if output:
        with open(output, "w", encoding="utf-8") as f:
            f.write(result.text_content)
        print(f"Saved to {output}")
    else:
        print(result.text_content)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python convert_to_markdown.py <file_or_url> [output.md]")
        sys.exit(1)
    source = sys.argv[1]
    output = sys.argv[2] if len(sys.argv) > 2 else None
    convert(source, output)
