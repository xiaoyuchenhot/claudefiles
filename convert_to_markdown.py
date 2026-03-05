#!/usr/bin/env python3
"""
Convert files and URLs to Markdown using Microsoft's MarkItDown.
Supports: PDF, DOCX, PPTX, XLSX, images, HTML, CSV, JSON, XML, ZIP, YouTube URLs, EPub

For URLs that require browser-like headers, use --browser flag.
"""

import sys
import os
import tempfile
import requests
from markitdown import MarkItDown

BROWSER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/121.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}


def convert_url_with_browser_headers(url: str) -> str:
    """Fetch URL with browser headers then convert HTML to markdown."""
    r = requests.get(url, headers=BROWSER_HEADERS, timeout=30, allow_redirects=True)
    r.raise_for_status()
    with tempfile.NamedTemporaryFile(suffix=".html", delete=False, mode="wb") as f:
        f.write(r.content)
        tmp = f.name
    try:
        md = MarkItDown()
        result = md.convert(tmp)
        return result.text_content
    finally:
        os.unlink(tmp)


def convert(source: str, output: str = None, browser: bool = False):
    if browser and source.startswith("http"):
        text = convert_url_with_browser_headers(source)
    else:
        md = MarkItDown()
        result = md.convert(source)
        text = result.text_content

    if output:
        with open(output, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"Saved to {output}")
    else:
        print(text)


if __name__ == "__main__":
    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print("Usage: python convert_to_markdown.py [--browser] <file_or_url> [output.md]")
        print("  --browser  Fetch URL with browser User-Agent headers")
        sys.exit(0 if args else 1)

    browser = "--browser" in args
    args = [a for a in args if a != "--browser"]

    source = args[0]
    output = args[1] if len(args) > 1 else None
    convert(source, output, browser=browser)
