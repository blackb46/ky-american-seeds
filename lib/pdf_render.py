"""Render PDF pages to base64 PNGs for the Anthropic vision API.

Uses PyMuPDF — pure Python, no system dependencies (works on Streamlit Cloud
without poppler).
"""
from __future__ import annotations
import base64
from io import BytesIO
from pathlib import Path

import fitz  # PyMuPDF


def render_pdf_pages(
    source: str | Path | bytes,
    dpi: int = 150,
    max_pages: int = 12,
    max_dimension: int = 1568,
) -> list[dict]:
    """Render every page of a PDF to PNG bytes.

    Returns a list of dicts shaped for Anthropic's vision API:
        {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": "..."}}

    `max_dimension` clamps the longest side. Anthropic recommends ≤1568px for
    best vision performance and lower token cost.
    """
    if isinstance(source, (str, Path)):
        doc = fitz.open(source)
    else:
        doc = fitz.open(stream=source, filetype="pdf")

    blocks: list[dict] = []
    try:
        zoom = dpi / 72.0
        for i, page in enumerate(doc):
            if i >= max_pages:
                break
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)

            # Downscale if needed (in pixels, not pts).
            if pix.width > max_dimension or pix.height > max_dimension:
                shrink = max_dimension / max(pix.width, pix.height)
                mat = fitz.Matrix(zoom * shrink, zoom * shrink)
                pix = page.get_pixmap(matrix=mat, alpha=False)

            png_bytes = pix.tobytes("png")
            b64 = base64.standard_b64encode(png_bytes).decode("ascii")
            blocks.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": b64,
                },
            })
    finally:
        doc.close()
    return blocks


def render_pdf_thumbnail(source: str | Path | bytes, page: int = 0, dpi: int = 80) -> bytes:
    """Return PNG bytes of a single page for UI preview thumbnails."""
    if isinstance(source, (str, Path)):
        doc = fitz.open(source)
    else:
        doc = fitz.open(stream=source, filetype="pdf")
    try:
        if page >= len(doc):
            page = 0
        zoom = dpi / 72.0
        pix = doc[page].get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        return pix.tobytes("png")
    finally:
        doc.close()
