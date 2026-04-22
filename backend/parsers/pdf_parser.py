"""
RightCut — PDF parser.
Uses pdfplumber (preferred) with pypdf fallback.
"""

from __future__ import annotations

import io
import logging
import re
from typing import Any

from models import ParsedDocument

logger = logging.getLogger(__name__)

MAX_CONTENT_CHARS = 50_000


async def parse_pdf(file_bytes: bytes, filename: str) -> ParsedDocument:
    """Extract text and tables from a PDF file."""
    content_parts: list[str] = []
    tables: list[list[list[str]]] = []
    page_count = 0

    # ── Try pdfplumber first (better table extraction) ────────────────────────
    try:
        import pdfplumber  # type: ignore

        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            page_count = len(pdf.pages)
            for page_num, page in enumerate(pdf.pages, start=1):
                # Extract text
                text = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
                if text.strip():
                    content_parts.append(f"[Page {page_num}]\n{text}")

                # Extract tables
                page_tables = page.extract_tables() or []
                for tbl in page_tables:
                    clean = [
                        [str(cell).strip() if cell else "" for cell in row]
                        for row in tbl
                        if any(cell for cell in row)
                    ]
                    if clean:
                        tables.append(clean)

        content = "\n\n".join(content_parts)[:MAX_CONTENT_CHARS]
        return ParsedDocument(
            filename=filename,
            content=content,
            tables=tables,
            page_count=page_count,
            file_type="pdf",
        )

    except ImportError:
        logger.info("pdfplumber not available, falling back to pypdf")
    except Exception as e:
        logger.warning(f"pdfplumber failed for {filename}: {e}, falling back to pypdf")

    # ── pypdf fallback ────────────────────────────────────────────────────────
    try:
        import pypdf  # type: ignore

        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        page_count = len(reader.pages)

        for page_num, page in enumerate(reader.pages, start=1):
            try:
                text = page.extract_text(extraction_mode="layout") or ""
            except TypeError:
                # older pypdf versions don't support extraction_mode
                text = page.extract_text() or ""

            if text.strip():
                content_parts.append(f"[Page {page_num}]\n{text}")
                # Heuristic table extraction: lines with 2+ consecutive spaces
                tbl = _extract_table_from_layout_text(text)
                if tbl:
                    tables.extend(tbl)

        content = "\n\n".join(content_parts)[:MAX_CONTENT_CHARS]
        return ParsedDocument(
            filename=filename,
            content=content,
            tables=tables,
            page_count=page_count,
            file_type="pdf",
        )

    except Exception as e:
        logger.error(f"pypdf also failed for {filename}: {e}")
        return ParsedDocument(
            filename=filename,
            content=f"[Error parsing PDF: {e}]",
            tables=[],
            page_count=0,
            file_type="pdf",
        )


def _extract_table_from_layout_text(text: str) -> list[list[list[str]]]:
    """
    Heuristic: detect tabular sections in layout-mode text.
    Returns a list of tables (each table is list of rows).
    """
    tables: list[list[list[str]]] = []
    current_table: list[list[str]] = []

    for line in text.split("\n"):
        # A line looks tabular if it has 2+ consecutive spaces (column separators)
        if re.search(r"  +", line.strip()):
            cells = re.split(r"  +", line.strip())
            cells = [c.strip() for c in cells if c.strip()]
            if len(cells) >= 2:
                current_table.append(cells)
            else:
                if len(current_table) >= 2:
                    tables.append(current_table)
                current_table = []
        else:
            if len(current_table) >= 2:
                tables.append(current_table)
            current_table = []

    if len(current_table) >= 2:
        tables.append(current_table)

    return tables
