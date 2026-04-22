"""
RightCut — CSV / XLSX parser using pandas.
"""

from __future__ import annotations

import io
import logging

from models import ParsedDocument

logger = logging.getLogger(__name__)

MAX_CONTENT_CHARS = 50_000
MAX_ROWS_PER_TABLE = 500


async def parse_csv(file_bytes: bytes, filename: str) -> ParsedDocument:
    """Parse CSV or XLSX into a ParsedDocument."""
    try:
        import pandas as pd  # type: ignore

        ext = filename.rsplit(".", 1)[-1].lower()

        if ext in ("xlsx", "xls"):
            return await _parse_excel(file_bytes, filename, pd)
        else:
            return await _parse_csv_file(file_bytes, filename, pd)

    except Exception as e:
        logger.error(f"CSV/XLSX parsing failed for {filename}: {e}")
        return ParsedDocument(
            filename=filename,
            content=f"[Error parsing file: {e}]",
            tables=[],
            file_type="csv",
        )


async def _parse_csv_file(file_bytes: bytes, filename: str, pd) -> ParsedDocument:
    # Try encodings in order
    df = None
    for encoding in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding=encoding)
            break
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue

    if df is None:
        return ParsedDocument(
            filename=filename,
            content="[Could not decode CSV file]",
            tables=[],
            file_type="csv",
        )

    table = _df_to_table(df)
    content = f"File: {filename}\nRows: {len(df)}, Columns: {len(df.columns)}\n\n"
    content += _table_preview(table)

    return ParsedDocument(
        filename=filename,
        content=content[:MAX_CONTENT_CHARS],
        tables=[table],
        page_count=None,
        file_type="csv",
    )


async def _parse_excel(file_bytes: bytes, filename: str, pd) -> ParsedDocument:
    sheets_dict = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, engine="openpyxl")

    all_tables: list[list[list[str]]] = []
    content_parts: list[str] = []

    for sheet_name, df in sheets_dict.items():
        table = _df_to_table(df)
        all_tables.append(table)
        content_parts.append(
            f"Sheet: {sheet_name} ({len(df)} rows × {len(df.columns)} cols)\n"
            + _table_preview(table)
        )

    content = "\n\n".join(content_parts)[:MAX_CONTENT_CHARS]

    return ParsedDocument(
        filename=filename,
        content=content,
        tables=all_tables,
        page_count=None,
        file_type="xlsx",
    )


def _df_to_table(df) -> list[list[str]]:
    """Convert a DataFrame to a list of rows (first row = headers)."""
    rows: list[list[str]] = []

    # Header row
    headers = [str(col) for col in df.columns]
    rows.append(headers)

    # Data rows (capped)
    for _, row in df.head(MAX_ROWS_PER_TABLE).iterrows():
        rows.append([_format_cell(v) for v in row.values])

    return rows


def _format_cell(value) -> str:
    import math
    if value is None:
        return ""
    try:
        import pandas as pd
        if pd.isna(value):
            return ""
    except Exception:
        pass
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return ""
        # Format nicely: no trailing zeros for integers
        if value == int(value):
            return str(int(value))
        return f"{value:.4g}"
    return str(value).strip()


def _table_preview(table: list[list[str]], max_rows: int = 10) -> str:
    if not table:
        return ""
    lines = []
    for row in table[:max_rows]:
        lines.append(" | ".join(row))
    if len(table) > max_rows:
        lines.append(f"... ({len(table) - max_rows} more rows)")
    return "\n".join(lines)
