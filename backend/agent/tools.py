"""
RightCut — Tool executor.
Dispatches Gemini function calls to actual implementations.
"""

from __future__ import annotations

import logging
from typing import Any

from excel.engine import WorkbookEngine
from models import ParsedDocument, ToolResult
from parsers.csv_parser import parse_csv
from parsers.docx_parser import parse_docx
from parsers.pdf_parser import parse_pdf

logger = logging.getLogger(__name__)


class ToolExecutor:
    """Routes tool_name → implementation, always returning a ToolResult."""

    def __init__(
        self,
        engine: WorkbookEngine,
        documents: dict[str, ParsedDocument],
    ) -> None:
        self.engine = engine
        self.documents = documents

    async def execute(self, tool_name: str, args: dict[str, Any]) -> ToolResult:
        handler = getattr(self, f"_{tool_name}", None)
        if handler is None:
            return ToolResult(
                summary=f"Unknown tool: {tool_name}",
                data={"error": f"Tool '{tool_name}' not found"},
                success=False,
            )
        try:
            return await handler(**args)
        except TypeError as e:
            logger.error(f"Tool {tool_name} called with bad args {args}: {e}")
            return ToolResult(
                summary=f"Bad arguments for {tool_name}: {e}",
                data={"error": str(e)},
                success=False,
            )
        except Exception as e:
            logger.exception(f"Tool {tool_name} failed: {e}")
            return ToolResult(
                summary=f"Tool {tool_name} failed: {str(e)[:200]}",
                data={"error": str(e)},
                success=False,
            )

    # ── Tool implementations ──────────────────────────────────────────────────

    async def _parse_document(
        self,
        file_id: str,
        extract_tables: bool = True,
    ) -> ToolResult:
        doc = self.documents.get(file_id)
        if not doc:
            return ToolResult(
                summary=f"File '{file_id}' not found. Was it uploaded?",
                data={"error": "file_not_found"},
                success=False,
            )

        table_count = len(doc.tables)
        content_len = len(doc.content)

        return ToolResult(
            summary=(
                f"Parsed '{doc.filename}': {content_len:,} characters, "
                f"{table_count} table(s) extracted"
                + (f", {doc.page_count} pages" if doc.page_count else "")
            ),
            data={
                "filename": doc.filename,
                "file_type": doc.file_type,
                "content": doc.content[:10_000],  # cap to stay in context window
                "tables": [
                    {"rows": t[:20], "total_rows": len(t)}
                    for t in (doc.tables if extract_tables else [])
                ],
                "page_count": doc.page_count,
                "truncated": content_len > 10_000,
            },
        )

    async def _create_sheet(
        self,
        sheet_name: str,
        headers: list[str],
    ) -> ToolResult:
        result = self.engine.create_sheet(sheet_name, headers)
        return ToolResult(
            summary=f"Created sheet '{sheet_name}' with {len(headers)} columns: {', '.join(headers[:5])}{'...' if len(headers) > 5 else ''}",
            data=result,
        )

    async def _insert_data(
        self,
        sheet_name: str,
        rows: list[list[str]],
        start_row: int = 2,
    ) -> ToolResult:
        # Coerce all values to strings (Gemini may pass ints/floats)
        str_rows = [[str(v) if v is not None else "" for v in row] for row in rows]
        result = self.engine.insert_data(sheet_name, str_rows, start_row)
        rows_written = result.get("rows_written", len(rows))
        return ToolResult(
            summary=f"Inserted {rows_written} row(s) into '{sheet_name}' starting at row {start_row}",
            data=result,
        )

    async def _add_formula(
        self,
        sheet_name: str,
        cell: str,
        formula: str,
        apply_to_range: str | None = None,
    ) -> ToolResult:
        result = self.engine.add_formula(sheet_name, cell, formula, apply_to_range)
        if not result.get("success", True):
            return ToolResult(
                summary=f"Formula error in '{sheet_name}'{cell}: {result.get('error')}",
                data=result,
                success=False,
            )
        cells = result.get("cells", [cell])
        return ToolResult(
            summary=f"Added formula '{formula}' to {len(cells)} cell(s) in '{sheet_name}'",
            data=result,
        )

    async def _edit_cell(
        self,
        sheet_name: str,
        cell: str,
        value: str,
    ) -> ToolResult:
        result = self.engine.edit_cell(sheet_name, cell, value)
        return ToolResult(
            summary=f"Updated {sheet_name}!{cell}: '{result['old']}' → '{result['new']}'",
            data=result,
        )

    async def _apply_formatting(
        self,
        sheet_name: str,
        cell_range: str,
        format_type: str,
        format_config: dict | None = None,
    ) -> ToolResult:
        result = self.engine.apply_formatting(sheet_name, cell_range, format_type, format_config)
        if not result.get("success", True):
            return ToolResult(
                summary=f"Formatting error: {result.get('error')}",
                data=result,
                success=False,
            )
        return ToolResult(
            summary=f"Applied '{format_type}' formatting to {sheet_name}!{cell_range}",
            data=result,
        )

    async def _add_citation(
        self,
        sheet_name: str,
        cell: str,
        source_file: str,
        source_location: str,
        excerpt: str = "",
    ) -> ToolResult:
        result = self.engine.add_citation(sheet_name, cell, source_file, source_location, excerpt)
        return ToolResult(
            summary=f"Citation added to {sheet_name}!{cell} → {source_file} ({source_location})",
            data=result,
        )

    async def _sort_range(
        self,
        sheet_name: str,
        sort_column: str,
        ascending: bool = True,
    ) -> ToolResult:
        result = self.engine.sort_range(sheet_name, sort_column, ascending)
        if not result.get("success", True):
            return ToolResult(
                summary=f"Sort failed: {result.get('error')}",
                data=result,
                success=False,
            )
        direction = "ascending" if ascending else "descending"
        return ToolResult(
            summary=f"Sorted '{sheet_name}' by '{sort_column}' {direction} ({result.get('rows_sorted', '?')} rows)",
            data=result,
        )

    async def _create_chart(
        self,
        sheet_name: str,
        chart_type: str,
        data_range: str,
        title: str = "",
        target_cell: str = "H2",
    ) -> ToolResult:
        result = self.engine.create_chart(sheet_name, chart_type, data_range, title, target_cell)
        if not result.get("success", True):
            return ToolResult(
                summary=f"Chart creation failed: {result.get('error')}",
                data=result,
                success=False,
            )
        return ToolResult(
            summary=f"Created {chart_type} chart '{title or sheet_name}' in '{sheet_name}' at {target_cell} (visible in .xlsx download)",
            data=result,
        )

    async def _create_model_scaffold(
        self,
        model_type: str,
        params: dict | None = None,
    ) -> ToolResult:
        result = self.engine.create_model_scaffold(model_type, params or {})
        if not result.get("success", True):
            return ToolResult(
                summary=f"Scaffold error: {result.get('error')}",
                data=result,
                success=False,
            )
        sheets = result.get("sheets_created", [])
        return ToolResult(
            summary=f"Built complete {model_type.upper()} scaffold: {', '.join(sheets)}. Formulas are circular-reference-free and correctly cross-referenced.",
            data=result,
        )

    async def _validate_workbook(self, check_hardcoded: bool = True) -> ToolResult:
        result = self.engine.validate_workbook(check_hardcoded)
        status = "PASS" if result.get("valid") else "WARNINGS"
        return ToolResult(
            summary=f"Validation [{status}]: {result.get('summary', '')}",
            data=result,
        )

    async def _get_sheet_state(self, sheet_name: str) -> ToolResult:
        result = self.engine.get_sheet_state(sheet_name)
        row_count = result.get("row_count", 0)
        headers = result.get("headers", [])
        return ToolResult(
            summary=f"Read '{sheet_name}': {row_count} rows, columns: {', '.join(headers[:6])}{'...' if len(headers) > 6 else ''}",
            data=result,
        )
