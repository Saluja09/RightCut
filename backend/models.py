"""
RightCut — Pydantic models and shared data types.
"""

from __future__ import annotations

import uuid
from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


# ── Spreadsheet primitives ────────────────────────────────────────────────────

class CellValue(BaseModel):
    value: str | int | float | None = None
    formula: str | None = None          # raw formula string e.g. "=SUM(A1:A3)"
    comment: str | None = None          # citation / annotation text
    number_format: str | None = None    # e.g. "$#,##0" / "0.0%"
    bold: bool = False
    bg_color: str | None = None         # hex without # e.g. "FF0000"
    font_color: str | None = None


class ChartMeta(BaseModel):
    chart_type: str                     # "bar" | "line" | "pie" | "scatter"
    title: str = ""
    data_range: str                     # e.g. "A1:B10"
    anchor_cell: str = "H2"            # top-left corner of chart placement


class SheetState(BaseModel):
    name: str
    headers: list[str] = []
    rows: list[list[CellValue]] = []
    charts: list[ChartMeta] = []
    frozen_rows: int = 1


class WorkbookState(BaseModel):
    sheets: list[SheetState] = []
    active_sheet: str = ""


# ── Document parsing ──────────────────────────────────────────────────────────

class ParsedDocument(BaseModel):
    file_id: str = Field(default_factory=lambda: str(uuid.uuid4())[:8])
    filename: str
    content: str                        # full extracted text (truncated to 50k)
    tables: list[list[list[str]]] = []  # list of tables, each is list of rows
    page_count: int | None = None
    file_type: str = "unknown"          # "pdf" | "docx" | "csv" | "xlsx"


# ── Agent tool calls ──────────────────────────────────────────────────────────

class ToolResult(BaseModel):
    summary: str                        # human-readable 1-2 sentence description
    data: dict[str, Any] = {}           # machine-readable result for Gemini
    success: bool = True
    error: str | None = None


class ToolStep(BaseModel):
    tool: str
    args: dict[str, Any] = {}
    result_summary: str
    duration_ms: int
    success: bool = True
    error: str | None = None


# ── WebSocket protocol ────────────────────────────────────────────────────────

class WSMessageType(str, Enum):
    TOOL_CALL = "tool_call"
    WORKBOOK_UPDATE = "workbook_update"
    AGENT_RESPONSE = "agent_response"
    NEW_TAB = "new_tab"
    ERROR = "error"
    STATUS = "status"
    THINKING = "thinking"
    PONG = "pong"


class TabInfo(BaseModel):
    id: str
    name: str
    type: str = "sheet"    # "sheet" | "document"


# ── Session ───────────────────────────────────────────────────────────────────

class ValidationStats(BaseModel):
    formula_count: int = 0
    hardcoded_count: int = 0
    total_cells: int = 0
    issues: list[str] = []
