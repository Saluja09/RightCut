"""
RightCut — FastAPI application entry point.
Handles WebSocket connections, file uploads, and workbook downloads.
"""

from __future__ import annotations

import asyncio
import datetime
import io
import logging
import os
import time
import uuid
from dataclasses import dataclass, field
from typing import Any

from fastapi import FastAPI, File, HTTPException, UploadFile, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from google.genai import types
from pydantic import BaseModel

from agent.orchestrator import AgentOrchestrator
from agent.prompts import CELL_EDIT_CONTEXT_TEMPLATE, DOCUMENT_UPLOAD_CONTEXT_TEMPLATE, get_system_prompt
from agent.tools import ToolExecutor
from config import CORS_ORIGINS, CORS_ORIGIN_REGEX, GEMINI_API_KEY, GEMINI_MODEL, MAX_TOOL_ITERATIONS, SESSION_TTL_HOURS, UPLOAD_MAX_SIZE_MB
from excel.engine import WorkbookEngine
from models import ParsedDocument

# ── Logging ───────────────────────────────────────────────────────────────────
LOG_LEVEL = os.environ.get("LOG_LEVEL", "INFO").upper()

logging.basicConfig(
    level=getattr(logging, LOG_LEVEL, logging.INFO),
    format="%(asctime)s │ %(levelname)-5s │ %(name)-20s │ %(message)s",
    datefmt="%H:%M:%S",
)
# Quiet noisy libraries
logging.getLogger("httpcore").setLevel(logging.WARNING)
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("urllib3").setLevel(logging.WARNING)
logging.getLogger("websockets").setLevel(logging.WARNING)
logging.getLogger("formulas").setLevel(logging.WARNING)
logging.getLogger("schedula").setLevel(logging.WARNING)

logger = logging.getLogger("rightcut")

# ── Application ───────────────────────────────────────────────────────────────
app = FastAPI(
    title="RightCut API",
    description="AI-powered spreadsheet agent for private markets",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ORIGINS,
    allow_origin_regex=CORS_ORIGIN_REGEX,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Shared state ──────────────────────────────────────────────────────────────
orchestrator = AgentOrchestrator(api_key=GEMINI_API_KEY)


@dataclass
class WorkspaceSession:
    engine: WorkbookEngine = field(default_factory=WorkbookEngine)
    conversation_history: list[types.Content] = field(default_factory=list)
    documents: dict[str, ParsedDocument] = field(default_factory=dict)
    cell_edit_buffer: list[dict] = field(default_factory=list)
    created_at: datetime.datetime = field(default_factory=datetime.datetime.utcnow)
    last_active: datetime.datetime = field(default_factory=datetime.datetime.utcnow)
    session_role: str = "general"   # "finance" | "general"


sessions: dict[str, WorkspaceSession] = {}


def get_or_create_session(session_id: str) -> WorkspaceSession:
    if session_id not in sessions:
        sessions[session_id] = WorkspaceSession()
        logger.info(f"Created session {session_id}")
    sessions[session_id].last_active = datetime.datetime.utcnow()
    return sessions[session_id]


# ── Background: session cleanup ───────────────────────────────────────────────

async def _cleanup_sessions() -> None:
    """Periodically remove expired sessions to prevent memory leaks."""
    while True:
        await asyncio.sleep(3600)  # run every hour
        cutoff = datetime.datetime.utcnow() - datetime.timedelta(hours=SESSION_TTL_HOURS)
        expired = [sid for sid, s in sessions.items() if s.last_active < cutoff]
        for sid in expired:
            del sessions[sid]
            logger.info(f"Expired session {sid}")


@app.on_event("startup")
async def startup() -> None:
    asyncio.create_task(_cleanup_sessions())
    logger.info(
        f"RightCut API started — model={GEMINI_MODEL}  "
        f"max_tools={MAX_TOOL_ITERATIONS}  session_ttl={SESSION_TTL_HOURS}h  "
        f"log_level={LOG_LEVEL}"
    )
    # Validate formulas library is available for xlsx export
    try:
        import formulas  # noqa: F401
        logger.info("Formula evaluation: formulas library available ✓")
    except ImportError:
        logger.warning("Formula evaluation: formulas library NOT installed — xlsx will lack cached values")


# ── Health check ──────────────────────────────────────────────────────────────

@app.get("/health")
async def health() -> dict:
    return {"status": "ok", "sessions": len(sessions)}


# ── Session restore ───────────────────────────────────────────────────────────

class RestoreMessageItem(BaseModel):
    role: str
    text: str | None = None
    timestamp: int | None = None


class RestoreRequest(BaseModel):
    workbook_state: dict | None = None   # WorkbookState JSON from frontend
    messages: list[RestoreMessageItem] = []
    role: str = "general"                # "finance" | "general" — persists agent persona


@app.post("/restore/{session_id}")
async def restore_session(session_id: str, body: RestoreRequest) -> dict:
    """
    Restore a backend session from a workbook snapshot + message history.
    Called by the frontend when switching to a previously saved session.
    Rebuilds the WorkbookEngine from the serialised WorkbookState so the
    agent operates on the correct workbook when the user sends the next message.
    """
    session = get_or_create_session(session_id)

    # ── 0. Restore role ────────────────────────────────────────────────────
    valid_roles = {"finance", "general"}
    session.session_role = body.role if body.role in valid_roles else "general"

    # ── 1. Rebuild WorkbookEngine from snapshot ────────────────────────────
    sheets_restored = 0
    if body.workbook_state:
        engine = WorkbookEngine()
        wb_sheets = body.workbook_state.get("sheets") or []

        for sheet_data in wb_sheets:
            sheet_name = sheet_data.get("name", "Sheet")
            headers = sheet_data.get("headers") or []
            rows_data = sheet_data.get("rows") or []

            # Create sheet with headers
            engine.create_sheet(sheet_name, headers)

            # Collect plain-value rows and formula cells separately
            plain_rows: list[list[str]] = []
            formula_cells: list[tuple[int, int, str]] = []  # (row_idx, col_idx, formula)

            for row_idx, row in enumerate(rows_data):
                plain_row: list[str] = []
                for col_idx, cell in enumerate(row):
                    if not cell:
                        plain_row.append("")
                        continue
                    formula = cell.get("formula")
                    if formula:
                        # Placeholder so insert_data keeps row count correct
                        plain_row.append("")
                        formula_cells.append((row_idx + 2, col_idx + 1, formula))  # +2: 1-indexed, skip header
                    else:
                        val = cell.get("value")
                        plain_row.append("" if val is None else str(val))
                plain_rows.append(plain_row)

            # Insert non-formula cell values
            if plain_rows:
                engine.insert_data(sheet_name, plain_rows, start_row=2)

            # Write formulas directly via openpyxl (bypasses insert_data string coercion)
            if formula_cells and sheet_name in engine.wb.sheetnames:
                ws = engine.wb[sheet_name]
                for r, c, formula in formula_cells:
                    ws.cell(row=r, column=c).value = formula

            # Restore charts — re-create openpyxl chart objects and repopulate _charts
            for chart_data in (sheet_data.get("charts") or []):
                try:
                    engine.create_chart(
                        sheet_name=sheet_name,
                        chart_type=chart_data.get("chart_type", "bar"),
                        data_range=chart_data.get("data_range", "A1:B2"),
                        title=chart_data.get("title", ""),
                        target_cell=chart_data.get("anchor_cell", "H2"),
                    )
                except Exception as chart_err:
                    logger.warning(f"Restore: could not recreate chart in '{sheet_name}': {chart_err}")

            sheets_restored += 1

        session.engine = engine
        logger.info(
            f"Restore session {session_id}: rebuilt {sheets_restored} sheet(s)"
        )

    # ── 2. Rebuild conversation_history from saved messages ────────────────
    if body.messages:
        history: list[types.Content] = []
        for msg in body.messages:
            role = msg.role
            text = (msg.text or "").strip()
            if not text:
                continue
            # Map frontend roles to Gemini roles
            if role == "user":
                history.append(
                    types.Content(role="user", parts=[types.Part.from_text(text=text)])
                )
            elif role in ("agent", "assistant"):
                history.append(
                    types.Content(role="model", parts=[types.Part.from_text(text=text)])
                )
            # Skip system/error/agent_pending messages — not relevant to Gemini context

        # ── Inject a workbook-state anchor at the END of restored history ──
        # This prevents the model from treating short follow-ups ("redo", "again")
        # as continuations of an old financial model task from prior turns.
        # The anchor describes the CURRENT workbook and acts as ground truth.
        if history and session.engine:
            sheet_names = session.engine.get_all_sheet_names()
            if sheet_names:
                sheets_summary = ", ".join(f'"{s}"' for s in sheet_names)
                anchor = (
                    f"[SESSION CONTEXT: The current workbook contains {len(sheet_names)} sheet(s): "
                    f"{sheets_summary}. "
                    "When the user's next message is a short follow-up (e.g. 'redo', 'do it again', "
                    "'try again', 'again', 'repeat'), interpret it as a request to redo the MOST RECENT "
                    "completed task visible in the current workbook — NOT as a continuation of any "
                    "financial modelling task from earlier in this conversation. "
                    "Always base your interpretation of short commands on the current workbook state.]"
                )
                history.append(
                    types.Content(role="user", parts=[types.Part.from_text(text=anchor)])
                )
                # Add a model acknowledgement so the conversation alternates correctly
                history.append(
                    types.Content(
                        role="model",
                        parts=[types.Part.from_text(
                            text=f"Understood. I can see the current workbook has {len(sheet_names)} sheet(s): "
                                 f"{sheets_summary}. I'll interpret follow-up commands relative to the "
                                 "current workbook state."
                        )],
                    )
                )

        session.conversation_history = history
        logger.info(
            f"Restore session {session_id}: rebuilt {len(history)} conversation turns"
        )

    return {
        "ok": True,
        "session_id": session_id,
        "sheets_restored": sheets_restored,
        "history_turns": len(session.conversation_history),
    }


# ── Session configuration ─────────────────────────────────────────────────────

class ConfigureRequest(BaseModel):
    role: str = "general"   # "finance" | "general"


@app.post("/configure/{session_id}")
async def configure_session(session_id: str, body: ConfigureRequest) -> dict:
    """Set session-level configuration such as the agent role/persona."""
    valid_roles = {"finance", "general"}
    role = body.role if body.role in valid_roles else "general"
    session = get_or_create_session(session_id)
    session.session_role = role
    logger.info(f"Session {session_id}: role set to '{role}'")
    return {"ok": True, "session_id": session_id, "role": role}


# ── File upload ───────────────────────────────────────────────────────────────

@app.post("/upload/{session_id}")
async def upload_file(session_id: str, file: UploadFile = File(...)) -> dict:
    """
    Upload a document (PDF, DOCX, CSV, XLSX) to a session.
    Returns file_id for use with the parse_document tool.
    """
    file_bytes = await file.read()
    size_mb = len(file_bytes) / (1024 * 1024)

    if size_mb > UPLOAD_MAX_SIZE_MB:
        raise HTTPException(
            status_code=413,
            detail=f"File too large ({size_mb:.1f} MB). Limit is {UPLOAD_MAX_SIZE_MB} MB.",
        )

    filename = file.filename or "unknown"
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    file_id = str(uuid.uuid4())[:8]

    # Parse based on extension
    try:
        if ext == "pdf":
            from parsers.pdf_parser import parse_pdf
            parsed = await parse_pdf(file_bytes, filename)
        elif ext in ("docx", "doc"):
            from parsers.docx_parser import parse_docx
            parsed = await parse_docx(file_bytes, filename)
        elif ext in ("csv", "xlsx", "xls"):
            from parsers.csv_parser import parse_csv
            parsed = await parse_csv(file_bytes, filename)
        else:
            raise HTTPException(status_code=415, detail=f"Unsupported file type: .{ext}")
    except HTTPException:
        raise
    except Exception as e:
        logger.exception(f"Parsing failed for {filename}: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to parse file: {e}")

    parsed.file_id = file_id
    session = get_or_create_session(session_id)
    session.documents[file_id] = parsed

    logger.info(f"Session {session_id}: uploaded {filename} → file_id={file_id}")

    return {
        "file_id": file_id,
        "filename": filename,
        "file_type": ext,
        "page_count": parsed.page_count,
        "table_count": len(parsed.tables),
        "size_mb": round(size_mb, 2),
    }


# ── Chat summary download ─────────────────────────────────────────────────────

@app.get("/summary/{session_id}")
async def download_summary(session_id: str, format: str = "md") -> StreamingResponse:
    """
    Generate and download a structured summary of the session's conversation.
    Includes key decisions, assumptions, figures, and model outputs.
    format: 'md' (default) or 'txt'
    """
    session = sessions.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")

    history = session.conversation_history
    if not history:
        raise HTTPException(status_code=404, detail="No conversation history yet")

    # Build transcript for Gemini to summarise
    lines: list[str] = []
    for turn in history:
        role = (turn.role or "unknown").upper()
        text_parts = [
            p.text for p in (turn.parts or [])
            if hasattr(p, "text") and p.text and not p.text.startswith("[CONTEXT:")
        ]
        if text_parts:
            lines.append(f"{role}: {''.join(text_parts)[:800]}")

    transcript = "\n\n".join(lines)

    # Also include current workbook structure for context
    sheet_names = session.engine.get_all_sheet_names()
    workbook_ctx = f"Sheets in workbook: {', '.join(sheet_names)}" if sheet_names else "No workbook built."

    prompt = f"""You are preparing a professional session summary for a financial analyst.

Based on the conversation below, produce a structured Markdown document with these sections:

# Session Summary

## Overview
One paragraph describing what was accomplished in this session.

## Model Built
What type of model was built (DCF, LBO, comps, etc.), for which company, and the key structure.

## Key Assumptions
A table or bullet list of all numerical assumptions used (WACC, growth rates, margins, multiples, etc.) with their values.

## Key Outputs & Figures
The critical outputs — intrinsic value per share, enterprise value, EBITDA, IRR, MOIC, etc. — in a table or bullet list.

## Decisions & Changes
Any notable decisions made during the session — model structure choices, assumption revisions, data sources used.

## Uploaded Documents
List any documents the user uploaded and how they were used.

## Caveats & Assumptions to Review
Any assumptions that were defaulted or that the analyst should validate.

---
{workbook_ctx}

CONVERSATION TRANSCRIPT:
{transcript}

Write only the Markdown document. Be factual and concise. Use real numbers from the conversation."""

    try:
        from google import genai as _genai
        from google.genai import types as _types
        client = _genai.Client(api_key=GEMINI_API_KEY)
        resp = await client.aio.models.generate_content(
            model=GEMINI_MODEL,
            contents=[_types.Content(role="user", parts=[_types.Part.from_text(text=prompt)])],
            config=_types.GenerateContentConfig(temperature=0.1, max_output_tokens=2048),
        )
        cands = resp.candidates or []
        cparts = (cands[0].content.parts if cands and cands[0].content else None) or []
        md_content = (cparts[0].text or "").strip() if cparts else ""
        if not md_content:
            raise ValueError("Model returned an empty summary")
    except Exception as e:
        logger.exception(f"Summary generation failed for session {session_id}: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to generate summary: {e}")

    if format == "txt":
        # Strip markdown syntax for plain text
        import re
        txt = re.sub(r"#{1,6}\s*", "", md_content)
        txt = re.sub(r"\*\*(.+?)\*\*", r"\1", txt)
        txt = re.sub(r"\*(.+?)\*", r"\1", txt)
        content_bytes = txt.encode("utf-8")
        media_type = "text/plain"
        filename = "rightcut_summary.txt"
    else:
        content_bytes = md_content.encode("utf-8")
        media_type = "text/markdown"
        filename = "rightcut_summary.md"

    return StreamingResponse(
        io.BytesIO(content_bytes),
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ── Workbook download ─────────────────────────────────────────────────────────

@app.get("/download/{session_id}")
async def download_workbook(session_id: str) -> StreamingResponse:
    """Download the current workbook as a .xlsx file."""
    session = sessions.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")

    if not session.engine.get_all_sheet_names():
        raise HTTPException(status_code=404, detail="No sheets in workbook yet")

    try:
        t0 = time.perf_counter()
        xlsx_bytes = session.engine.to_bytes()
        elapsed = round((time.perf_counter() - t0) * 1000)
        sheets = session.engine.get_all_sheet_names()
        logger.info(
            f"[{session_id[:8]}] download: {len(xlsx_bytes):,} bytes, "
            f"{len(sheets)} sheet(s), formula eval {elapsed}ms"
        )
    except Exception as e:
        logger.exception(f"[{session_id[:8]}] download FAILED: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to generate workbook: {e}")

    return StreamingResponse(
        io.BytesIO(xlsx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="rightcut_analysis.xlsx"'},
    )


# ── WebSocket ─────────────────────────────────────────────────────────────────

@app.websocket("/ws/{session_id}")
async def websocket_endpoint(websocket: WebSocket, session_id: str) -> None:
    await websocket.accept()
    session = get_or_create_session(session_id)
    logger.info(f"WebSocket connected: session={session_id}")

    has_workbook = bool(session.engine.get_all_sheet_names())

    # Tell the frontend whether the backend session already has a workbook.
    # If not, the frontend will call POST /restore/{session_id} to rebuild it.
    await websocket.send_json({
        "type": "session_ready",
        "session_id": session_id,
        "has_workbook": has_workbook,
    })

    # Send initial workbook state if sheets exist
    if has_workbook:
        from excel.serializer import serialize_workbook
        wb_state = serialize_workbook(session.engine.wb, session.engine._charts)
        await websocket.send_json({"type": "workbook_update", "state": wb_state})

    try:
        while True:
            data = await websocket.receive_json()
            msg_type = data.get("type")
            session.last_active = datetime.datetime.utcnow()

            if msg_type == "ping":
                await websocket.send_json({"type": "pong"})

            elif msg_type == "user_message":
                text = (data.get("text") or "").strip()
                if not text:
                    continue

                msg_preview = text[:80] + ("..." if len(text) > 80 else "")
                logger.info(f"[{session_id[:8]}] ← user_message: {msg_preview}")

                # Inject any buffered cell edits as context
                edits_flushed = len(session.cell_edit_buffer)
                _flush_cell_edit_buffer(session)
                if edits_flushed:
                    logger.info(f"[{session_id[:8]}]   flushed {edits_flushed} buffered cell edit(s)")

                # Inject document upload context messages for any new files
                uploaded_files: list[dict] = data.get("files", [])
                for f in uploaded_files:
                    fid = f.get("file_id")
                    doc = session.documents.get(fid)
                    if doc:
                        logger.info(f"[{session_id[:8]}]   attaching doc context: {doc.filename} (id={fid})")
                        ctx = DOCUMENT_UPLOAD_CONTEXT_TEMPLATE.format(
                            filename=doc.filename,
                            file_id=fid,
                            file_type=doc.file_type,
                        )
                        session.conversation_history.append(
                            types.Content(
                                role="user",
                                parts=[types.Part.from_text(text=ctx)],
                            )
                        )

                executor = ToolExecutor(session.engine, session.documents)

                t0 = time.perf_counter()
                try:
                    await orchestrator.run(
                        user_message=text,
                        conversation_history=session.conversation_history,
                        executor=executor,
                        workbook=session.engine,
                        websocket=websocket,
                        system_prompt=get_system_prompt(session.session_role),
                    )
                    elapsed = time.perf_counter() - t0
                    sheets = session.engine.get_all_sheet_names()
                    logger.info(
                        f"[{session_id[:8]}] → agent done in {elapsed:.1f}s  "
                        f"sheets={sheets}  history_turns={len(session.conversation_history)}"
                    )
                except Exception as e:
                    elapsed = time.perf_counter() - t0
                    logger.exception(
                        f"[{session_id[:8]}] ✗ agent error after {elapsed:.1f}s: {e}"
                    )
                    err_str = str(e)
                    if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str or "quota" in err_str.lower():
                        user_msg = "API quota exceeded. Please wait a moment and try again, or check your Gemini API plan."
                    elif "503" in err_str or "unavailable" in err_str.lower() or "overloaded" in err_str.lower():
                        user_msg = "The AI service is temporarily unavailable. Please try again in a moment."
                    else:
                        user_msg = "Something went wrong. Please try again."
                    await websocket.send_json({
                        "type": "error",
                        "message": user_msg,
                    })

            elif msg_type == "cell_edit":
                # User edited a cell directly in the spreadsheet
                sheet = data.get("sheet", "")
                cell = data.get("cell", "")
                old_val = data.get("old", "")
                new_val = data.get("new", "")

                if sheet and cell:
                    logger.debug(
                        f"[{session_id[:8]}] cell_edit: {sheet}!{cell} "
                        f"'{str(old_val)[:30]}' → '{str(new_val)[:30]}'"
                    )
                    # Apply to workbook immediately
                    session.engine.apply_user_edit(sheet, cell, new_val)
                    # Buffer for context injection
                    session.cell_edit_buffer.append({
                        "sheet": sheet,
                        "cell": cell,
                        "old": old_val,
                        "new": new_val,
                        "timestamp": datetime.datetime.utcnow().isoformat(),
                    })

    except WebSocketDisconnect:
        logger.info(f"[{session_id[:8]}] ws disconnected")
    except Exception as e:
        logger.exception(f"[{session_id[:8]}] ws error: {e}")
        try:
            await websocket.send_json({"type": "error", "message": str(e)[:300]})
        except Exception:
            pass


# ── Helpers ───────────────────────────────────────────────────────────────────

def _flush_cell_edit_buffer(session: WorkspaceSession) -> None:
    """Inject buffered cell edits into conversation history as context."""
    if not session.cell_edit_buffer:
        return

    for edit in session.cell_edit_buffer:
        ctx = CELL_EDIT_CONTEXT_TEMPLATE.format(
            cell=edit["cell"],
            sheet=edit["sheet"],
            old_value=edit["old"],
            new_value=edit["new"],
            timestamp=edit.get("timestamp", ""),
        )
        session.conversation_history.append(
            types.Content(
                role="user",
                parts=[types.Part.from_text(text=ctx)],
            )
        )

    session.cell_edit_buffer.clear()


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
