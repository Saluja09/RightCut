"""
RightCut — FastAPI application entry point.
Handles WebSocket connections, file uploads, and workbook downloads.
"""

from __future__ import annotations

import asyncio
import datetime
import io
import logging
import uuid
from dataclasses import dataclass, field
from typing import Any

from fastapi import FastAPI, File, HTTPException, UploadFile, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from google.genai import types

from agent.orchestrator import AgentOrchestrator
from agent.prompts import CELL_EDIT_CONTEXT_TEMPLATE, DOCUMENT_UPLOAD_CONTEXT_TEMPLATE
from agent.tools import ToolExecutor
from config import CORS_ORIGINS, GEMINI_API_KEY, SESSION_TTL_HOURS, UPLOAD_MAX_SIZE_MB
from excel.engine import WorkbookEngine
from models import ParsedDocument

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)

# ── Application ───────────────────────────────────────────────────────────────
app = FastAPI(
    title="RightCut API",
    description="AI-powered spreadsheet agent for private markets",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ORIGINS,
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
    logger.info("RightCut API started")


# ── Health check ──────────────────────────────────────────────────────────────

@app.get("/health")
async def health() -> dict:
    return {"status": "ok", "sessions": len(sessions)}


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
        xlsx_bytes = session.engine.to_bytes()
    except Exception as e:
        logger.exception(f"Failed to serialize workbook for session {session_id}: {e}")
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

    # Send initial workbook state if sheets exist
    if session.engine.get_all_sheet_names():
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

                # Inject any buffered cell edits as context
                _flush_cell_edit_buffer(session)

                # Inject document upload context messages for any new files
                uploaded_files: list[dict] = data.get("files", [])
                for f in uploaded_files:
                    fid = f.get("file_id")
                    doc = session.documents.get(fid)
                    if doc:
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

                try:
                    await orchestrator.run(
                        user_message=text,
                        conversation_history=session.conversation_history,
                        executor=executor,
                        workbook=session.engine,
                        websocket=websocket,
                    )
                except Exception as e:
                    logger.exception(f"Agent error in session {session_id}: {e}")
                    await websocket.send_json({
                        "type": "error",
                        "message": f"Agent encountered an error: {str(e)[:300]}",
                    })

            elif msg_type == "cell_edit":
                # User edited a cell directly in the spreadsheet
                sheet = data.get("sheet", "")
                cell = data.get("cell", "")
                old_val = data.get("old", "")
                new_val = data.get("new", "")

                if sheet and cell:
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
        logger.info(f"WebSocket disconnected: session={session_id}")
    except Exception as e:
        logger.exception(f"WebSocket error in session {session_id}: {e}")
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
