# RightCut — AI Spreadsheet Agent 

RightCut is an agentic AI system that builds, audits, and self-corrects institutional-quality `.xlsx` workbooks through natural conversation. It doesn't just generate spreadsheets — it evaluates every formula it writes, catches its own errors, and fixes them before you see the output.

Built for the Brightriver AI take-home. Zero infrastructure cost.

---

## Product Features

### Core
- **Conversational spreadsheet building** — describe what you need, watch it appear cell-by-cell in a live Excel-like grid
- **Formula-first outputs** — every derived metric uses real Excel formulas (`=SUM`, `=IRR`, `=VLOOKUP`, cross-sheet refs), never hardcoded values
- **Live preview** — AG Grid renders the workbook in real time with HyperFormula evaluating formulas in-browser, including cross-sheet references
- **Direct cell editing** — click any cell to edit; changes sync back to the agent's context so it stays aware of your modifications
- **Download .xlsx** — full workbook with all formatting, formulas, and native Excel charts preserved
- **Document upload** — parse PDFs, DOCX, CSV, XLSX and extract data with source citations on every cell
- **Chart generation** — creates Chart.js visualizations in the app and native Excel charts in the download
- **Session history** — Supabase persistence for logged-in users, localStorage fallback for guests
- **Dark mode** — full light/dark theme support across the grid, chat, and chrome

### Agentic Self-Correction (audit_sheet)
The key differentiator. After building any sheet, the agent calls `audit_sheet` which:
1. **Evaluates every formula** using the Python `formulas` library — computes actual values, not just syntax checks
2. **Detects errors** — `#DIV/0!`, `#REF!`, `#NAME?`, `NaN`, `Inf`, circular references, missing cross-sheet refs, unbalanced parentheses
3. **Returns a value snapshot** — the LLM sees `"EBITDA": "3000 ← =B4-B6"` and can verify the math makes sense
4. **Self-corrects** — if issues are found, the LLM calls `edit_cell`/`add_formula` to fix them, then re-audits
5. **All within one turn** — the user never sees broken intermediate states

```
User: "Build a quarterly P&L"

Orchestrator tool loop:
  Iteration 1: create_sheet → insert_data (7 rows with formulas)
  Iteration 2: apply_formatting × 9 (headers, number formats, colors)
  Iteration 3: audit_sheet → 23 formulas ALL EVALUATED → 0 errors ✓
               validate_workbook → 0 hardcoded values ✓
  Iteration 4: → final text response to user
```

### Professional Formatting
Every output is styled — the system prompt mandates:
- Navy headers with white text
- Zebra-striped data rows
- Green highlight on key output rows (EBITDA, Net Profit)
- Dark navy on final answer rows (IRR, Intrinsic Value)
- Muted grey for supporting metrics
- Proper number formats (`$#,##0`, `0.0%`, `0.0"x"`)

### Token Management
3-layer compaction pipeline keeps long sessions within context limits:
1. **Tool result compression** — strips bulky data from write-only tool responses (40-50% token reduction)
2. **LLM summarization** — older turns condensed into summaries
3. **Sliding window** — hard trim as last resort

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        FRONTEND (React + Vite)                  │
│                                                                 │
│  ┌──────────────────┐  ┌──────────────────────────────────────┐ │
│  │  Chat Panel      │  │  Preview Panel                       │ │
│  │  • Messages      │  │  • AG Grid (Excel-like)              │ │
│  │  • Tool timeline │  │  • HyperFormula (live eval)          │ │
│  │  • File upload   │  │  • Chart.js visualizations           │ │
│  │                  │  │  • Formula bar + row numbers         │ │
│  └──────────────────┘  └──────────────────────────────────────┘ │
│                    WebSocket (bidirectional)                     │
└──────────────────────────┬──────────────────────────────────────┘
                           │
┌──────────────────────────▼──────────────────────────────────────┐
│                        BACKEND (FastAPI)                        │
│                                                                 │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │  Orchestrator (agentic tool loop, up to 20 iterations)  │   │
│  │                                                         │   │
│  │  Gemini ←→ Tool Executor ←→ WorkbookEngine (openpyxl)  │   │
│  │                 │                                       │   │
│  │                 ├── create_sheet, insert_data            │   │
│  │                 ├── add_formula, edit_cell               │   │
│  │                 ├── apply_formatting (12 format types)   │   │
│  │                 ├── audit_sheet (formulas lib eval)      │   │
│  │                 ├── create_chart, sort_range             │   │
│  │                 ├── parse_document, add_citation         │   │
│  │                 ├── clean_data (14 operations)           │   │
│  │                 └── validate_workbook, get_sheet_state   │   │
│  └─────────────────────────────────────────────────────────┘   │
│                                                                 │
│  Serializer: openpyxl → CellValue (colors, fonts, formats)     │
│              → JSON → WebSocket → Frontend renders              │
└─────────────────────────────────────────────────────────────────┘
```

### Data Flow (Unidirectional)
```
Backend (source of truth) → WebSocket workbook_update → Zustand store → AG Grid render
                                                                      → HyperFormula eval
User cell edit → WebSocket cell_edit → Backend applies → workbook_update → re-render
```

The frontend **never** renders workbook data from localStorage — only from the backend via WebSocket.

---

## Tech Stack

| Layer | Technology |
|---|---|
| LLM | Gemini 3.1 Pro Preview (configurable via `.env`) |
| Backend | FastAPI + Python 3.14 + WebSockets |
| Excel Engine | openpyxl (formatting, formulas, charts) |
| Formula Evaluation | `formulas` library (audit_sheet) |
| Doc Parsing | pdfplumber + pypdf + python-docx + pandas |
| Frontend | React 18 + Vite |
| Spreadsheet Grid | AG Grid Community v35 |
| Formula Display | HyperFormula (in-browser eval) |
| Charts | Chart.js (live) + openpyxl charts (.xlsx) |
| State | Zustand |
| Auth | Supabase Auth (optional, guest mode supported) |
| Hosting | Render (backend) + Netlify (frontend) |

### Cost

| Model | Per Request | Per Session (5 msgs) |
|---|---|---|
| gemini-2.5-flash | ~$0.005 | ~$0.03–0.05 |
| gemini-2.5-pro | ~$0.05 | ~$0.30–0.50 |
| gemini-3.1-pro | ~$0.05 | ~$0.30–0.50 |

The audit_sheet overhead is <10% of total cost per request.

---

## Agent Tools (15)

| Tool | Purpose | Read/Write |
|---|---|---|
| `create_sheet` | Create sheet with column headers | Write |
| `insert_data` | Insert rows with formula support | Write |
| `add_formula` | Add formula to cell or range | Write |
| `edit_cell` | Update a specific cell | Write |
| `apply_formatting` | 12 format types (headers, zebra, output rows, colors, borders, number formats) | Write |
| `create_chart` | Bar, line, pie, scatter charts | Write |
| `sort_range` | Sort rows by column | Write |
| `clean_data` | 14 cleaning operations (trim, dedup, case, split, fill, etc.) | Write |
| `create_model_scaffold` | One-call DCF or LBO model with cross-sheet refs | Write |
| `add_citation` | Cell comment linking to source document | Write |
| `parse_document` | Extract text + tables from PDF/DOCX/CSV/XLSX | Read |
| `get_sheet_state` | Read current sheet data + user edits | Read |
| `get_all_sheet_names` | List sheets in workbook | Read |
| `audit_sheet` | Evaluate all formulas, detect errors, return value snapshot | Read |
| `validate_workbook` | Check for hardcoded values in formula columns | Read |

---

## Key Design Decisions

**1. Agentic self-correction over blind generation**
The LLM evaluates its own formulas via `audit_sheet` before responding. This catches `#DIV/0!`, `#REF!`, circular refs, and nonsensical values. The fix loop runs within the same turn — the user never sees broken output. This is what makes it an agent, not a template generator.

**2. Server-side formula evaluation for audit**
LLMs are bad at arithmetic. Instead of asking the model to mentally compute `=SUM(B2:B6)`, the `formulas` library does exact computation, and the LLM reads the result and reasons about whether it makes sense. Play to each system's strengths.

**3. Backend as single source of truth**
The frontend never renders workbook data from localStorage or Supabase directly. Saved snapshots restore the backend via POST `/restore`, then the backend pushes `workbook_update` over WebSocket. This prevents stale data and cross-session contamination.

**4. Unified system prompt**
One prompt handles everything — budgets, DCFs, LBOs, comps tables, data cleaning. No persona selection, no prescriptive formulas. The prompt describes the rendering pipeline (openpyxl → serializer → AG Grid → HyperFormula → Chart.js) so the LLM understands how its output reaches the user and can exploit every capability.

**5. Formula-first, not value-first**
The system prompt mandates formulas over hardcoded values. `validate_workbook` flags numeric cells in calculated columns. This ensures downloaded `.xlsx` files are dynamic — change an assumption and everything recalculates.

**6. Incremental WebSocket updates**
Every mutating tool call pushes a `workbook_update` to the frontend immediately — the user sees each step happen live (sheet created → data inserted → formatting applied), not just the final result.

---

## Known Bugs

- **Cell colors not rendering in AG Grid** — Backend applies openpyxl formatting correctly (confirmed via serializer tests), and `cellStyle` callbacks are wired up in `sheetSync.js`, but data row colors (zebra stripes, output_row green, section_header blue) don't appear visually in the browser. Headers render correctly via CSS variables. The pipeline is: `bg_color` in CellValue → `__meta_${colIdx}` in rowData → `cellStyle` callback. Investigation narrowed to either AG Grid v35 `cellStyle` behavior or a rendering timing issue.
- **Gemini 3.1 Pro intermittent `httpx.ReadError`** — The model occasionally drops the connection after ~18 seconds. Retries with exponential backoff handle most cases, but some requests fail on the first attempt.
- **Multiple WebSocket connections on session switch** — Switching sessions rapidly can create 3-4 simultaneous WS connections before the old ones close. Functionally harmless (stale session guard rejects mismatched messages) but noisy in logs.

---

## Future Implementations

### Short-term
- **Fix cell color rendering** — Trace why AG Grid v35 `cellStyle` returns styles but they don't paint. May need to switch to `cellClassRules` or `getRowStyle` for AG Grid v35 compatibility.
- **Streaming tool output** — Show partial results as the agent works (e.g., row-by-row insertion animation) instead of bulk updates after each tool call.
- **Audit summary in chat** — Surface the audit results (formulas evaluated, errors found/fixed) as a collapsible card in the chat UI so users can see the verification happened.

### Medium-term
- **Multi-turn audit memory** — Track which formulas were flagged and fixed across turns, so the agent doesn't re-introduce a previously corrected error.
- **Sensitivity tables** — Dedicated tool for building 2D sensitivity/scenario matrices (e.g., IRR across entry multiples × exit multiples).
- **Collaborative editing** — Multiple users on the same session with conflict resolution.
- **Template library** — Pre-built scaffolds for common models (comps, DCF, LBO, cap table, waterfall) that users can customize.
- **Version history** — Workbook snapshots at each agent turn, with diff view and rollback.

### Long-term
- **Excel file import → edit** — Upload an existing `.xlsx`, agent reads it via `get_sheet_state`, user asks for modifications.
- **Data connectors** — Pull live data from APIs (PitchBook, Capital IQ, Bloomberg) directly into sheets.
- **PDF → structured extraction** — Better table extraction from CIMs with layout-aware parsing (current pdfplumber approach struggles with multi-column layouts).
- **Conditional formatting in preview** — Render openpyxl conditional formatting rules (color scales, data bars, icon sets) in AG Grid, not just in the downloaded `.xlsx`.

---

## Setup

### Prerequisites
- Python 3.11+
- Node.js 20+
- A Gemini API key from [Google AI Studio](https://aistudio.google.com/apikey)

### 1. Clone and configure

```bash
git clone <repo>
cd rightcut
cp .env.example .env
# Edit .env — add your GEMINI_API_KEY
# Optionally set GEMINI_MODEL (default: gemini-2.5-pro)
```

### 2. Backend

```bash
cd backend
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```

### 3. Frontend

```bash
cd frontend
npm install
npm run dev
```

Open http://localhost:5173

---

## Project Structure

```
rightcut/
├── backend/
│   ├── main.py                 # FastAPI app, WebSocket, upload, download, restore
│   ├── agent/
│   │   ├── orchestrator.py     # Gemini agentic loop (up to 20 tool iterations)
│   │   ├── tools.py            # Tool dispatcher → WorkbookEngine methods
│   │   ├── tool_schemas.py     # 15 Gemini FunctionDeclaration objects
│   │   ├── prompts.py          # Unified system prompt + templates
│   │   └── compaction.py       # 3-layer token compaction pipeline
│   ├── excel/
│   │   ├── engine.py           # WorkbookEngine — all mutations + audit_sheet
│   │   ├── serializer.py       # Workbook → JSON (colors, fonts, formats)
│   │   └── formulas.py         # Formula helpers, format string inference
│   ├── parsers/
│   │   ├── pdf_parser.py       # pdfplumber + pypdf fallback
│   │   ├── docx_parser.py      # python-docx
│   │   └── csv_parser.py       # pandas CSV/XLSX
│   ├── config.py               # Environment variables + CORS config
│   └── models.py               # Pydantic models (CellValue, ChartMeta, etc.)
├── frontend/
│   ├── src/
│   │   ├── App.jsx             # Root layout, session management
│   │   ├── components/
│   │   │   ├── ChatPanel.jsx       # Message input + file upload
│   │   │   ├── MessageBubble.jsx   # User/agent/error messages + sheet ref cards
│   │   │   ├── ToolTimeline.jsx    # Collapsible tool call accordion
│   │   │   ├── PreviewPanel.jsx    # Tab container + download
│   │   │   ├── SpreadsheetView.jsx # AG Grid wrapper + formula bar
│   │   │   ├── ChartView.jsx       # Chart.js renderer from chart metadata
│   │   │   ├── DocumentView.jsx    # Uploaded document preview
│   │   │   ├── TabBar.jsx          # Sheet + chart tab switcher
│   │   │   ├── LeftSidebar.jsx     # Session history sidebar
│   │   │   └── HistoryDrawer.jsx   # Session list drawer
│   │   ├── hooks/
│   │   │   ├── useWebSocket.js     # WS lifecycle, reconnection, session restore
│   │   │   └── useWorkbook.js      # Derived workbook state
│   │   ├── stores/
│   │   │   ├── workspaceStore.js   # Zustand (messages, workbook, tabs, status)
│   │   │   └── historyStore.js     # Supabase/localStorage persistence
│   │   └── utils/
│   │       ├── sheetSync.js        # Sheet → AG Grid conversion + HyperFormula eval
│   │       └── api.js              # API/WS URL helpers
│   └── package.json
├── .env.example
└── README.md
```

---

## WebSocket Protocol

```
Client → Server:
  { type: "user_message", text, files? }
  { type: "cell_edit", sheet, cell, old, new }
  { type: "ping" }

Server → Client (streamed during agent loop):
  { type: "session_ready", has_workbook }
  { type: "thinking" }
  { type: "tool_call", step: { tool, args, result_summary, duration_ms, success } }
  { type: "workbook_update", state: WorkbookState }
  { type: "new_tab", tab: { id, name, type } }
  { type: "agent_response", text, timeline }
  { type: "history_compacted", strategy, tokens_before, tokens_after }
  { type: "error", message }
  { type: "pong" }
```
