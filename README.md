# RightCut — AI Spreadsheet Agent for Private Markets

> Take-home assignment for [Brightriver AI](https://www.brightriver.ai)

RightCut is an AI-powered spreadsheet agent that builds, edits, and analyzes institutional-quality `.xlsx` workbooks through natural conversation. Upload deal documents, ask for analyses, and watch the agent construct formula-driven spreadsheets in real time — with a live preview, direct cell editing, and a transparent tool-call timeline for every action.

---

## Demo Scenarios

**1. Build a comps table from scratch**
> "Build a SaaS comps table with 8 companies including Salesforce, HubSpot, and Zendesk. Include EV, Revenue, EBITDA, EV/EBITDA, EV/Revenue, and CAGR. Apply color scale formatting to the EV/EBITDA column."

**2. Upload a CIM and extract financials**
> Upload a PDF → "Extract the key financial metrics from this CIM and build a deal sheet with a Summary, Financials, and Returns Analysis tab."

**3. Returns analysis**
> "Create an IRR/MOIC matrix for a $150M entry at 8x EBITDA. Model 3 exit scenarios at 10x, 12x, and 14x EBITDA after 5 years."

---

## Stack — Zero Cost

| Layer | Technology |
|---|---|
| LLM | Gemini 2.5 Flash (free tier, Google AI Studio) |
| Backend | FastAPI + Python 3.11 |
| Excel Engine | openpyxl |
| Doc Parsing | pdfplumber + pypdf + python-docx + pandas |
| Frontend | React 18 + Vite |
| Spreadsheet Grid | Jspreadsheet CE (Community Edition) |
| State | Zustand |
| Real-time | WebSockets |

---

## Setup

### Prerequisites

- Python 3.11+
- Node.js 20+
- A free Gemini API key from [Google AI Studio](https://aistudio.google.com/apikey)

### 1. Clone and configure

```bash
git clone <repo>
cd rightcut

# Create .env from template
cp .env.example .env
# Edit .env and add your GEMINI_API_KEY
```

### 2. Backend

```bash
cd backend
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
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

### Docker (optional)

```bash
docker-compose up --build
```

---

## Architecture

```
┌──────────────────────────────────────────────────────────────┐
│  RightCut                                           [●] Live │
├────────────────────────┬─────────────────────────────────────┤
│  LEFT 40%              │  [Comparables] [Financials] [doc.pdf]│
│  Chat Panel            │                                      │
│                        │  RIGHT 60%                           │
│  • Message bubbles     │  Preview Panel                       │
│  • Tool call timeline  │                                      │
│    (collapsible)       │  Active tab shows either:            │
│  • Sync badges when    │  • Spreadsheet (Jspreadsheet CE)     │
│    user edits sheet    │  • Document preview (uploaded file)  │
│                        │                                      │
│  [📎][__input____ ][▶] │  Cells are directly editable.        │
│                        │  Edits sync back to agent context.   │
├────────────────────────┴─────────────────────────────────────┤
│  ● Live · 3 sheets · 12 formulas · All formulas OK           │
└──────────────────────────────────────────────────────────────┘
```

### Key Design Decisions

**Formula-first outputs** — Every derived metric uses an Excel formula (`=IRR(...)`, `=((E2/B2)^(1/5))-1`), never a hardcoded value. This is what separates RightCut from generic LLM wrappers.

**Citation discipline** — Every cell containing data extracted from a document gets a cell comment linking it back to the source file and page. Institutional investors need every number to be traceable.

**Tool-call timeline** — The collapsible accordion in the chat panel shows every tool the agent called, with arguments, result summaries, and timing. Full transparency builds trust with analysts.

**Bidirectional sync** — User edits cells directly in the preview panel → changes are immediately applied to the in-memory workbook AND injected as context into the next agent message. The agent stays aware of what the user is seeing.

### Agent Tools (11)

| Tool | Purpose |
|---|---|
| `parse_document` | Extract text + tables from PDF/DOCX/CSV/XLSX |
| `create_sheet` | Create a new sheet with styled headers |
| `insert_data` | Insert rows (formula strings supported) |
| `add_formula` | Add Excel formula to a cell or range |
| `edit_cell` | Update a specific cell |
| `apply_formatting` | Color scales, data bars, bold headers, number formats |
| `add_citation` | Cell comment linking to source document |
| `sort_range` | Sort rows by a column |
| `create_chart` | Embed chart in .xlsx (bar, line, pie, scatter) |
| `validate_workbook` | Check for hardcoded values, formula integrity |
| `get_sheet_state` | Read current sheet state including user edits |

---

## WebSocket Protocol

```typescript
// Client → Server
{ type: "user_message", text: string, files?: { file_id, filename }[] }
{ type: "cell_edit", sheet: string, cell: string, old: string, new: string }
{ type: "ping" }

// Server → Client (streamed during agent loop)
{ type: "thinking", iteration: number }
{ type: "tool_call", step: { tool, args, result_summary, duration_ms, success } }
{ type: "workbook_update", state: WorkbookState }
{ type: "new_tab", tab: { id, name, type } }
{ type: "agent_response", text: string, timeline: ToolStep[] }
{ type: "error", message: string }
{ type: "pong" }
```

---

## Project Structure

```
rightcut/
├── backend/
│   ├── main.py                 # FastAPI app, WebSocket, upload, download
│   ├── agent/
│   │   ├── orchestrator.py     # Gemini agentic loop (AFC disabled, manual streaming)
│   │   ├── tools.py            # Tool dispatcher → implementations
│   │   ├── tool_schemas.py     # Gemini FunctionDeclaration objects
│   │   └── prompts.py          # System prompt + message templates
│   ├── excel/
│   │   ├── engine.py           # WorkbookEngine (all mutations)
│   │   ├── serializer.py       # Workbook → JSON for frontend
│   │   └── formulas.py         # Formula helpers, format strings
│   ├── parsers/
│   │   ├── pdf_parser.py       # pdfplumber + pypdf fallback
│   │   ├── docx_parser.py      # python-docx
│   │   └── csv_parser.py       # pandas CSV/XLSX
│   ├── config.py               # Env vars
│   ├── models.py               # Pydantic models
│   └── requirements.txt
├── frontend/
│   ├── src/
│   │   ├── App.jsx             # Root — 40/60 split layout
│   │   ├── components/
│   │   │   ├── ChatPanel.jsx       # Message list + input
│   │   │   ├── MessageBubble.jsx   # User/agent/error messages
│   │   │   ├── ToolTimeline.jsx    # Collapsible tool call accordion
│   │   │   ├── PreviewPanel.jsx    # Tab container
│   │   │   ├── SpreadsheetView.jsx # Jspreadsheet CE wrapper
│   │   │   ├── DocumentView.jsx    # Uploaded doc preview
│   │   │   ├── TabBar.jsx          # Tab switcher + download button
│   │   │   └── StatusBar.jsx       # Bottom status bar
│   │   ├── hooks/
│   │   │   ├── useWebSocket.js     # WS lifecycle + reconnection
│   │   │   └── useWorkbook.js      # Workbook derived state
│   │   ├── stores/
│   │   │   └── workspaceStore.js   # Zustand store (all UI state)
│   │   └── utils/
│   │       └── sheetSync.js        # Coordinate conversion utilities
│   ├── vite.config.js          # Proxy config (no CORS in dev)
│   └── package.json
├── docker-compose.yml
├── .env.example
└── README.md
```
