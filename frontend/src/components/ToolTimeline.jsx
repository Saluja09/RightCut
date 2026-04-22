/**
 * RightCut — Tool call timeline.
 * Collapsed by default (shows step count badge). Expands on click.
 */
import { useState } from 'react'
import {
  CheckCircle2, XCircle, Loader2, ChevronDown, ChevronUp,
  FileSearch, Table2, PenLine, FunctionSquare, Pencil,
  Paintbrush, Link2, ArrowUpDown, BarChart2, ShieldCheck,
  Rows3, Settings2
} from 'lucide-react'

const TOOL_META = {
  parse_document:    { Icon: FileSearch,    label: 'Parse document'    },
  create_sheet:      { Icon: Table2,        label: 'Create sheet'      },
  set_range:         { Icon: PenLine,       label: 'Set range'         },
  add_formula:       { Icon: FunctionSquare,label: 'Add formula'       },
  set_cell:          { Icon: Pencil,        label: 'Set cell'          },
  format_cells:      { Icon: Paintbrush,    label: 'Format cells'      },
  cite_source:       { Icon: Link2,         label: 'Cite source'       },
  sort_range:        { Icon: ArrowUpDown,   label: 'Sort range'        },
  add_chart:         { Icon: BarChart2,     label: 'Add chart'         },
  validate_workbook: { Icon: ShieldCheck,   label: 'Validate workbook' },
  freeze_panes:      { Icon: Rows3,         label: 'Freeze panes'      },
}

function getMeta(toolName) {
  return TOOL_META[toolName] || { Icon: Settings2, label: toolName?.replace(/_/g, ' ') || 'Tool call' }
}

function summarise(toolName, args) {
  if (!args) return ''
  if (toolName === 'set_cell' || toolName === 'set_range') {
    const cell = args.cell || args.start_cell || ''
    const val  = args.value ?? args.formula ?? ''
    const sheet = args.sheet_name ? `${args.sheet_name} · ` : ''
    return `${sheet}${cell}${val ? ` → ${String(val).slice(0, 40)}` : ''}`
  }
  if (toolName === 'create_sheet') return args.sheet_name || ''
  if (toolName === 'format_cells') {
    const sheet = args.sheet_name || ''
    const range = args.range || ''
    return `${sheet} ${range}`.trim()
  }
  if (toolName === 'add_formula') {
    return `${args.cell || ''} ${args.formula ? '= ' + String(args.formula).slice(0, 40) : ''}`.trim()
  }
  if (toolName === 'parse_document') return args.file_id || ''
  if (toolName === 'add_chart') return args.chart_type || ''
  if (toolName === 'sort_range') return `${args.sheet_name || ''} ${args.range || ''}`.trim()
  const first = Object.values(args).find((v) => typeof v === 'string')
  return first ? String(first).slice(0, 50) : ''
}

export default function ToolTimeline({ steps = [], isLive = false }) {
  const [expanded, setExpanded] = useState(false)

  if (!steps || steps.length === 0) return null

  const doneCount  = steps.filter((s) => s.success === true).length
  const errorCount = steps.filter((s) => s.success === false).length
  const totalMs    = steps.reduce((a, s) => a + (s.duration_ms || 0), 0)

  // While live, show the last few steps inline (no collapse)
  if (isLive) {
    const recent = steps.slice(-4)
    return (
      <div className="tl-root">
        <div className="tl-steps">
          {recent.map((step, i) => (
            <StepRow key={i} step={step} />
          ))}
        </div>
      </div>
    )
  }

  // Completed — show collapsed summary badge by default
  return (
    <div className="tl-root">
      <button className="tl-summary-bar" onClick={() => setExpanded((v) => !v)}>
        <span className="tl-summary-counts">
          <span className="tl-badge tl-badge--done">{doneCount}×</span>
          {errorCount > 0 && <span className="tl-badge tl-badge--err">{errorCount} err</span>}
          <span className="tl-summary-label">
            {steps.length} step{steps.length !== 1 ? 's' : ''}
            {totalMs >= 1000 && ` · ${(totalMs / 1000).toFixed(1)}s`}
          </span>
        </span>
        <span className="tl-summary-chevron">
          {expanded ? <ChevronUp size={11} /> : <ChevronDown size={11} />}
        </span>
      </button>

      {expanded && (
        <div className="tl-steps tl-steps--expanded">
          {steps.map((step, i) => (
            <StepRow key={i} step={step} />
          ))}
        </div>
      )}
    </div>
  )
}

function StepRow({ step }) {
  const [open, setOpen] = useState(false)
  const { Icon, label } = getMeta(step.tool)
  const summary = summarise(step.tool, step.args)
  const isErr   = step.success === false
  const isLive  = step.success === undefined

  return (
    <div className={`tl-step ${isErr ? 'tl-step--err' : ''}`}>
      <div className="tl-step-row" onClick={() => setOpen((v) => !v)}>
        <span className="tl-status">
          {isLive
            ? <Loader2 size={12} className="tl-spin" />
            : isErr
            ? <XCircle size={12} className="tl-err-ico" />
            : <CheckCircle2 size={12} className="tl-ok-ico" />
          }
        </span>
        <span className="tl-tool-ico"><Icon size={11} /></span>
        <span className="tl-label">{label}</span>
        {summary && <span className="tl-summary">{summary}</span>}
        {step.duration_ms != null && (
          <span className="tl-dur">{step.duration_ms}ms</span>
        )}
      </div>
      {open && (step.args || step.error) && (
        <div className="tl-detail">
          {step.args && Object.keys(step.args).length > 0 && (
            <pre className="tl-code">{JSON.stringify(step.args, null, 2)}</pre>
          )}
          {step.error && <div className="tl-err-msg">{step.error}</div>}
        </div>
      )}
    </div>
  )
}
