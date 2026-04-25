/**
 * RightCut — Tool call timeline.
 * Collapsed: "Ran N tools ›"  Expanded: vertical connector + icon + label
 * Modelled after Claude Code's tool call presentation.
 */
import { useState } from 'react'
import {
  FileSearch, Table2, FunctionSquare, Pencil,
  Paintbrush, Link2, ArrowUpDown, BarChart2, ShieldCheck,
  Settings2, Database, Layers, ChevronRight
} from 'lucide-react'

const TOOL_META = {
  parse_document:       { Icon: FileSearch,      label: 'Read document'        },
  create_sheet:         { Icon: Table2,           label: 'Create sheet'         },
  insert_data:          { Icon: Database,         label: 'Insert data'          },
  add_formula:          { Icon: FunctionSquare,   label: 'Add formula'          },
  edit_cell:            { Icon: Pencil,           label: 'Edit cell'            },
  apply_formatting:     { Icon: Paintbrush,       label: 'Apply formatting'     },
  add_citation:         { Icon: Link2,            label: 'Add citation'         },
  sort_range:           { Icon: ArrowUpDown,      label: 'Sort range'           },
  create_chart:         { Icon: BarChart2,        label: 'Create chart'         },
  validate_workbook:    { Icon: ShieldCheck,      label: 'Validate workbook'    },
  get_sheet_state:      { Icon: Layers,           label: 'Read sheet'           },
  get_all_sheet_names:  { Icon: Layers,           label: 'List sheets'          },
  create_model_scaffold:{ Icon: Layers,           label: 'Build model scaffold' },
  clean_data:           { Icon: Settings2,        label: 'Clean data'           },
}

function getMeta(toolName) {
  return TOOL_META[toolName] || {
    Icon: Settings2,
    label: toolName?.replace(/_/g, ' ') || 'Tool call',
  }
}

function getSubtitle(toolName, args) {
  if (!args) return null
  if (toolName === 'create_sheet' || toolName === 'get_sheet_state')
    return args.sheet_name || null
  if (toolName === 'insert_data')
    return args.sheet_name ? `${args.sheet_name}` : null
  if (toolName === 'add_formula')
    return args.cell ? `${args.sheet_name ? args.sheet_name + '!' : ''}${args.cell}` : null
  if (toolName === 'edit_cell')
    return args.cell ? `${args.sheet_name ? args.sheet_name + '!' : ''}${args.cell}` : null
  if (toolName === 'apply_formatting')
    return args.format_type ? `${args.format_type}${args.cell_range ? ' · ' + args.cell_range : ''}` : null
  if (toolName === 'parse_document')
    return args.file_id || null
  if (toolName === 'create_chart')
    return args.chart_type || null
  if (toolName === 'create_model_scaffold')
    return args.model_type ? args.model_type.toUpperCase() : null
  if (toolName === 'sort_range')
    return args.sort_column ? `by ${args.sort_column}` : null
  return null
}

export default function ToolTimeline({ steps = [], isLive = false }) {
  const [expanded, setExpanded] = useState(false)

  if (!steps || steps.length === 0) return null

  const errorCount = steps.filter((s) => s.success === false).length
  const label = isLive
    ? `Running tools…`
    : `Ran ${steps.length} tool${steps.length !== 1 ? 's' : ''}${errorCount > 0 ? ` · ${errorCount} error${errorCount > 1 ? 's' : ''}` : ''}`

  return (
    <div className="tl-root">
      <button
        className={`tl-header ${errorCount > 0 ? 'tl-header--err' : ''}`}
        onClick={() => setExpanded((v) => !v)}
        aria-expanded={expanded}
      >
        <span className="tl-header-label">{label}</span>
        <ChevronRight
          size={12}
          className={`tl-header-chevron ${expanded ? 'tl-header-chevron--open' : ''}`}
        />
      </button>

      {(expanded || isLive) && (
        <div className="tl-steps">
          {steps.map((step, i) => {
            const isLast = i === steps.length - 1
            const isPending = step.success === undefined
            const isErr = step.success === false
            const { Icon, label: toolLabel } = getMeta(step.tool)
            const subtitle = getSubtitle(step.tool, step.args)

            return (
              <div key={`${step.tool}_${i}`} className="tl-step">
                {/* Vertical connector line */}
                <div className="tl-connector">
                  <div className={`tl-dot ${isErr ? 'tl-dot--err' : isPending ? 'tl-dot--pending' : 'tl-dot--done'}`} />
                  {!isLast && <div className="tl-line" />}
                </div>

                {/* Step content */}
                <div className="tl-step-body">
                  <div className="tl-step-top">
                    <Icon size={12} className={`tl-step-icon ${isErr ? 'tl-step-icon--err' : ''}`} />
                    <span className={`tl-step-label ${isErr ? 'tl-step-label--err' : ''} ${isPending ? 'tl-step-label--pending' : ''}`}>
                      {toolLabel}
                    </span>
                  </div>
                  {subtitle && (
                    <div className="tl-step-subtitle">{subtitle}</div>
                  )}
                  {isErr && step.error && (
                    <div className="tl-step-error">{step.error}</div>
                  )}
                </div>
              </div>
            )
          })}
        </div>
      )}
    </div>
  )
}
