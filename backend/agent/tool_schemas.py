"""
RightCut — Gemini function declarations for all 11 agent tools.
Uses google-genai types.Schema / types.FunctionDeclaration.
"""

from google.genai import types


def _s(t: types.Type, description: str, **kwargs) -> types.Schema:
    return types.Schema(type=t, description=description, **kwargs)


def _str(desc: str) -> types.Schema:
    return _s(types.Type.STRING, desc)


def _bool(desc: str) -> types.Schema:
    return _s(types.Type.BOOLEAN, desc)


def _int(desc: str) -> types.Schema:
    return _s(types.Type.INTEGER, desc)


def _obj(desc: str, props: dict | None = None, required: list[str] | None = None) -> types.Schema:
    kwargs = {}
    if props:
        kwargs["properties"] = props
    if required:
        kwargs["required"] = required
    return types.Schema(type=types.Type.OBJECT, description=desc, **kwargs)


def _arr(desc: str, items: types.Schema) -> types.Schema:
    return types.Schema(type=types.Type.ARRAY, description=desc, items=items)


# ── Tool declarations ──────────────────────────────────────────────────────────

PARSE_DOCUMENT = types.FunctionDeclaration(
    name="parse_document",
    description=(
        "Extract text and tables from an uploaded document (PDF, DOCX, CSV, or XLSX). "
        "ALWAYS call this before using any document data. Returns structured text and tables."
    ),
    parameters=_obj(
        "parse_document parameters",
        props={
            "file_id": _str("ID of the uploaded file to parse (returned by the upload endpoint)"),
            "extract_tables": _bool("Whether to extract tables separately. Default true."),
        },
        required=["file_id"],
    ),
)

CREATE_SHEET = types.FunctionDeclaration(
    name="create_sheet",
    description=(
        "Create a new sheet in the workbook with the specified column headers. "
        "If the sheet already exists it will be replaced. "
        "Always call this before insert_data on a new sheet."
    ),
    parameters=_obj(
        "create_sheet parameters",
        props={
            "sheet_name": _str("Name for the new sheet (e.g. 'Comparables', 'Financials')"),
            "headers": _arr(
                "Column headers for the sheet",
                _str("A single column header string"),
            ),
        },
        required=["sheet_name", "headers"],
    ),
)

INSERT_DATA = types.FunctionDeclaration(
    name="insert_data",
    description=(
        "Insert rows of data into an existing sheet. "
        "Use formulas (starting with '=') wherever possible instead of hardcoded values. "
        "start_row defaults to 2 (first row after header)."
    ),
    parameters=_obj(
        "insert_data parameters",
        props={
            "sheet_name": _str("Target sheet name"),
            "rows": _arr(
                "Array of rows to insert. Each row is an array of cell values (strings). "
                "Use '=FORMULA' for calculated cells.",
                _arr("A single row", _str("A cell value or formula string")),
            ),
            "start_row": _int("Row number to start inserting at (1-indexed). Default 2."),
        },
        required=["sheet_name", "rows"],
    ),
)

ADD_FORMULA = types.FunctionDeclaration(
    name="add_formula",
    description=(
        "Add an Excel formula to a specific cell. "
        "Always prefer formulas over hardcoded values. "
        "Supports all standard Excel formulas: SUM, AVERAGE, RANK, IF, VLOOKUP, IRR, XIRR, etc. "
        "Optionally apply the same formula pattern across a range."
    ),
    parameters=_obj(
        "add_formula parameters",
        props={
            "sheet_name": _str("Target sheet name"),
            "cell": _str("Target cell reference e.g. 'B3', 'C10'"),
            "formula": _str("Excel formula starting with '=' e.g. '=SUM(B2:B10)'"),
            "apply_to_range": _str(
                "Optional: apply the formula pattern to this range e.g. 'B3:B12'. "
                "Row references will be shifted automatically."
            ),
        },
        required=["sheet_name", "cell", "formula"],
    ),
)

EDIT_CELL = types.FunctionDeclaration(
    name="edit_cell",
    description=(
        "Edit the value of a specific cell. "
        "Use for targeted corrections or updates. "
        "Call get_sheet_state first to see the current value."
    ),
    parameters=_obj(
        "edit_cell parameters",
        props={
            "sheet_name": _str("Target sheet name"),
            "cell": _str("Cell reference e.g. 'A1', 'C5'"),
            "value": _str("New value or formula for the cell"),
        },
        required=["sheet_name", "cell", "value"],
    ),
)

APPLY_FORMATTING = types.FunctionDeclaration(
    name="apply_formatting",
    description=(
        "Apply formatting to a range of cells. "
        "Use section_header for logical section dividers within a sheet. "
        "Use subtotal_row for totals/summary lines. "
        "Use output_row (green) for key outputs like EBITDA, FCFF, Net Profit. "
        "Use final_answer_row (dark navy) for THE most important result (IRR, intrinsic value, MOIC). "
        "Use muted_row for sub-commentary rows like growth % and margin %. "
        "Use zebra_stripe for comps/sensitivity tables. "
        "Use color_scale for rankings. "
        "Use background_color with FF0000 for red flags."
    ),
    parameters=_obj(
        "apply_formatting parameters",
        props={
            "sheet_name": _str("Target sheet name"),
            "cell_range": _str("Cell range e.g. 'A1:F10', 'C2:C20'"),
            "format_type": types.Schema(
                type=types.Type.STRING,
                description="Type of formatting to apply",
                enum=[
                    "color_scale",      # green-yellow-red traffic light (good for rankings)
                    "data_bar",         # blue data bar proportional to value
                    "bold_header",      # navy background, white bold text — for main column header rows
                    "section_header",   # light-blue bg, dark bold text — for section divider rows within a sheet
                    "subtotal_row",     # bold text only — for totals, sub-totals, computed summary rows
                    "output_row",       # light green bg, bold — for key output rows (FCFF, EBITDA, Net Profit)
                    "final_answer_row", # dark navy bg, white bold — for THE single most important result row
                    "muted_row",        # grey font — for sub-commentary rows (growth %, margin % labels)
                    "zebra_stripe",     # alternating grey/white rows — for comps tables and sensitivity tables
                    "number_format",    # custom number format string
                    "border",           # thin border around cells
                    "font_color",       # change font color
                    "background_color", # fill background color (use for red flags: FF0000)
                ],
            ),
            "format_config": _obj(
                "Format-specific configuration. "
                "For number_format: {format: '$#,##0'}. "
                "For font_color/background_color: {color: 'FF0000'}. "
                "For border: {style: 'thin'}. "
                "Not needed for section_header, subtotal_row, output_row, final_answer_row, muted_row, zebra_stripe.",
            ),
        },
        required=["sheet_name", "cell_range", "format_type"],
    ),
)

ADD_CITATION = types.FunctionDeclaration(
    name="add_citation",
    description=(
        "Add a cell comment linking the cell's value to its source document. "
        "REQUIRED for every cell containing data extracted from an uploaded document."
    ),
    parameters=_obj(
        "add_citation parameters",
        props={
            "sheet_name": _str("Target sheet name"),
            "cell": _str("Cell reference to annotate"),
            "source_file": _str("Filename of the source document"),
            "source_location": _str("Page number, section, or paragraph reference e.g. 'p.12', 'Section 3.2'"),
            "excerpt": _str("Brief quote or description from the source (max 200 chars)"),
        },
        required=["sheet_name", "cell", "source_file", "source_location"],
    ),
)

SORT_RANGE = types.FunctionDeclaration(
    name="sort_range",
    description=(
        "Sort all data rows in a sheet by a specified column. "
        "The header row (row 1) is preserved. "
        "Useful for ranking comps by EV/EBITDA, revenue, etc."
    ),
    parameters=_obj(
        "sort_range parameters",
        props={
            "sheet_name": _str("Target sheet name"),
            "sort_column": _str(
                "Column to sort by — either the header name (e.g. 'EV/EBITDA') "
                "or a letter (e.g. 'C')"
            ),
            "ascending": _bool("True for ascending (lowest first), False for descending. Default True."),
        },
        required=["sheet_name", "sort_column"],
    ),
)

CREATE_CHART = types.FunctionDeclaration(
    name="create_chart",
    description=(
        "Create a chart from data in the workbook. "
        "The chart is embedded in the .xlsx download. "
        "A placeholder is shown in the live preview."
    ),
    parameters=_obj(
        "create_chart parameters",
        props={
            "sheet_name": _str("Sheet containing the data"),
            "chart_type": types.Schema(
                type=types.Type.STRING,
                description="Chart type",
                enum=["bar", "line", "pie", "scatter"],
            ),
            "data_range": _str("Data range including headers e.g. 'A1:C10'"),
            "title": _str("Chart title"),
            "target_cell": _str("Cell where the top-left of the chart is placed e.g. 'H2'"),
        },
        required=["sheet_name", "chart_type", "data_range"],
    ),
)

CREATE_MODEL_SCAFFOLD = types.FunctionDeclaration(
    name="create_model_scaffold",
    description=(
        "Build a complete, professionally formatted financial model scaffold in one call. "
        "Use this INSTEAD of manually calling create_sheet + insert_data + add_formula for standard models. "
        "Produces correctly cross-referenced formulas, proper section headers, color-coded rows — "
        "no circular references guaranteed. "
        "Supported model_type: 'dcf' or 'lbo'. "
        "For LBO: builds Cover, Assumptions, Debt Schedule, Income Statement, and Returns sheets "
        "with MOIC and IRR automatically calculated."
    ),
    parameters=_obj(
        "create_model_scaffold parameters",
        props={
            "model_type": types.Schema(
                type=types.Type.STRING,
                description="Type of model to build: 'dcf' or 'lbo'",
                enum=["dcf", "lbo"],
            ),
            "params": _obj(
                "Model parameters. "
                "For DCF: base_ebitda, revenue_growth (e.g. 0.10), ebitda_margin (e.g. 0.20), "
                "da_pct (e.g. 0.05), capex_pct (e.g. 0.06), wc_pct (e.g. 0.01), "
                "tax_rate (e.g. 0.25), wacc (e.g. 0.10), terminal_growth (e.g. 0.03), "
                "net_debt, shares_outstanding, company_name, years (default 5), currency (default USD). "
                "For LBO: entry_ebitda, entry_multiple (e.g. 8.0), revenue_growth (e.g. 0.10), "
                "ebitda_margin (e.g. 0.25), debt_pct (e.g. 0.60), interest_rate (e.g. 0.07), "
                "amort_pct (e.g. 0.05), exit_multiple (e.g. 10.0), tax_rate (e.g. 0.25), "
                "company_name, years (default 5), currency (default USD).",
            ),
        },
        required=["model_type"],
    ),
)

CLEAN_DATA = types.FunctionDeclaration(
    name="clean_data",
    description=(
        "Apply a data-cleaning operation to a sheet. "
        "Use this to fix messy data: trim whitespace, remove duplicates, remove blank rows, "
        "change case, find & replace, convert text to numbers or dates, split columns, fill down, "
        "extract numbers from text, fix number formatting, or remove special characters. "
        "Always call get_sheet_state first to understand the data structure."
    ),
    parameters=_obj(
        "clean_data parameters",
        props={
            "sheet_name": _str("Target sheet name"),
            "operation": types.Schema(
                type=types.Type.STRING,
                description="The cleaning operation to apply",
                enum=[
                    "trim_whitespace",       # Remove extra whitespace from text cells
                    "remove_duplicates",     # Delete exact duplicate rows
                    "remove_blank_rows",     # Delete rows where the target column (or all columns) is empty
                    "to_uppercase",          # Convert text to UPPERCASE
                    "to_lowercase",          # Convert text to lowercase
                    "to_titlecase",          # Convert text to Title Case
                    "find_replace",          # Find and replace text (requires find_text)
                    "convert_to_number",     # Strip $, commas, % and convert to numeric
                    "convert_to_date",       # Parse text as dates (requires dateutil)
                    "remove_special_chars",  # Remove non-alphanumeric characters
                    "fill_down",             # Fill blank cells with the value above (requires column)
                    "extract_numbers",       # Extract first number from mixed text (requires column)
                    "split_column",          # Split column on delimiter into two columns (requires column)
                    "fix_number_format",     # Parse and reformat numbers with consistent Excel number_format
                    "standardize_text",      # Normalize unicode, collapse whitespace
                ],
            ),
            "column": _str(
                "Optional: column header name (e.g. 'Name') or letter (e.g. 'B') to target. "
                "If omitted, the operation applies to all text columns."
            ),
            "find_text": _str("Text to find (required for find_replace operation)"),
            "replace_text": _str("Replacement text (used by find_replace; defaults to empty string)"),
            "delimiter": _str("Delimiter for split_column (defaults to space)"),
            "new_column_name": _str("Header name for the new column created by split_column"),
        },
        required=["sheet_name", "operation"],
    ),
)

VALIDATE_WORKBOOK = types.FunctionDeclaration(
    name="validate_workbook",
    description=(
        "Validate the workbook for quality: checks formulas, flags hardcoded values "
        "in calculated columns, and returns a summary report. "
        "Call this after completing a set of changes."
    ),
    parameters=_obj(
        "validate_workbook parameters",
        props={
            "check_hardcoded": _bool(
                "If true, flag numeric cells in calculated-looking columns. Default true."
            ),
        },
    ),
)

GET_SHEET_STATE = types.FunctionDeclaration(
    name="get_sheet_state",
    description=(
        "Read the current state of a sheet including headers, data, and any user edits. "
        "ALWAYS call this before editing an existing sheet."
    ),
    parameters=_obj(
        "get_sheet_state parameters",
        props={
            "sheet_name": _str("Name of the sheet to read"),
        },
        required=["sheet_name"],
    ),
)

GET_ALL_SHEET_NAMES = types.FunctionDeclaration(
    name="get_all_sheet_names",
    description=(
        "List all sheet names currently in the workbook. "
        "Use this when you need to know what sheets exist before deciding which to edit or clean."
    ),
    parameters=_obj("get_all_sheet_names parameters"),
)

# ── Combined tool object passed to Gemini ─────────────────────────────────────

RIGHTCUT_TOOL = types.Tool(
    function_declarations=[
        CREATE_MODEL_SCAFFOLD,
        PARSE_DOCUMENT,
        CREATE_SHEET,
        INSERT_DATA,
        ADD_FORMULA,
        EDIT_CELL,
        APPLY_FORMATTING,
        ADD_CITATION,
        SORT_RANGE,
        CREATE_CHART,
        CLEAN_DATA,
        VALIDATE_WORKBOOK,
        GET_SHEET_STATE,
        GET_ALL_SHEET_NAMES,
    ]
)
