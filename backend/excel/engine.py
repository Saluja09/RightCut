"""
RightCut — WorkbookEngine
In-memory openpyxl workbook manager. All mutations happen here.
One instance per WebSocket session.
"""

from __future__ import annotations

import io
import logging
from typing import Any

import openpyxl
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.comments import Comment
from openpyxl.formatting.rule import (
    ColorScaleRule,
    DataBarRule,
    FormatObject,
    Rule,
)
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
    numbers,
)
from openpyxl.utils import column_index_from_string, get_column_letter

from .formulas import infer_format, validate_formula
from models import ChartMeta, ValidationStats

logger = logging.getLogger(__name__)

# Thin border side used everywhere
_THIN = Side(style="thin")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


class WorkbookEngine:
    """Thread-safe-enough for single-session async usage."""

    def __init__(self) -> None:
        self.wb: Workbook = Workbook()
        # Remove the default blank sheet
        if "Sheet" in self.wb.sheetnames:
            del self.wb["Sheet"]
        # chart metadata keyed by sheet name
        self._charts: dict[str, list[ChartMeta]] = {}

    # ── Sheet lifecycle ───────────────────────────────────────────────────────

    def create_sheet(self, sheet_name: str, headers: list[str]) -> dict:
        """Create (or reset) a sheet with bold styled headers."""
        if sheet_name in self.wb.sheetnames:
            del self.wb[sheet_name]

        ws = self.wb.create_sheet(title=sheet_name)
        self._charts.setdefault(sheet_name, [])

        # Write and style the header row
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(fill_type="solid", fgColor="1F4E79")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = _BORDER
            ws.column_dimensions[get_column_letter(col_idx)].width = max(len(header) + 4, 14)

        ws.freeze_panes = "A2"

        return {
            "sheet_name": sheet_name,
            "headers": headers,
            "created": True,
        }

    def get_all_sheet_names(self) -> list[str]:
        return list(self.wb.sheetnames)

    # ── Model scaffold ────────────────────────────────────────────────────────

    def create_model_scaffold(
        self,
        model_type: str,
        params: dict,
    ) -> dict:
        """
        Build a complete, correctly-structured model scaffold with proper
        formulas, references, and formatting in one call.
        Supported model_type: "dcf"
        params for DCF:
          base_ebitda, revenue_growth, ebitda_margin, da_pct, capex_pct,
          wc_pct, tax_rate, wacc, terminal_growth, net_debt, shares_outstanding,
          company_name, years (default 5), currency (default USD)
        """
        if model_type.lower() == "dcf":
            return self._build_dcf_scaffold(params)
        return {"error": f"Unknown model_type: {model_type}", "success": False}

    def _build_dcf_scaffold(self, p: dict) -> dict:
        from datetime import date

        # ── Parameters with defaults ──────────────────────────────────────────
        company     = p.get("company_name", "Company")
        currency    = p.get("currency", "USD")
        base_ebitda = float(p.get("base_ebitda", 50_000_000))
        rev_growth  = float(p.get("revenue_growth", 0.10))
        ebitda_margin = float(p.get("ebitda_margin", 0.20))
        da_pct      = float(p.get("da_pct", 0.05))
        capex_pct   = float(p.get("capex_pct", 0.06))
        wc_pct      = float(p.get("wc_pct", 0.01))
        tax_rate    = float(p.get("tax_rate", 0.25))
        wacc        = float(p.get("wacc", 0.10))
        term_g      = float(p.get("terminal_growth", 0.03))
        net_debt    = float(p.get("net_debt", 0))
        shares      = float(p.get("shares_outstanding", 10_000_000))
        n_years     = int(p.get("years", 5))

        # Derived base year
        base_revenue = base_ebitda / ebitda_margin

        # ── Clear existing sheets ─────────────────────────────────────────────
        for sn in ["Cover", "Assumptions", "Income Statement", "DCF Valuation"]:
            if sn in self.wb.sheetnames:
                del self.wb[sn]
            self._charts[sn] = []

        today = date.today().strftime("%Y-%m-%d")
        base_year = date.today().year - 1
        year_labels = [f"FY{base_year}A"] + [f"FY{base_year+i}E" for i in range(1, n_years + 1)]

        # ════════════════════════════════════════════════════════════════════
        # SHEET 1 — Cover
        # ════════════════════════════════════════════════════════════════════
        ws_cov = self.wb.create_sheet("Cover")
        ws_cov.column_dimensions["A"].width = 28
        ws_cov.column_dimensions["B"].width = 35

        # Row 1 = header (Label | Value) — serializer reads this as column headers
        for c, hdr in enumerate(["Label", "Value"], start=1):
            cell = ws_cov.cell(row=1, column=c, value=hdr)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(fill_type="solid", fgColor="1F4E79")
            cell.alignment = Alignment(horizontal="center")
            cell.border = _BORDER

        cover_rows = [
            ("Model Title",       f"DCF Valuation — {company}"),
            ("Model Date",        today),
            ("Currency",          currency),
            ("Projection Period", f"{n_years} Years"),
            ("Valuation Method",  "Discounted Cash Flow — Gordon Growth Model"),
            ("Analyst",           "RightCut AI"),
        ]
        # Data rows start at row 2
        for r, (label, value) in enumerate(cover_rows, start=2):
            ws_cov.cell(row=r, column=1, value=label).border = _BORDER
            ws_cov.cell(row=r, column=2, value=value).border = _BORDER

        # Navy highlight on Model Title row (row 2)
        for c in (1, 2):
            cell = ws_cov.cell(row=2, column=c)
            cell.font = Font(bold=True, color="FFFFFF", size=12)
            cell.fill = PatternFill(fill_type="solid", fgColor="1F3864")
            cell.alignment = Alignment(horizontal="left", vertical="center")

        # ════════════════════════════════════════════════════════════════════
        # SHEET 2 — Assumptions  (values in column B, row 2 onward)
        # Row map (1-indexed):
        #   1=header, 2=section, 3=rev_growth, 4=section, 5=ebitda_margin,
        #   6=da_pct, 7=capex_pct, 8=wc_pct, 9=tax_rate,
        #   10=section, 11=wacc, 12=term_g,
        #   13=section, 14=base_rev, 15=base_ebitda, 16=net_debt, 17=shares
        # ════════════════════════════════════════════════════════════════════
        ws_ass = self.wb.create_sheet("Assumptions")
        ws_ass.column_dimensions["A"].width = 32
        ws_ass.column_dimensions["B"].width = 18

        # Header row
        for c, hdr in enumerate(["Assumption", "Value"], start=1):
            cell = ws_ass.cell(row=1, column=c, value=hdr)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(fill_type="solid", fgColor="1F4E79")
            cell.alignment = Alignment(horizontal="center")
            cell.border = _BORDER

        _SEC = "D9E2F3"   # light blue section header bg
        _DARK = "1F3864"  # dark navy text

        ass_data = [
            # (row, label, value, is_section, number_format)
            (2,  "REVENUE ASSUMPTIONS",    None,         True,  None),
            (3,  "Revenue Growth Rate",    rev_growth,   False, "0.0%"),
            (4,  "MARGIN ASSUMPTIONS",     None,         True,  None),
            (5,  "EBITDA Margin",          ebitda_margin,False, "0.0%"),
            (6,  "D&A as % of Revenue",    da_pct,       False, "0.0%"),
            (7,  "CapEx as % of Revenue",  capex_pct,    False, "0.0%"),
            (8,  "Change in WC as % Rev",  wc_pct,       False, "0.0%"),
            (9,  "Corporate Tax Rate",     tax_rate,     False, "0.0%"),
            (10, "WACC / DISCOUNT ASSUMPTIONS", None,    True,  None),
            (11, "WACC",                   wacc,         False, "0.0%"),
            (12, "Terminal Growth Rate",   term_g,       False, "0.0%"),
            (13, "BALANCE SHEET INPUTS",   None,         True,  None),
            (14, "Base Revenue (FY2024A)", base_revenue, False, "#,##0"),
            (15, "Base EBITDA (FY2024A)",  base_ebitda,  False, "#,##0"),
            (16, "Net Debt",               net_debt,     False, "#,##0"),
            (17, "Shares Outstanding",     shares,       False, "#,##0"),
        ]

        # Named row pointers (for formula references)
        # B3=rev_growth, B5=ebitda_margin, B6=da_pct, B7=capex_pct,
        # B8=wc_pct, B9=tax_rate, B11=wacc, B12=term_g,
        # B14=base_rev, B15=base_ebitda, B16=net_debt, B17=shares
        A = {
            "rev_growth":    "Assumptions!$B$3",
            "ebitda_margin": "Assumptions!$B$5",
            "da_pct":        "Assumptions!$B$6",
            "capex_pct":     "Assumptions!$B$7",
            "wc_pct":        "Assumptions!$B$8",
            "tax_rate":      "Assumptions!$B$9",
            "wacc":          "Assumptions!$B$11",
            "term_g":        "Assumptions!$B$12",
            "base_rev":      "Assumptions!$B$14",
            "base_ebitda":   "Assumptions!$B$15",
            "net_debt":      "Assumptions!$B$16",
            "shares":        "Assumptions!$B$17",
        }

        for row, label, value, is_section, fmt in ass_data:
            cell_a = ws_ass.cell(row=row, column=1, value=label)
            cell_b = ws_ass.cell(row=row, column=2, value=value)
            cell_a.border = _BORDER
            cell_b.border = _BORDER
            if is_section:
                for c in (cell_a, cell_b):
                    c.font = Font(bold=True, color=_DARK, size=11)
                    c.fill = PatternFill(fill_type="solid", fgColor=_SEC)
                    c.alignment = Alignment(horizontal="left", vertical="center")
            elif fmt:
                cell_b.number_format = fmt

        # ════════════════════════════════════════════════════════════════════
        # SHEET 3 — Income Statement
        # Col A=label, B=base year, C..=projected years
        # Row map: 1=header, 2=rev, 3=ebitda, 4=da, 5=ebit, 6=nopat,
        #          7=capex, 8=da_addback, 9=dwc, 10=FCFF
        # ════════════════════════════════════════════════════════════════════
        ws_is = self.wb.create_sheet("Income Statement")
        ws_is.column_dimensions["A"].width = 30
        for i in range(1, n_years + 2):
            ws_is.column_dimensions[get_column_letter(i + 1)].width = 14

        # Header row
        is_headers = ["Line Item"] + year_labels
        for c, hdr in enumerate(is_headers, start=1):
            cell = ws_is.cell(row=1, column=c, value=hdr)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(fill_type="solid", fgColor="1F4E79")
            cell.alignment = Alignment(horizontal="center")
            cell.border = _BORDER

        # Column letters: B=base, C=yr1, D=yr2, ...
        col_base = "B"   # FY2024A
        col_yrs  = [get_column_letter(i + 3) for i in range(n_years)]  # C, D, E, F, G

        # Build all IS rows directly (no lambdas — avoids closure bugs)
        # col_yrs[0]=C (yr1), col_yrs[1]=D (yr2), ...
        # Prior year for revenue: base=B, yr1 prior=B, yr2 prior=C, ...
        rev_prior = [col_base] + col_yrs[:-1]  # B, C, D, E, F

        # Row 2: Revenue
        ws_is.cell(row=2, column=1, value="Revenue").border = _BORDER
        ws_is.cell(row=2, column=2, value=base_revenue).border = _BORDER
        ws_is.cell(row=2, column=2).number_format = "#,##0"
        for i, col in enumerate(col_yrs):
            prior = rev_prior[i]
            cell = ws_is.cell(row=2, column=3 + i, value=f"={prior}2*(1+{A['rev_growth']})")
            cell.number_format = "#,##0"
            cell.border = _BORDER

        def _is_rows_3_to_10():
            for row_num, label, b_tmpl, proj_tmpl in [
                (3,  "EBITDA",                   f"={col_base}2*{A['ebitda_margin']}", "{col}2*{em}"),
                (4,  "D&A",                       f"={col_base}2*{A['da_pct']}",       "{col}2*{dp}"),
                (5,  "EBIT",                      f"={col_base}3-{col_base}4",          "{col}3-{col}4"),
                (6,  "NOPAT",                     f"={col_base}5*(1-{A['tax_rate']})",  "{col}5*(1-{tr})"),
                (7,  "Capital Expenditures",      f"=-{col_base}2*{A['capex_pct']}",   "-{col}2*{cp}"),
                (8,  "D&A Add-back",              f"={col_base}4",                      "{col}4"),
                (9,  "Change in Working Capital", f"=-{col_base}2*{A['wc_pct']}",      "-{col}2*{wp}"),
                (10, "FCFF",                      f"={col_base}6+{col_base}8+{col_base}7+{col_base}9",
                                                                                         "{col}6+{col}8+{col}7+{col}9"),
            ]:
                ws_is.cell(row=row_num, column=1, value=label).border = _BORDER
                ws_is.cell(row=row_num, column=2, value=b_tmpl).border = _BORDER
                ws_is.cell(row=row_num, column=2).number_format = "#,##0"
                for i, col in enumerate(col_yrs):
                    formula = "=" + proj_tmpl.format(
                        col=col,
                        em=A['ebitda_margin'], dp=A['da_pct'], tr=A['tax_rate'],
                        cp=A['capex_pct'], wp=A['wc_pct'],
                    )
                    cell = ws_is.cell(row=row_num, column=3 + i, value=formula)
                    cell.number_format = "#,##0"
                    cell.border = _BORDER

        _is_rows_3_to_10()

        # Formatting
        _green_fill = PatternFill(fill_type="solid", fgColor="E2EFDA")
        _bold_green = Font(bold=True, color="375623", size=11)
        for c in range(1, n_years + 3):
            # EBITDA row (3) — output_row
            cell3 = ws_is.cell(row=3, column=c)
            cell3.fill = _green_fill
            cell3.font = _bold_green
            # FCFF row (10) — output_row
            cell10 = ws_is.cell(row=10, column=c)
            cell10.fill = _green_fill
            cell10.font = _bold_green
            # EBIT row (5) — subtotal_row
            ws_is.cell(row=5, column=c).font = Font(bold=True, size=11)

        ws_is.freeze_panes = "B2"

        # ════════════════════════════════════════════════════════════════════
        # SHEET 4 — DCF Valuation
        # Only projected years (no base year column)
        # Col A=label, B=yr1 … F=yr5
        # Row map:
        #   1=header
        #   2=FCFF (from IS)
        #   3=discount factor
        #   4=PV FCFF
        #   5=blank
        #   6=section: TERMINAL VALUE
        #   7=terminal FCFF
        #   8=terminal value
        #   9=PV terminal value
        #   10=blank
        #   11=section: EQUITY VALUE BRIDGE
        #   12=sum PV FCFFs
        #   13=enterprise value
        #   14=less net debt
        #   15=equity value
        #   16=shares outstanding
        #   17=INTRINSIC VALUE / SHARE  ← final_answer_row
        # ════════════════════════════════════════════════════════════════════
        ws_dcf = self.wb.create_sheet("DCF Valuation")
        ws_dcf.column_dimensions["A"].width = 30
        for i in range(1, n_years + 1):
            ws_dcf.column_dimensions[get_column_letter(i + 1)].width = 14

        # Projected year column letters: B=yr1, C=yr2, ...
        dcf_cols = [get_column_letter(i + 2) for i in range(n_years)]  # B, C, D, E, F

        # Header row (projected years only)
        dcf_headers = ["Metric"] + year_labels[1:]  # skip base year
        for c, hdr in enumerate(dcf_headers, start=1):
            cell = ws_dcf.cell(row=1, column=c, value=hdr)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(fill_type="solid", fgColor="1F4E79")
            cell.alignment = Alignment(horizontal="center")
            cell.border = _BORDER

        def dcf_cell(row: int, col_letter: str, value, fmt="#,##0"):
            cell = ws_dcf.cell(row=row, column=ord(col_letter) - ord("A") + 1, value=value)
            cell.border = _BORDER
            if fmt:
                cell.number_format = fmt
            return cell

        def dcf_label(row: int, label: str):
            cell = ws_dcf.cell(row=row, column=1, value=label)
            cell.border = _BORDER
            return cell

        # FCFF from Income Statement (projected years = IS cols C, D, E, F, G)
        dcf_label(2, "FCFF")
        is_proj_cols = [get_column_letter(i + 3) for i in range(n_years)]  # C, D, E, F, G
        for i, (dcf_col, is_col) in enumerate(zip(dcf_cols, is_proj_cols)):
            dcf_cell(2, dcf_col, f"='Income Statement'!{is_col}10")

        # Discount factor = 1/(1+WACC)^n
        dcf_label(3, "Discount Factor")
        for i, col in enumerate(dcf_cols):
            dcf_cell(3, col, f"=1/(1+{A['wacc']})^{i+1}", "0.0000")

        # PV of FCFF = FCFF * DF
        dcf_label(4, "PV of FCFF")
        for i, col in enumerate(dcf_cols):
            dcf_cell(4, col, f"={col}2*{col}3")

        # Blank row 5
        dcf_label(5, "")

        # Section header row 6
        dcf_label(6, "TERMINAL VALUE CALCULATION")
        for c in range(1, n_years + 2):
            cell = ws_dcf.cell(row=6, column=c)
            cell.font = Font(bold=True, color=_DARK, size=11)
            cell.fill = PatternFill(fill_type="solid", fgColor=_SEC)
            cell.border = _BORDER

        # Terminal FCFF = last projected FCFF * (1+g)
        last_col = dcf_cols[-1]
        dcf_label(7, "Terminal FCFF")
        dcf_cell(7, "B", f"={last_col}2*(1+{A['term_g']})")

        # Terminal Value = Terminal FCFF / (WACC - g)
        dcf_label(8, "Terminal Value")
        dcf_cell(8, "B", f"=B7/({A['wacc']}-{A['term_g']})")

        # PV of Terminal Value
        dcf_label(9, "PV of Terminal Value")
        dcf_cell(9, "B", f"=B8/(1+{A['wacc']})^{n_years}")

        # Blank row 10
        dcf_label(10, "")

        # Section header row 11
        dcf_label(11, "EQUITY VALUE BRIDGE")
        for c in range(1, n_years + 2):
            cell = ws_dcf.cell(row=11, column=c)
            cell.font = Font(bold=True, color=_DARK, size=11)
            cell.fill = PatternFill(fill_type="solid", fgColor=_SEC)
            cell.border = _BORDER

        # Sum of PV FCFFs
        pv_range = f"B4:{get_column_letter(n_years + 1)}4"
        dcf_label(12, "Sum of PV FCFFs")
        dcf_cell(12, "B", f"=SUM({pv_range})")

        # Enterprise Value
        dcf_label(13, "Enterprise Value")
        dcf_cell(13, "B", "=B12+B9")

        # Less: Net Debt
        dcf_label(14, "Less: Net Debt")
        dcf_cell(14, "B", f"={A['net_debt']}")

        # Equity Value
        dcf_label(15, "Equity Value")
        dcf_cell(15, "B", "=B13-B14")

        # Shares Outstanding
        dcf_label(16, "Shares Outstanding")
        dcf_cell(16, "B", f"={A['shares']}", "#,##0")

        # Intrinsic Value / Share — FINAL ANSWER
        dcf_label(17, "Intrinsic Value / Share")
        dcf_cell(17, "B", "=B15/B16", "#,##0.00")

        # Formatting for DCF
        _green_fill = PatternFill(fill_type="solid", fgColor="E2EFDA")
        _bold_green = Font(bold=True, color="375623", size=11)
        _navy_fill = PatternFill(fill_type="solid", fgColor="1F3864")
        _white_bold = Font(bold=True, color="FFFFFF", size=12)

        for c in range(1, n_years + 2):
            # PV FCFF row (4) — output_row
            ws_dcf.cell(row=4, column=c).fill = _green_fill
            ws_dcf.cell(row=4, column=c).font = _bold_green
            # Equity Value row (15) — output_row
            ws_dcf.cell(row=15, column=c).fill = _green_fill
            ws_dcf.cell(row=15, column=c).font = _bold_green

        # Final answer row (17) — dark navy
        for c in range(1, 3):  # just label + value columns
            cell = ws_dcf.cell(row=17, column=c)
            cell.fill = _navy_fill
            cell.font = _white_bold
            cell.alignment = Alignment(horizontal="left" if c == 1 else "right", vertical="center")

        ws_dcf.freeze_panes = "B2"

        # ── Summary ───────────────────────────────────────────────────────────
        sheets_created = ["Cover", "Assumptions", "Income Statement", "DCF Valuation"]
        for sn in sheets_created:
            self._charts.setdefault(sn, [])

        return {
            "sheets_created": sheets_created,
            "assumption_rows": {k: v for k, v in {
                "revenue_growth_row": 3, "ebitda_margin_row": 5, "da_pct_row": 6,
                "capex_pct_row": 7, "wc_pct_row": 8, "tax_rate_row": 9,
                "wacc_row": 11, "terminal_growth_row": 12,
                "base_revenue_row": 14, "base_ebitda_row": 15,
                "net_debt_row": 16, "shares_row": 17,
            }.items()},
            "income_statement_rows": {
                "revenue": 2, "ebitda": 3, "da": 4, "ebit": 5,
                "nopat": 6, "capex": 7, "da_addback": 8, "wc": 9, "fcff": 10,
            },
            "dcf_rows": {
                "fcff": 2, "discount_factor": 3, "pv_fcff": 4,
                "terminal_fcff": 7, "terminal_value": 8, "pv_tv": 9,
                "sum_pv_fcff": 12, "enterprise_value": 13,
                "net_debt": 14, "equity_value": 15,
                "shares": 16, "intrinsic_value_per_share": 17,
            },
            "note": "All formulas are circular-reference-free. Assumptions sheet drives IS and DCF. Edit Assumptions!B3-B17 to update model.",
        }

    # ── Data insertion ────────────────────────────────────────────────────────

    def insert_data(
        self,
        sheet_name: str,
        rows: list[list[str]],
        start_row: int = 2,
    ) -> dict:
        ws = self._get_sheet(sheet_name)
        max_col = ws.max_column or 1

        rows_written = 0
        for row_idx, row_data in enumerate(rows, start=start_row):
            for col_idx, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = value if value != "" else None
                cell.border = _BORDER
                cell.alignment = Alignment(vertical="center")

                # Auto-detect number format from header
                header_cell = ws.cell(row=1, column=col_idx)
                if header_cell.value:
                    fmt = infer_format(str(header_cell.value))
                    if fmt:
                        cell.number_format = fmt

            rows_written += 1

        return {
            "sheet_name": sheet_name,
            "rows_written": rows_written,
            "start_row": start_row,
        }

    # ── Formula insertion ─────────────────────────────────────────────────────

    def add_formula(
        self,
        sheet_name: str,
        cell: str,
        formula: str,
        apply_to_range: str | None = None,
    ) -> dict:
        ws = self._get_sheet(sheet_name)

        if not validate_formula(formula):
            return {"error": f"Invalid formula: {formula}", "success": False}

        ws[cell] = formula
        ws[cell].border = _BORDER

        applied_cells = [cell]

        if apply_to_range:
            # Derive the column offset pattern and apply across the range
            try:
                from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
                anchor_col_letter, anchor_row = coordinate_from_string(cell)
                anchor_col = column_index_from_string(anchor_col_letter)

                # Parse the target range
                start_ref, end_ref = apply_to_range.split(":")
                start_col_letter, start_row = coordinate_from_string(start_ref)
                end_col_letter, end_row = coordinate_from_string(end_ref)
                start_col = column_index_from_string(start_col_letter)
                end_col = column_index_from_string(end_col_letter)

                for r in range(start_row, end_row + 1):
                    for c in range(start_col, end_col + 1):
                        if r == anchor_row and c == anchor_col:
                            continue  # already written
                        # Shift formula row references
                        row_delta = r - anchor_row
                        shifted = _shift_formula_rows(formula, row_delta)
                        target_cell = f"{get_column_letter(c)}{r}"
                        ws[target_cell] = shifted
                        ws[target_cell].border = _BORDER
                        applied_cells.append(target_cell)
            except Exception as e:
                logger.warning(f"apply_to_range failed: {e}")

        return {
            "sheet_name": sheet_name,
            "formula": formula,
            "cells": applied_cells,
        }

    # ── Cell editing ──────────────────────────────────────────────────────────

    def edit_cell(self, sheet_name: str, cell: str, value: str) -> dict:
        ws = self._get_sheet(sheet_name)
        old_value = ws[cell].value
        ws[cell].value = value
        ws[cell].border = _BORDER
        return {"sheet_name": sheet_name, "cell": cell, "old": str(old_value), "new": value}

    def apply_user_edit(self, sheet_name: str, cell: str, value: str) -> None:
        """Apply a user-initiated cell edit (no response needed)."""
        try:
            ws = self._get_sheet(sheet_name)
            ws[cell].value = value
        except Exception as e:
            logger.warning(f"apply_user_edit failed for {sheet_name}!{cell}: {e}")

    # ── Formatting ────────────────────────────────────────────────────────────

    def apply_formatting(
        self,
        sheet_name: str,
        cell_range: str,
        format_type: str,
        format_config: dict | None = None,
    ) -> dict:
        ws = self._get_sheet(sheet_name)
        cfg = format_config or {}

        if format_type == "color_scale":
            rule = ColorScaleRule(
                start_type="min",
                start_color="F8696B",
                mid_type="percentile",
                mid_value=50,
                mid_color="FFEB84",
                end_type="max",
                end_color="63BE7B",
            )
            ws.conditional_formatting.add(cell_range, rule)

        elif format_type == "data_bar":
            rule = DataBarRule(
                start_type="min",
                start_value=0,
                end_type="max",
                end_value=100,
                color="638EC6",
            )
            ws.conditional_formatting.add(cell_range, rule)

        elif format_type == "bold_header":
            for row in ws[cell_range]:
                for cell in row:
                    cell.font = Font(bold=True, color="FFFFFF", size=11)
                    cell.fill = PatternFill(fill_type="solid", fgColor="1F4E79")
                    cell.alignment = Alignment(horizontal="center")
                    cell.border = _BORDER

        elif format_type == "number_format":
            fmt = cfg.get("format", "#,##0")
            for row in ws[cell_range]:
                for cell in row:
                    cell.number_format = fmt

        elif format_type == "border":
            style = cfg.get("style", "thin")
            side = Side(style=style)
            border = Border(left=side, right=side, top=side, bottom=side)
            for row in ws[cell_range]:
                for cell in row:
                    cell.border = border

        elif format_type == "font_color":
            color = cfg.get("color", "FF0000").lstrip("#")
            for row in ws[cell_range]:
                for cell in row:
                    cell.font = Font(color=color)

        elif format_type == "background_color":
            color = cfg.get("color", "FFFF00").lstrip("#")
            for row in ws[cell_range]:
                for cell in row:
                    cell.fill = PatternFill(fill_type="solid", fgColor=color)

        elif format_type == "section_header":
            # Light-blue section divider row with bold dark text — use between logical blocks
            for row in ws[cell_range]:
                for cell in row:
                    cell.font = Font(bold=True, color="1F3864", size=11)
                    cell.fill = PatternFill(fill_type="solid", fgColor="D9E2F3")
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    cell.border = _BORDER

        elif format_type == "subtotal_row":
            # Bold text, no background — for totals, sub-totals, computed summary rows
            for row in ws[cell_range]:
                for cell in row:
                    cell.font = Font(bold=True, size=11)
                    cell.border = _BORDER

        elif format_type == "output_row":
            # Light green background — for key output rows (FCFF, Net Profit, EBITDA)
            for row in ws[cell_range]:
                for cell in row:
                    cell.font = Font(bold=True, color="375623", size=11)
                    cell.fill = PatternFill(fill_type="solid", fgColor="E2EFDA")
                    cell.border = _BORDER

        elif format_type == "final_answer_row":
            # Dark navy background with white text — for the single most important output row
            for row in ws[cell_range]:
                for cell in row:
                    cell.font = Font(bold=True, color="FFFFFF", size=12)
                    cell.fill = PatternFill(fill_type="solid", fgColor="1F3864")
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    cell.border = _BORDER

        elif format_type == "zebra_stripe":
            # Alternate light-gray rows for tables (apply to even rows within range)
            for row in ws[cell_range]:
                for cell in row:
                    if cell.row % 2 == 0:
                        cell.fill = PatternFill(fill_type="solid", fgColor="F2F2F2")

        elif format_type == "muted_row":
            # Grey font, lighter weight — for sub-commentary rows (growth %, margin %)
            for row in ws[cell_range]:
                for cell in row:
                    cell.font = Font(color="555555", size=10)

        else:
            return {"error": f"Unknown format_type: {format_type}", "success": False}

        return {"sheet_name": sheet_name, "range": cell_range, "format_type": format_type}

    # ── Citations ─────────────────────────────────────────────────────────────

    def add_citation(
        self,
        sheet_name: str,
        cell: str,
        source_file: str,
        source_location: str,
        excerpt: str = "",
    ) -> dict:
        ws = self._get_sheet(sheet_name)
        comment_text = f"Source: {source_file}\nRef: {source_location}"
        if excerpt:
            comment_text += f"\n\"{excerpt[:200]}\""
        comment = Comment(comment_text, "RightCut Agent")
        ws[cell].comment = comment
        return {
            "sheet_name": sheet_name,
            "cell": cell,
            "citation": comment_text,
        }

    # ── Sorting ───────────────────────────────────────────────────────────────

    def sort_range(
        self,
        sheet_name: str,
        sort_column: str,
        ascending: bool = True,
    ) -> dict:
        ws = self._get_sheet(sheet_name)
        max_row = ws.max_row
        max_col = ws.max_column

        if max_row < 3:
            return {"sheet_name": sheet_name, "sorted": False, "reason": "not enough rows"}

        # Find the sort column index
        sort_col_idx: int | None = None
        for col_idx in range(1, max_col + 1):
            header = ws.cell(row=1, column=col_idx).value
            if header and str(header).strip().lower() == sort_column.strip().lower():
                sort_col_idx = col_idx
                break

        if sort_col_idx is None:
            # Try treating sort_column as a letter reference
            try:
                sort_col_idx = column_index_from_string(sort_column.upper())
            except Exception:
                return {"error": f"Column '{sort_column}' not found", "success": False}

        # Read data rows (skip header)
        data_rows: list[list[Any]] = []
        for row_idx in range(2, max_row + 1):
            row_data = [ws.cell(row=row_idx, column=c).value for c in range(1, max_col + 1)]
            data_rows.append(row_data)

        # Sort — try numeric, fall back to string
        def sort_key(row: list[Any]) -> tuple:
            val = row[sort_col_idx - 1]
            if val is None:
                return (1, "")
            try:
                return (0, float(str(val).replace("$", "").replace(",", "").replace("%", "")))
            except (ValueError, TypeError):
                return (0, str(val).lower())

        data_rows.sort(key=sort_key, reverse=not ascending)

        # Write back
        for row_idx, row_data in enumerate(data_rows, start=2):
            for col_idx, value in enumerate(row_data, start=1):
                ws.cell(row=row_idx, column=col_idx).value = value

        return {
            "sheet_name": sheet_name,
            "sort_column": sort_column,
            "ascending": ascending,
            "rows_sorted": len(data_rows),
        }

    # ── Charts ────────────────────────────────────────────────────────────────

    def create_chart(
        self,
        sheet_name: str,
        chart_type: str,
        data_range: str,
        title: str = "",
        target_cell: str = "H2",
    ) -> dict:
        ws = self._get_sheet(sheet_name)

        # Parse data range
        try:
            start_ref, end_ref = data_range.split(":")
            from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
            start_col_letter, start_row = coordinate_from_string(start_ref)
            end_col_letter, end_row = coordinate_from_string(end_ref)
            min_col = column_index_from_string(start_col_letter)
            max_col = column_index_from_string(end_col_letter)
        except Exception as e:
            return {"error": f"Invalid data_range '{data_range}': {e}", "success": False}

        # Build chart
        chart_map = {
            "bar": BarChart,
            "line": LineChart,
            "pie": PieChart,
            "scatter": ScatterChart,
        }
        ChartClass = chart_map.get(chart_type.lower(), BarChart)
        chart = ChartClass()
        chart.title = title or sheet_name
        chart.style = 10
        chart.width = 18
        chart.height = 10

        if chart_type.lower() == "bar":
            chart.type = "col"   # vertical bars (horizontal is the confusing default)
            chart.grouping = "clustered"

        data_ref = Reference(
            ws,
            min_col=min_col,
            min_row=start_row,
            max_col=max_col,
            max_row=end_row,
        )

        if chart_type.lower() == "scatter":
            chart.add_data(data_ref)
        else:
            chart.add_data(data_ref, titles_from_data=True)

        # Categories from first column if more than one column
        if max_col > min_col:
            cats = Reference(ws, min_col=min_col, min_row=start_row + 1, max_row=end_row)
            chart.set_categories(cats)

        ws.add_chart(chart, target_cell)

        # Track chart metadata for frontend
        meta = ChartMeta(
            chart_type=chart_type,
            title=title or sheet_name,
            data_range=data_range,
            anchor_cell=target_cell,
        )
        self._charts.setdefault(sheet_name, []).append(meta)

        return {
            "sheet_name": sheet_name,
            "chart_type": chart_type,
            "title": chart.title,
            "anchor": target_cell,
            "note": "Chart embedded in .xlsx. Download to view rendered chart.",
        }

    # ── Validation ────────────────────────────────────────────────────────────

    def validate_workbook(self, check_hardcoded: bool = True) -> dict:
        stats = ValidationStats()
        hardcoded_cells: list[str] = []

        for sheet_name in self.wb.sheetnames:
            ws = self.wb[sheet_name]
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if cell.value is None:
                        continue
                    stats.total_cells += 1
                    val_str = str(cell.value)
                    if val_str.startswith("="):
                        stats.formula_count += 1
                    elif check_hardcoded:
                        # Flag numeric cells that look like they could be formulas
                        try:
                            float(val_str.replace("$", "").replace(",", "").replace("%", ""))
                            # It's a number — flag if the header suggests it should be a formula
                            header = ws.cell(row=1, column=cell.column).value
                            if header and any(
                                k in str(header).lower()
                                for k in ("cagr", "margin", "irr", "moic", "growth", "ratio")
                            ):
                                ref = f"{sheet_name}!{cell.coordinate}"
                                hardcoded_cells.append(ref)
                                stats.hardcoded_count += 1
                        except (ValueError, TypeError):
                            pass

        stats.issues = [f"Hardcoded value in {c} (consider using a formula)" for c in hardcoded_cells[:10]]

        return {
            "formula_count": stats.formula_count,
            "hardcoded_count": stats.hardcoded_count,
            "total_cells": stats.total_cells,
            "issues": stats.issues,
            "valid": stats.hardcoded_count == 0,
            "summary": (
                f"{stats.formula_count} formulas, "
                f"{stats.hardcoded_count} potentially hardcoded values, "
                f"{stats.total_cells} total data cells"
            ),
        }

    # ── Sheet state (for agent) ───────────────────────────────────────────────

    def get_sheet_state(self, sheet_name: str) -> dict:
        ws = self._get_sheet(sheet_name)
        headers = []
        rows = []

        for row_idx, row in enumerate(ws.iter_rows(values_only=False), start=1):
            row_data = []
            for cell in row:
                val = cell.value
                if val is None:
                    row_data.append({"value": None, "formula": None})
                elif isinstance(val, str) and val.startswith("="):
                    row_data.append({"value": None, "formula": val})
                else:
                    row_data.append({"value": str(val), "formula": None})

            if row_idx == 1:
                headers = [str(c["value"] or "") for c in row_data]
            else:
                rows.append(row_data)

        return {
            "sheet_name": sheet_name,
            "headers": headers,
            "row_count": len(rows),
            "rows": rows[:50],  # cap at 50 rows for context window
        }

    # ── Download ──────────────────────────────────────────────────────────────

    def to_bytes(self) -> bytes:
        buf = io.BytesIO()
        self.wb.save(buf)
        buf.seek(0)
        return buf.read()

    # ── Internal helpers ──────────────────────────────────────────────────────

    def _get_sheet(self, sheet_name: str):
        if sheet_name not in self.wb.sheetnames:
            # Auto-create if referenced but not yet created
            ws = self.wb.create_sheet(title=sheet_name)
            self._charts[sheet_name] = []
            return ws
        return self.wb[sheet_name]


def _shift_formula_rows(formula: str, delta: int) -> str:
    """Naively shift absolute row numbers in a formula by delta rows."""
    import re
    if delta == 0:
        return formula

    def replace_ref(m: re.Match) -> str:
        col = m.group(1)
        row = int(m.group(2)) + delta
        return f"{col}{row}"

    # Match cell references like A1, B3, AA10 (no $ prefix for simplicity)
    return re.sub(r"([A-Z]+)(\d+)", replace_ref, formula)
