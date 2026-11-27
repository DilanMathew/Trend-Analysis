"""Builds Advanced_Milestone_Trend_Analysis.xlsx.

Run:  python tools/build_workbook.py
Then: recalculate (LibreOffice / Excel) - Excel recalculates automatically on open.
"""
import os
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, ScatterChart, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.legend import Legend
from openpyxl.chart.marker import DataPoint, Marker
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.series import SeriesLabel, StrRef
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (CharacterProperties, Font as DFont, Paragraph, ParagraphProperties,
                                   RegularTextRun)
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableColumn, TableFormula, TableStyleInfo

sys.path.insert(0, os.path.dirname(__file__))
import sample_data  # noqa: E402

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "Advanced_Milestone_Trend_Analysis.xlsx")

# ----------------------------------------------------------------------------- style
FONT = "Arial"
NAVY = "1F3864"
ACCENT = "2F5597"
BG = "F3F5F9"
GRID = "D9DEE7"
TXT = "404040"
MUTED = "7F8691"
INPUT_FILL = "FFF7D6"      # pale yellow: editable cells
CALC_HDR = "5B6577"
RED, GREEN, GREYC = "C0392B", "1E8449", "7F8C8D"
STATUS_COL = {"Completed": "1F4E79", "Delayed": "C0392B", "Ahead of Schedule": "1E8449",
              "On Schedule": "5B9BD5", "Not Yet Started": "A6A6A6"}
STATUSES = ["Completed", "Delayed", "Ahead of Schedule", "On Schedule", "Not Yet Started"]
SERIES_COLORS = ["1F77B4", "FF7F0E", "2CA02C", "D62728", "9467BD", "8C564B", "E377C2", "17BECF",
                 "BCBD22", "393B79", "AD494A", "637939", "E7969C", "7B4173", "3182BD"]
NSLOT = 15          # milestone lines on the MTA chart
NREP = 40           # reporting periods per project handled by the chart engine
DA_FIRST, DA_N = 16, 200
DA_LAST = DA_FIRST + DA_N - 1
DATE = "dd-mmm-yyyy"
DATE_S = "dd-mmm-yy"
DATE_HIDE0 = "dd-mmm-yyyy;;;"
DAYS = '+0" d";-0" d";0" d"'

thin = Side(style="thin", color=GRID)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
UNLOCK = Protection(locked=False)


def f(size=10, bold=False, color=TXT, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def setw(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def title_block(ws, title, subtitle, last_col="P"):
    ws["A1"].value = None
    ws["B1"] = title
    ws["B1"].font = f(18, True, NAVY)
    ws["B2"] = subtitle
    ws["B2"].font = f(10, False, MUTED, italic=True)
    ws.row_dimensions[1].height = 30
    ws.sheet_view.showGridLines = False


def section(ws, ref, text, span=None):
    c = ws[ref]
    c.value = text
    c.font = f(11, True, "FFFFFF")
    c.fill = fill(NAVY)
    c.alignment = Alignment(vertical="center", indent=1)
    if span:
        ws.merge_cells(f"{ref}:{span}")


def hdr(cell, calc=False):
    cell.font = f(10, True, "FFFFFF")
    cell.fill = fill(CALC_HDR if calc else NAVY)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER


def inp(cell):
    cell.fill = fill(INPUT_FILL)
    cell.protection = UNLOCK
    cell.border = BORDER


def add_name(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def add_table(ws, name, ref, columns, style="TableStyleMedium2"):
    """columns: list of (header, calculated_formula_or_None)."""
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True)
    cols = []
    for i, (h, fml) in enumerate(columns, 1):
        tc = TableColumn(id=i, name=h)
        if fml:
            tc.calculatedColumnFormula = TableFormula(attr_text=fml.lstrip("="))
        cols.append(tc)
    t.tableColumns = cols
    t.autoFilter = AutoFilter(ref=ref)
    ws.add_table(t)
    return t


# chart helpers ---------------------------------------------------------------
def rich(sz=800, bold=False, color="595959"):
    cp = CharacterProperties(sz=sz, b=bold, solidFill=color, latin=DFont(typeface=FONT))
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def ctitle(text, sz=1100):
    cp = CharacterProperties(sz=sz, b=True, solidFill=NAVY, latin=DFont(typeface=FONT))
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=text)])
    return Title(tx=Text(rich=RichText(p=[para])), overlay=False)


def gridlines():
    return ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill="E3E6EB", w=6350)))


def style_axes(ch, xfmt=None, yfmt=None, xtitle=None, ytitle=None):
    for ax, fmt, t in ((ch.x_axis, xfmt, xtitle), (ch.y_axis, yfmt, ytitle)):
        ax.delete = False
        ax.txPr = rich(800)
        if fmt:
            ax.number_format = fmt
            ax.numFmt.sourceLinked = False
        if t:
            ax.title = ctitle(t, 850)
        ax.spPr = GraphicalProperties(ln=LineProperties(solidFill="BFBFBF"))


def finish_chart(ch, w, h):
    ch.width, ch.height = w, h
    ch.roundedCorners = False
    ch.display_blanks = "gap"
    ch.visible_cells_only = False   # keep working if the Engine sheet is hidden
    if ch.legend is not None:
        ch.legend.overlay = False
        ch.legend.txPr = rich(800)


def series_title(s, ref):
    s.tx = SeriesLabel(strRef=StrRef(ref))


# ----------------------------------------------------------------------------- data refs
T = "tblMilestoneData"
MID, REP, FC, ACT = f"{T}[Milestone ID]", f"{T}[Reporting Date]", f"{T}[Forecast Completion Date]", f"{T}[Actual Completion Date]"
PID, KEY, MOVE = f"{T}[Project ID]", f"{T}[Obs Key]", f"{T}[Movement Days]"


def tr(col, table=T):
    return f"{table}[[#This Row],[{col}]]"


def main():
    projects, milestones, obs = sample_data.build()
    wb = Workbook()
    ws_dash = wb.active
    ws_dash.title = "Dashboard"
    ws_data = wb.create_sheet("Milestone Data")
    ws_set = wb.create_sheet("Project Setup")
    ws_reg = wb.create_sheet("Milestone Register")
    ws_rh = wb.create_sheet("Reporting History")
    ws_da = wb.create_sheet("Delay Analysis")
    ws_rpt = wb.create_sheet("Status Report")
    ws_ins = wb.create_sheet("Instructions")
    ws_eng = wb.create_sheet("Engine")

    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = {"Dashboard": NAVY, "Engine": "A6A6A6", "Instructions": "7F7F7F"}.get(ws.title, ACCENT)

    # ========================================================================= PROJECT SETUP
    ws = ws_set
    title_block(ws, "Project Setup", "Projects, analysis thresholds and reference lists. Yellow cells are editable; the tables expand automatically.")
    section(ws, "B4", "Projects  (tblProjects)", "J4")
    pcols = ["Project ID", "Project Name", "Project Manager", "Sponsor", "Department", "Start Date",
             "Planned Finish Date", "Reporting Frequency", "Description"]
    for j, h in enumerate(pcols):
        ws.cell(5, 2 + j, h)
        hdr(ws.cell(5, 2 + j))
    for i, p in enumerate(projects):
        r = 6 + i
        vals = [p["id"], p["name"], p["pm"], p["sponsor"], p["dept"], p["start"], p["finish"], p["freq"], p["desc"]]
        for j, v in enumerate(vals):
            c = ws.cell(r, 2 + j, v)
            c.font = f()
            if j in (5, 6):
                c.number_format = DATE
    add_table(ws, "tblProjects", f"B5:J{5 + len(projects)}", [(h, None) for h in pcols])
    setw(ws, {"A": 2, "B": 12, "C": 30, "D": 18, "E": 24, "F": 22, "G": 13, "H": 15, "I": 16, "J": 60})

    section(ws, "B12", "Analysis thresholds  (named cells used by every calculation)", "J12")
    cfg = [
        ("cfgTolerance", "On-schedule tolerance (calendar days)", 3,
         "A milestone is On Schedule while |current forecast - baseline| <= this value."),
        ("cfgSigMove", "Significant movement threshold (days)", 10,
         "A single-period forecast change of at least this many days is highlighted on the MTA chart."),
        ("cfgSlipCount", "Repeated slippage: number of slips", 3,
         "Risk flag when a forecast moved later at least this many times."),
        ("cfgHighVar", "High baseline variance (days)", 20,
         "Risk flag when current forecast is at least this many days after baseline."),
        ("cfgWindow", "Approaching deadline window (days)", 21,
         "Risk flag when an open milestone is due within this many days (or overdue)."),
        ("cfgAmberPct", "Overall status AMBER: share of delayed milestones", 0.10,
         "Project shows AMBER at or above this share of delayed milestones."),
        ("cfgRedPct", "Overall status RED: share of delayed milestones", 0.30,
         "Project shows RED at or above this share of delayed milestones."),
    ]
    for col, h in (("B", "Parameter"), ("D", "Value"), ("E", "Name"), ("F", "How it is used")):
        hdr(ws[f"{col}13"])
        ws[f"{col}13"] = h
    ws.merge_cells("B13:C13")
    ws.merge_cells("F13:J13")
    for i, (nm, label, val, note) in enumerate(cfg):
        r = 14 + i
        ws.cell(r, 2, label).font = f()
        ws.merge_cells(f"B{r}:C{r}")
        c = ws.cell(r, 4, val)
        inp(c)
        c.font = f(10, True, "0000FF")
        c.number_format = "0%" if isinstance(val, float) else "0"
        ws.cell(r, 5, nm).font = f(9, False, MUTED)
        ws.cell(r, 6, note).font = f(9, False, MUTED)
        ws.merge_cells(f"F{r}:J{r}")
        add_name(wb, nm, f"'Project Setup'!$D${r}")

    # reference lists
    section(ws, "B23", "Reference lists  (drive the dropdowns)", "G23")
    lists = [("tblCategories", "Category", sample_data.CATEGORIES, "B"),
             ("tblPriorities", "Priority", sample_data.PRIORITIES, "D"),
             ("tblDepartments", "Department", sample_data.DEPARTMENTS, "F")]
    for tname, h, vals, col in lists:
        ws[f"{col}24"] = h
        hdr(ws[f"{col}24"])
        for i, v in enumerate(vals):
            ws[f"{col}{25 + i}"] = v
            ws[f"{col}{25 + i}"].font = f()
        add_table(ws, tname, f"{col}24:{col}{24 + len(vals)}", [(h, None)], "TableStyleLight9")
    ws["H24"] = "Statuses (fixed)"
    hdr(ws["H24"], True)
    for i, s in enumerate(STATUSES):
        ws[f"H{25 + i}"] = s
        ws[f"H{25 + i}"].font = f()
    ws.column_dimensions["H"].width = 18
    add_name(wb, "lstCategory", "tblCategories[Category]")
    add_name(wb, "lstPriority", "tblPriorities[Priority]")
    add_name(wb, "lstDepartment", "tblDepartments[Department]")
    add_name(wb, "lstProjectID", "tblProjects[Project ID]")
    add_name(wb, "lstProjectName", "tblProjects[Project Name]")
    ws.freeze_panes = "A6"

    # ========================================================================= MILESTONE REGISTER
    ws = ws_reg
    title_block(ws, "Milestone Register  (tblMilestones)",
                "Master list of milestones - one row per unique Milestone ID. Navy headers = input, grey headers = calculated.")
    reg_cols = [
        ("Project ID", None), ("Project Name",
         f'=IFERROR(INDEX(tblProjects[Project Name],MATCH({tr("Project ID","tblMilestones")},tblProjects[Project ID],0)),"")'),
        ("Milestone ID", None), ("Milestone Name", None), ("Milestone Description", None), ("Baseline Date", None),
        ("Responsible Person", None), ("Department", None), ("Priority", None), ("Category", None),
        ("Current Status",
         f'=IF({tr("Milestone ID","tblMilestones")}="","",IF(COUNTIF({MID},{tr("Milestone ID","tblMilestones")})=0,"Not Yet Started",'
         f'IFERROR(INDEX({T}[Status],MATCH({tr("Milestone ID","tblMilestones")}&"|"&TEXT(_xlfn.MAXIFS({REP},{MID},{tr("Milestone ID","tblMilestones")}),"0"),{KEY},0)),"Check data")))'),
    ]
    HR = 4
    for j, (h, fml) in enumerate(reg_cols):
        ws.cell(HR, 1 + j, h)
        hdr(ws.cell(HR, 1 + j), calc=bool(fml))
    for i, m in enumerate(milestones):
        r = HR + 1 + i
        vals = [m["project"], None, m["id"], m["name"], m["desc"], m["base"], m["owner"], m["dept"], m["prio"], m["cat"], None]
        for j, v in enumerate(vals):
            fml = reg_cols[j][1]
            c = ws.cell(r, 1 + j, fml if fml else v)
            c.font = f()
            if j == 5:
                c.number_format = DATE
    reg_last = HR + len(milestones)
    add_table(ws, "tblMilestones", f"A{HR}:K{reg_last}", reg_cols)
    setw(ws, {"A": 11, "B": 28, "C": 12, "D": 34, "E": 52, "F": 13, "G": 18, "H": 22, "I": 10, "J": 14, "K": 17})
    ws.freeze_panes = "E5"
    add_name(wb, "lstMilestoneID", "tblMilestones[Milestone ID]")
    big = 2000
    dv = DataValidation(type="list", formula1="lstProjectID", allow_blank=False, showErrorMessage=True,
                        errorTitle="Unknown project", error="Choose a Project ID defined on the Project Setup sheet.")
    dv.add(f"A5:A{big}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="custom", formula1=f"AND(LEN(C5)>0,COUNTIF($C$5:$C${big},C5)=1)", showErrorMessage=True,
                        errorTitle="Duplicate Milestone ID", error="Each milestone needs a unique, non-blank Milestone ID.")
    dv.add(f"C5:C{big}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="date", operator="between", formula1="DATE(2020,1,1)", formula2="DATE(2040,12,31)",
                        showErrorMessage=True, errorTitle="Invalid date", error="Enter a valid baseline date (2020-2040).")
    dv.add(f"F5:F{big}")
    ws.add_data_validation(dv)
    for col, nm in (("H", "lstDepartment"), ("I", "lstPriority"), ("J", "lstCategory")):
        dv = DataValidation(type="list", formula1=f"{nm}", showErrorMessage=True, error="Pick a value from the list.")
        dv.add(f"{col}5:{col}{big}")
        ws.add_data_validation(dv)
    ws.conditional_formatting.add(f"C5:C{big}", FormulaRule(formula=[f'AND(C5<>"",COUNTIF($C$5:$C${big},C5)>1)'],
                                                            fill=fill("F8D7DA"), font=Font(color=RED, bold=True)))
    status_cf(ws, f"K5:K{big}", "K5")

    # ========================================================================= MILESTONE DATA
    ws = ws_data
    title_block(ws, "Milestone Data  (tblMilestoneData)",
                "One row per milestone per reporting date. Enter only the 5 navy columns; grey columns calculate automatically. Never edit or delete past rows - history drives the MTA chart.")
    ws["B3"] = "Example of a new row:  M-108 | 30-Nov-2026 | 08-Jan-2027 | (blank) | Dependency slipped (+7 d)"
    ws["B3"].font = f(9, False, MUTED, True)
    tolS = "cfgTolerance"
    data_cols = [
        ("Milestone ID", None), ("Reporting Date", None), ("Forecast Completion Date", None),
        ("Actual Completion Date", None), ("Comments", None),
        ("Project ID", f'=IFERROR(INDEX(tblMilestones[Project ID],MATCH({tr("Milestone ID")},tblMilestones[Milestone ID],0)),"")'),
        ("Project Name", f'=IFERROR(INDEX(tblProjects[Project Name],MATCH({tr("Project ID")},tblProjects[Project ID],0)),"")'),
        ("Milestone Name", f'=IFERROR(INDEX(tblMilestones[Milestone Name],MATCH({tr("Milestone ID")},tblMilestones[Milestone ID],0)),"")'),
        ("Milestone Category", f'=IFERROR(INDEX(tblMilestones[Category],MATCH({tr("Milestone ID")},tblMilestones[Milestone ID],0)),"")'),
        ("Original Baseline Date", f'=IFERROR(INDEX(tblMilestones[Baseline Date],MATCH({tr("Milestone ID")},tblMilestones[Milestone ID],0)),"")'),
        ("Milestone Owner", f'=IFERROR(INDEX(tblMilestones[Responsible Person],MATCH({tr("Milestone ID")},tblMilestones[Milestone ID],0)),"")'),
        ("Priority", f'=IFERROR(INDEX(tblMilestones[Priority],MATCH({tr("Milestone ID")},tblMilestones[Milestone ID],0)),"")'),
        ("Status", f'=IF(OR({tr("Milestone ID")}="",NOT(ISNUMBER({tr("Forecast Completion Date")}))),"",IF(N({tr("Actual Completion Date")})>0,"Completed",'
                   f'IF(NOT(ISNUMBER({tr("Original Baseline Date")})),"",IF({tr("Forecast Completion Date")}-{tr("Original Baseline Date")}>{tolS},"Delayed",'
                   f'IF({tr("Forecast Completion Date")}-{tr("Original Baseline Date")}<-{tolS},"Ahead of Schedule","On Schedule")))))'),
        ("Obs Key", f'={tr("Milestone ID")}&"|"&TEXT({tr("Reporting Date")},"0")'),
        ("Previous Reporting Date", f'=_xlfn.MAXIFS({REP},{MID},{tr("Milestone ID")},{REP},"<"&{tr("Reporting Date")})'),
        ("Previous Forecast Date", f'=IF({tr("Previous Reporting Date")}=0,"",SUMIFS({FC},{MID},{tr("Milestone ID")},{REP},{tr("Previous Reporting Date")}))'),
        ("Movement Days", f'=IF(OR({tr("Previous Forecast Date")}="",NOT(ISNUMBER({tr("Forecast Completion Date")}))),"",{tr("Forecast Completion Date")}-{tr("Previous Forecast Date")})'),
        ("Movement Indicator", f'=IF({tr("Milestone ID")}="","",IF({tr("Movement Days")}="","◆ First report",IF({tr("Movement Days")}>0,"▲ Later",IF({tr("Movement Days")}<0,"▼ Earlier","● No change"))))'),
        ("Abs Movement Days", f'=IF(ISNUMBER({tr("Movement Days")}),ABS({tr("Movement Days")}),0)'),
        ("Slip Flag", f'=IF(ISNUMBER({tr("Movement Days")}),IF({tr("Movement Days")}>0,1,0),0)'),
        ("Data Check", f'=IF({tr("Milestone ID")}="","Missing Milestone ID",IF(ISNA(MATCH({tr("Milestone ID")},tblMilestones[Milestone ID],0)),"Unknown Milestone ID",'
                       f'IF(NOT(ISNUMBER({tr("Reporting Date")})),"Invalid reporting date",IF(NOT(ISNUMBER({tr("Forecast Completion Date")})),"Missing / invalid forecast",'
                       f'IF(COUNTIFS({MID},{tr("Milestone ID")},{REP},{tr("Reporting Date")})>1,"DUPLICATE observation",'
                       f'IF(AND(N({tr("Actual Completion Date")})>0,N({tr("Actual Completion Date")})>{tr("Reporting Date")}),"Actual after reporting date","OK"))))))'),
    ]
    DH = 4
    for j, (h, fml) in enumerate(data_cols):
        ws.cell(DH, 1 + j, h)
        hdr(ws.cell(DH, 1 + j), calc=bool(fml))
    ws.row_dimensions[DH].height = 32
    fmts = {1: DATE, 2: DATE, 3: DATE, 9: DATE, 14: DATE_HIDE0, 15: DATE, 16: DAYS}
    obs_sorted = sorted(obs, key=lambda o: (o["rep"], o["mid"]))
    for i, o in enumerate(obs_sorted):
        r = DH + 1 + i
        vals = [o["mid"], o["rep"], o["fc"], o["actual"], o["comment"]]
        for j, (h, fml) in enumerate(data_cols):
            c = ws.cell(r, 1 + j, vals[j] if j < 5 else fml)
            c.font = f()
            if j in fmts:
                c.number_format = fmts[j]
    data_last = DH + len(obs_sorted)
    add_table(ws, "tblMilestoneData", f"A{DH}:{L(len(data_cols))}{data_last}", data_cols)
    setw(ws, {"A": 12, "B": 13, "C": 14, "D": 14, "E": 36, "F": 10, "G": 26, "H": 32, "I": 14, "J": 13, "K": 16,
              "L": 9, "M": 17, "N": 15, "O": 13, "P": 13, "Q": 11, "R": 15, "S": 10, "T": 8, "U": 22})
    ws.freeze_panes = "C5"
    big = 5000
    dv = DataValidation(type="list", formula1="lstMilestoneID", showErrorMessage=True, errorTitle="Unknown milestone",
                        error="Pick a Milestone ID from the Milestone Register (add new milestones there first).")
    dv.add(f"A5:A{big}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="custom",
                        formula1=f"AND(ISNUMBER(B5),B5>=DATE(2020,1,1),B5<=DATE(2040,12,31),COUNTIFS($A$5:$A${big},A5,$B$5:$B${big},B5)=1)",
                        showErrorMessage=True, errorTitle="Invalid or duplicate reporting date",
                        error="Enter a real date (2020-2040). Each milestone may have only ONE observation per reporting date.")
    dv.add(f"B5:B{big}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="date", operator="between", formula1="DATE(2020,1,1)", formula2="DATE(2040,12,31)",
                        showErrorMessage=True, errorTitle="Invalid date", error="Enter a valid forecast date (2020-2040).")
    dv.add(f"C5:C{big}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="custom", formula1="AND(ISNUMBER(D5),D5>=DATE(2020,1,1),D5<=B5)", showErrorMessage=True,
                        errorTitle="Invalid actual date", error="Actual completion must be a real date on or before the reporting date. Leave blank if not complete.")
    dv.add(f"D5:D{big}")
    ws.add_data_validation(dv)
    ws.conditional_formatting.add(f"U5:U{big}", FormulaRule(formula=['AND(U5<>"",U5<>"OK")'], fill=fill("F8D7DA"), font=Font(color=RED, bold=True)))
    movement_cf(ws, f"R5:R{big}", "R5")
    status_cf(ws, f"M5:M{big}", "M5")
    # required-field check for partially entered rows
    ws.conditional_formatting.add(f"A5:C{big}", FormulaRule(formula=['AND(COUNTA($A5:$E5)>0,A5="")'], fill=fill("F8D7DA")))

    # ========================================================================= DASHBOARD (inputs first: names used by Engine/DA)
    ws = ws_dash
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 80
    ws.column_dimensions["A"].width = 2
    for c in range(2, 27):
        ws.column_dimensions[L(c)].width = 8.6
    ws.column_dimensions["AA"].width = 2
    for r in range(1, 80):
        for c in range(1, 28):
            ws.cell(r, c).fill = fill(BG)
    ws["B1"] = "Milestone Trend Analysis  |  Executive Dashboard"
    ws["B1"].font = f(20, True, NAVY)
    ws.row_dimensions[1].height = 34
    ws["B2"] = '="Portfolio view as of "&IF(ISNUMBER(selPeriod),TEXT(selPeriod,"dd-mmm-yyyy"),"each project\'s latest report")&"   |   Filters: "&selProject&" / "&selCategory&" / "&selStatus'
    ws["B2"].font = f(10, False, MUTED, True)

    filters = [("B", "E", "Project", "All Projects", "selProject", "=lstProjectFilter"),
               ("G", "I", "Reporting Period (as-of)", "(Latest)", "selPeriod", "=lstPeriod"),
               ("K", "N", "Milestone Category", "All Categories", "selCategory", "=lstCategoryFilter"),
               ("P", "S", "Milestone Status", "All Statuses", "selStatus", "=lstStatusFilter")]
    for c1, c2, label, default, nm, src in filters:
        ws[f"{c1}4"] = label
        ws[f"{c1}4"].font = f(9, True, MUTED)
        ws.merge_cells(f"{c1}5:{c2}5")
        c = ws[f"{c1}5"]
        c.value = default
        c.font = f(11, True, NAVY)
        c.fill = fill("FFFFFF")
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c.protection = UNLOCK
        c.number_format = DATE
        for cc in range(ws[f"{c1}5"].column, ws[f"{c2}5"].column + 1):
            ws.cell(5, cc).border = Border(top=Side(style="thin", color=ACCENT), bottom=Side(style="thin", color=ACCENT),
                                           left=Side(style="thin", color=ACCENT) if cc == ws[f"{c1}5"].column else None,
                                           right=Side(style="thin", color=ACCENT) if cc == ws[f"{c2}5"].column else None)
        add_name(wb, nm, f"Dashboard!${c1}$5")
        dv = DataValidation(type="list", formula1=src.lstrip("="), showErrorMessage=True, errorTitle="Filter", error="Choose a value from the dropdown list.")
        dv.add(f"{c1}5")
        ws.add_data_validation(dv)
    ws.row_dimensions[5].height = 22
    ws["U4"] = "Tip"
    ws["U4"].font = f(9, True, MUTED)
    ws["U5"] = "Pick filters in the white boxes. The MTA chart shows one project at a time (the first project when 'All Projects' is selected)."
    ws["U5"].font = f(8, False, MUTED, True)
    ws["U5"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("U5:Z6")

    # ========================================================================= ENGINE (selections, lists, chart data)
    ws = ws_eng
    title_block(ws, "Chart Engine", "Helper calculations for dropdown lists and charts. Do not edit - everything here is formula-driven.")
    params = [
        ("engProjID", "MTA chart project ID",
         '=IF(selProject="All Projects",INDEX(tblProjects[Project ID],1),IFERROR(INDEX(tblProjects[Project ID],MATCH(selProject,tblProjects[Project Name],0)),INDEX(tblProjects[Project ID],1)))', None),
        ("engProjName", "MTA chart project name", '=IFERROR(INDEX(tblProjects[Project Name],MATCH(engProjID,tblProjects[Project ID],0)),"")', None),
        ("engAsOf", "MTA chart as-of date", f'=IF(ISNUMBER(selPeriod),selPeriod,_xlfn.MAXIFS({REP},{PID},engProjID))', DATE),
        ("engProjCrit", "Project criterion for lists", '=IF(selProject="All Projects","*",engProjID)', None),
        ("rhProjID", "Reporting History project ID", '=IFERROR(INDEX(tblProjects[Project ID],MATCH(rhProject,tblProjects[Project Name],0)),"")', None),
        ("rptProjID", "Status Report project ID", '=IFERROR(INDEX(tblProjects[Project ID],MATCH(rptProject,tblProjects[Project Name],0)),"")', None),
        ("rptAsOf", "Status Report as-of date", f'=IF(ISNUMBER(selPeriod),selPeriod,_xlfn.MAXIFS({REP},{PID},rptProjID))', DATE),
    ]
    for i, (nm, label, fml, fmt) in enumerate(params):
        r = 3 + i
        ws.cell(r, 1, label).font = f(9, False, MUTED)
        c = ws.cell(r, 2, fml)
        c.font = f(10, True)
        if fmt:
            c.number_format = fmt
        ws.cell(r, 3, nm).font = f(8, False, MUTED)
        add_name(wb, nm, f"Engine!$B${r}")

    # lists ---------------------------------------------------------------
    LR = 12
    for col, h in (("A", "Project filter list"), ("C", "Reporting period list"), ("E", "Category filter list"), ("G", "Status filter list")):
        ws[f"{col}{LR - 1}"] = h
        hdr(ws[f"{col}{LR - 1}"], True)
    ws[f"A{LR}"] = "All Projects"
    for k in range(1, 31):
        ws[f"A{LR + k}"] = f'=IFERROR(INDEX(tblProjects[Project Name],{k})&"","")'
    ws[f"C{LR}"] = "(Latest)"
    ws[f"C{LR + 1}"] = f'=IF(_xlfn.MINIFS({REP},{PID},engProjCrit)=0,"",_xlfn.MINIFS({REP},{PID},engProjCrit))'
    for k in range(2, 61):
        p = f"C{LR + k - 1}"
        ws[f"C{LR + k}"] = (f'=IF({p}="","",IF(_xlfn.MINIFS({REP},{PID},engProjCrit,{REP},">"&{p})=0,"",'
                            f'_xlfn.MINIFS({REP},{PID},engProjCrit,{REP},">"&{p})))')
    for k in range(1, 61):
        ws[f"C{LR + k}"].number_format = DATE
    ws[f"E{LR}"] = "All Categories"
    for k in range(1, 31):
        ws[f"E{LR + k}"] = f'=IFERROR(INDEX(tblCategories[Category],{k})&"","")'
    ws[f"G{LR}"] = "All Statuses"
    for k, s in enumerate(STATUSES, 1):
        ws[f"G{LR + k}"] = s
    add_name(wb, "lstProjectFilter", f'Engine!$A${LR}:INDEX(Engine!$A${LR}:$A${LR + 30},COUNTIF(Engine!$A${LR}:$A${LR + 30},"?*"))')
    add_name(wb, "lstPeriod", f"Engine!$C${LR}:INDEX(Engine!$C${LR}:$C${LR + 60},1+COUNT(Engine!$C${LR + 1}:$C${LR + 60}))")
    add_name(wb, "lstCategoryFilter", f'Engine!$E${LR}:INDEX(Engine!$E${LR}:$E${LR + 30},COUNTIF(Engine!$E${LR}:$E${LR + 30},"?*"))')
    add_name(wb, "lstStatusFilter", f"Engine!$G${LR}:$G${LR + len(STATUSES)}")

    # Delay Analysis column map (needed by engine formulas) ----------------
    da_keys = ["num", "id", "proj", "projname", "name", "cat", "prio", "owner", "base", "asof", "nobs", "firstrep",
               "firstfc", "prevrep", "currep", "prevfc", "curfc", "actual", "basevar", "move", "totdelay",
               "movefirst", "slips", "maxmove", "projlatest", "daystofc", "status", "indicator", "r_slip", "r_sig",
               "r_var", "r_miss", "r_dead", "riskscore", "risklevel", "risknotes", "infilter", "finish",
               "topscore", "chartelig", "chartrank", "rhrank", "rptrank", "rptdelay", "rptdelayrank", "rptchg",
               "rptchgrank", "rptrisk", "rptriskrank"]
    DC = {k: L(i + 1) for i, k in enumerate(da_keys)}
    for k, col in DC.items():
        add_name(wb, f"da_{k}", f"'Delay Analysis'!${col}${DA_FIRST}:${col}${DA_LAST}")

    # slot table ------------------------------------------------------------
    SR = 13  # first slot row
    slot_hdr = ["Slot", "DA row", "Milestone ID", "Milestone Name", "Series label", "Show", "Colour", "Current / actual",
                "Last movement", "Baseline", "Status", "Baseline (chart)", "Current (chart)", "Max move"]
    for j, h in enumerate(slot_hdr):
        ws.cell(SR - 1, 9 + j, h)
        hdr(ws.cell(SR - 1, 9 + j), True)
    for k in range(1, NSLOT + 1):
        r = SR + k - 1
        ws[f"I{r}"] = k
        ws[f"J{r}"] = f'=IFERROR(MATCH(I{r},da_chartrank,0),"")'
        ws[f"K{r}"] = f'=IF(J{r}="","",INDEX(da_id,J{r}))'
        ws[f"L{r}"] = f'=IF(J{r}="","",INDEX(da_name,J{r}))'
        ws[f"M{r}"] = f'=IF(K{r}="","",K{r}&"  "&L{r})'
        ws[f"N{r}"] = f'=IF(Dashboard!$Z${14 + k}="No","No","Yes")'
        ws[f"O{r}"] = SERIES_COLORS[k - 1]
        ws[f"P{r}"] = f'=IF(J{r}="","",INDEX(da_finish,J{r}))'
        ws[f"Q{r}"] = f'=IF(J{r}="","",INDEX(da_indicator,J{r}))'
        ws[f"R{r}"] = f'=IF(J{r}="","",INDEX(da_base,J{r}))'
        ws[f"S{r}"] = f'=IF(J{r}="","",INDEX(da_status,J{r}))'
        ws[f"T{r}"] = f'=IF(J{r}="",NA(),INDEX(da_base,J{r}))'
        ws[f"U{r}"] = f'=IF(J{r}="",NA(),IF(INDEX(da_finish,J{r})="",NA(),INDEX(da_finish,J{r})))'
        ws[f"V{r}"] = f'=IF(J{r}="","",INDEX(da_maxmove,J{r}))'
        for col in "PRTU":
            ws[f"{col}{r}"].number_format = DATE
    SLOT_LAST = SR + NSLOT - 1

    # small chart tables ----------------------------------------------------
    BR = 77
    ws[f"A{BR - 1}"] = "Delay distribution (total delay vs baseline, in-filter, started milestones)"
    ws[f"A{BR - 1}"].font = f(9, True)
    for j, h in enumerate(["Bucket", "From (d)", "To (d)", "Milestones"]):
        ws.cell(BR, 1 + j, h)
        hdr(ws.cell(BR, 1 + j), True)
    bins = [("Early > 14 d", -99999, -15), ("Early 1-14 d", -14, -1), ("On baseline", 0, 0), ("Late 1-14 d", 1, 14),
            ("Late 15-30 d", 15, 30), ("Late 31-60 d", 31, 60), ("Late > 60 d", 61, 99999)]
    for i, (lab, lo, hi) in enumerate(bins):
        r = BR + 1 + i
        ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"] = lab, lo, hi
        ws[f"D{r}"] = f'=COUNTIFS(da_totdelay,">="&B{r},da_totdelay,"<="&C{r},da_infilter,1)'
    BIN_LAST = BR + len(bins)

    ws[f"F{BR - 1}"] = "Status breakdown (in-filter)"
    ws[f"F{BR - 1}"].font = f(9, True)
    ws[f"F{BR}"], ws[f"G{BR}"] = "Status", "Milestones"
    hdr(ws[f"F{BR}"], True)
    hdr(ws[f"G{BR}"], True)
    for i, s in enumerate(STATUSES):
        ws[f"F{BR + 1 + i}"] = s
        ws[f"G{BR + 1 + i}"] = f'=COUNTIFS(da_status,F{BR + 1 + i},da_infilter,1)'
    ST_LAST = BR + len(STATUSES)

    ws[f"I{BR - 1}"] = "Top 5 most delayed open milestones (in-filter)"
    ws[f"I{BR - 1}"].font = f(9, True)
    for j, h in enumerate(["Rank", "Score", "DA row", "Milestone", "Total delay (d)"]):
        ws.cell(BR, 9 + j, h)
        hdr(ws.cell(BR, 9 + j), True)
    for k in range(1, 6):
        r = BR + k
        ws[f"I{r}"] = k
        ws[f"J{r}"] = f'=IFERROR(LARGE(da_topscore,I{r}),"")'
        ws[f"K{r}"] = f'=IF(J{r}="","",MATCH(J{r},da_topscore,0))'
        ws[f"L{r}"] = f'=IF(K{r}="","",INDEX(da_id,K{r})&"  "&INDEX(da_name,K{r}))'
        ws[f"M{r}"] = f'=IF(K{r}="",NA(),INDEX(da_totdelay,K{r}))'

    # diagonal
    GR = 98  # first grid data row
    GL = GR + NREP - 1
    DR = 88
    ws[f"A{DR - 1}"] = "Completion line (forecast = reporting date)"
    ws[f"A{DR - 1}"].font = f(9, True)
    ws[f"A{DR}"] = f'=IF(COUNT(A{GR}:A{GL})=0,NA(),MIN(A{GR}:A{GL}))'
    ws[f"B{DR}"] = f"=A{DR}"
    ws[f"A{DR + 1}"] = f'=IF(COUNT(A{GR}:A{GL})=0,NA(),MAX(A{GR}:A{GL}))'
    ws[f"B{DR + 1}"] = f"=A{DR + 1}"
    ws[f"C{DR}"] = "Completion line"
    for rr in (DR, DR + 1):
        for col in "AB":
            ws[f"{col}{rr}"].number_format = DATE

    # MTA grid --------------------------------------------------------------
    ws[f"A{GR - 2}"] = "MTA grid: rows = reporting dates of the chart project (<= as-of), columns = milestone slots, values = forecast date (NA() error value = no observation, not plotted)"
    ws[f"A{GR - 2}"].font = f(9, True)
    ws[f"A{GR - 1}"], ws[f"B{GR - 1}"] = "Reporting date", "X (date)"
    hdr(ws[f"A{GR - 1}"], True)
    hdr(ws[f"B{GR - 1}"], True)
    for k in range(1, NSLOT + 1):
        c = ws.cell(GR - 1, 2 + k, f"=M{SR + k - 1}")
        hdr(c, True)
        c2 = ws.cell(GR - 1, 18 + k, f'="Move "&I{SR + k - 1}')
        hdr(c2, True)
    for j, h in enumerate(["Slipped", "Pulled in", "Unchanged"]):
        ws.cell(GR - 1, 35 + j, h)
        hdr(ws.cell(GR - 1, 35 + j), True)
    for i in range(NREP):
        r = GR + i
        if i == 0:
            ws[f"A{r}"] = (f'=IF(_xlfn.MINIFS({REP},{PID},engProjID,{REP},"<="&engAsOf)=0,"",'
                           f'_xlfn.MINIFS({REP},{PID},engProjID,{REP},"<="&engAsOf))')
        else:
            p = f"A{r - 1}"
            ws[f"A{r}"] = (f'=IF({p}="","",IF(_xlfn.MINIFS({REP},{PID},engProjID,{REP},">"&{p},{REP},"<="&engAsOf)=0,"",'
                           f'_xlfn.MINIFS({REP},{PID},engProjID,{REP},">"&{p},{REP},"<="&engAsOf)))')
        ws[f"B{r}"] = f'=IF(A{r}="",NA(),A{r})'
        ws[f"A{r}"].number_format = DATE
        ws[f"B{r}"].number_format = DATE
        for k in range(1, NSLOT + 1):
            s = SR + k - 1
            keyf = f'$K${s}&"|"&TEXT($A{r},"0")'
            c = ws.cell(r, 2 + k, f'=IF(OR($A{r}="",$K${s}="",$N${s}<>"Yes"),NA(),IFERROR(INDEX({FC},MATCH({keyf},{KEY},0)),NA()))')
            c.number_format = DATE
            ws.cell(r, 18 + k, f'=IF(OR($A{r}="",$K${s}="",$N${s}<>"Yes"),NA(),IFERROR(INDEX({MOVE},MATCH({keyf},{KEY},0)),NA()))')
        mv = f"S{r}:AG{r}"
        ws.cell(r, 35, f'=IF($A{r}="",NA(),COUNTIF({mv},">0"))')
        ws.cell(r, 36, f'=IF($A{r}="",NA(),COUNTIF({mv},"<0"))')
        ws.cell(r, 37, f'=IF($A{r}="",NA(),COUNTIF({mv},0))')

    # flattened points for overlay series ------------------------------------
    FR = GL + 5
    ws[f"A{FR - 2}"] = "Flattened grid for overlay series (significant movement / milestone achieved)"
    ws[f"A{FR - 2}"].font = f(9, True)
    for j, h in enumerate(["#", "Row", "Slot", "X", "Y", "Significant", "Achieved"]):
        ws.cell(FR - 1, 1 + j, h)
        hdr(ws.cell(FR - 1, 1 + j), True)
    NPTS = NREP * NSLOT
    for i in range(1, NPTS + 1):
        r = FR + i - 1
        ws[f"A{r}"] = i
        ws[f"B{r}"] = f"=INT((A{r}-1)/{NSLOT})+1"
        ws[f"C{r}"] = f"=MOD(A{r}-1,{NSLOT})+1"
        ws[f"D{r}"] = f"=INDEX($B${GR}:$B${GL},B{r})"
        ws[f"E{r}"] = f"=INDEX($C${GR}:$Q${GL},B{r},C{r})"
        mvi = f"INDEX($S${GR}:$AG${GL},B{r},C{r})"
        ws[f"F{r}"] = f'=IF(ISNA(E{r}),NA(),IF(ISNUMBER({mvi}),IF(ABS({mvi})>=cfgSigMove,E{r},NA()),NA()))'
        ws[f"G{r}"] = (f'=IF(ISNA(E{r}),NA(),IF(N(INDEX({ACT},MATCH(INDEX($K${SR}:$K${SLOT_LAST},C{r})&"|"&TEXT(D{r},"0"),{KEY},0)))>0,E{r},NA()))')
        for col in "DEFG":
            ws[f"{col}{r}"].number_format = DATE
    FL = FR + NPTS - 1
    setw(ws, {"A": 22, "B": 14, "C": 13, "D": 12, "E": 15, "F": 17, "G": 14, "L": 30, "M": 34})

    # ========================================================================= DELAY ANALYSIS
    ws = ws_da
    title_block(ws, "Delay Analysis",
                "One row per milestone in the register (capacity 200). All values are calculated as of the Dashboard reporting period; delays are in calendar days.")
    c = lambda k, r: f"${DC[k]}{r}"  # noqa: E731
    da_spec = [
        ("num", "#", lambda r: f"=ROW()-{DA_FIRST - 1}", "0"),
        ("id", "Milestone ID", lambda r: f'=IFERROR(INDEX(tblMilestones[Milestone ID],{c("num", r)})&"","")', None),
        ("proj", "Project ID", lambda r: f'=IF({c("id", r)}="","",INDEX(tblMilestones[Project ID],{c("num", r)})&"")', None),
        ("projname", "Project Name", lambda r: f'=IF({c("id", r)}="","",IFERROR(INDEX(tblProjects[Project Name],MATCH({c("proj", r)},tblProjects[Project ID],0)),"(unknown project)"))', None),
        ("name", "Milestone Name", lambda r: f'=IF({c("id", r)}="","",INDEX(tblMilestones[Milestone Name],{c("num", r)})&"")', None),
        ("cat", "Category", lambda r: f'=IF({c("id", r)}="","",INDEX(tblMilestones[Category],{c("num", r)})&"")', None),
        ("prio", "Priority", lambda r: f'=IF({c("id", r)}="","",INDEX(tblMilestones[Priority],{c("num", r)})&"")', None),
        ("owner", "Owner", lambda r: f'=IF({c("id", r)}="","",INDEX(tblMilestones[Responsible Person],{c("num", r)})&"")', None),
        ("base", "Original Baseline", lambda r: f'=IF({c("id", r)}="","",INDEX(tblMilestones[Baseline Date],{c("num", r)}))', DATE),
        ("asof", "As-of Date", lambda r: f'=IF({c("id", r)}="","",IF(ISNUMBER(selPeriod),selPeriod,_xlfn.MAXIFS({REP},{PID},{c("proj", r)})))', DATE_HIDE0),
        ("nobs", "Observations", lambda r: f'=IF({c("id", r)}="",0,COUNTIFS({MID},{c("id", r)},{REP},"<="&{c("asof", r)}))', "0"),
        ("firstrep", "First Report", lambda r: f'=IF({c("nobs", r)}=0,"",_xlfn.MINIFS({REP},{MID},{c("id", r)},{REP},"<="&{c("asof", r)}))', DATE),
        ("firstfc", "First Forecast", lambda r: f'=IF({c("firstrep", r)}="","",SUMIFS({FC},{MID},{c("id", r)},{REP},{c("firstrep", r)}))', DATE),
        ("prevrep", "Previous Report", lambda r: f'=IF({c("nobs", r)}<2,"",_xlfn.MAXIFS({REP},{MID},{c("id", r)},{REP},"<"&{c("currep", r)}))', DATE),
        ("currep", "Latest Report", lambda r: f'=IF({c("nobs", r)}=0,"",_xlfn.MAXIFS({REP},{MID},{c("id", r)},{REP},"<="&{c("asof", r)}))', DATE),
        ("prevfc", "Previous Forecast", lambda r: f'=IF({c("prevrep", r)}="","",SUMIFS({FC},{MID},{c("id", r)},{REP},{c("prevrep", r)}))', DATE),
        ("curfc", "Current Forecast", lambda r: f'=IF({c("currep", r)}="","",SUMIFS({FC},{MID},{c("id", r)},{REP},{c("currep", r)}))', DATE),
        ("actual", "Actual Completion", lambda r: f'=IF({c("nobs", r)}=0,"",IF(_xlfn.MAXIFS({ACT},{MID},{c("id", r)},{REP},"<="&{c("asof", r)})=0,"",_xlfn.MAXIFS({ACT},{MID},{c("id", r)},{REP},"<="&{c("asof", r)})))', DATE),
        ("basevar", "Baseline Variance (d)", lambda r: f'=IF({c("curfc", r)}="","",{c("curfc", r)}-{c("base", r)})', DAYS),
        ("move", "Forecast Movement (d)", lambda r: f'=IF({c("prevfc", r)}="","",{c("curfc", r)}-{c("prevfc", r)})', DAYS),
        ("totdelay", "Total Delay (d)", lambda r: f'=IF({c("curfc", r)}="","",IF({c("actual", r)}<>"",{c("actual", r)},{c("curfc", r)})-{c("base", r)})', DAYS),
        ("movefirst", "Movement since First (d)", lambda r: f'=IF({c("firstfc", r)}="","",{c("curfc", r)}-{c("firstfc", r)})', DAYS),
        ("slips", "Slip Count", lambda r: f'=IF({c("nobs", r)}=0,0,COUNTIFS({MID},{c("id", r)},{REP},"<="&{c("asof", r)},{T}[Slip Flag],1))', "0"),
        ("maxmove", "Largest Single Move (d)", lambda r: f'=IF({c("nobs", r)}=0,0,_xlfn.MAXIFS({T}[Abs Movement Days],{MID},{c("id", r)},{REP},"<="&{c("asof", r)}))', "0"),
        ("projlatest", "Project Latest Report", lambda r: f'=IF({c("id", r)}="","",_xlfn.MAXIFS({REP},{PID},{c("proj", r)},{REP},"<="&{c("asof", r)}))', DATE_HIDE0),
        ("daystofc", "Days to Forecast", lambda r: f'=IF(OR({c("curfc", r)}="",{c("actual", r)}<>""),"",{c("curfc", r)}-{c("asof", r)})', "0"),
        ("status", "Schedule Status", lambda r: (f'=IF({c("id", r)}="","",IF({c("nobs", r)}=0,"Not Yet Started",IF({c("actual", r)}<>"","Completed",'
                                                 f'IF({c("basevar", r)}>cfgTolerance,"Delayed",IF({c("basevar", r)}<-cfgTolerance,"Ahead of Schedule","On Schedule")))))'), None),
        ("indicator", "Movement Indicator", lambda r: (f'=IF({c("nobs", r)}=0,"",IF({c("move", r)}="","◆ First report",IF({c("move", r)}>0,"▲ Later",'
                                                       f'IF({c("move", r)}<0,"▼ Earlier","● No change"))))'), None),
        ("r_slip", "Risk: Repeated Slippage", lambda r: f'=IF(AND({c("id", r)}<>"",{c("actual", r)}="",{c("slips", r)}>=cfgSlipCount),1,0)', "0;;"),
        ("r_sig", "Risk: Significant Change", lambda r: f'=IF(AND({c("nobs", r)}>0,{c("actual", r)}="",{c("maxmove", r)}>=cfgSigMove),1,0)', "0;;"),
        ("r_var", "Risk: High Baseline Variance", lambda r: f'=IF(AND({c("actual", r)}="",ISNUMBER({c("basevar", r)})),IF({c("basevar", r)}>=cfgHighVar,1,0),0)', "0;;"),
        ("r_miss", "Risk: Missing Update", lambda r: f'=IF(AND({c("nobs", r)}>0,{c("actual", r)}=""),IF({c("projlatest", r)}>{c("currep", r)},1,0),0)', "0;;"),
        ("r_dead", "Risk: Approaching Deadline", lambda r: f'=IF(ISNUMBER({c("daystofc", r)}),IF({c("daystofc", r)}<=cfgWindow,1,0),0)', "0;;"),
        ("riskscore", "Risk Score", lambda r: f'=SUM({c("r_slip", r)}:{c("r_dead", r)})', "0"),
        ("risklevel", "Risk Level", lambda r: f'=IF({c("id", r)}="","",IF({c("riskscore", r)}>=3,"High",IF({c("riskscore", r)}=2,"Medium",IF({c("riskscore", r)}=1,"Low","None"))))', None),
        ("risknotes", "Risk Observations", lambda r: (
            f'=_xlfn.TEXTJOIN("; ",TRUE,IF({c("r_slip", r)}=1,"Repeated slippage ("&{c("slips", r)}&" slips)",""),'
            f'IF({c("r_sig", r)}=1,"Large single move ("&{c("maxmove", r)}&" d)",""),'
            f'IF({c("r_var", r)}=1,"High baseline variance (+"&{c("basevar", r)}&" d)",""),'
            f'IF({c("r_miss", r)}=1,"Missing latest update",""),'
            f'IF({c("r_dead", r)}=1,IF({c("daystofc", r)}<0,"Overdue by "&-{c("daystofc", r)}&" d","Due in "&{c("daystofc", r)}&" d"),""))'), None),
        ("infilter", "In Dashboard Filter", lambda r: (f'=IF({c("id", r)}="",0,IF(AND(OR(selProject="All Projects",{c("projname", r)}=selProject),'
                                                       f'OR(selCategory="All Categories",{c("cat", r)}=selCategory),OR(selStatus="All Statuses",{c("status", r)}=selStatus)),1,0))'), "0"),
        ("finish", "Forecast / Actual Finish", lambda r: f'=IF({c("actual", r)}<>"",{c("actual", r)},{c("curfc", r)})', DATE),
        ("topscore", "Top-Delay Score", lambda r: f'=IF(AND({c("infilter", r)}=1,{c("nobs", r)}>0,{c("actual", r)}=""),IF({c("totdelay", r)}>0,{c("totdelay", r)}+{c("num", r)}/10000,""),"")', "0.0000"),
        ("chartelig", "MTA Eligible", lambda r: (f'=IF({c("id", r)}="",0,IF(AND({c("proj", r)}=engProjID,OR(selCategory="All Categories",{c("cat", r)}=selCategory),'
                                                 f'OR(selStatus="All Statuses",{c("status", r)}=selStatus)),1,0))'), "0"),
        ("chartrank", "MTA Slot", lambda r: f'=IF({c("chartelig", r)}=1,COUNTIF(${DC["chartelig"]}${DA_FIRST}:{c("chartelig", r)},1),"")', "0"),
        ("rhrank", "Entry-Form Row", lambda r: f'=IF(AND({c("id", r)}<>"",{c("proj", r)}=rhProjID),COUNTIF(${DC["proj"]}${DA_FIRST}:{c("proj", r)},rhProjID),"")', "0"),
        ("rptrank", "Report Row", lambda r: f'=IF(AND({c("id", r)}<>"",{c("proj", r)}=rptProjID),COUNTIF(${DC["proj"]}${DA_FIRST}:{c("proj", r)},rptProjID),"")', "0"),
        ("rptdelay", "Report: Delayed", lambda r: f'=IF(AND({c("proj", r)}=rptProjID,{c("status", r)}="Delayed"),1,0)', "0"),
        ("rptdelayrank", "Report: Delayed Row", lambda r: f'=IF({c("rptdelay", r)}=1,COUNTIF(${DC["rptdelay"]}${DA_FIRST}:{c("rptdelay", r)},1),"")', "0"),
        ("rptchg", "Report: Changed", lambda r: f'=IF(AND({c("proj", r)}=rptProjID,ISNUMBER({c("move", r)}),{c("currep", r)}={c("projlatest", r)}),IF({c("move", r)}<>0,1,0),0)', "0"),
        ("rptchgrank", "Report: Changed Row", lambda r: f'=IF({c("rptchg", r)}=1,COUNTIF(${DC["rptchg"]}${DA_FIRST}:{c("rptchg", r)},1),"")', "0"),
        ("rptrisk", "Report: Risk", lambda r: f'=IF(AND({c("proj", r)}=rptProjID,{c("riskscore", r)}>0),1,0)', "0"),
        ("rptriskrank", "Report: Risk Row", lambda r: f'=IF({c("rptrisk", r)}=1,COUNTIF(${DC["rptrisk"]}${DA_FIRST}:{c("rptrisk", r)},1),"")', "0"),
    ]
    assert [k for k, *_ in da_spec] == da_keys
    HRow = DA_FIRST - 1
    helper_from = da_keys.index("infilter")
    for j, (k, h, fn, fmt) in enumerate(da_spec):
        cell = ws.cell(HRow, j + 1, h)
        hdr(cell, calc=j >= helper_from)
        for r in range(DA_FIRST, DA_LAST + 1):
            cc = ws.cell(r, j + 1, fn(r))
            cc.font = f(9)
            if fmt:
                cc.number_format = fmt
    ws.row_dimensions[HRow].height = 42
    widths = {"num": 4, "id": 9, "proj": 9, "projname": 22, "name": 30, "cat": 12, "prio": 8, "owner": 15,
              "risknotes": 48, "status": 16, "indicator": 14, "risklevel": 9}
    for k, col in DC.items():
        ws.column_dimensions[col].width = widths.get(k, 11)
    ws.freeze_panes = f"F{DA_FIRST}"
    ws.auto_filter.ref = f"A{HRow}:{DC['risknotes']}{DA_LAST}"
    status_cf(ws, f"{DC['status']}{DA_FIRST}:{DC['status']}{DA_LAST}", f"{DC['status']}{DA_FIRST}")
    movement_cf(ws, f"{DC['indicator']}{DA_FIRST}:{DC['indicator']}{DA_LAST}", f"{DC['indicator']}{DA_FIRST}")
    for k in ("r_slip", "r_sig", "r_var", "r_miss", "r_dead"):
        rng = f"{DC[k]}{DA_FIRST}:{DC[k]}{DA_LAST}"
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"{DC[k]}{DA_FIRST}=1"], fill=fill("FCE4D6"), font=Font(color="C55A11", bold=True)))
    rl = f"{DC['risklevel']}{DA_FIRST}"
    rng = f"{DC['risklevel']}{DA_FIRST}:{DC['risklevel']}{DA_LAST}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{rl}="High"'], fill=fill(RED), font=Font(color="FFFFFF", bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{rl}="Medium"'], fill=fill("F4B183")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{rl}="Low"'], fill=fill("FFE699")))
    for k in ("basevar", "move", "totdelay", "movefirst"):
        a = f"{DC[k]}{DA_FIRST}"
        rng = f"{DC[k]}{DA_FIRST}:{DC[k]}{DA_LAST}"
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({a}),{a}>0)"], font=Font(color=RED, bold=True)))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({a}),{a}<0)"], font=Font(color=GREEN, bold=True)))
    ws[f"{DC['infilter']}{HRow - 1}"] = "Helper columns used by the Dashboard, Reporting History and Status Report"
    ws[f"{DC['infilter']}{HRow - 1}"].font = f(8, False, MUTED, True)

    # historical comparison panel -----------------------------------------
    section(ws, "B4", "Historical comparison - select a milestone", "J4")
    ws["B5"] = "Milestone ID"
    ws["B5"].font = f(10, True)
    ws["D5"] = milestones[5]["id"]
    inp(ws["D5"])
    ws["D5"].font = f(11, True, "0000FF")
    add_name(wb, "cmpID", "'Delay Analysis'!$D$5")
    dv = DataValidation(type="list", formula1="lstMilestoneID", showErrorMessage=True, error="Pick a Milestone ID.")
    dv.add("D5")
    ws.add_data_validation(dv)
    ws["E5"] = '=IFERROR(INDEX(da_name,MATCH(cmpID,da_id,0))&"  ("&INDEX(da_projname,MATCH(cmpID,da_id,0))&")","Unknown milestone")'
    ws["E5"].font = f(10, True, NAVY)
    ws["B6"] = '="Status: "&IFERROR(INDEX(da_status,MATCH(cmpID,da_id,0))&"   |   Risk: "&INDEX(da_risklevel,MATCH(cmpID,da_id,0))&"   |   Observations: "&INDEX(da_nobs,MATCH(cmpID,da_id,0)),"")'
    ws["B6"].font = f(9, False, MUTED)
    for j, h in enumerate(["Point in history", "", "Reporting date", "Date", "vs Baseline (d)", "vs Previous point (d)"]):
        if h:
            ws.cell(7, 2 + j, h)
            hdr(ws.cell(7, 2 + j))
    ws.merge_cells("B7:C7")
    cmp_rows = [("Original baseline", None, "base"), ("First forecast", "firstrep", "firstfc"),
                ("Previous forecast", "prevrep", "prevfc"), ("Current forecast", "currep", "curfc"),
                ("Actual completion", None, "actual")]
    for i, (lab, repk, valk) in enumerate(cmp_rows):
        r = 8 + i
        ws.cell(r, 2, lab).font = f(10, True)
        ws.merge_cells(f"B{r}:C{r}")
        m = "MATCH(cmpID,da_id,0)"
        if repk:
            ws.cell(r, 4, f'=IFERROR(INDEX(da_{repk},{m}),"")').number_format = DATE
        ws.cell(r, 5, f'=IFERROR(INDEX(da_{valk},{m}),"")').number_format = DATE
        ws.cell(r, 6, f'=IF(OR(E{r}="",$E$8=""),"",E{r}-$E$8)').number_format = DAYS
        if i > 0:
            ws.cell(r, 7, f'=IF(OR(E{r}="",E{r - 1}=""),"",E{r}-E{r - 1})').number_format = DAYS
        for cc in range(2, 8):
            ws.cell(r, cc).border = BORDER
            if cc > 3:
                ws.cell(r, cc).font = f(10)
    for col in "FG":
        rng = f"{col}8:{col}12"
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({col}8),{col}8>0)"], font=Font(color=RED, bold=True)))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({col}8),{col}8<0)"], font=Font(color=GREEN, bold=True)))

    # ========================================================================= DASHBOARD body
    ws = ws_dash
    kpis = [
        ("Total Milestones", "=COUNTIF(da_infilter,1)", "0", '="across "&B9&" project(s)"'),
        ("Total Projects", "=SUMPRODUCT(--(COUNTIFS(da_proj,tblProjects[Project ID],da_infilter,1)>0))", "0", '="of "&ROWS(tblProjects[Project ID])&" in setup"'),
        ("Completed", '=COUNTIFS(da_infilter,1,da_status,"Completed")', "0", None),
        ("Delayed", '=COUNTIFS(da_infilter,1,da_status,"Delayed")', "0", None),
        ("Ahead of Schedule", '=COUNTIFS(da_infilter,1,da_status,"Ahead of Schedule")', "0", None),
        ("On Schedule", '=COUNTIFS(da_infilter,1,da_status,"On Schedule")', "0", None),
        ("Average Delay", '=IFERROR(AVERAGEIFS(da_totdelay,da_infilter,1),0)', '+0.0" d";-0.0" d";0" d"', '="vs baseline, "&COUNTIFS(da_infilter,1,da_status,"<>Not Yet Started")&" started milestones"'),
        ("Latest Forecast Completion", '=IF(COUNTIF(da_infilter,1)=0,"",_xlfn.MAXIFS(da_finish,da_infilter,1))', DATE, '="latest forecast/actual milestone date"'),
    ]
    kpi_colors = [NAVY, NAVY, STATUS_COL["Completed"], STATUS_COL["Delayed"], STATUS_COL["Ahead of Schedule"],
                  STATUS_COL["On Schedule"], "C55A11", NAVY]
    for i, (lab, fml, fmt, sub) in enumerate(kpis):
        c1 = 2 + 3 * i
        c2 = c1 + 2
        for r in range(7, 11):
            for cc in range(c1, c2 + 1):
                cell = ws.cell(r, cc)
                cell.fill = fill("FFFFFF")
                cell.border = Border(top=Side(style="thick", color=kpi_colors[i]) if r == 7 else None,
                                     bottom=Side(style="thin", color=GRID) if r == 10 else None,
                                     left=Side(style="thin", color=GRID) if cc == c1 else None,
                                     right=Side(style="thin", color=GRID) if cc == c2 else None)
        ws.cell(7, c1, lab.upper()).font = f(8, True, MUTED)
        ws.cell(7, c1).alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(start_row=7, start_column=c1, end_row=7, end_column=c2)
        v = ws.cell(8, c1, fml)
        v.font = f(20 if i < 7 else 15, True, kpi_colors[i])
        v.number_format = fmt
        v.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(start_row=8, start_column=c1, end_row=9, end_column=c2)
        if sub is None:
            sub = f'=IF($B$8=0,"",TEXT({L(c1)}8/$B$8,"0%")&" of milestones")'
        s = ws.cell(10, c1, sub)
        s.font = f(8, False, MUTED)
        s.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.merge_cells(start_row=10, start_column=c1, end_row=10, end_column=c2)
    ws.row_dimensions[7].height = 20
    ws.row_dimensions[8].height = 22
    ws.row_dimensions[9].height = 18
    ws.row_dimensions[10].height = 20
    # fix "across n projects" reference to Total Projects KPI (E8)
    ws["B10"] = '="across "&E8&" project(s)"'

    ws["B12"] = '="Milestone Trend Analysis  -  "&engProjID&"  "&engProjName&"   (as of "&TEXT(engAsOf,"dd-mmm-yyyy")&")"'
    section(ws, "B12", ws["B12"].value, "Z12")
    ws.row_dimensions[12].height = 22

    # key panel
    ws["S13"] = "Milestone key  (set Show = No to hide a line)"
    ws["S13"].font = f(9, True, NAVY)
    ws.merge_cells("S13:Z13")
    for col, h in (("S", ""), ("T", "Milestone"), ("X", "Forecast"), ("Y", "Last move"), ("Z", "Show")):
        ws[f"{col}14"] = h
        hdr(ws[f"{col}14"])
    ws.merge_cells("T14:W14")
    for k in range(1, NSLOT + 1):
        r = 14 + k
        s = SR + k - 1
        sw = ws[f"S{r}"]
        sw.value = f'=IF(Engine!$K${s}="","","■")'
        sw.font = Font(name=FONT, size=14, bold=True, color=SERIES_COLORS[k - 1])
        sw.alignment = Alignment(horizontal="center", vertical="center")
        ws[f"T{r}"] = f"=Engine!$M${s}"
        ws[f"T{r}"].font = f(8)
        ws.merge_cells(f"T{r}:W{r}")
        ws[f"X{r}"] = f"=Engine!$P${s}"
        ws[f"X{r}"].number_format = DATE_S
        ws[f"X{r}"].font = f(8)
        ws[f"Y{r}"] = f"=Engine!$Q${s}"
        ws[f"Y{r}"].font = f(8)
        z = ws[f"Z{r}"]
        z.value = "Yes"
        z.font = f(8, True, ACCENT)
        z.alignment = Alignment(horizontal="center")
        z.protection = UNLOCK
        for col in "STUVWXYZ":
            ws[f"{col}{r}"].fill = fill("FFFFFF")
            ws[f"{col}{r}"].border = Border(bottom=Side(style="hair", color=GRID))
    dv = DataValidation(type="list", formula1='"Yes,No"', showErrorMessage=True, error="Choose Yes or No.")
    dv.add(f"Z15:Z{14 + NSLOT}")
    ws.add_data_validation(dv)
    movement_cf(ws, f"Y15:Y{14 + NSLOT}", "Y15")
    ws.conditional_formatting.add(f"Z15:Z{14 + NSLOT}", FormulaRule(formula=['Z15="No"'], font=Font(color="BFBFBF")))
    ws["S31"] = "How to read the MTA chart"
    ws["S31"].font = f(9, True, NAVY)
    notes = ["• X axis = reporting date; Y axis = forecast completion date.",
             "• Line rising = milestone slipping later; falling = pulled in; flat = stable forecast.",
             "• Dashed diagonal = completion line (forecast = report date). Milestones end on it when achieved (◆).",
             "• Red rings = a forecast that moved by ≥ the significant-movement threshold in one period.",
             "• A missing report is not plotted: no marker is drawn and no date is invented."]
    for i, n in enumerate(notes):
        cc = ws[f"S{32 + i}"]
        cc.value = n
        cc.font = f(8, False, TXT)
        cc.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(f"S{32 + i}:Z{32 + i}")
        ws.row_dimensions[32 + i].height = 24

    # ---------------------------------------------------------------- MTA chart
    ch = ScatterChart()
    ch.scatterStyle = "lineMarker"
    ch.style = 2
    xref = Reference(ws_eng, min_col=2, min_row=GR, max_row=GL)
    for k in range(1, NSLOT + 1):
        yref = Reference(ws_eng, min_col=2 + k, min_row=GR, max_row=GL)
        s = Series(yref, xref)
        series_title(s, f"Engine!${L(2 + k)}${GR - 1}")
        col = SERIES_COLORS[k - 1]
        s.graphicalProperties.line.solidFill = col
        s.graphicalProperties.line.width = 22000
        s.marker = Marker(symbol="circle", size=6)
        s.marker.graphicalProperties = GraphicalProperties(solidFill=col)
        s.marker.graphicalProperties.line.solidFill = "FFFFFF"
        s.smooth = False
        ch.series.append(s)
    # completion diagonal
    s = Series(Reference(ws_eng, min_col=2, min_row=DR, max_row=DR + 1), Reference(ws_eng, min_col=1, min_row=DR, max_row=DR + 1))
    series_title(s, f"Engine!$C${DR}")
    s.graphicalProperties.line.solidFill = "8C8C8C"
    s.graphicalProperties.line.dashStyle = "dash"
    s.graphicalProperties.line.width = 15875
    s.marker = Marker(symbol="none")
    s.smooth = False
    ch.series.append(s)
    # significant movement overlay
    ws_eng[f"F{FR - 1}"] = "Significant movement"
    ws_eng[f"G{FR - 1}"] = "Milestone achieved"
    for colidx, sym, size, color, nofill in ((6, "circle", 12, "E00000", True), (7, "diamond", 9, "222222", False)):
        s = Series(Reference(ws_eng, min_col=colidx, min_row=FR, max_row=FL), Reference(ws_eng, min_col=4, min_row=FR, max_row=FL))
        series_title(s, f"Engine!${L(colidx)}${FR - 1}")
        s.graphicalProperties.line.noFill = True
        s.marker = Marker(symbol=sym, size=size)
        gp = GraphicalProperties()
        if nofill:
            gp.noFill = True
            gp.line.solidFill = color
            gp.line.width = 25400
        else:
            gp.solidFill = color
            gp.line.solidFill = "FFFFFF"
        s.marker.graphicalProperties = gp
        s.smooth = False
        ch.series.append(s)
    style_axes(ch, DATE_S, DATE_S, "Reporting date", "Forecast completion date")
    ch.x_axis.majorGridlines = gridlines()
    ch.y_axis.majorGridlines = gridlines()
    ch.legend = Legend()
    ch.legend.position = "b"
    finish_chart(ch, 28.8, 13.9)
    ch.title = None
    ws.add_chart(ch, "B13")

    # ---------------------------------------------------------------- insight charts
    section(ws, "B40", "Schedule insights", "Z40")
    # Baseline vs current
    lc = LineChart()
    cats = Reference(ws_eng, min_col=11, min_row=SR, max_row=SLOT_LAST)
    for colidx, name, color, sym in ((20, "Original baseline", "A6A6A6", "dash"), (21, "Current forecast / actual", NAVY, "circle")):
        s = Series(Reference(ws_eng, min_col=colidx, min_row=SR, max_row=SLOT_LAST), title=name)
        s.graphicalProperties.line.noFill = True
        s.marker = Marker(symbol=sym, size=9 if sym == "dash" else 7)
        s.marker.graphicalProperties = GraphicalProperties(solidFill=color)
        s.marker.graphicalProperties.line.solidFill = color
        s.smooth = False
        lc.series.append(s)
    lc.set_categories(cats)
    lc.title = ctitle("Baseline vs current forecast  (MTA project)")
    style_axes(lc, None, DATE_S)
    lc.y_axis.majorGridlines = gridlines()
    lc.legend.position = "b"
    finish_chart(lc, 20.2, 9.3)
    ws.add_chart(lc, "B41")

    # forecast movement over time
    mc = ScatterChart()
    mc.scatterStyle = "lineMarker"
    for colidx, color in ((35, RED), (36, GREEN), (37, "A6A6A6")):
        s = Series(Reference(ws_eng, min_col=colidx, min_row=GR, max_row=GL), xref)
        series_title(s, f"Engine!${L(colidx)}${GR - 1}")
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.width = 22000
        s.marker = Marker(symbol="circle", size=6)
        s.marker.graphicalProperties = GraphicalProperties(solidFill=color)
        s.smooth = False
        mc.series.append(s)
    mc.title = ctitle("Forecast movement per reporting period  (visible MTA milestones)")
    style_axes(mc, DATE_S, "0", None, "Milestones")
    mc.y_axis.majorGridlines = gridlines()
    mc.y_axis.scaling.min = 0
    mc.legend.position = "b"
    finish_chart(mc, 21.9, 9.3)
    ws.add_chart(mc, "N41")

    # delay distribution
    bc = BarChart()
    bc.type = "col"
    s = Series(Reference(ws_eng, min_col=4, min_row=BR + 1, max_row=BIN_LAST), title="Milestones")
    s.graphicalProperties.solidFill = ACCENT
    s.graphicalProperties.line.noFill = True
    for i, colr in enumerate([GREEN, "70AD47", "5B9BD5", "F4B183", "ED7D31", "C55A11", RED]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = colr
        pt.graphicalProperties.line.noFill = True
        s.dPt.append(pt)
    bc.series.append(s)
    bc.set_categories(Reference(ws_eng, min_col=1, min_row=BR + 1, max_row=BIN_LAST))
    bc.title = ctitle("Delay distribution vs baseline")
    style_axes(bc, None, "0")
    bc.y_axis.majorGridlines = gridlines()
    bc.dataLabels = DataLabelList(showVal=True, showCatName=False, showSerName=False, showLegendKey=False, showPercent=False)
    bc.dataLabels.txPr = rich(800, True)
    bc.legend = None
    bc.gapWidth = 60
    finish_chart(bc, 13.5, 9.3)
    ws.add_chart(bc, "B60")

    # status breakdown
    sc = BarChart()
    sc.type = "bar"
    s = Series(Reference(ws_eng, min_col=7, min_row=BR + 1, max_row=ST_LAST), title="Milestones")
    for i, st in enumerate(STATUSES):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = STATUS_COL[st]
        pt.graphicalProperties.line.noFill = True
        s.dPt.append(pt)
    sc.series.append(s)
    sc.set_categories(Reference(ws_eng, min_col=6, min_row=BR + 1, max_row=ST_LAST))
    sc.title = ctitle("Milestone status breakdown")
    style_axes(sc, None, "0")
    sc.x_axis.scaling.orientation = "maxMin"
    sc.y_axis.majorGridlines = gridlines()
    sc.dataLabels = DataLabelList(showVal=True, showCatName=False, showSerName=False, showLegendKey=False, showPercent=False)
    sc.dataLabels.txPr = rich(800, True)
    sc.legend = None
    sc.gapWidth = 50
    finish_chart(sc, 13.5, 9.3)
    ws.add_chart(sc, "J60")

    # top 5
    tc = BarChart()
    tc.type = "bar"
    s = Series(Reference(ws_eng, min_col=13, min_row=BR + 1, max_row=BR + 5), title="Total delay (days)")
    s.graphicalProperties.solidFill = RED
    s.graphicalProperties.line.noFill = True
    tc.series.append(s)
    tc.set_categories(Reference(ws_eng, min_col=12, min_row=BR + 1, max_row=BR + 5))
    tc.title = ctitle("Top 5 most delayed open milestones (days)")
    style_axes(tc, None, "0")
    tc.x_axis.scaling.orientation = "maxMin"
    tc.y_axis.majorGridlines = gridlines()
    tc.dataLabels = DataLabelList(showVal=True, showCatName=False, showSerName=False, showLegendKey=False, showPercent=False)
    tc.dataLabels.txPr = rich(800, True)
    tc.legend = None
    tc.gapWidth = 50
    finish_chart(tc, 15.2, 9.3)
    ws.add_chart(tc, "R60")
    ws["B79"] = "Data: Milestone Data (tblMilestoneData) and Milestone Register (tblMilestones). Calculations: Delay Analysis. Chart series: Engine sheet."
    ws["B79"].font = f(8, False, MUTED, True)

    # ========================================================================= REPORTING HISTORY
    ws = ws_rh
    title_block(ws, "Reporting History  -  record a new snapshot",
                "1) Choose project and reporting date.  2) Enter the latest forecast for each milestone.  3) Copy the green 'ready to post' block and Paste Values below the last row of Milestone Data.")
    ws["B4"] = "Project"
    ws["B5"] = "Reporting date"
    for rr in (4, 5):
        ws[f"B{rr}"].font = f(10, True)
        ws.merge_cells(f"C{rr}:D{rr}")
        inp(ws[f"C{rr}"])
        ws[f"C{rr}"].font = f(11, True, "0000FF")
    ws["C4"] = projects[0]["name"]
    ws["C5"] = sample_data.D(2026, 11, 30)
    ws["C5"].number_format = DATE
    add_name(wb, "rhProject", "'Reporting History'!$C$4")
    add_name(wb, "rhDate", "'Reporting History'!$C$5")
    dv = DataValidation(type="list", formula1="lstProjectName", showErrorMessage=True, error="Choose a project.")
    dv.add("C4")
    ws.add_data_validation(dv)
    dv = DataValidation(type="date", operator="between", formula1="DATE(2020,1,1)", formula2="DATE(2040,12,31)",
                        showErrorMessage=True, errorTitle="Invalid date", error="Enter a valid reporting date.")
    dv.add("C5")
    ws.add_data_validation(dv)
    ws["F4"] = f'=IF(rhProjID="","Select a project.","Project "&rhProjID&": "&COUNTIF(da_proj,rhProjID)&" milestones, last report "&IF(_xlfn.MAXIFS({REP},{PID},rhProjID)=0,"none",TEXT(_xlfn.MAXIFS({REP},{PID},rhProjID),"dd-mmm-yyyy")))'
    ws["F5"] = (f'=IF(NOT(ISNUMBER(rhDate)),"Enter the reporting date of this snapshot (a real date).",IF(COUNTIFS({PID},rhProjID,{REP},rhDate)>0,'
                f'"Warning: "&COUNTIFS({PID},rhProjID,{REP},rhDate)&" observation(s) already exist for this date - those milestones will be flagged as duplicates.",'
                f'"OK - new reporting date. Ready to post: "&COUNTIF($L$10:$L$34,1)&" row(s)."))')
    ws["F4"].font = f(9, False, MUTED)
    ws["F5"].font = f(9, True, ACCENT)
    rh_hdr = ["#", "Milestone ID", "Milestone Name", "Baseline", "Last reported", "Last forecast", "NEW forecast",
              "Actual (if done)", "Comments", "Move (d)", "Check", "Ready", "Seq"]
    for j, h in enumerate(rh_hdr):
        ws.cell(9, 1 + j, h)
        hdr(ws.cell(9, 1 + j), calc=j not in (6, 7, 8))
    for col in "GHI":
        ws[f"{col}9"].fill = fill("B8860B")
    ws.row_dimensions[9].height = 30
    for k in range(1, 26):
        r = 9 + k
        ws[f"A{r}"] = k
        ws[f"B{r}"] = f'=IFERROR(INDEX(da_id,MATCH(A{r},da_rhrank,0)),"")'
        ws[f"C{r}"] = f'=IF(B{r}="","",INDEX(da_name,MATCH(A{r},da_rhrank,0)))'
        ws[f"D{r}"] = f'=IF(B{r}="","",INDEX(da_base,MATCH(A{r},da_rhrank,0)))'
        lastrep = f'_xlfn.MAXIFS({REP},{MID},B{r},{REP},"<"&rhDate)'
        ws[f"E{r}"] = f'=IF(OR(B{r}="",NOT(ISNUMBER(rhDate))),"",IF({lastrep}=0,"",{lastrep}))'
        ws[f"F{r}"] = f'=IF(E{r}="","",SUMIFS({FC},{MID},B{r},{REP},E{r}))'
        for col in "GHI":
            inp(ws[f"{col}{r}"])
            ws[f"{col}{r}"].font = f(10, False, "0000FF")
        ws[f"J{r}"] = f'=IF(OR(G{r}="",F{r}=""),"",IF(ISNUMBER(G{r}),G{r}-F{r},""))'
        done = f'_xlfn.MAXIFS({ACT},{MID},B{r})>0'
        ws[f"K{r}"] = (f'=IF(B{r}="","",IF(G{r}="",IF({done},"Completed - no update needed",IF(H{r}="","— awaiting forecast","✖ Forecast required")),'
                       f'IF(NOT(ISNUMBER(rhDate)),"✖ Enter reporting date",IF(COUNTIFS({MID},B{r},{REP},rhDate)>0,"✖ Duplicate - already recorded",'
                       f'IF(NOT(ISNUMBER(G{r})),"✖ Invalid forecast date",IF(AND(H{r}<>"",NOT(ISNUMBER(H{r}))),"✖ Invalid actual date",'
                       f'IF(AND(ISNUMBER(H{r}),H{r}>rhDate),"✖ Actual after reporting date",IF(AND(H{r}="",G{r}<rhDate),"⚠ Forecast is in the past - ready","✔ Ready"))))))))')
        ws[f"L{r}"] = f'=IF(OR(LEFT(K{r},1)="✔",LEFT(K{r},1)="⚠"),1,0)'
        ws[f"M{r}"] = f'=IF(L{r}=1,COUNTIF($L$10:L{r},1),"")'
        for col, fm in (("D", DATE), ("E", DATE), ("F", DATE), ("G", DATE), ("H", DATE), ("J", DAYS)):
            ws[f"{col}{r}"].number_format = fm
        for col in "ABCDEFJKLM":
            ws[f"{col}{r}"].font = f(9, False, MUTED if col in "LM" else TXT)
            ws[f"{col}{r}"].border = BORDER
        # paste-ready block O:S
        m = f"MATCH(ROWS(O$10:O{r}),$M$10:$M$34,0)"
        ws[f"O{r}"] = f'=IFERROR(INDEX($B$10:$B$34,{m}),"")'
        ws[f"P{r}"] = f'=IF(O{r}="","",rhDate)'
        ws[f"Q{r}"] = f'=IF(O{r}="","",INDEX($G$10:$G$34,{m}))'
        ws[f"R{r}"] = f'=IF(O{r}="","",IF(INDEX($H$10:$H$34,{m})="","",INDEX($H$10:$H$34,{m})))'
        ws[f"S{r}"] = f'=IF(O{r}="","",IF(INDEX($I$10:$I$34,{m})="","",INDEX($I$10:$I$34,{m})))'
        for col in "OPQRS":
            ws[f"{col}{r}"].fill = fill("E2F0D9")
            ws[f"{col}{r}"].border = BORDER
            ws[f"{col}{r}"].font = f(9)
        for col in "PQR":
            ws[f"{col}{r}"].number_format = DATE
    for j, h in enumerate(["Milestone ID", "Reporting Date", "Forecast Completion Date", "Actual Completion Date", "Comments"]):
        ws.cell(9, 15 + j, h)
        hdr(ws.cell(9, 15 + j))
        ws.cell(9, 15 + j).fill = fill("548235")
    ws["O8"] = "Ready to post  →  copy O10:S(last filled row), then on Milestone Data paste as VALUES into the first empty row below the table."
    ws["O8"].font = f(9, True, "548235")
    ws["A8"] = "Snapshot entry (yellow = input)"
    ws["A8"].font = f(9, True, NAVY)
    ws.conditional_formatting.add("K10:K34", FormulaRule(formula=['LEFT(K10,1)="✖"'], font=Font(color=RED, bold=True), fill=fill("F8D7DA")))
    ws.conditional_formatting.add("K10:K34", FormulaRule(formula=['LEFT(K10,1)="✔"'], font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add("K10:K34", FormulaRule(formula=['LEFT(K10,1)="⚠"'], font=Font(color="C55A11", bold=True)))
    ws.conditional_formatting.add("J10:J34", FormulaRule(formula=["AND(ISNUMBER(J10),J10>0)"], font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add("J10:J34", FormulaRule(formula=["AND(ISNUMBER(J10),J10<0)"], font=Font(color=GREEN, bold=True)))

    # history matrix
    HM = 38
    section(ws, f"A{HM}", "Forecast history matrix  -  every reported forecast for the selected project (history is never overwritten)", f"AD{HM}")
    ws[f"A{HM + 1}"] = "Colour = movement vs the previous reported forecast (first report vs baseline):  red = later, green = earlier, grey = unchanged. Blank = no report for that milestone on that date."
    ws[f"A{HM + 1}"].font = f(8, False, MUTED, True)
    HR2 = HM + 2
    ws[f"A{HR2}"], ws[f"B{HR2}"], ws[f"C{HR2}"], ws[f"D{HR2}"] = "#", "Milestone ID", "Milestone Name", "Baseline"
    for col in "ABCD":
        hdr(ws[f"{col}{HR2}"])
    NC = 26
    for j in range(NC):
        col = L(5 + j)
        cell = ws[f"{col}{HR2}"]
        if j == 0:
            cell.value = f'=IF(_xlfn.MINIFS({REP},{PID},rhProjID)=0,"",_xlfn.MINIFS({REP},{PID},rhProjID))'
        else:
            p = f"{L(4 + j)}{HR2}"
            cell.value = f'=IF({p}="","",IF(_xlfn.MINIFS({REP},{PID},rhProjID,{REP},">"&{p})=0,"",_xlfn.MINIFS({REP},{PID},rhProjID,{REP},">"&{p})))'
        hdr(cell)
        cell.number_format = DATE_S
        ws.column_dimensions[col].width = max(ws.column_dimensions[col].width or 0, 10)
    ws.row_dimensions[HR2].height = 30
    last_mcol = L(4 + NC)
    for k in range(1, 26):
        r = HR2 + k
        ws[f"A{r}"] = k
        ws[f"B{r}"] = f"=B{9 + k}"
        ws[f"C{r}"] = f"=C{9 + k}"
        ws[f"D{r}"] = f"=D{9 + k}"
        ws[f"D{r}"].number_format = DATE_S
        for j in range(NC):
            col = L(5 + j)
            ws[f"{col}{r}"] = f'=IF(OR($B{r}="",{col}${HR2}=""),"",IFERROR(INDEX({FC},MATCH($B{r}&"|"&TEXT({col}${HR2},"0"),{KEY},0)),""))'
            ws[f"{col}{r}"].number_format = DATE_S
        for cc in range(1, 5 + NC):
            ws.cell(r, cc).font = f(8)
            ws.cell(r, cc).border = BORDER
    mr = f"E{HR2 + 1}:{last_mcol}{HR2 + 25}"
    a = f"E{HR2 + 1}"
    prev = f"LOOKUP(2,1/ISNUMBER($D{HR2 + 1}:D{HR2 + 1}),$D{HR2 + 1}:D{HR2 + 1})"
    ws.conditional_formatting.add(mr, FormulaRule(formula=[f"AND(ISNUMBER({a}),{a}>{prev})"], fill=fill("F8D7DA"), font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(mr, FormulaRule(formula=[f"AND(ISNUMBER({a}),{a}<{prev})"], fill=fill("D9F0E3"), font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(mr, FormulaRule(formula=[f"AND(ISNUMBER({a}),{a}={prev})"], fill=fill("EDEDED"), font=Font(color="595959")))
    setw(ws, {"A": 4, "B": 15, "C": 30, "D": 11, "E": 11, "F": 11, "G": 12, "H": 12, "I": 26, "J": 8, "K": 28,
              "L": 6, "M": 5, "N": 2, "O": 12, "P": 12, "Q": 14, "R": 14, "S": 26})
    ws.freeze_panes = "D10"

    # ========================================================================= STATUS REPORT
    ws = ws_rpt
    ws.sheet_view.showGridLines = False
    setw(ws, {"A": 2, "B": 10, "C": 30, "D": 15, "E": 13, "F": 13, "G": 11, "H": 13, "I": 15, "J": 10, "K": 2})
    ws["B1"] = "Project Schedule Status Report"
    ws["B1"].font = f(18, True, NAVY)
    ws.row_dimensions[1].height = 30
    ws["B2"] = '="Generated from the Milestone Trend Analysis workbook  |  As of "&IF(ISNUMBER(rptAsOf),TEXT(rptAsOf,"dd-mmm-yyyy"),"-")&"  (reporting period follows the Dashboard filter)"'
    ws["B2"].font = f(9, False, MUTED, True)
    ws["B4"] = "Project"
    ws["B4"].font = f(10, True)
    ws.merge_cells("C4:D4")
    ws["C4"] = projects[1]["name"]
    inp(ws["C4"])
    ws["C4"].font = f(11, True, "0000FF")
    add_name(wb, "rptProject", "'Status Report'!$C$4")
    dv = DataValidation(type="list", formula1="lstProjectName", showErrorMessage=True, error="Choose a project.")
    dv.add("C4")
    ws.add_data_validation(dv)
    ws["F4"] = "Overall schedule status"
    ws["F4"].font = f(10, True)
    ws.merge_cells("F4:G4")
    ws.merge_cells("H4:I4")
    pinfo = lambda col: f'IFERROR(INDEX(tblProjects[{col}],MATCH(rptProjID,tblProjects[Project ID],0)),"")'  # noqa: E731
    section(ws, "B6", "1. Project summary", "J6")
    summ = [
        ("Project ID", "=rptProjID", None), ("Project manager", f"={pinfo('Project Manager')}", None),
        ("Sponsor", f"={pinfo('Sponsor')}", None), ("Reporting frequency", f"={pinfo('Reporting Frequency')}", None),
        ("Planned finish", f"={pinfo('Planned Finish Date')}", DATE),
        ("Forecast finish", '=IF(COUNTIF(da_proj,rptProjID)=0,"",_xlfn.MAXIFS(da_finish,da_proj,rptProjID))', DATE),
        ("Finish variance", '=IF(OR(E12="",E11=""),"",E12-E11)', DAYS),
    ]
    summ2 = [
        ("Milestones", "=COUNTIF(da_proj,rptProjID)", "0"),
        ("Completed", '=COUNTIFS(da_proj,rptProjID,da_status,"Completed")', "0"),
        ("Delayed", '=COUNTIFS(da_proj,rptProjID,da_status,"Delayed")', "0"),
        ("Ahead of schedule", '=COUNTIFS(da_proj,rptProjID,da_status,"Ahead of Schedule")', "0"),
        ("On schedule", '=COUNTIFS(da_proj,rptProjID,da_status,"On Schedule")', "0"),
        ("Not yet started", '=COUNTIFS(da_proj,rptProjID,da_status,"Not Yet Started")', "0"),
        ("High-risk milestones", '=COUNTIFS(da_proj,rptProjID,da_risklevel,"High")', "0"),
    ]
    for i, ((l1, f1, m1), (l2, f2, m2)) in enumerate(zip(summ, summ2)):
        r = 7 + i
        ws.merge_cells(f"B{r}:C{r}")
        ws[f"B{r}"] = l1
        ws[f"B{r}"].font = f(9, True, MUTED)
        ws.merge_cells(f"D{r}:E{r}")
        ws[f"D{r}"] = f1
        ws[f"D{r}"].font = f(10, True)
        ws[f"D{r}"].alignment = Alignment(horizontal="left")
        if m1:
            ws[f"D{r}"].number_format = m1
        ws.merge_cells(f"G{r}:H{r}")
        ws[f"G{r}"] = l2
        ws[f"G{r}"].font = f(9, True, MUTED)
        ws[f"I{r}"] = f2
        ws[f"I{r}"].font = f(10, True)
        ws[f"I{r}"].number_format = m2
        for cc in "BCDEGHI":
            ws[f"{cc}{r}"].border = Border(bottom=Side(style="hair", color=GRID))
    # finish variance formula refers to D-column merged cells
    ws["D13"] = '=IF(OR(D12="",D11=""),"",D12-D11)'
    ws["D13"].number_format = DAYS
    ws["H4"] = ('=IF(I7=0,"-",IF(OR(I9/I7>=cfgRedPct,N(D13)>cfgHighVar),"RED",IF(OR(I9/I7>=cfgAmberPct,I13>0,N(D13)>cfgTolerance),"AMBER","GREEN")))')
    ws["H4"].font = f(14, True, "FFFFFF")
    ws["H4"].alignment = Alignment(horizontal="center", vertical="center")
    ws.conditional_formatting.add("H4:I4", FormulaRule(formula=['$H$4="RED"'], fill=fill(RED)))
    ws.conditional_formatting.add("H4:I4", FormulaRule(formula=['$H$4="AMBER"'], fill=fill("ED7D31")))
    ws.conditional_formatting.add("H4:I4", FormulaRule(formula=['$H$4="GREEN"'], fill=fill(GREEN)))
    ws.row_dimensions[4].height = 26
    ws["B14"] = ('="Rule: RED if delayed share ≥ "&TEXT(cfgRedPct,"0%")&" or finish variance > "&cfgHighVar&" d; AMBER if delayed share ≥ "'
                 '&TEXT(cfgAmberPct,"0%")&", any high-risk milestone, or finish variance > "&cfgTolerance&" d; otherwise GREEN."')
    ws["B14"].font = f(8, False, MUTED, True)

    def report_list(start, title, rankname, n, cols):
        section(ws, f"B{start}", title, f"J{start}")
        for j, (h, *_rest) in enumerate(cols):
            ws.cell(start + 1, 2 + j, h)
            hdr(ws.cell(start + 1, 2 + j))
        for k in range(1, n + 1):
            r = start + 1 + k
            m = f"MATCH({k},{rankname},0)"
            for j, (h, key, fmt) in enumerate(cols):
                cell = ws.cell(r, 2 + j)
                if key is None:
                    continue
                cell.value = f'=IFERROR(INDEX(da_{key},{m}),"")'
                cell.font = f(9)
                cell.border = Border(bottom=Side(style="hair", color=GRID))
                if fmt:
                    cell.number_format = fmt
                    cell.alignment = Alignment(horizontal="center")
        return start + 1 + n

    perf_cols = [("ID", "id", None), ("Milestone", "name", None), ("Owner", "owner", None), ("Baseline", "base", DATE),
                 ("Forecast / actual", "finish", DATE), ("Variance", "totdelay", DAYS), ("Last move", "indicator", None),
                 ("Status", "status", None), ("Risk", "risklevel", None)]
    end = report_list(16, "2. Milestone performance", "da_rptrank", 20, perf_cols)
    status_cf(ws, f"I18:I{end}", "I18")
    movement_cf(ws, f"H18:H{end}", "H18")
    ws.conditional_formatting.add(f"G18:G{end}", FormulaRule(formula=["AND(ISNUMBER(G18),G18>0)"], font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(f"G18:G{end}", FormulaRule(formula=["AND(ISNUMBER(G18),G18<0)"], font=Font(color=GREEN, bold=True)))
    s2 = end + 2
    del_cols = [("ID", "id", None), ("Milestone", "name", None), ("Owner", "owner", None), ("Baseline", "base", DATE),
                ("Current forecast", "curfc", DATE), ("Variance", "basevar", DAYS), ("Slips", "slips", "0"),
                ("Priority", "prio", None), ("Risk", "risklevel", None)]
    end2 = report_list(s2, "3. Delayed milestones", "da_rptdelayrank", 8, del_cols)
    ws.cell(s2 + 1, 12).value = None
    ws[f"B{end2 + 1}"] = f'=IF(COUNTIF(da_rptdelay,1)=0,"No delayed milestones.",IF(COUNTIF(da_rptdelay,1)>8,"+ "&(COUNTIF(da_rptdelay,1)-8)&" more - see Delay Analysis.",""))'
    ws[f"B{end2 + 1}"].font = f(8, False, MUTED, True)
    s3 = end2 + 3
    chg_cols = [("ID", "id", None), ("Milestone", "name", None), ("Previous report", "prevrep", DATE),
                ("Previous forecast", "prevfc", DATE), ("Current forecast", "curfc", DATE), ("Movement", "move", DAYS),
                ("Direction", "indicator", None), ("Status", "status", None), (None, None, None)]
    chg_cols = [c_ for c_ in chg_cols if c_[0]]
    end3 = report_list(s3, "4. Forecast changes in the latest reporting period", "da_rptchgrank", 8, chg_cols)
    movement_cf(ws, f"H{s3 + 2}:H{end3}", f"H{s3 + 2}")
    ws[f"B{end3 + 1}"] = '=IF(COUNTIF(da_rptchg,1)=0,"No forecast changes in the latest report.","")'
    ws[f"B{end3 + 1}"].font = f(8, False, MUTED, True)
    s4 = end3 + 3
    section(ws, f"B{s4}", "5. Risk observations", f"J{s4}")
    for j, h in enumerate(["ID", "Milestone", "Risk", "Observations"]):
        col = "BCDE"[j]
        ws[f"{col}{s4 + 1}"] = h
        hdr(ws[f"{col}{s4 + 1}"])
    ws.merge_cells(f"E{s4 + 1}:J{s4 + 1}")
    for k in range(1, 11):
        r = s4 + 1 + k
        m = f"MATCH({k},da_rptriskrank,0)"
        ws[f"B{r}"] = f'=IFERROR(INDEX(da_id,{m}),"")'
        ws[f"C{r}"] = f'=IFERROR(INDEX(da_name,{m}),"")'
        ws[f"D{r}"] = f'=IFERROR(INDEX(da_risklevel,{m}),"")'
        ws[f"E{r}"] = f'=IFERROR(INDEX(da_risknotes,{m}),"")'
        ws.merge_cells(f"E{r}:J{r}")
        for col in "BCDE":
            ws[f"{col}{r}"].font = f(9)
            ws[f"{col}{r}"].border = Border(bottom=Side(style="hair", color=GRID))
        ws[f"E{r}"].alignment = Alignment(wrap_text=True)
    end4 = s4 + 11
    ws.conditional_formatting.add(f"D{s4 + 2}:D{end4}", FormulaRule(formula=[f'D{s4 + 2}="High"'], font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(f"D{s4 + 2}:D{end4}", FormulaRule(formula=[f'D{s4 + 2}="Medium"'], font=Font(color="C55A11", bold=True)))
    ws[f"B{end4 + 1}"] = '=IF(COUNTIF(da_rptrisk,1)=0,"No risk observations.","")'
    ws[f"B{end4 + 1}"].font = f(8, False, MUTED, True)
    ws.print_area = f"A1:K{end4 + 2}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.oddFooter.center.text = "Page &P of &N"
    ws.oddHeader.right.text = "&D"

    # ========================================================================= INSTRUCTIONS
    build_instructions(ws_ins)

    # ========================================================================= protection & print
    for wsx, pw in ((ws_dash, None), (ws_da, None), (ws_eng, None), (ws_rpt, None), (ws_rh, None)):
        wsx.protection.sheet = True
        wsx.protection.autoFilter = False
        wsx.protection.sort = False
    ws_dash.print_area = "A1:AA79"
    ws_dash.page_setup.orientation = "landscape"
    ws_dash.page_setup.fitToWidth = 1
    ws_dash.page_setup.fitToHeight = 1
    ws_dash.sheet_properties.pageSetUpPr.fitToPage = True
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(OUT)
    print("saved", OUT, len(milestones), "milestones", len(obs), "observations")


def status_cf(ws, rng, first):
    for st, col in STATUS_COL.items():
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="{st}"'], font=Font(color=col, bold=True)))


def movement_cf(ws, rng, first):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({first},1)="▲"'], font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({first},1)="▼"'], font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({first},1)="●"'], font=Font(color=GREYC, bold=True)))


INSTRUCTIONS = [
    ("h", "Advanced Milestone Trend Analysis - User Guide"),
    ("p", "This workbook tracks how forecast milestone completion dates change from one reporting period to the next. It follows Milestone Trend Analysis (MTA) practice: each milestone is a line, the X axis is the reporting date, the Y axis is the forecast completion date."),
    ("s", "1. Sheet overview"),
    ("b", "Dashboard - KPI cards, the MTA chart, five supporting charts and the four filters (Project, Reporting Period, Category, Status)."),
    ("b", "Milestone Data (tblMilestoneData) - the forecast history. One row = one milestone forecast observation on one reporting date."),
    ("b", "Project Setup - projects table (tblProjects), analysis thresholds (named cells) and dropdown reference lists."),
    ("b", "Milestone Register (tblMilestones) - master list of milestones with unique Milestone IDs. Feeds every dropdown and lookup."),
    ("b", "Reporting History - snapshot entry form (validates and builds paste-ready rows) and a forecast history matrix per project."),
    ("b", "Delay Analysis - one row per milestone: baseline/first/previous/current/actual dates, variances, status, movement and risk flags. Also a single-milestone historical comparison panel."),
    ("b", "Status Report - printable one-project status report (A4 portrait, fits one page wide)."),
    ("b", "Engine - helper calculations for lists and chart series. Formula-only; do not edit."),
    ("s", "2. Colour legend"),
    ("b", "Yellow cells / navy headers = inputs.  Grey headers = calculated columns (do not type over them)."),
    ("b", "Movement: ▲ red = forecast moved later; ▼ green = moved earlier; ● grey = no change; ◆ = first report."),
    ("b", "Status colours: Completed navy, Delayed red, Ahead of Schedule green, On Schedule blue, Not Yet Started grey."),
    ("s", "3. Setting up a new project"),
    ("n", "1. Project Setup: add a row directly below tblProjects (Project ID such as PRJ-004, name, manager, dates, frequency). The table expands by itself."),
    ("n", "2. Milestone Register: add one row per milestone. Pick the Project ID from the dropdown and give each milestone a unique ID (duplicates are rejected and highlighted). Enter the original baseline date."),
    ("n", "3. Current Status in the register fills in automatically once forecasts are recorded."),
    ("s", "4. Recording a reporting snapshot (every reporting cycle)"),
    ("p", "Option A - Reporting History form (recommended):"),
    ("n", "1. Choose the Project and type the Reporting Date (C4:C5)."),
    ("n", "2. For each milestone type the NEW forecast date (and the actual date if the milestone is complete, plus optional comments). The last forecast and the movement are shown for reference."),
    ("n", "3. Check column K: ✔ Ready, ⚠ warning (still postable), ✖ must be fixed. Duplicates (same milestone + reporting date already recorded) are blocked."),
    ("n", "4. Copy the green block O10:S(last filled row). On Milestone Data select column A of the first empty row directly below the table and use Paste Special > Values. The table expands and every calculated column fills automatically."),
    ("n", "5. Clear the yellow inputs (G:I) ready for the next cycle."),
    ("p", "Option B - direct entry: type a new row at the bottom of Milestone Data (Milestone ID, Reporting Date, Forecast, Actual if complete, Comments). Dropdowns and validation prevent unknown IDs, invalid dates and duplicate milestone/date pairs."),
    ("p", "Never edit or delete earlier rows: they are the forecast history. A correction to a past observation should be made deliberately and recorded in Comments."),
    ("p", "A milestone is Completed once an observation carries an Actual Completion Date. Record the actual date as the forecast date in that observation too, so the line ends on the completion line."),
    ("s", "5. Using the Dashboard"),
    ("b", "Project: 'All Projects' or one project. KPIs and the distribution, status and top-5 charts cover every project in the filter. The MTA chart, its key panel, the baseline-vs-current and the movement charts show one project: the selected project, or the first project when 'All Projects' is chosen."),
    ("b", "Reporting Period: '(Latest)' uses each project's latest report. Choosing a date turns the whole workbook back to that point in time: only observations on or before it are used (status, variances, risks, chart)."),
    ("b", "Category / Status: limit KPIs and charts to matching milestones. Status is evaluated as of the chosen reporting period."),
    ("b", "Show / hide lines: set Show = No in the key panel next to the chart (Z15:Z29). In Excel 365 you can also use the chart filter button."),
    ("b", "The chart takes up to 15 milestones per project (after filtering) and up to 40 reporting periods per project. New reporting dates and milestones appear automatically."),
    ("s", "6. Calculations (Delay Analysis)"),
    ("b", "Baseline Variance = Current Forecast - Original Baseline (calendar days)."),
    ("b", "Forecast Movement = Current Forecast - Previous Forecast (the previous observation of that milestone, whatever the reporting frequency)."),
    ("b", "Total Delay = (Actual Completion if completed, otherwise Current Forecast) - Original Baseline."),
    ("b", "Schedule Status: Not Yet Started (no observation as of the period) > Completed (actual date recorded) > Delayed (variance > tolerance) > Ahead of Schedule (variance < -tolerance) > On Schedule. Completed milestones are never labelled Delayed, however much their forecasts moved earlier on."),
    ("b", "Risk flags (open milestones): Repeated Slippage (slip count ≥ threshold), Significant Change (largest single move ≥ threshold), High Baseline Variance, Missing Update (the project issued a later report without this milestone), Approaching Deadline (due within the window or overdue). Risk Level: 0 None, 1 Low, 2 Medium, 3+ High."),
    ("b", "Thresholds live on Project Setup (cfgTolerance, cfgSigMove, cfgSlipCount, cfgHighVar, cfgWindow, cfgAmberPct, cfgRedPct) and can be changed at any time."),
    ("s", "7. Technical architecture"),
    ("b", "Data layer: three Excel Tables (tblProjects, tblMilestones, tblMilestoneData). Structured references let formulas and dropdowns grow with the tables. Lookup columns in tblMilestoneData are table calculated columns, so they fill automatically for new rows."),
    ("b", "Observation key: tblMilestoneData[Obs Key] = Milestone ID & '|' & reporting-date serial. One INDEX/MATCH on this key returns the forecast for any milestone/date pair without inventing values."),
    ("b", "Previous forecast per observation: MAXIFS(Reporting Date, Milestone ID = this, Reporting Date < this) then SUMIFS of the forecast on that date. This works for irregular reporting frequencies and missed reports."),
    ("b", "Delay Analysis: 200 pre-built formula rows read the register by position (INDEX(tblMilestones[...], n)). As-of logic uses MINIFS/MAXIFS with '<=' as-of criteria. Each column is exposed as a named range (da_id, da_status, da_totdelay...) used by the Dashboard, Engine and Status Report."),
    ("b", "Dashboard filters are named cells (selProject, selPeriod, selCategory, selStatus). Dropdown lists are dynamic named ranges (lstProjectFilter, lstPeriod...) built with INDEX so they grow without volatile OFFSET/INDIRECT."),
    ("b", "Sorted unique reporting dates are built without dynamic arrays: each row is MINIFS(date, date > previous row), which returns the next distinct date."),
    ("b", "MTA chart: an XY (scatter) chart, so both axes are true date axes and the spacing between points matches the real gaps between reports (weekly, bi-weekly, monthly or irregular). The Engine grid holds one row per reporting date and one column per milestone slot. A cell with no observation returns the NA() error value, which Excel does not plot. A missing report never shows up as zero and is never filled with an invented forecast. Overlay series mark significant moves (red rings) and achieved milestones (diamonds). A dashed diagonal is the completion line."),
    ("b", "Slots: Delay Analysis gives each milestone that passes the MTA filters a running number (MTA Slot). Engine slot k picks the milestone numbered k, so the chart series never need re-pointing."),
    ("b", "Why fixed ranges instead of FILTER/SORT/UNIQUE spills: chart series and data validation need fixed addresses, and spills cannot live inside Excel Tables. The workbook therefore uses dynamic-array-safe functions that also work in Excel 2019 and in LibreOffice: INDEX/MATCH (same result as XLOOKUP), MAXIFS/MINIFS, COUNTIFS/SUMIFS, TEXTJOIN. No volatile functions (OFFSET, INDIRECT, TODAY) drive calculations."),
    ("b", "Protection: Dashboard, Delay Analysis, Reporting History, Status Report and Engine are protected without a password (Review > Unprotect Sheet), with input cells unlocked. Data sheets are not protected, because Excel cannot expand a table on a protected sheet."),
    ("s", "8. Optional VBA (not required)"),
    ("p", "The workbook is a macro-free .xlsx. vba/MTA_Snapshot.bas in the project repository has an optional 'PostSnapshot' macro. It appends the ready rows from Reporting History to tblMilestoneData, rejects duplicates and clears the form, which replaces the copy / paste-values step. To use it: save as .xlsm, press Alt+F11, File > Import File, then assign the macro to a button. The non-VBA alternative is Option A above."),
    ("s", "9. Capacity and limits"),
    ("b", "200 milestones in total (Delay Analysis rows); 15 milestone lines and 40 reporting periods per project on the MTA chart; 25 milestones per project on the Reporting History form; 26 reporting dates in the history matrix; 20 milestones in the printable report table. Each limit is a block of copied formula rows and can be extended by copying the last row down."),
]


def build_instructions(ws):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 140
    r = 1
    for kind, text in INSTRUCTIONS:
        c = ws.cell(r, 2, text if kind != "b" else "•  " + text)
        if kind == "h":
            c.font = f(18, True, NAVY)
            ws.row_dimensions[r].height = 30
        elif kind == "s":
            r += 1
            c = ws.cell(r, 2, text)
            c.font = f(12, True, "FFFFFF")
            c.fill = fill(NAVY)
            c.alignment = Alignment(indent=1, vertical="center")
            ws.cell(r - 1, 2).value = None
            ws.row_dimensions[r].height = 20
        else:
            c.font = f(10)
            c.alignment = Alignment(wrap_text=True, vertical="top", indent=1 if kind in ("b", "n") else 0)
            ws.row_dimensions[r].height = 15 * max(1, -(-len(text) // 150))
        r += 1


if __name__ == "__main__":
    main()
