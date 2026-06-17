"""
Generate Template_simple.xlsx — the presentation template for simple_proposal().

Run once (or whenever template layout changes):
    uv run python create_simple_template.py

Deploy generated file to SharePoint @tools/resources/ via bid setup.
No cell merging used — text overflows naturally from the anchor cell.

Page 1 layout:
  Rows 1–5   Entity header block  ← print_title_rows="1:5" (repeats on every page)
  Row  6     Spacer
  Row  7     Proposal type        ("COMMERCIAL PROPOSAL" / "TECHNICAL PROPOSAL")
  Row  8     Spacer
  Rows 9–14  Metadata left         C = "Label Value" (Attention to … Project Name)
  Rows 9–12  Metadata right        F = "Label Value" from Config A28:B35 (overflows G/H)
  Row  18    Spacer
  Row  19    Column header        (blue fill, white bold text; Python updates F/G with currency)
  Row  20+   BOQ data

Page 2+ layout:
  Rows 1–5 repeated (entity header), then BOQ data continues.

Column widths:
  A  No          (5)
  B  SN          (4)
  C  Description (55)   ← also holds metadata labels (pre-styled bold in template)
  D  Qty         (5)    ← also holds metadata values (overflow into E/F/G)
  E  Unit        (5)
  F  Unit Price  (11)   — hidden in technical mode
  G  Total       (12)   — hidden in technical mode
  H  Scope       (8)
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins

OUTPUT = "Template_simple.xlsx"

BLUE  = "FF005BBF"   # ARGB — Jason Blue, fully opaque
GREY  = "FF595959"   # ARGB — dark grey, fully opaque
WHITE = "FFFFFFFF"   # ARGB — white, fully opaque
FONT  = "Aptos"      # entity header rows only
BODY  = "Arial"      # all data/body content


def af(name, bold=False, size=10, color="000000", italic=False):
    return Font(name=name, bold=bold, size=size, color=color, italic=italic)


def fill(hex_color):
    return PatternFill(fill_type="solid", fgColor=hex_color)


def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Proposal"

# ---------------------------------------------------------------------------
# Column widths
# ---------------------------------------------------------------------------
col_widths = {
    "A": 5,   # No
    "B": 4,   # SN
    "C": 55,  # Description  ← also metadata labels
    "D": 5,   # Qty          ← also metadata values (overflow right)
    "E": 5,   # Unit
    "F": 11,  # Unit Price
    "G": 12,  # Total
    "H": 8,   # Scope
}
for col, w in col_widths.items():
    ws.column_dimensions[col].width = w

# ===========================================================================
# REPEAT BLOCK — rows 1–5  (print_title_rows = "1:5")
# ===========================================================================

# ROW 1 — Entity name
ws.row_dimensions[1].height = 22
c = ws["A1"]
c.value = "[ENTITY NAME]"
c.font = af(FONT, bold=True, size=14, color=BLUE)
c.alignment = align()

# ROW 2 — Address
ws.row_dimensions[2].height = 13
c = ws["A2"]
c.value = "[ADDRESS]"
c.font = af(FONT, size=9, color=GREY)
c.alignment = align()

# ROW 3 — Contact (extra height adds gap between text and blue rule below)
ws.row_dimensions[3].height = 20
c = ws["A3"]
c.value = "[TEL  |  FAX  |  WEBSITE  |  Co. Reg. No.]"
c.font = af(FONT, size=9, color=GREY)
c.alignment = align()

# ROW 4 — Blue rule
ws.row_dimensions[4].height = 2
blue_fill = fill(BLUE)
for col in range(1, 9):
    ws.cell(row=4, column=col).fill = blue_fill

# ROW 5 — Spacer
ws.row_dimensions[5].height = 3

# ===========================================================================
# DATA AREA — rows 6+ (page 1 only for metadata; page 2+ sees rows 1–5 then BOQ)
# ===========================================================================

# ROW 6 — Spacer
ws.row_dimensions[6].height = 8

# ROW 7 — Proposal type anchor  (Python writes "COMMERCIAL PROPOSAL" etc. to C7)
ws.row_dimensions[7].height = 17
c = ws["C7"]
c.value = "[PROPOSAL TYPE]"
c.font = af(BODY, bold=True, size=12)
c.alignment = align()

# ROW 8 — Spacer
ws.row_dimensions[8].height = 5

# ROWS 9–16 — Metadata two-column layout
# Left (C): 6 fields from Config B21–B26
# Right (F): dynamic fields from Config A28:B35 (max 4 visible)
# Python writes "Label Value" combined text at runtime.
LEFT_META_LABELS = [
    "Attention to:",
    "Designation:",
    "Customer:",
    "Client Reference:",
    "Ref Doc No:",
    "Project Name:",
]
RIGHT_META_LABELS = [
    "Sales:",
    "Jason Ref:",
    "Revision Num:",
    "Date:",
]
_META_START = 9
for i, label in enumerate(LEFT_META_LABELS):
    row = _META_START + i
    ws.row_dimensions[row].height = 14
    c = ws.cell(row=row, column=3)   # column C
    c.value = label
    c.font = af(BODY, bold=False, size=10)
    c.alignment = align()

for i, label in enumerate(RIGHT_META_LABELS):
    row = _META_START + i
    c = ws.cell(row=row, column=6)   # column F
    c.value = label
    c.font = af(BODY, bold=False, size=10)
    c.alignment = align("left")      # left so text overflows right into G/H

# Derived row numbers (must match functions.py constants)
_META_END    = _META_START + len(LEFT_META_LABELS) - 1   # 14
_SPACER2_ROW = _META_END + 1                               # 15
_HDR_ROW     = _SPACER2_ROW + 1                            # 16
_DATA_ROW    = _HDR_ROW + 1                                # 17

# ROW _SPACER2_ROW — Spacer before column header (Python writes header dynamically)
ws.row_dimensions[_SPACER2_ROW].height = 6

# ROW _DATA_ROW — BOQ anchor placeholder (Python inserts column header + BOQ above)
ws.row_dimensions[_DATA_ROW].height = 14
c = ws.cell(row=_DATA_ROW, column=3)
c.value = "(Python inserts column header + BOQ rows here)"
c.font = af(FONT, size=8, color="FFAAAAAA")

# ---------------------------------------------------------------------------
# Page setup
# ---------------------------------------------------------------------------
ws.page_setup.orientation = "portrait"
ws.page_setup.paperSize   = 9      # A4
ws.page_setup.fitToPage   = True
ws.page_setup.fitToWidth  = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True

ws.page_margins = PageMargins(
    left=0.55, right=0.55,
    top=0.75,  bottom=0.75,
    header=0.3, footer=0.3,
)

ws.freeze_panes = "A6"         # freeze entity header when viewing in Excel
ws.print_title_rows = "1:5"    # repeat entity header on every printed page / PDF page

# Pre-style all 8 data columns.
# xlwings value writes don't disturb cell formatting, so these persist at runtime.
# On Mac, xlwings vertical_alignment silently fails — template pre-styling is the fix.
# Rows 9–20 (metadata + dynamic header area): center vertical alignment so the
# header row is always middle-aligned regardless of which row Python picks.
# Rows 21+ (BOQ data area): top alignment for wrapped multi-line descriptions.
_col_h = ["right", "center", "left", "right", "center", "right", "right", "center"]
for _r in range(_META_START, _HDR_ROW + 1):       # rows 9–16: center vertical
    for _ci, h in enumerate(_col_h, 1):
        ws.cell(row=_r, column=_ci).alignment = align(h, v="center")
# Column F in metadata rows overrides right→left so right-column text overflows into G/H
for _r in range(_META_START, _HDR_ROW + 1):
    ws.cell(row=_r, column=6).alignment = align("left", v="center")
for _r in range(_DATA_ROW, _META_START + 320):    # rows 17–328: top vertical
    for _ci, h in enumerate(_col_h, 1):
        ws.cell(row=_r, column=_ci).alignment = align(h, v="top")

# Footer: centered "Page X of Y", Arial 10
ws.oddFooter.center.text = '&"Arial,Regular"&10Page &P of &N'

wb.save(OUTPUT)
print(f"Created {OUTPUT}")
