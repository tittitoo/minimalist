"""
Read formatting (fills, fonts, borders) from PERSONAL.XLSB Design sheet rows.
Run with Excel open (PERSONAL.XLSB auto-loaded).

Two-step process:
  Step 1: Run this script — it detects no BorderDump sheet and prints the VBA
          code + instructions.  Paste the VBA into PERSONAL.XLSB and run it.
  Step 2: Run this script again — BorderDump sheet now exists, full output printed.

Usage:
    uv run python read_design_rows.py
"""
import sys
from pathlib import Path
import xlwings as xw
import openpyxl

# Design rows used by copy_design_row() and apply_lastrow_border()
DESIGN_ROWS = [5, 7, 8, 9, 11, 13, 15, 17, 18, 19, 21]

SAMPLE_COLS = list("ABCDEFGH")

_DUMP_SHEET = "BorderDump"

# ---------------------------------------------------------------------------
# VBA sub to run once inside PERSONAL.XLSB
# ---------------------------------------------------------------------------
_VBA_CODE = '''\
Sub DumpDesignBorders()
    Dim wsD As Worksheet, wsO As Worksheet
    Set wsD = ThisWorkbook.Sheets("Design")

    On Error Resume Next
    Set wsO = ThisWorkbook.Sheets("BorderDump")
    On Error GoTo 0
    If wsO Is Nothing Then
        Set wsO = ThisWorkbook.Sheets.Add()
        wsO.Name = "BorderDump"
    Else
        wsO.Cells.Clear
    End If

    Dim designRows As Variant
    designRows = Array(5, 7, 8, 9, 11, 13, 15, 17, 18, 19, 21)

    Dim edgeNames  As Variant: edgeNames  = Array("top", "bottom", "left", "right", "inside_h")
    Dim edgeConsts As Variant: edgeConsts = Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight, xlInsideHorizontal)

    Dim outRow As Long: outRow = 1
    wsO.Cells(outRow, 1).Value = "row"
    wsO.Cells(outRow, 2).Value = "col"
    wsO.Cells(outRow, 3).Value = "edge"
    wsO.Cells(outRow, 4).Value = "style"
    wsO.Cells(outRow, 5).Value = "weight"
    wsO.Cells(outRow, 6).Value = "colorR"
    wsO.Cells(outRow, 7).Value = "colorG"
    wsO.Cells(outRow, 8).Value = "colorB"
    outRow = 2

    Dim r As Variant
    For Each r In designRows
        Dim c As Integer
        For c = 1 To 8
            Dim i As Integer
            For i = 0 To 4
                Dim bdr As Border
                Set bdr = wsD.Cells(r, c).Borders(edgeConsts(i))
                If bdr.LineStyle <> xlNone Then
                    Dim bgr As Long: bgr = bdr.Color
                    wsO.Cells(outRow, 1).Value = r
                    wsO.Cells(outRow, 2).Value = c
                    wsO.Cells(outRow, 3).Value = edgeNames(i)
                    wsO.Cells(outRow, 4).Value = bdr.LineStyle
                    wsO.Cells(outRow, 5).Value = bdr.Weight
                    wsO.Cells(outRow, 6).Value = bgr And 255
                    wsO.Cells(outRow, 7).Value = (bgr \\ 256) And 255
                    wsO.Cells(outRow, 8).Value = (bgr \\ 65536) And 255
                    outRow = outRow + 1
                End If
            Next i
        Next c
    Next r
    MsgBox "Done — " & (outRow - 2) & " borders written to 'BorderDump' sheet."
End Sub
'''


def _color_hex(rgb):
    if rgb is None:
        return "None"
    if isinstance(rgb, tuple):
        r, g, b = rgb
        return f"#{r:02X}{g:02X}{b:02X}"
    return str(rgb)


# ---------------------------------------------------------------------------
# Read BorderDump sheet (written by VBA)
# ---------------------------------------------------------------------------
def read_border_dump(personal_wb):
    """Read border data from the BorderDump sheet in PERSONAL.XLSB."""
    dump = personal_wb.sheets[_DUMP_SHEET]

    row_borders = {}  # row_num → {edge → {style, weight, color}}
    r = 2
    while True:
        val = dump.range(f"A{r}").value
        if val is None:
            break
        row_num = int(val)
        col     = int(dump.range(f"B{r}").value)
        edge    = dump.range(f"C{r}").value
        style   = int(dump.range(f"D{r}").value)
        weight  = int(dump.range(f"E{r}").value)
        cr      = int(dump.range(f"F{r}").value or 0)
        cg      = int(dump.range(f"G{r}").value or 0)
        cb      = int(dump.range(f"H{r}").value or 0)

        # Collapse per-column duplicates: keep first occurrence of each edge
        row_borders.setdefault(row_num, {})
        if edge not in row_borders[row_num]:
            row_borders[row_num][edge] = {
                "style":  style,
                "weight": weight,
                "color":  f"#{cr:02X}{cg:02X}{cb:02X}",
            }
        r += 1

    return row_borders


# ---------------------------------------------------------------------------
# Read fills + fonts via xlwings
# ---------------------------------------------------------------------------
def read_fills_fonts(design_sheet):
    result = {}
    for row_num in DESIGN_ROWS:
        fills, fonts = {}, {}
        for col in SAMPLE_COLS:
            cell = design_sheet.range(f"{col}{row_num}")
            try:
                fills[col] = cell.color
            except Exception:
                fills[col] = None
            try:
                fonts[col] = {
                    "bold":  cell.font.bold,
                    "color": cell.font.color,
                    "size":  cell.font.size,
                }
            except Exception:
                fonts[col] = {}
        result[row_num] = {"fills": fills, "fonts": fonts}
    return result


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------
def print_row_summary(row_num, fills, fonts, borders):
    print(f"Row {row_num}:")
    non_none = {k: v for k, v in fills.items() if v is not None}
    if not non_none:
        print("  fill:       None (all columns)")
    elif len(set(str(v) for v in non_none.values())) == 1:
        print(f"  fill:       {_color_hex(next(iter(non_none.values())))} (col{'s' if len(non_none) > 1 else ''} {','.join(non_none)})")
    else:
        for col, v in non_none.items():
            print(f"  fill[{col}]:    {_color_hex(v)}")

    unique_bold  = set(str(f.get("bold"))  for f in fonts.values() if isinstance(f, dict))
    unique_color = set(str(f.get("color")) for f in fonts.values() if isinstance(f, dict))
    print(f"  font_bold:  {', '.join(sorted(unique_bold))}")
    print(f"  font_color: {', '.join(sorted(str(x) for x in unique_color))}")

    style_names  = {1: "continuous", 2: "dashed", 4: "dotted"}
    weight_names = {1: "hairline", 2: "thin", 3: "medium", 4: "thick"}
    if borders:
        for edge, info in sorted(borders.items()):
            sn = style_names.get(info["style"],  str(info["style"]))
            wn = weight_names.get(info["weight"], str(info["weight"]))
            print(f"  border[{edge:9s}]: {sn} {wn} {info['color']}")
    else:
        print("  borders:    (none)")
    print()


def emit_python_dict(fills_fonts, all_borders):
    print("=" * 60)
    print("# DESIGN_ROW_STYLES — paste into functions.py")
    print("DESIGN_ROW_STYLES = {")
    for row_num in DESIGN_ROWS:
        ff      = fills_fonts.get(row_num, {})
        borders = all_borders.get(row_num, {})
        fills   = ff.get("fills", {})
        fonts   = ff.get("fonts", {})

        non_none = {k: v for k, v in fills.items() if v is not None}
        if not non_none:
            fill_str = "None"
        elif len(non_none) == 1:
            fill_str = f'"{_color_hex(next(iter(non_none.values())))}"'
        else:
            fill_str = "{" + ", ".join(f'"{k}": "{_color_hex(v)}"' for k, v in non_none.items()) + "}"

        bold_vals = [f.get("bold") for f in fonts.values() if isinstance(f, dict)]
        bold_str  = "True" if True in bold_vals else "False"

        if borders:
            items = [
                f'"{e}": {{"style": {i["style"]}, "weight": {i["weight"]}, "color": "{i["color"]}"}}'
                for e, i in sorted(borders.items())
            ]
            border_str = "{" + ", ".join(items) + "}"
        else:
            border_str = "{}"

        print(f'    {row_num}: {{"fill": {fill_str}, "font_bold": {bold_str}, "borders": {border_str}}},')
    print("}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    app = xw.apps.active
    if app is None:
        print("ERROR: Excel is not running.")
        sys.exit(1)

    personal_wb  = None
    design_sheet = None
    for book in app.books:
        try:
            if book.name == "PERSONAL.XLSB":
                personal_wb  = book
                design_sheet = book.sheets["Design"]
                print("Found PERSONAL.XLSB → Design sheet")
                break
        except Exception:
            continue

    if design_sheet is None:
        print("ERROR: Could not find PERSONAL.XLSB or its Design sheet.")
        sys.exit(1)

    # Check for BorderDump sheet
    has_dump = _DUMP_SHEET in [s.name for s in personal_wb.sheets]

    if not has_dump:
        print(f"\nNo '{_DUMP_SHEET}' sheet found in PERSONAL.XLSB.")
        print("Please do the following once:")
        print()
        print("  1. Open Excel VBA Editor  (Option + F11)")
        print("  2. In the Project panel, expand PERSONAL.XLSB → Modules")
        print("  3. Double-click any existing module (or insert a new one)")
        print("  4. Paste the code below at the bottom of the module")
        print("  5. Click inside DumpDesignBorders and press F5 to run")
        print("  6. Re-run:  uv run python read_design_rows.py")
        print()
        print("-" * 60)
        # Write VBA to a file for easy copy-paste
        bas_path = Path(__file__).parent / "DumpDesignBorders.bas"
        bas_path.write_text(_VBA_CODE)
        print(f"VBA code saved to: {bas_path}")
        print("(Or copy from below)\n")
        print(_VBA_CODE)
        sys.exit(0)

    print(f"Found '{_DUMP_SHEET}' sheet — reading border data...")
    all_borders = read_border_dump(personal_wb)
    fills_fonts = read_fills_fonts(design_sheet)

    print(f"\nDesign rows {DESIGN_ROWS}:\n")
    for row_num in DESIGN_ROWS:
        ff = fills_fonts.get(row_num, {})
        print_row_summary(
            row_num,
            ff.get("fills", {}),
            ff.get("fonts", {}),
            all_borders.get(row_num, {}),
        )

    emit_python_dict(fills_fonts, all_borders)


if __name__ == "__main__":
    main()
