' PERSONAL.XLSB — Module1
' Authoritative source for all VBA macros callable from Python or Excel shortcuts.
' Note: Workbook_Open belongs in the ThisWorkbook module, not here.

' ============================================================
' Keyboard shortcuts
' Call SetupShortcuts from Workbook_Open in ThisWorkbook so shortcuts
' survive future Module1 replacements (procedure attributes are fragile).
' ============================================================
Sub SetupShortcuts()
    ' Ctrl+E/J/M work via OnKey on Mac (no OS conflict).
    Application.OnKey "^e", "formula"
    Application.OnKey "^j", "hide_columns"
    Application.OnKey "^m", "unhide_columns"
    ' add_row / delete_row shortcuts TBD — Ctrl+W/Q/I/D all intercepted by Mac Excel
End Sub

' ============================================================
' Performance helper
' Call appTGGL False before bulk ops, appTGGL (True) after.
' ============================================================
Public Sub appTGGL(Optional bTGGL As Boolean = True)
    With Application
        .ScreenUpdating = bTGGL
        .EnableEvents = bTGGL
        .DisplayAlerts = bTGGL
        .AutoRecover.Enabled = bTGGL
        .Calculation = IIf(bTGGL, xlCalculationAutomatic, xlCalculationManual)
        .CutCopyMode = False
        .StatusBar = vbNullString
    End With
    Debug.Print Timer
End Sub

' ============================================================
' Conditional Formatting — proposal view (AL column)
' Primary path is Python apply_conditional_format; this is the fallback.
' ============================================================
Sub conditional_format()
    Dim wasUpdating As Boolean
    wasUpdating = Application.ScreenUpdating
    If wasUpdating Then appTGGL bTGGL:=False

    Dim activeRange As Range
    Set activeRange = Selection
    Cells.FormatConditions.Delete

    ' --- Column C: row-type styles ---
    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=AL1=""Deleted"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Strikethrough = True
    End With
    Selection.FormatConditions(1).StopIfTrue = True

    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=AL1=""Comment"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = False
        .Italic = True
        .Color = -52732
    End With
    Selection.FormatConditions(1).StopIfTrue = True

    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=AL1=""Subtitle"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = False
        .Italic = True
        .Underline = xlUnderlineStyleSingle
    End With
    Selection.FormatConditions(1).StopIfTrue = True

    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=AL1=""Title"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
    End With
    Selection.FormatConditions(1).StopIfTrue = True

    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=AL1=""Subsystem"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -7137279
    End With
    Selection.FormatConditions(1).StopIfTrue = True

    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=AL1=""System"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -7137279
    End With
    Selection.FormatConditions(1).StopIfTrue = True

    ' --- Columns D:G: bold when Title row has a number in D ---
    Columns("D:G").Select
    Range("D:G").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=AND($AL1=""Title"",ISNUMBER($D1))"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    activeRange.Select
    If wasUpdating Then appTGGL
End Sub

' ============================================================
' Conditional Formatting — internal costing view (R column)
' ============================================================
Sub conditional_format_internal_costing()
appTGGL bTGGL:=False
    Cells.FormatConditions.Delete

    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=R1=""Deleted"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Strikethrough = True
    End With
    Selection.FormatConditions(1).StopIfTrue = True

    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=R1=""Comment"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = False
        .Italic = True
        .Color = -52732
    End With
    Selection.FormatConditions(1).StopIfTrue = True

    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=R1=""Subtitle"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = False
        .Italic = True
        .Underline = xlUnderlineStyleSingle
    End With
    Selection.FormatConditions(1).StopIfTrue = True

    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=R1=""Title"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
    End With
    Selection.FormatConditions(1).StopIfTrue = True

    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=R1=""Subsystem"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -7137279
    End With
    Selection.FormatConditions(1).StopIfTrue = True

    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=R1=""System"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -7137279
    End With
    Selection.FormatConditions(1).StopIfTrue = True

appTGGL
End Sub

' ============================================================
' Column and row borders
' ============================================================
Sub format_column_border()
appTGGL bTGGL:=False
    Dim activeRange As Range
    Set activeRange = Selection

    Columns("A:A").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -52732
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0.599993896298105
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone

    Columns("B:B").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0.599993896298105
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0.599993896298105
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone

    Columns("C:C").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0.599993896298105
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0.599993896298105
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone

    Columns("D:D").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0.599993896298105
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0.599993896298105
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone

    Columns("E:E").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0.599993896298105
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0.599993896298105
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone

    Columns("F:F").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0.599993896298105
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0.599993896298105
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone

    Columns("G:G").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0.599993896298105
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0.599993896298105
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone

    Columns("H:H").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 4
        .TintAndShade = 0.599993896298105
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -52732
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone

    Columns("I:BD").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With

    Rows("1:1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    Rows("2:2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -52732
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -52732
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -52732
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    activeRange.Select
appTGGL
End Sub

Sub format_col_a_left_border()
    ' Left teal border on column A from row 2 down — shared by technical and commercial proposals.
    With ActiveSheet.Range("A2:A1048576").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -52732
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Sub format_col_f_right_border()
    ' Right teal border on column F from row 2 down — last data column in a technical proposal.
    With ActiveSheet.Range("F2:F1048576").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -52732
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Sub format_col_h_right_border()
    ' Right teal border on column H from row 2 down — last data column in a commercial proposal.
    With ActiveSheet.Range("H2:H1048576").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -52732
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Sub remove_h_borders()
    Dim wasUpdating As Boolean
    wasUpdating = Application.ScreenUpdating
    If wasUpdating Then appTGGL bTGGL:=False
    Dim lastRow As Long
    lastRow = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Dim DataRange As Range
    Set DataRange = Range("A3:H" & lastRow - 2)
    DataRange.Borders(xlInsideHorizontal).LineStyle = xlNone
    If wasUpdating Then appTGGL
End Sub

Sub pagebreak_borders()
' PageBreakPreview is required to enumerate breaks correctly.
    ActiveWindow.View = xlPageBreakPreview
    For Each pgbr In ActiveSheet.HPageBreaks
        With pgbr.Location
            With Range(.Offset(-1, 0), .Offset(-1, 7)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = -52732
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
    Next
    ActiveWindow.View = xlNormalView
End Sub

Sub remove_pagebreak_borders()
    Dim wasUpdating As Boolean
    wasUpdating = Application.ScreenUpdating
    If wasUpdating Then appTGGL bTGGL:=False
    For Each pgbr In ActiveSheet.HPageBreaks
        With pgbr.Location.Offset(-1, 0).EntireRow.Borders(xlEdgeBottom)
            .LineStyle = xlNone
        End With
    Next
    If wasUpdating Then appTGGL
End Sub

' ============================================================
' Shading — internal costing view
' ============================================================
Sub shaded()
    Dim activeRange As Range
    Set activeRange = Selection
    Columns("I:BD").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    Rows("2:2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -52732
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -52732
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -52732
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    activeRange.Select
End Sub

Sub unshaded()
    Dim activeRange As Range
    Set activeRange = Selection
    Columns("I:BD").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Rows("2:2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -52732
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -52732
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -52732
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    activeRange.Select
End Sub

' ============================================================
' Row operations — keyboard shortcuts
' ============================================================
Sub add_row()
' Shortcut: Ctrl+W (Windows) / Ctrl+G (Mac)
    Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
End Sub

Sub delete_row()
' Shortcut: Ctrl+Q (Windows) / Ctrl+L (Mac)
    appTGGL bTGGL:=False
    Selection.EntireRow.Delete
    appTGGL
End Sub

' ============================================================
' Validation helpers
' ============================================================
Sub put_systems_validation_formula()
    ActiveSheet.Range("F5:F100").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Config!$A$96:$A$200"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    ActiveSheet.Range("F5").Select
End Sub

Sub put_currency_proposal_validation_formula()
    ActiveSheet.Range("B12").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Config!$B$95:$B$96"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    ActiveSheet.Range("B13").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Config!$C$95:$C$96"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    ActiveSheet.Range("F4").Select
End Sub

Sub put_checklists_validation_formula()
    ActiveSheet.Range("G5").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Config!$E$96:$E$200"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    ActiveSheet.Range("G5").Select
End Sub

' ============================================================
' Subtotal borders — called from Python on Mac
' ============================================================
Sub apply_subtotal_borders(row As Long)
    ' Apply thin blue (#0432FF) top and bottom borders to the full subtotal row.
    ' Called from Python (apply_lastrow_border) on Mac — faster than clipboard copy.
    ' Applies to the entire row, consistent with the Windows COM implementation.
    With ActiveSheet.Rows(row).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(4, 50, 255)
    End With
    With ActiveSheet.Rows(row).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(4, 50, 255)
    End With
End Sub

' ============================================================
' Column layout — called from Python hide_columns()
' ============================================================
Sub commercial_prepare_sheet(sheetName As String)
    ' Called from Python commercial() for each data sheet.
    ' Runs entirely inside Excel — no large Apple Event data transfers.
    ' DO NOT call appTGGL — Python already set ScreenUpdating=False.
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)

    Dim lastRow As Long
    lastRow = ws.Cells(1500, 7).End(xlUp).Row  ' Column G = 7

    If lastRow < 3 Then Exit Sub  ' No data rows to process

    ' Freeze A:H values (formula -> static value, inside Excel process)
    ws.Range("A3:H" & lastRow).Value = ws.Range("A3:H" & lastRow).Value

    ' Re-add G price formula and subtotal
    ws.Range("G3:G" & lastRow - 1).Formula = _
        "=IF(AND(F3<>"""", H3<>""OPTION"", H3<>""INCLUDED"", H3<>""WAIVED""), D3*F3,"""")"
    ws.Range("G" & lastRow).Formula = "=SUM(G3:G" & lastRow - 1 & ")"

    ' Delete pricing columns beyond H (after deletion AL shifts to column I)
    ws.Columns("AM:BD").Delete
    ws.Columns("I:AK").Delete

    ' Save row-type labels now at column I (was AL)
    Dim alValues As Variant
    alValues = ws.Range("I1:I" & lastRow).Value

    ' Remove from I, write to AL, hide
    ws.Columns("I").Delete
    ws.Range("AL1:AL" & lastRow).Value = alValues
    ws.Columns("AL").ColumnWidth = 0
End Sub

Sub hide_proposal_columns()
    ' Replaces ~24 individual xlwings/appscript calls with a single VBA execution.
    With ActiveSheet
        ' Hidden columns (width 0)
        .Columns("AI:AL").ColumnWidth = 0
        .Columns("AC:AD").ColumnWidth = 0
        .Columns("AF").ColumnWidth = 0
        .Columns("S:AA").ColumnWidth = 0
        .Columns("Q").ColumnWidth = 0
        .Columns("O").ColumnWidth = 0

        ' Fixed-width columns
        .Columns("A").ColumnWidth = 5
        .Columns("C").ColumnWidth = 55
        .Columns("C").WrapText = True
        .Columns("I").ColumnWidth = 10
        .Columns("I").WrapText = False
        .Columns("P").ColumnWidth = 20
        .Columns("P").WrapText = False

        ' AutoFit columns
        .Columns("B").AutoFit
        .Columns("D:H").AutoFit
        .Columns("J:K").AutoFit
        .Columns("L").AutoFit
        .Columns("M:N").AutoFit
        .Columns("R").AutoFit
        .Columns("T").AutoFit
        .Columns("AB").AutoFit
        .Columns("AE").AutoFit
        .Columns("AG:AH").AutoFit
        .Columns("AM:AP").AutoFit
    End With
End Sub
