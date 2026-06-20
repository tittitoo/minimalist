' PERSONAL.XLSB — Module1
' Authoritative source for all VBA macros callable from Python or Excel shortcuts.
' Note: Workbook_Open belongs in the ThisWorkbook module, not here.

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
appTGGL bTGGL:=False

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
appTGGL
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

Sub remove_h_borders()
appTGGL bTGGL:=False
    Dim lastRow As Long
    lastRow = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Dim DataRange As Range
    Set DataRange = Range("A3:H" & lastRow - 2)
    DataRange.Borders(xlInsideHorizontal).LineStyle = xlNone
appTGGL
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
    appTGGL bTGGL:=False
    For Each pgbr In ActiveSheet.HPageBreaks
        With pgbr.Location.Offset(-1, 0).EntireRow.Borders(xlEdgeBottom)
            .LineStyle = xlNone
        End With
    Next
    appTGGL
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
' Keyboard Shortcut: Ctrl+w
    Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
End Sub

Sub delete_row()
' Keyboard Shortcut: Ctrl+q
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
