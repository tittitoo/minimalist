' Public function to disable updating to make things faster

Public Sub appTGGL(Optional bTGGL As Boolean = True)
    With AppliderlineStyleSingle
        '.Strikethrough = False
        '.TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=G1=""Title"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        '.Strikethrough = False
        '.TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=G1=""Subsystem"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -7137279
    End With
    With Selection.FormatConditions(1).Interior
        .Pattern = xlNone
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=G1=""System"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -7137279
        '.Strikethrough = False
        '.TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .Pattern = xlNone
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    Range("A1").Select
    appTGGL
End Sub
Sub conditional_format_internal_costing()
'
' conditional_format_internal_costing
'
'
    appTGGL bTGGL:=False
    Cells.FormatConditions.Delete
    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=R1=""Deleted"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Strikethrough = True
        .TintAndShade = 0
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
        '.Strikethrough = False
        .Color = -52732
        '.TintAndShade = 0
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
        '.Strikethrough = False
        '.TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=R1=""Title"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold =       .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    Columns("C:         .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = -52732
            End With
        End With

    Next          .ThemeColor = xlThemeColorDark1
        .TintAnContinuous
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
'
' Unshaded
'

'
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
Sub format_colums(xlInsideVertical).LineStyle = xlNone
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
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
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

Sub put_systems_validation_formula()
'
' Put formula for system names validation
'

'
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
'
' Put currency and proposal validation formula
'

'
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
'
' Put formula for system names validation
'

'
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
' Make workbook open hidden
  Private Sub Workbook_Open()
      Windows(ThisWorkbook.Name).Visible = False
  End Sub





n_border()
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
    Selection.BorderdShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
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
        .LineStyle = xl
ActiveWindow.View = xlNormalView
    
End Sub
Sub remove_pagebreak_borders()
'Clear current pagebreaks
appTGGL bTGGL:=False
  For Each pgbr In ActiveSheet.HPageBreaks
    
        With pgbr.Location.Offset(-1, 0).EntireRow.Borders(xlEdgeBottom)
            .LineStyle = xlNone
        End With

    Next
appTGGL
End Sub
Sub remove_h_borders()
appTGGL bTGGL:=False

Dim lastRow As Long

lastRow = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

Dim DataRange As Range
Set DataRange = Range("A3:H" & lastRow - 2)

DataRange.Borders(xlInsideHorizontal).LineStyle = xlNone
' Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
appTGGL
End Sub
Sub shaded()
'
' Shade the region
'

'
    Dim activeRange As Range
    Set activeRange = Selection
    Columns("I:BD").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
     C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=R1=""System"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -7137279
        '.Strikethrough = False
        '.TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .Pattern = xlNone
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    Range("A1").Select
    appTGGL
End Sub
Sub pagebreak_borders()
' Set horizontal borders in pagebreaks
' Setting to PageBreakPreview ensures the border is drawn
ActiveWindow.View = xlPageBreakPreview
  For Each pgbr In ActiveSheet.HPageBreaks
        With pgbr.Location
            With Range(.Offset(-1, 0), .Offset(-1, 7)).Borders(xlEdgeBottom)
    True
        .Italic = False
        '.Strikethrough = False
        '.TintAndShade = 0
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
    With Selection.FormatConditions(1).Interior
        .Pattern = xlNone
 UnderlineStyleSingle
        '.Strikethrough = False
        '.TintAndShade = 0
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
        '.Strikethrough = False
        '.TintAndShade = 0
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
    With Selection.FormatConditions(1).Interior
        .Pattern = xlNone
        .TintAndShade = 0
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
        '.Strikethrough = False
        '.TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .Pattern = xlNone
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    activeRange.Select
    appTGGL
End Sub
Sub conditional_format_technical()
'
' Conditional_Format_Technical Macro
'
'
appTGGL bTGGL:=False
    Cells.FormatConditions.Delete
    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=G1=""Deleted"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Strikethrough = True
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=G1=""Comment"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = False
        .Italic = True
        '.Strikethrough = False
        .Color = -52732
        '.TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    Columns("C:C").Select
    Range("C:C").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=G1=""Subtitle"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = False
        .Italic = True
        .Underline = xlUncation
        .ScreenUpdating = bTGGL
        .EnableEvents = bTGGL
        .DisplayAlerts = bTGGL
        .AutoRecover.Enabled = bTGGL   'no interruptions with an auto-save
        .Calculation = IIf(bTGGL, xlCalculationAutomatic, xlCalculationManual)
        .CutCopyMode = False
        .StatusBar = vbNullString
    End With
    Debug.Print Timer
End Sub


Sub add_row()
'
' add_row Macro
' Add row
'
' Keyboard Shortcut: Ctrl+w

    Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
End Sub
Sub delete_row()
'
' delete_row Macro
' Delete Row
'
' Keyboard Shortcut: Ctrl+q
'
    appTGGL bTGGL:=False
    Selection.EntireRow.Delete
    appTGGL
End Sub
Sub conditional_format()
'
' Conditional_Format Macro
'
'
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
