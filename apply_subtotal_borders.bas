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
