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
                    wsO.Cells(outRow, 7).Value = (bgr \ 256) And 255
                    wsO.Cells(outRow, 8).Value = (bgr \ 65536) And 255
                    outRow = outRow + 1
                End If
            Next i
        Next c
    Next r
    MsgBox "Done — " & (outRow - 2) & " borders written to 'BorderDump' sheet."
End Sub
