Sub hide_proposal_columns()
    ' Called from Python hide_columns() on ActiveSheet.
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
