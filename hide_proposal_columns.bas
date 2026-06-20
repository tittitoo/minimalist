Sub hide_proposal_columns()
    ' Called from Python hide_columns() on ActiveSheet.
    ' Replaces ~24 individual xlwings/COM calls with a single VBA execution.
    ' Currency columns fixed at 15 — accommodates up to 1,000,000.00 with cents.
    With ActiveSheet
        ' Hidden columns (width 0)
        .Columns("O").ColumnWidth = 0
        .Columns("Q").ColumnWidth = 0
        .Columns("S:AA").ColumnWidth = 0
        .Columns("AC:AD").ColumnWidth = 0
        .Columns("AF").ColumnWidth = 0
        .Columns("AI:AL").ColumnWidth = 0

        ' Fixed-width columns
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 6
        .Columns("C").ColumnWidth = 55
        .Columns("C").WrapText = True
        .Columns("I").ColumnWidth = 10
        .Columns("I").WrapText = False
        .Columns("P").ColumnWidth = 20
        .Columns("P").WrapText = False

        ' Currency/value columns
        .Columns("D:H").ColumnWidth = 15
        .Columns("J:K").ColumnWidth = 15
        .Columns("L").ColumnWidth = 15
        .Columns("M:N").ColumnWidth = 15
        .Columns("R").ColumnWidth = 15
        .Columns("T").ColumnWidth = 15
        .Columns("AB").ColumnWidth = 15
        .Columns("AE").ColumnWidth = 15
        .Columns("AG:AH").ColumnWidth = 15
        .Columns("AM:AP").ColumnWidth = 15
    End With
End Sub
