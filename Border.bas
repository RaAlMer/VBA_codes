Attribute VB_Name = "Border"
Sub Border()

    Dim rng As Range
    Dim LastCol As Long
    Dim LastRow As Long
    Dim Ws As Worksheet

    With ThisWorkbook.Sheets("Sheet1")
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row 'Last row in column 1
        LastCol = .Cells(2, .Columns.Count).End(xlToLeft).Column 'Last column in row 2
        Set rng = Range("a3", .Cells(LastRow, LastCol))
    End With

    'Table borders
    With rng
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    'Align data in each column
    Range("a3:a" & LastRow).HorizontalAlignment = xlLeft
    Range("c3:c" & LastRow).HorizontalAlignment = xlLeft
    Range("d3:d" & LastRow).HorizontalAlignment = xlCenter
    Range("f3:f" & LastRow).HorizontalAlignment = xlCenter
    Range("g3:g" & LastRow).HorizontalAlignment = xlCenter

    'Changes second column format and align it
    Range("b3:b" & LastRow).Select
    With Selection
        .NumberFormat = "0"
        .Value = .Value
    End With
    Range("b3:b" & LastRow).HorizontalAlignment = xlLeft

    'Unselect every cell selected
    Range("A1").Select
    Application.CutCopyMode = False

    'Changes the Sheet1 name
    Set Ws = Worksheets("Sheet1")
    Ws.Name = Left(Range("A3").Value, 6) & " Sheet"

End Sub
