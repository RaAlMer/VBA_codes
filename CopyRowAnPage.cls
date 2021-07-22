VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub ID()

    Dim rID As Variant
    Dim LastRow As Long, LastColumn As Long, LastRow2 As Long, LastCol2 As Long
    Dim FindRowExp As Range, FindRowNumber As Long
    Dim rngCopy As Range, rngPaste As Range
    Dim Rng As Range
    
    On Error GoTo ErrHandler 
    
    'Last 4 digits from the ID
    rID = InputBox("Last four digits from the ID", "ID XXXX")
        
    If IsNumeric(rID) Then 'Check it's a number
        If Len(rID) = 4 Then
            GoTo BuscID
        Else: MsgBox "Wrong number, input valid number", vbCritical, "Wrong number"
            Exit Sub
        End If
    Else: MsgBox "Wrong number, input valid number", vbCritical, "Wrong number"
        Exit Sub
    End If
    
BuscID:
    
    With Sheets("SEARCH ID")
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row 'Last row in column 1
        LastColumn = .Cells(3, .Columns.Count).End(xlToLeft).Column 'Last column in row 3
    End With
    
    With Sheets("IDs")
        Set FindRowExp = .Range("D:D").Find(What:=rID, LookIn:=xlValues) 'Look for the ID row we need
        FindRowNumber = FindRowExp.Row 'Look for the row number
        .Cells(FindRowNumber, 4).Copy
        Sheets("SEARCH ID").Cells(LastRow + 1, 1).PasteSpecial xlPasteValues 'Copy the ID
        .Cells(FindRowNumber, 6).Copy
        Sheets("SEARCH ID").Cells(LastRow + 1, 4).PasteSpecial xlPasteValues 'Copy the number
        .Cells(FindRowNumber, 5).Copy
        Sheets("SEARCH ID").Cells(LastRow + 1, 5).PasteSpecial xlPasteValues 'Copy the name
        .Cells(FindRowNumber, 9).Copy
        Sheets("SEARCH ID").Cells(LastRow + 1, 6).PasteSpecial xlPasteValues 'Copy the title
        .Cells(FindRowNumber, 10).Copy
        Sheets("SEARCH ID").Cells(LastRow + 1, 7).PasteSpecial xlPasteValues 'Copy the process
        .Cells(FindRowNumber, 22).Copy
        Sheets("SEARCH ID").Cells(LastRow + 1, 9).PasteSpecial xlPasteValues 'Copy the first year
        .Cells(FindRowNumber, 23).Copy
        Sheets("SEARCH ID").Cells(LastRow + 1, 10).PasteSpecial xlPasteValues 'Copy the second year
        .Cells(FindRowNumber, 24).Copy
        Sheets("SEARCH ID").Cells(LastRow + 1, 11).PasteSpecial xlPasteValues 'Copy the third year
        .Cells(FindRowNumber, 11).Copy
        Sheets("SEARCH ID").Cells(LastRow + 1, 12).PasteSpecial xlPasteValues 'Copy the enterprise
    End With
    
    'Format the row
    With Sheets("SEARCH ID")
        LastRow2 = .Cells(.Rows.Count, 1).End(xlUp).Row
        LastCol2 = .Cells(3, .Columns.Count).End(xlToLeft).Column
        Set Rng = Range("a4", .Cells(LastRow2, LastCol2))
        .Rows(LastRow2).Font.Size = 10
        .Rows(LastRow2).Font.Name = "Arial"
        .Rows(LastRow2).WrapText = True
    End With
    
    'Table borders
    With Rng
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    'Format money
    Range("i4:i" & LastRow2).NumberFormat = "#,##0.00 €_)"
    Range("j4:j" & LastRow2).NumberFormat = "#,##0.00 €_)"
    Range("k4:k" & LastRow2).NumberFormat = "#,##0.00 €_)"
    
    'Align data in columns
    Range("a4:a" & LastRow2).HorizontalAlignment = xlCenter
    Range("b4:b" & LastRow2).HorizontalAlignment = xlCenter
    Range("c4:c" & LastRow2).HorizontalAlignment = xlCenter
    Range("d4:d" & LastRow2).HorizontalAlignment = xlCenter
    Range("e4:e" & LastRow2).HorizontalAlignment = xlCenter
    Range("f4:f" & LastRow2).HorizontalAlignment = xlLeft
    Range("g4:g" & LastRow2).HorizontalAlignment = xlCenter
    Range("h4:h" & LastRow2).HorizontalAlignment = xlLeft
    Range("i4:i" & LastRow2).HorizontalAlignment = xlLeft
    Range("j4:j" & LastRow2).HorizontalAlignment = xlLeft
    Range("k4:k" & LastRow2).HorizontalAlignment = xlLeft
    Range("l4:l" & LastRow2).HorizontalAlignment = xlLeft
    Range("m4:m" & LastRow2).HorizontalAlignment = xlCenter
    Range("n4:n" & LastRow2).HorizontalAlignment = xlCenter
    With Sheets("SEARCH ID")
        .Rows(LastRow2).VerticalAlignment = xlVAlignCenter
    End With
    
    'Adjust text to row
    Rows(LastRow + 1).AutoFit
    
    'Put today's date (cell in yellow = PROVISIONAL)
    Cells(LastRow + 1, 2).Value = Date
    Cells(LastRow + 1, 2).Interior.ColorIndex = 6
    Cells(1, 3).Value = Date

    Exit Sub
    
ErrHandler:

    'Error message
    MsgBox "Oh, an error ocurred." & vbCrLf & vbCrLf & "Check the ID exits or it's correct.", vbCritical, "Unknown error"

End Sub

