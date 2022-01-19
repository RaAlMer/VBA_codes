Option Explicit

Public Sub GET_ID()

    Dim rID As Variant
    Dim LastRow As Long, LastColumn As Long, LastRow2 As Long, LastCol2 As Long
    Dim FindRowExp As Range, FindRowNumber As Long
    Dim rngCopy As Range, rngPaste As Range
    Dim Rng As Range
    
    On Error GoTo ErrHandler 'Si aparece un error
    
  'Introduce the last 4 digits of the ID
  rID = InputBox("Introduce the last 4 digits of the ID", "ID XXXX")
        
    If IsNumeric(rID) Then 'Comprueba que es un número
        If Len(rID) = 4 Then
            GoTo SearchID
    Else: MsgBox "Wrong number, introduce valid number", vbCritical, "Wrong number"
            Exit Sub
        End If
    Else: MsgBox "Wrong number, introduce valid number", vbCritical, "Wrong number"
        Exit Sub
    End If
    
SearchID:
    
      With Sheets("LISTED")
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row 'Find last row used in column 1
        LastColumn = .Cells(3, .Columns.Count).End(xlToLeft).Column 'Find last column used in row 3
    End With
    
With Sheets("ID")
Set FindRowExp = .Range("D:D").Find(What:=rID, LookIn:=xlValues) 'Find row from the ID needed
  FindRowNumber = FindRowExp.Row 'Find row number
        .Cells(FindRowNumber, 4).Copy
        Sheets("LISTED").Cells(LastRow + 1, 1).PasteSpecial xlPasteValues 'Copy the ID from the Exp chosen
        .Cells(FindRowNumber, 6).Copy
        Sheets("LISTED").Cells(LastRow + 1, 4).PasteSpecial xlPasteValues 'Copy the number from the Exp chosen
        .Cells(FindRowNumber, 5).Copy
        Sheets("LISTED").Cells(LastRow + 1, 5).PasteSpecial xlPasteValues 'Copy the VI from the Exp chosen
        .Cells(FindRowNumber, 9).Copy
        Sheets("LISTED").Cells(LastRow + 1, 6).PasteSpecial xlPasteValues 'Copy the title from the Exp chosen
        .Cells(FindRowNumber, 10).Copy
        Sheets("LISTED").Cells(LastRow + 1, 7).PasteSpecial xlPasteValues 'Copy the contract from the Exp chosen
        .Cells(FindRowNumber, 11).Copy
        Sheets("LISTED").Cells(LastRow + 1, 12).PasteSpecial xlPasteValues 'Copy the company from the Exp chosen
    End With
    
    'Format row
    With Sheets("LISTED")
        LastRow2 = .Cells(.Rows.Count, 1).End(xlUp).Row 'Find last row in column 1
        LastCol2 = .Cells(3, .Columns.Count).End(xlToLeft).Column 'Find last column in row 3
        Set Rng = Range("a4", .Cells(LastRow2, LastCol2))
        .Rows(LastRow2).Font.Size = 10
        .Rows(LastRow2).Font.Name = "Arial"
        .Rows(LastRow2).WrapText = True 'Adjust text to row
    End With
    
    'Borders from table
    With Rng
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    'Align data in each column
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
    With Sheets("LISTED")
        .Rows(LastRow2).VerticalAlignment = xlVAlignCenter
    End With
    
    'Adjust row to text
    Rows(LastRow + 1).AutoFit
    
    'Updates date from the exp (yellow cell as PROVISIONAL)
    Cells(LastRow + 1, 2).Value = Date
    Cells(LastRow + 1, 2).Interior.ColorIndex = 6
    Cells(1, 3).Value = Date

    Exit Sub 'To exit before the error
    
ErrHandler:

    'Error message if the same Excel book is chosen
     MsgBox "Doh, there's been an error." & vbCrLf & vbCrLf & "Check every data is introduced correctly or if it exists.", vbCritical, "Unknown error"

End Sub

