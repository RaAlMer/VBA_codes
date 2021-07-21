Sub CopiaDatosArchivoLicitacion()
 Dim shO As Worksheet, shC As Worksheet, lastRowO As Long, lastRowC As Long
 Dim arrO As Variant, arrC As Variant, lastColO As Long, lastColC As Long
 Dim El As Variant, arrTransf As Variant, strH As String, copyCell As Range
 Dim wbNumb As Variant, wb As Workbook, Ws As Worksheet, strWB As String
 Dim WbC As Workbook, sh As Worksheet, strWh As String, shNunb As Variant

 Dim I As Long, j As Long
 
 On Error GoTo ErrHandler
 
 Set shC = ActiveSheet
WbSelection:
 For I = 1 To Workbooks.Count
    strWB = strWB & Workbooks(I).Name & " - " & I & vbCrLf
 Next I

 wbNumb = InputBox("Please, write the correct number of the workbook to choose the data to copy:" & vbCrLf & _
                vbCrLf & strWB, "Choose the book in which the data is going to be copied by columns!", 1)
    
    If wbNumb <= Workbooks.Count And wbNumb >= 1 Then 'If the number chosen is less than 1 or bigger than the number of worksbooks in this file
        GoTo WbSelection2
    ElseIf wbNumb = "" Then MsgBox "You canceled the operation", vbCritical, "Canceled" 'If you cancel it or leave it empty
        Exit Sub
    ElseIf wbNumb > Workbooks.Count Or wbNumb < 1 Then MsgBox "Wrong number, choose a coorect number", vbExclamation, "Wrong number"
        Exit Sub
    End If
    
WbSelection2:
    If IsNumeric(wbNumb) Then 'If a correct number is chosen
        Set WbC = Workbooks(CLng(wbNumb))
    Else
        MsgBox "Please select the number located at the right of the chosen book!": GoTo WbSelection 'If the input is not numeric
    End If
    
WsSelection:
    For I = 1 To WbC.Worksheets.Count
        strWh = strWh & WbC.Worksheets(I).Name & " - " & I & vbCrLf
    Next
  shNunb = InputBox("Please, write the correct number of the worksheet to choose the data to copy:" & vbCrLf & _
          vbCrLf & strWh, "Choose the sheet in which the data is going to be copied by columns!", 1)
          
     If shNunb <= WbC.Worksheets.Count And shNunb >= 1 Then 'If the number chosen is less than 1 or bigger than the number of worksheets in this file
        GoTo WsSelection2
     ElseIf shNunb = "" Then MsgBox "You canceled the operation", vbCritical, "Canceled" 'If you cancel it or leave it empty
        Exit Sub
     ElseIf shNunb > WbC.Worksheets.Count Or shNunb < 1 Then MsgBox "Wrong number, choose a coorect number", vbExclamation, "Wrong number"
        Exit Sub
     End If

WsSelection2:
    If IsNumeric(shNunb) Then 'If a correct number is chosen
        Set shO = WbC.Worksheets(CLng(shNunb))
    Else
        MsgBox "Please select the number located at the right of the chosen sheet!": GoTo WbSelection 'If the input is not numeric
    End If

 arrC = Split("Id|Number|Part|Seller|Description|Units|Total", "|")
 lastColO = shO.Cells(1, Cells.Columns.Count).End(xlToLeft).Column
 arrO = shO.Range(shO.Cells(1, 1), shO.Cells(1, lastColO)).Value
 'Copy the columns:
 For j = 0 To UBound(arrC)
    If arrC(j) = "Number" Then strH = "number" Else strH = arrC(j)
    If arrC(j) = "Seller" Then strH = "seller" Else strH = arrC(j)
    If arrC(j) = "Units" Then strH = "uni" Else strH = arrC(j)
    For I = 1 To UBound(arrO, 2)
        If arrO(1, I) = strH Then
            lastRowO = shO.Cells(Rows.Count, I).End(xlUp).Row     'last row of the found orig header column
            lastRowC = shC.Cells(Rows.Count, j + 1).End(xlUp).Row 'last row of toCopy sheet header column
            arrTransf = shO.Range(shO.Cells(2, I), shO.Cells(lastRowO, I)).Value
            Set copyCell = shC.Range(shC.Range("A2"), shC.Cells(1, lastColO)).Find(arrC(j))
            If copyCell Is Nothing Then MsgBox "There is no column named """ & _
                                            arrC(j) & """ in the sheet to copy.": Exit Sub
            copyCell.Offset(1, 0).Resize(UBound(arrTransf, 1), UBound(arrTransf, 2)).Value = arrTransf
        End If
    Next I
 Next j
 
 For j = 0 To UBound(arrC)
    If arrC(j) = "Total" Then strH = "Total quantity" Else strH = arrC(j)
    For I = 1 To UBound(arrO, 2)
        If arrO(1, I) = strH Then
            lastRowO = shO.Cells(Rows.Count, I).End(xlUp).Row     'last row of the found orig header column
            lastRowC = shC.Cells(Rows.Count, j + 1).End(xlUp).Row 'last row of toCopy sheet header column
            arrTransf = shO.Range(shO.Cells(2, I), shO.Cells(lastRowO, I)).Value
            Set copyCell = shC.Range(shC.Range("A2"), shC.Cells(1, lastColO)).Find(arrC(j))
            If copyCell Is Nothing Then MsgBox "No hay una columna llamada """ & _
                                            arrC(j) & """ en la hoja para copiar.": Exit Sub
            copyCell.Offset(1, 0).Resize(UBound(arrTransf, 1), UBound(arrTransf, 2)).Value = arrTransf
        End If
    Next I
 Next j
 
 Exit Sub 'Exit before the error
ErrHandler:

    'Error message
    MsgBox "Oh, there is been an error because you choose" & vbCrLf & "the same workbook you are working on.", vbCritical, "Error: Same workbook" '& vbCrLf '_
        '& vbCrLf & "Error Code : " & Err.Number & " , " & Err.Description
End Sub