Attribute VB_Name = "Módulo2"
Sub Bordes()

Dim rng As Range
Dim LastCol As Long
Dim LastRow As Long
Dim Ws As Worksheet

With ThisWorkbook.Sheets("ANEXO 1 ") 'Cambiar ANEXO 1 para trabajar con otra hoja
    LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row 'Encontrará la última fila usada en la columna 1
    LastCol = .Cells(2, .Columns.Count).End(xlToLeft).Column 'Encontrará la última columna usada en la fila 2
    Set rng = Range("a3", .Cells(LastRow, LastCol))
End With

'Pone los bordes de la tabla
With rng
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
End With

'Alinea los datos de cada columna de la tabla
Range("a3:a" & LastRow).HorizontalAlignment = xlLeft
Range("c3:c" & LastRow).HorizontalAlignment = xlLeft
Range("d3:d" & LastRow).HorizontalAlignment = xlCenter
Range("f3:f" & LastRow).HorizontalAlignment = xlCenter
Range("g3:g" & LastRow).HorizontalAlignment = xlCenter

'Cambia el formato de la segunda columna para que no esté en notación científica y lo alinea
Range("b3:b" & LastRow).Select
With Selection
    .NumberFormat = "0"
    .Value = .Value
End With
Range("b3:b" & LastRow).HorizontalAlignment = xlLeft

'Deselecciona cualquier celda que haya seleccionada
Range("A1").Select
Application.CutCopyMode = False

'Cambia el nombre de la hoja al adecuado
  Set Ws = Worksheets("ANEXO 1 ")

  Ws.Name = Left(Range("A3").Value, 6) & " ANEXO 1"

End Sub
