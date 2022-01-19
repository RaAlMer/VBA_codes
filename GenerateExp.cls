VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub PAAD_ID()

    Dim rID As Variant
    Dim LastRow As Long, LastColumn As Long, LastRow2 As Long, LastCol2 As Long
    Dim FindRowExp As Range, FindRowNumber As Long
    Dim rngCopy As Range, rngPaste As Range
    Dim Rng As Range
    
    On Error GoTo ErrHandler 'Si aparece un error
    
    'Introduce los 4 últimos dígitos del ID PAAD
    rID = InputBox("Introduce los cuatro últimos dígitos del ID PAAD", "ID PAAD XXXX")
        
    If IsNumeric(rID) Then 'Comprueba que es un número
        If Len(rID) = 4 Then
            GoTo BuscID
        Else: MsgBox "Número incorrecto, introduzca número valido", vbCritical, "Número incorrecto"
            Exit Sub
        End If
    Else: MsgBox "Número incorrecto, introduzca número valido", vbCritical, "Número incorrecto"
        Exit Sub
    End If
    
BuscID:
    
    With Sheets("TRAMITADOS SEGEF")
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row 'Encontrará la última fila usada en la columna 1
        LastColumn = .Cells(3, .Columns.Count).End(xlToLeft).Column 'Encontrará la última columna usada en la fila 3
    End With
    
    With Sheets("PAAD")
        Set FindRowExp = .Range("D:D").Find(What:=rID, LookIn:=xlValues) 'Busca la fila del ID que necesitamos
        FindRowNumber = FindRowExp.Row 'Busca el número de la fila
        .Cells(FindRowNumber, 4).Copy
        Sheets("TRAMITADOS SEGEF").Cells(LastRow + 1, 1).PasteSpecial xlPasteValues 'Copia el ID PAAD del Expediente elegido
        .Cells(FindRowNumber, 6).Copy
        Sheets("TRAMITADOS SEGEF").Cells(LastRow + 1, 4).PasteSpecial xlPasteValues 'Copia el nº del Expediente elegido
        .Cells(FindRowNumber, 5).Copy
        Sheets("TRAMITADOS SEGEF").Cells(LastRow + 1, 5).PasteSpecial xlPasteValues 'Copia el VI del Expediente elegido
        .Cells(FindRowNumber, 9).Copy
        Sheets("TRAMITADOS SEGEF").Cells(LastRow + 1, 6).PasteSpecial xlPasteValues 'Copia el Título del Expediente elegido
        .Cells(FindRowNumber, 10).Copy
        Sheets("TRAMITADOS SEGEF").Cells(LastRow + 1, 7).PasteSpecial xlPasteValues 'Copia el Proc Contrat del Expediente elegido
'        .Cells(FindRowNumber, 22).Copy
'        Sheets("TRAMITADOS SEGEF").Cells(LastRow + 1, 9).PasteSpecial xlPasteValues 'Copia la anualidad 2022 del Expediente elegido
'        .Cells(FindRowNumber, 23).Copy
'        Sheets("TRAMITADOS SEGEF").Cells(LastRow + 1, 10).PasteSpecial xlPasteValues 'Copia la anualidad 2023 del Expediente elegido
'        .Cells(FindRowNumber, 24).Copy
'        Sheets("TRAMITADOS SEGEF").Cells(LastRow + 1, 11).PasteSpecial xlPasteValues 'Copia la anualidad 2024 del Expediente elegido
        .Cells(FindRowNumber, 11).Copy
        Sheets("TRAMITADOS SEGEF").Cells(LastRow + 1, 12).PasteSpecial xlPasteValues 'Copia el ADJUDICATARIO del Expediente elegido
    End With
    
    'Poner formato a la fila
    With Sheets("TRAMITADOS SEGEF")
        LastRow2 = .Cells(.Rows.Count, 1).End(xlUp).Row 'Encontrará la última fila usada en la columna 1
        LastCol2 = .Cells(3, .Columns.Count).End(xlToLeft).Column 'Encontrará la última columna usada en la fila 3
        Set Rng = Range("a4", .Cells(LastRow2, LastCol2))
        .Rows(LastRow2).Font.Size = 10
        .Rows(LastRow2).Font.Name = "Arial"
        .Rows(LastRow2).WrapText = True 'Ajusta el texto a la celda
    End With
    
    'Pone los bordes de la tabla
    With Rng
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
'    'Pone formato moneda a las columnas de anualidades
'    Range("i4:i" & LastRow2).NumberFormat = "#,##0.00 €_)"
'    Range("j4:j" & LastRow2).NumberFormat = "#,##0.00 €_)"
'    Range("k4:k" & LastRow2).NumberFormat = "#,##0.00 €_)"
    
    'Alinea los datos de cada columna de la tabla
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
    With Sheets("TRAMITADOS SEGEF")
        .Rows(LastRow2).VerticalAlignment = xlVAlignCenter
    End With
    
    'Ajusta la fila al texto
    Rows(LastRow + 1).AutoFit
    
    'Actualiza la fecha del expediente (poniendo la casilla en amarillo = PROVISIONAL) y la de cambio
    Cells(LastRow + 1, 2).Value = Date
    Cells(LastRow + 1, 2).Interior.ColorIndex = 6
    Cells(1, 3).Value = Date

    Exit Sub 'Para que salga cuando compila todo sin mostrar el error de aquí debajo
    
ErrHandler:

    'Mensaje de error si eliges como libro para obtener datos este mismo Excel
    MsgBox "Ups, ha habido un error." & vbCrLf & vbCrLf & "Comprueba que has introducido bien los números del ID PAAD o que existen.", vbCritical, "Error desconocido"

End Sub

