VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton2_Click()

    Dim xOLE As Object
    
On Error Resume Next

'Guarda el excel xlsm en xlsx, es decir sin macros
With Application.FileDialog(msoFileDialogSaveAs)
    .Title = "Por favor elige donde guardar el archivo y un nombre" 'Elige el título para nuestro excel
    .ButtonName = "Guardar Excel" 'Nombre del botón
    .InitialFileName = "C:\Users\admin\Desktop\" & Left(Range("A3").Value, 6) & " ARCHIVO LICITACIÓN" 'Da un nombre y destino inicial donde guardar el archivo
    If .Show = 0 Then
        MsgBox "El archivo no se ha guardado.", vbCritical 'Muestra mensaje de error si el archivo no se guarda
        Exit Sub
    End If
    Application.DisplayAlerts = False
    .Execute
    Application.DisplayAlerts = True
End With

'Borra todos los botones que había en la macro del excel
    ActiveSheet.Buttons.Delete
    For Each xOLE In ActiveSheet.OLEObjects
        If TypeName(xOLE.Object) = "CommandButton" Then
            xOLE.Delete
        End If
    Next

'Guarda automaticamente el nuevo xlsx para que no salga la consola prompt que avisa de que no se guardan las macros
    If ThisWorkbook.Saved = False Then
    ThisWorkbook.Save
    End If

End Sub
