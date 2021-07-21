VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SaveSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub svExcel_Click()

    Dim xOLE As Object
    
On Error Resume Next

'Saves the Excel xlsm to xlsx, without macros
With Application.FileDialog(msoFileDialogSaveAs)
    .Title = "Please choose where to save the file and a name" 'Excel title
    .ButtonName = "Save Excel"
    .InitialFileName = "C:\Users\admin\Desktop\" & Left(Range("A3").Value, 6) & " SAVED EXCEL" 'Initial name for the path
    If .Show = 0 Then
        MsgBox "File not saved.", vbCritical 'Error message
        Exit Sub
    End If
    Application.DisplayAlerts = False
    .Execute
    Application.DisplayAlerts = True
End With

'Delete every button from the Excel file
ActiveSheet.Buttons.Delete
For Each xOLE In ActiveSheet.OLEObjects
    If TypeName(xOLE.Object) = "CommandButton" Then
        xOLE.Delete
    End If
Next

'Saves automatically the new xlsx so the alert dialog box doesn't show
    If ThisWorkbook.Saved = False Then
    ThisWorkbook.Save
    End If

End Sub
