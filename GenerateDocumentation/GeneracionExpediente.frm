VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GeneracionExpediente 
   Caption         =   "GENERACIÓN DE EXPEDIENTE"
   ClientHeight    =   7428
   ClientLeft      =   36
   ClientTop       =   372
   ClientWidth     =   8856
   OleObjectBlob   =   "GeneracionExpediente.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "GeneracionExpediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variables de la macro
Public strPath As String, strPathSub As String 'Path de los archivos plantilla y de las subcarpetas
Public rEXP As Variant 'Número de expediente
Public wrdApp As Object, wrdDoc As Object
Public FindRowExp As Range, FindRowNumber As Long 'Busqueda de línea de Expediente
Public FindRowSAG As Range, FindRowNumberSAG As Long 'Busqueda de línea de SAG
Public FindRowAM As Range 'Busqueda de línea de Lote AM
Public FindRowREC As Range, FindRowNumberREC As Long 'Busqueda de línea de REC
Public FindRowOMP As Range, FindRowNumberOMP As Long 'Busqueda de línea de OMP (Informe 668)
Public FindRowRepu As Range 'Busqueda de repuestos (Informe y más documentos)
Public MyRange As Word.Range 'Rango para busqueda de marcadores (bookmarks)
Public myRange2 As Word.Range 'Rango para busqueda de marcadores (bookmarks) (para PDC)
Public myRangeOMP As Word.Range 'Rango para busqueda de marcadores (bookmarks) (para Informe 668)
Public dashpos As Long 'Posición del guión en los títulos de Expediente
Public myTable As Word.Table 'Tablas del documento Word
Public NumAnual As Integer 'Número de anualidades
Public AppPres As String 'Aplicación Presupuestaria
Public Coronel As String 'Coronel o Comandante Acctal.
Public CoronelJef As String 'Coronel o Coronel Acctal. (PELOCHE)
Public objCC As ContentControl 'Casillas (CheckBox)
Public AcuerdoM As String, LoteAM As String, EmpresaAM As String 'Acuerdo Marco, su lote y su empresa
Public Titulo As String 'Título del Expediente
Public PBLSinForm As Long, ProrrogaSinForm As Long, ModifSinForm As Long 'Presupuesto Base de Licitación, Prórrogas y Modificaciones
Public PBL As Variant, Prorroga As Variant, Modif As Variant 'Presupuesto Base de Licitación, Prórrogas y Modificaciones
Public ValRefPre As Double, ValRefCal As Double, ValRef As Double 'Cálculos para el Valor de Referencia
Public LastCol As Long 'Última columna de la pestaña de SAGs (PDC)
Public LastCol2 As Long 'Última columna de la pestaña de SAGs (Informe 668)
Public LastColSAG As Long 'Última columna de la pestaña de SAGs (Informe)
Public FirstRowOMP As Long 'Primera fila donde hay coincidencia del SAG (Informe 668)
Public OMPRow As Long 'La fila de la OMP (Informe 668)
Public OMPfound As Boolean 'Misión 668 encontrada en el Informe (Informe 668)
Public i As Integer, j As Integer, k As Integer, l As Integer, m As Integer, n As Integer, o As Integer, p As Integer, _
    q As Integer, r As Integer, s As Integer, t As Integer, u As Integer, v As Integer, w As Integer ' Variables de los distintos bucles
Public str As String 'Celda donde se encuentra el nombre del REC (PDC)
Public openPos As Integer 'Punto "." después del rango y antes del nombre del REC (PDC)
Public closePos As Integer 'Paréntesis "(" después del nombre del REC y antes de la Maestranza (PDC)
Public midBit As String 'El nombre del REC (PDC)
Public PosFin As Integer 'Punto "." después del rango del REC (PDC)
Public RangStr As String 'Celda donde se encuentra el rango del REC (PDC)
Public Rang As String 'El rango del REC (PDC)
Public EcoPunt As Double 'Valor económico del punto (Memoria Criterios Adj.)
Public Redu As Double, VCC As Double 'Reducción sobre el PBL (Memoria Criterios Adj.)
Const MAXSKIPS As Long = 1  'Se salta 1 número que sea >0 (Informe)
Public Skips As Long 'Variable de saltos (Informe)
Public iCol As Long 'Columnas de la fila de SAGs (Informe)
Public Esp As String, Esp2 As String 'Espaciadores para los SAGs (Informe, Memoria, PCAP, PPT)
Public myRangeSAG As Range 'Rango de nuestra fila de SAGs (Informe)
Public countSAG As Integer 'Cuenta el número de SAGs en la línea (Informe)
Public cll As Range 'Celdas de la línea de SAGs (Informe)
Public NoCeroDir As String 'Dirección de la celda del penúltimo SAG con porcentaje >0 (Informe)
Public Repuesto As String 'Repuestos que compra el expediente
Public SistArm As String 'Número de Sistemas de Armas del expediente
Public Flota As String 'Flota que depende del número de SAGs
Dim errMsg As String 'Mensaje de Error por si surge un error
Dim errButton 'Tipo de Error
Dim errTitle As String 'Título del Error

'Elegir donde se encuentran las plantillas de los archivos
Public dlgInputTemplate As FileDialog 'Cuadro de diálogo para elegir la ruta donde se encuentran las plantillas
Public sFolderPathForLoad As String 'Ruta desde donde carga todas las plantillas

'Elegir donde guardar los archivos
Public dlgSaveFolder As FileDialog 'Cuadro de diálogo para elegir la ruta donde guardar los archivos
Public sFolderPathForSave As String 'Ruta donde guarda todos los archivos generados

Private Sub btnBorrar_Click()
    'Borrar elección documentación
    txtNumExp.Text = ""
    CheckBoxApen.Value = False
    CheckBoxArcLic.Value = False
    CheckBoxCl9.Value = False
    CheckBoxCl15.Value = False
    CheckBoxComAcct.Value = False
    CheckBoxCorAcc.Value = False
    CheckBoxCritAdj.Value = False
    CheckBoxCumpFact.Value = False
    CheckBoxCumpOfer.Value = False
    CheckBoxFactElec.Value = False
    CheckBoxIn.Value = False
    CheckBoxIn668.Value = False
    CheckBoxLotes.Value = False
    CheckBoxPCAP.Value = False
    CheckBoxPDC.Value = False
    CheckBoxPPT.Value = False
    CheckBoxPresup.Value = False
    CheckBoxSegPed.Value = False
    CheckBoxSeguPed.Value = False
    CheckBoxTrAnt.Value = False
    ComboBoxFinan.Text = ""
    ComboBoxLey.Text = ""
    ComboBoxTipExp.Text = ""
    OptionButtonDAM.Value = False
    OptionButtonNCP.Value = False
    OptionButtonNSP.Value = False
    'Borrar elección repuestos
    OptionButtonAsien.Value = False
    OptionButtonBal.Value = False
    OptionButtonBat.Value = False
    OptionButtonCel.Value = False
    OptionButtonElec.Value = False
    OptionButtonEst.Value = False
    OptionButtonFren.Value = False
    OptionButtonHel.Value = False
    OptionButtonIlu.Value = False
    OptionButtonMot.Value = False
    OptionButtonNeu.Value = False
    OptionButtonOtros.Value = False
    OptionButtonSeg.Value = False
    OptionButtonTren.Value = False
    OptionButtonTub.Value = False
    OptionButtonUni.Value = False
    'Borrar elección AM
    txtAM.Text = ""
    txtLote.Text = ""
    ComboBoxEmp.Text = ""
    MultiPage1.Pages(2).Visible = False 'Esconde la página AM
    'Borrar elección NSP
    ComboBoxEmp2.Text = ""
    MultiPage1.Pages(3).Visible = False 'Esconde la página NSP
    'Desactivar controles documentación
    CheckBoxApen.Enabled = False
    CheckBoxArcLic.Enabled = False
    CheckBoxCl9.Enabled = False
    CheckBoxCl15.Enabled = False
    CheckBoxComAcct.Enabled = False
    CheckBoxCorAcc.Enabled = False
    CheckBoxCritAdj.Enabled = False
    CheckBoxCumpFact.Enabled = False
    CheckBoxCumpOfer.Enabled = False
    CheckBoxFactElec.Enabled = False
    CheckBoxIn.Enabled = False
    CheckBoxIn668.Enabled = False
    CheckBoxLotes.Enabled = False
    CheckBoxPCAP.Enabled = False
    CheckBoxPDC.Enabled = False
    CheckBoxPPT.Enabled = False
    CheckBoxPresup.Enabled = False
    CheckBoxSegPed.Enabled = False
    CheckBoxSeguPed.Enabled = False
    CheckBoxTrAnt.Enabled = False
    ComboBoxFinan.Enabled = False
    ComboBoxLey.Enabled = False
    ComboBoxTipExp.Enabled = False
    lblFinan.Enabled = False
    lblLey.Enabled = False
    lblTipExp.Enabled = False
    btnCrearExpediente.Enabled = False
    'Desactivar botones repuestos
    OptionButtonAsien.Enabled = False
    OptionButtonBal.Enabled = False
    OptionButtonBat.Enabled = False
    OptionButtonCel.Enabled = False
    OptionButtonElec.Enabled = False
    OptionButtonEst.Enabled = False
    OptionButtonFren.Enabled = False
    OptionButtonHel.Enabled = False
    OptionButtonIlu.Enabled = False
    OptionButtonMot.Enabled = False
    OptionButtonNeu.Enabled = False
    OptionButtonOtros.Enabled = False
    OptionButtonSeg.Enabled = False
    OptionButtonTren.Enabled = False
    OptionButtonTub.Enabled = False
    OptionButtonUni.Enabled = False
    
    MultiPage1.Value = 0
    txtNumExp.SetFocus 'Pone el cursor donde se introduce el número del expediente
    
End Sub

Private Sub btnCancelar_Click()
    Unload Me
    MsgBox "Has cancelado la operación", vbExclamation, "Operación cancelada"
End Sub

Private Sub btnCrearExpediente_Click()

    On Error GoTo ErrHandler 'Si aparece un error
    
    rEXP = txtNumExp.Text
    
    If IsNumeric(rEXP) Then 'Comprueba que es un número
        If Len(rEXP) = 6 Then
            'Continúa con el código
        Else: MsgBox "Número incorrecto, introduzca número valido", vbCritical, "Número incorrecto"
            Exit Sub
        End If
    ElseIf StrPtr(rEXP) = 0 Then 'Cancelas el cuadro de diálogo
        MsgBox "¡No has introducido ningún número!", vbExclamation, "Nº Expediente no introducido"
        Exit Sub
    ElseIf rEXP = vbNullString Then 'Dejas el input vacío
        MsgBox "¡No has introducido ningún número!", vbExclamation, "Nº Expediente no introducido"
        Exit Sub
    Else: MsgBox "Número incorrecto, introduzca número valido", vbCritical, "Número incorrecto"
        Exit Sub
    End If
    
    'Error en el número de Expediente
    errMsg = "Ups, ha habido un error." & vbCrLf & vbCrLf & "Comprueba que has introducido bien los números del expediente o que existen."
    errButton = vbCritical
    errTitle = "Error desconocido"
    
    'Abrir el cuadro de diálogo de elegir carpeta donde se encuentran las plantillas
    Set dlgInputTemplate = Application.FileDialog(msoFileDialogFolderPicker)
    With dlgInputTemplate
        .Title = "Elige donde se encuentran las plantillas de los distintos archivos" 'Título del cuadro de dialogo
        .AllowMultiSelect = False 'Elegir más de una carpeta
        .InitialFileName = ThisWorkbook.Path & "\" 'Ruta donde guarda los archivos (inicialmente)
        .ButtonName = "Cargar plantillas" 'Nombre del botón del cuadro de dialogo
        If .Show <> -1 Then 'Si das a cancelar
            MsgBox "Has cancelado la operación", vbExclamation, "Operación cancelada"
            Exit Sub
        End If
        sFolderPathForLoad = .SelectedItems(1)
    End With
    Set dlgInputTemplate = Nothing

    'Abrir el cuadro de diálogo de elegir carpeta donde guardar archivos
    Set dlgSaveFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With dlgSaveFolder
        .Title = "Elige donde guardar los archivos del Expediente" 'Título del cuadro de dialogo
        .AllowMultiSelect = False 'Elegir más de una carpeta
        .InitialFileName = ThisWorkbook.Path & "\" 'Ruta donde guarda los archivos (inicialmente)
        .ButtonName = "Guardar archivos" 'Nombre del botón del cuadro de dialogo
        If .Show <> -1 Then 'Si das a cancelar
            MsgBox "Has cancelado la operación", vbExclamation, "Operación cancelada"
            Exit Sub
        End If
        sFolderPathForSave = .SelectedItems(1)
    End With
    Set dlgSaveFolder = Nothing
    
    
    'Acuerdo Marco (AM)
    If OptionButtonDAM.Value = True Then
'        With Sheets("PAAD")
'            Set FindRowExp = .Range("C:C").Find(What:=rEXP, LookIn:=xlValues) 'Busca la fila del NºExp que necesitamos
'            FindRowNumber = FindRowExp.Row 'Busca el número de la fila del Expediente
'            If .Cells(FindRowNumber, 5) = "C" Then
'                Set FindRowExp = .Range("C:C").Find(What:=rEXP, After:=.Cells(FindRowNumber, 6), LookIn:=xlValues) 'Busca la fila del NºExp que necesitamos
'                FindRowNumber = FindRowExp.Row 'Busca el número de la fila del Expediente
'            End If
'            AcuerdoM = Right(.Cells(FindRowNumber, 9), 6)
'        End With
        AcuerdoM = txtAM.Text 'Coge el AM del que hemos rellenado
    End If
    
    'LOTE AM y EMPRESA
    If OptionButtonDAM.Value = True Then 'DAM
        LoteAM = txtLote.Text
        EmpresaAM = ComboBoxEmp.Text
    ElseIf OptionButtonNSP.Value = True Then 'NSP
        EmpresaAM = ComboBoxEmp2.Text
    Else 'NCP
        'No hacer nada
    End If
    
    'Error en el número de Expediente
    errMsg = "Ups, ha habido un error." & vbCrLf & vbCrLf & "Comprueba que las casillas marcadas son correctas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    
    'Título del Expediente
    With Sheets("PAAD")
        Set FindRowExp = .Range("F:F").Find(What:=rEXP, LookIn:=xlValues) 'Busca la fila del NºExp que necesitamos
        FindRowNumber = FindRowExp.Row 'Busca el número de la fila del Expediente
        If .Cells(FindRowNumber, 5) = "C" Then 'Por si es un Compromiso
            Set FindRowExp = .Range("F:F").Find(What:=rEXP, After:=.Cells(FindRowNumber, 6), LookIn:=xlValues) 'Busca la fila del NºExp que necesitamos después del Compromiso
            FindRowNumber = FindRowExp.Row 'Busca el número de la fila del Expediente
        End If
        dashpos = InStr(1, .Cells(FindRowNumber, 9), "- AM") 'Posición del "- AM"
        If Not dashpos = 0 Then 'Si es 0 significa que no hay "- AM" (no es AM)
            Titulo = Left(.Cells(FindRowNumber, 9), dashpos - 2) 'Título
        Else
            Titulo = .Cells(FindRowNumber, 9).Value 'Título si no hay "- AM"
        End If
    End With
    
    'Anualidades del Expediente
    With Sheets("PAAD")
        If Not .Cells(FindRowNumber, 24).Value = 0 Then '3 anualidades
            NumAnual = 3
        ElseIf Not .Cells(FindRowNumber, 23).Value = 0 Then '2 anualidades
            NumAnual = 2
        Else '1 anualidad
            NumAnual = 1
        End If
    End With
    
    'Presupuesto Base de Licitación (PBL)
    With Sheets("PAAD")
        PBLSinForm = .Cells(FindRowNumber, 25).Value
        PBL = FormatCurrency(PBLSinForm, 0, , , vbTrue) 'Pone formato € sin decimales
    End With
    
    'Prórrogas y Modificaciones
    With Sheets("PAAD")
        If OptionButtonDAM.Value = True Then 'DAM (No hay Prórrogas ni Modificaciones)
            'Prórroga (igual al PBL)
            ProrrogaSinForm = 0
            Prorroga = FormatCurrency(ProrrogaSinForm, 0, , , vbTrue) 'Pone formato € sin decimales
            'Modificación (20% del PBL)
            ModifSinForm = 0
            Modif = FormatCurrency(ModifSinForm, 0, , , vbTrue) 'Pone formato € sin decimales
        ElseIf OptionButtonNCP.Value = True Then 'NCP
            'Prórroga (igual al PBL)
            ProrrogaSinForm = .Cells(FindRowNumber, 25).Value
            Prorroga = FormatCurrency(ProrrogaSinForm, 0, , , vbTrue) 'Pone formato € sin decimales
            'Modificación (20% del PBL)
            ModifSinForm = 0.2 * (.Cells(FindRowNumber, 25).Value)
            Modif = FormatCurrency(ModifSinForm, 0, , , vbTrue) 'Pone formato € sin decimales
        ElseIf OptionButtonNSP.Value = True Then 'NSP
            'Prórroga (igual al PBL)
            ProrrogaSinForm = .Cells(FindRowNumber, 25).Value
            Prorroga = FormatCurrency(ProrrogaSinForm, 0, , , vbTrue) 'Pone formato € sin decimales
            'Modificación (20% del PBL)
            ModifSinForm = 0.2 * (.Cells(FindRowNumber, 25).Value)
            Modif = FormatCurrency(ModifSinForm, 0, , , vbTrue) 'Pone formato € sin decimales
        End If
    End With
    
    'Valor de Referencia
    With Sheets("PAAD")
        If .Cells(FindRowNumber, 25).Value <= 10000 Then
            '0 € - 10.000€ VR Recomendado Inferior: 70%, Superior: 70%
            ValRefPre = (70 / 100) * .Cells(FindRowNumber, 25).Value
            ValRef = FormatCurrency(ValRefPre, 0, , , vbTrue)
        ElseIf .Cells(FindRowNumber, 25).Value > 10000 And .Cells(FindRowNumber, 25).Value <= 100000 Then
            '10.000 € - 100.000€ VR Recomendado Inferior: 70%, Superior: 75%
            ValRefCal = (((75 - 70) / (100000 - 10000)) * .Cells(FindRowNumber, 25).Value) + 70
            ValRefPre = (ValRefCal / 100) * .Cells(FindRowNumber, 25).Value
            ValRef = FormatCurrency(ValRefPre, 0, , , vbTrue)
        ElseIf .Cells(FindRowNumber, 25).Value > 100000 And .Cells(FindRowNumber, 25).Value <= 1000000 Then
            '100.000 € - 1.000.000€ VR Recomendado Inferior: 75%, Superior: 80%
            ValRefCal = (((80 - 75) / (1000000 - 100000)) * .Cells(FindRowNumber, 25).Value) + 75
            ValRefPre = (ValRefCal / 100) * .Cells(FindRowNumber, 25).Value
            ValRef = FormatCurrency(ValRefPre, 0, , , vbTrue)
        Else
            '1.000.000 € - 10.000.000€ VR Recomendado Inferior: 80%, Superior: 85%
            ValRefCal = (((85 - 80) / (10000000 - 1000000)) * .Cells(FindRowNumber, 25).Value) + 80
            ValRefPre = (ValRefCal / 100) * .Cells(FindRowNumber, 25).Value
            ValRef = FormatCurrency(ValRefPre, 0, , , vbTrue)
        End If
    End With

    'Aplicación Presupuestaria
    If ComboBoxFinan.Text = "660" Then '660
        AppPres = "14.022.122N.1.660"
    ElseIf ComboBoxFinan.Text = "668" Then '668
        AppPres = "14.003.122M.1.668"
    End If
    
    'Repuestos a adquirir por el Expediente
        'Asientos, Célula, Elementos estructurales, Motor, Hélice, Frenos, Tuberías, Iluminación, _
            Radiobalizas, Neumáticos, Elementos de unión, Equipos eléctricos, Tren de aterrizaje, Baterías y Sistemas de seguridad.
        If OptionButtonAsien.Value = True Then
            Repuesto = "de asientos lanzables"
        ElseIf OptionButtonCel.Value = True Then
            Repuesto = "de célula"
        ElseIf OptionButtonEst.Value = True Then
            Repuesto = "de elementos estructurales y accesorios"
        ElseIf OptionButtonMot.Value = True Then
            Repuesto = "de motor"
        ElseIf OptionButtonHel.Value = True Then
            Repuesto = "de hélices"
        ElseIf OptionButtonFren.Value = True Then
            Repuesto = "del sistema de frenos"
        ElseIf OptionButtonTub.Value = True Then
            Repuesto = "de tuberías y accesorios"
        ElseIf OptionButtonIlu.Value = True Then
            Repuesto = "de sistemas de iluminación"
        ElseIf OptionButtonBal.Value = True Then
            Repuesto = "de radiobalizas personales de supervivencia"
        ElseIf OptionButtonNeu.Value = True Then
            Repuesto = "de neumáticos"
        ElseIf OptionButtonUni.Value = True Then
            Repuesto = "de elementos de unión"
        ElseIf OptionButtonElec.Value = True Then
            Repuesto = "de equipos eléctricos, electrónicos y generadores"
        ElseIf OptionButtonTren.Value = True Then
            Repuesto = "de tren de aterrizaje"
        ElseIf OptionButtonBat.Value = True Then
            Repuesto = "de baterías"
        ElseIf OptionButtonSeg.Value = True Then
            Repuesto = "de los distintos Sistemas de Seguridad"
        ElseIf OptionButtonOtros.Value = True Then
            Repuesto = ""
        End If
        
    'Error en los repuestos
    errMsg = "Ups, ha habido un error." & vbCrLf & vbCrLf & "Comprueba que has elegido un repuesto."
    errButton = vbCritical
    errTitle = "Error desconocido"
    
    'Número de Sistemas de Armas
    With Sheets("SAG")
        Set FindRowSAG = .Range("C:C").Find(What:=rEXP, LookIn:=xlValues) 'Busca la fila del NºExp y SAG que necesitamos
        FindRowNumberSAG = FindRowSAG.Row 'Busca el número de la fila del Expediente y SAG
        If .Cells(FindRowNumberSAG, 2) = "C" Then
            Set FindRowSAG = .Range("C:C").Find(What:=rEXP, After:=.Cells(FindRowNumberSAG, 3), LookIn:=xlValues) 'Busca la fila del NºExp y SAG que necesitamos
            FindRowNumberSAG = FindRowSAG.Row 'Busca el número de la fila del Expediente y SAG
        End If
        'Set MyRange = wrdDoc.Content
        LastColSAG = .Cells(FindRowNumberSAG, .Columns.Count).End(xlToLeft).Column 'Encontrará la última columna usada en la fila de nuestro NºExp
        Set myRangeSAG = .Range(.Cells(FindRowNumberSAG, 14), .Cells(FindRowNumberSAG, LastColSAG - 1)) 'Rango de nuestra fila
        countSAG = 0 'Contador de SAGs con porcentaje mayor de 0
        For Each cll In myRangeSAG 'Para cada SAG en nuestro expediente
            If cll.Value > 0 Then 'Cuenta cuantos SAGs hay en nuestro expediente
                countSAG = countSAG + 1
            End If
        Next
        For iCol = LastColSAG - 1 To 1 Step -1 'Recorre la fila de SAGs de derecha a izquierda para buscar el penúltimo SAG
            If .Cells(FindRowNumberSAG, iCol).Value > 0 And Skips < MAXSKIPS Then 'Cuando encuentra el último con porcentaje >0 se lo salta
                Skips = Skips + 1
            ElseIf .Cells(FindRowNumberSAG, iCol).Value > 0 Then 'Cuando encuentra el penúltimo
                NoCeroDir = .Cells(FindRowNumberSAG, iCol).Address 'Dirección de la celda del penúltimo SAG de la fila
                Exit For 'Sale al encontrar el penúltimo
            End If
        Next iCol
    End With
    If countSAG = 1 Then 'Si sólo hay un SAG
        SistArm = "Sistema de Armas"
    ElseIf countSAG <= 3 Then 'Si hay más de 1 SAG pero menos de 4
        SistArm = "Sistemas de Armas"
    ElseIf countSAG > 3 Then
        SistArm = "diversos Sistemas de Armas"
    End If
    
    'Flotas
    If countSAG = 1 Then 'Si sólo hay un SAG
        Flota = "esta flota"
    ElseIf countSAG > 1 Then 'Si hay más de 1 SAG
        Flota = "estas flotas"
    End If
    
    'Coronel o Comandante Acctal.
    If CheckBoxComAcct.Value = True Then
        Coronel = "COMANDANTE JEFE ACCTAL."
    Else
        Coronel = "CORONEL JEFE"
    End If
    
    'Coronel o Coronel Acctal. (PELOCHE)
    If CheckBoxCorAcc.Value = True Then
        CoronelJef = "CORONEL JEFE ACCTAL. DEL ORGANO DE APOYO A LA DIRECCIÓN DE SOSTENIMIENTO"
    Else
        CoronelJef = "CORONEL JEFE DEL ORGANO DE APOYO" & vbCrLf & "A LA DIRECCIÓN DE SOSTENIMIENTO"
    End If
    
PrincipalCode:

    'ANEXO CLÁUSULA 9
    'Error
    errMsg = "Ha habido un error en el Anexo Cláusula 9." & vbCrLf & vbCrLf & "Comprueba que has marcado bien las casillas que necesitabas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    Call AnexoClausula9

     '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

    'ANEXO CLÁUSULA 15
    'Error
    errMsg = "Ha habido un error en el Anexo Cláusula 15." & vbCrLf & vbCrLf & "Comprueba que has marcado bien las casillas que necesitabas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    Call AnexoClausula15
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'INFORME 668
    'Error
    errMsg = "Ha habido un error en el Informe 668." & vbCrLf & vbCrLf & "Comprueba que has marcado bien las casillas que necesitabas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    Call Informe668
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'INFORME
    'Error
    errMsg = "Ha habido un error en el Informe." & vbCrLf & vbCrLf & "Comprueba que has marcado bien las casillas que necesitabas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    Call Informe
     
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

    'APÉNDICE ADICIONAL
    'Error
    errMsg = "Ha habido un error en el Apéndice Adicional." & vbCrLf & vbCrLf & "Comprueba que has marcado bien las casillas que necesitabas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    Call ApendiceAdicional
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'MEMORIA CRITERIOS ADJUDICACIÓN
    'Error
    errMsg = "Ha habido un error en la Memoria de Criterios Adjudicación." & vbCrLf & vbCrLf & "Comprueba que has marcado bien las casillas que necesitabas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    Call MemoriaCriteriosAdjudicacion
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'MODELO DE CUMPLIMENTACION DE FACTURA
    'Error
    errMsg = "Ha habido un error en el Modelo Cumplimentación Factura." & vbCrLf & vbCrLf & "Comprueba que has marcado bien las casillas que necesitabas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    Call ModeloCumplimentacionFactura
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'MODELO DE CUMPLIMENTACION DE OFERTA
    'Error
    errMsg = "Ha habido un error en el Modelo Cumplimentación Oferta." & vbCrLf & vbCrLf & "Comprueba que has marcado bien las casillas que necesitabas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    Call ModeloCumplimentacionOferta
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'MODELO DE SEGUIMIENTO DE PEDIDOS
    'Error
    errMsg = "Ha habido un error en el Modelo Seguimiento Pedidos." & vbCrLf & vbCrLf & "Comprueba que has marcado bien las casillas que necesitabas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    Call ModeloSeguimientoPedidos

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'PEDIDO DE CONTRATACIÓN (PDC)
    'Error
    errMsg = "Ha habido un error en el PDC." & vbCrLf & vbCrLf & "Comprueba que has marcado bien las casillas que necesitabas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    Call PDC

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'PRESUPUESTO
    'Error
    errMsg = "Ha habido un error en el Presupuesto." & vbCrLf & vbCrLf & "Comprueba que has marcado bien las casillas que necesitabas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    Call Presupuesto
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'PROPUESTA INCLUSION PCAP
    'Error
    errMsg = "Ha habido un error en la Propuesta PCAP." & vbCrLf & vbCrLf & "Comprueba que has marcado bien las casillas que necesitabas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    Call PropuestaInclusionPCAP
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'PLIEGO PRESCRIPCIONES TÉCNICAS (PPT)
    'Error
    errMsg = "Ha habido un error en el PPT." & vbCrLf & vbCrLf & "Comprueba que has marcado bien las casillas que necesitabas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    Call PPT
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'ARCHIVO LICITACION (Carpeta)
    'Error
    errMsg = "Ha habido un error en la carpeta Archivo Licitación." & vbCrLf & vbCrLf & "Comprueba que has marcado bien las casillas que necesitabas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    Call ArchivoLicitacionCARPETA
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'MODELO FACTURA ELECTRONICA (Carpeta y archivo)
    errMsg = "Ha habido un error en la carpeta Factura Electrónica." & vbCrLf & vbCrLf & "Comprueba que has marcado bien las casillas que necesitabas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    Call ModeloFacturaElectronicaCARPETA
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'MODELO SEGUIMIENTO PEDIDOS (Carpeta y archivo)
    errMsg = "Ha habido un error en la carpeta Modelo Seguimiento Pedidos." & vbCrLf & vbCrLf & "Comprueba que has marcado bien las casillas que necesitabas."
    errButton = vbCritical
    errTitle = "Error desconocido"
    Call ModeloSeguimientoPedidosCARPETA
    
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'Mensaje de finalización
     MsgBox "Todos los archivos han sido creados.", vbInformation, "Tarea completada"
    
    'Salir de la Macro
    Exit Sub 'Para que salga cuando compila todo sin mostrar el error de aquí debajo

ErrHandler:

    'Mensaje de error si eliges como libro para obtener datos este mismo Excel
    MsgBox errMsg, errButton, errTitle

    'Cerramos el Word
    If Not wrdApp Is Nothing Then
        wrdApp.Quit
        Set wrdApp = Nothing
    Else
        'No hay Words abiertos
    End If

End Sub

Sub AnexoClausula9()
    'ANEXO CLÁUSULA 9
    If CheckBoxCl9.Value = True Then

        strPath = sFolderPathForLoad & "\PLANTILLA Anexo clausula 9.dotx" 'Ruta elegida por el usuario para la plantilla
        'Abrimos la plantilla
        Set wrdApp = CreateObject("Word.Application")
        'wrdApp.Visible = True 'Si es True se ven los Words
        Set wrdDoc = wrdApp.Documents.Add(Template:=strPath, NewTemplate:=False, DocumentType:=0) 'Abrir Plantilla en ruta

        'Rango por si fuera necesario
        Set MyRange = wrdDoc.Content

        'Rellenamos el Número de Expediente
        wrdDoc.Bookmarks("NumExp").Range.Text = rEXP 'Cogemos el nºexp que se ha tecleado

        'Guardamos el documento en Word y PDF
        wrdDoc.SaveAs (sFolderPathForSave & "\" & rEXP & " Anexo a clausula 9.docx") 'Guarda el Word
        wrdDoc.ExportAsFixedFormat OutputFileName:=sFolderPathForSave & "\" & _
            rEXP & " Anexo a clausula 9.pdf", ExportFormat:=wdExportFormatPDF 'Guarda el PDF

        'Cerramos el Word
        wrdApp.Quit
        Set wrdApp = Nothing

    Else
        'Continúa con el código
    End If
End Sub

Sub AnexoClausula15()
    'ANEXO CLÁUSULA 15
    If CheckBoxCl15.Value = True Then

        strPath = sFolderPathForLoad & "\PLANTILLA Anexo cláusula 15.dotx"
        'Abrimos la plantilla
        Set wrdApp = CreateObject("Word.Application")
        Set wrdDoc = wrdApp.Documents.Add(Template:=strPath, NewTemplate:=False, DocumentType:=0) 'Abrir Plantilla en ruta

        'Rango por si fuera necesario
        Set MyRange = wrdDoc.Content

        'Rellenamos el Número de Expediente
        wrdDoc.Bookmarks("NumExp").Range.Text = rEXP 'Cogemos el nºexp que se ha tecleado

        'Guardamos el documento en Word y PDF
        wrdDoc.SaveAs (sFolderPathForSave & "\" & rEXP & " Anexo a clausula 15.docx") 'Guarda el Word
        wrdDoc.ExportAsFixedFormat OutputFileName:=sFolderPathForSave & "\" & _
            rEXP & " Anexo a clausula 15.pdf", ExportFormat:=wdExportFormatPDF 'Guarda el PDF

        'Cerramos el Word
        wrdApp.Quit
        Set wrdApp = Nothing

    Else
        'Continúa con el código
    End If
End Sub

Sub Informe668()
    'INFORME 668
    If CheckBoxIn668.Value = True Then

        'Elegimos si es DAM, NCP o NSP
        If OptionButtonDAM.Value = True Then 'DAM
            strPath = sFolderPathForLoad & "\PLANTILLA INFORME 668 DERIVADO AM.dotx"
        ElseIf OptionButtonNCP.Value = True Then 'NCP
            strPath = sFolderPathForLoad & "\PLANTILLA INFORME 668.dotx"
        ElseIf OptionButtonNSP.Value = True Then 'NSP
            strPath = sFolderPathForLoad & "\PLANTILLA INFORME 668.dotx"
        End If

        'Abrimos la plantilla
        Set wrdApp = CreateObject("Word.Application")
        Set wrdDoc = wrdApp.Documents.Add(Template:=strPath, NewTemplate:=False, DocumentType:=0) 'Abrir Plantilla en ruta
        'wrdApp.Visible = True

        'Rango por si fuera necesario
        Set MyRange = wrdDoc.Content
        Set myRange2 = wrdDoc.Content

        'Rellenamos el Número AM
        If OptionButtonDAM.Value = True Then
           wrdDoc.Bookmarks("NumAM").Range.Text = AcuerdoM
        End If

        'Rellenamos el LOTE AM y la EMPRESA
        If OptionButtonDAM.Value = True Then
            If Not LoteAM = "" Then
                wrdDoc.Bookmarks("LoteAM").Range.Text = "Lote " & LoteAM 'Lote AM
                If EmpresaAM = "Otra empresa" Then
                    wrdDoc.Bookmarks("EmpresaAM").Delete
                Else
                    wrdDoc.Bookmarks("EmpresaAM").Range.Text = EmpresaAM 'Empresa AM
                End If
            Else 'Resto de AM no tiene LOTES
                Set MyRange = wrdDoc.Bookmarks("LoteAM").Range 'Borra el marcador y el espacio que dejaría moviendo su rango
                With MyRange
                    .MoveStart Unit:=wdCharacter, Count:=-1
                    .Select
                End With
                MyRange = ""
                If EmpresaAM = "Otra empresa" Then
                    wrdDoc.Bookmarks("EmpresaAM").Delete
                Else
                    wrdDoc.Bookmarks("EmpresaAM").Range.Text = EmpresaAM 'Empresa AM
                End If
            End If
        End If

        'Rellenamos el Número de Expediente
        wrdDoc.Bookmarks("NumExp").Range.Text = rEXP 'Cogemos el nºexp que se ha tecleado

        'Título del Expediente
        wrdDoc.Bookmarks("Titulo").Range.Text = Titulo 'Cogemos el título ya calculado previamente

        'Presupuesto Base de Licitación (PBL)
        wrdDoc.Bookmarks("PBL").Range.Text = PBL

        'Misiones OMP
        With Sheets("SAG")
            Set FindRowSAG = .Range("C:C").Find(What:=rEXP, LookIn:=xlValues) 'Busca la fila del NºExp y SAG que necesitamos
            FindRowNumberSAG = FindRowSAG.Row 'Busca el número de la fila del Expediente y SAG
            If .Cells(FindRowNumberSAG, 2) = "C" Then 'Si es un Compromiso busca la siguiente coincidencia
                Set FindRowSAG = .Range("C:C").Find(What:=rEXP, After:=.Cells(FindRowNumberSAG, 3), LookIn:=xlValues) 'Busca la fila del NºExp y SAG que necesitamos
                FindRowNumberSAG = FindRowSAG.Row 'Busca el número de la fila del Expediente y SAG
            End If
            LastCol2 = .Cells(FindRowNumberSAG, .Columns.Count).End(xlToLeft).Column 'Encontrará la última columna usada en la fila 1
            For s = 14 To LastCol2 - 1 'El -1 es para que no coja la última columna donde vienen los 100%
                If Not .Cells(FindRowNumberSAG, s).Value = 0 Then 'Si el porcentaje del SAG es distinto de 0
                    Set FindRowOMP = Sheets("MISIONES 668").Range("B:B").Find(What:=.Cells(1, s), LookIn:=xlValues) 'Busca la fila de SAG en las OMPs
                    If FindRowOMP Is Nothing = False Then 'Si ese SAG tiene alguna OMP asociada
                        FirstRowOMP = FindRowOMP.Row 'Primera fila donde hay coincidencia del SAG
                        Do While Not FindRowOMP Is Nothing 'Hace un loop buscando todas las misiones del SAG
                            FindRowNumberOMP = FindRowOMP.Row 'Busca el número de la fila de la OMP
                            If Sheets("MISIONES 668").Cells(FindRowNumberOMP, 1).Value = "" Then 'Si la fila de la OMP está vacía
                                OMPRow = Sheets("MISIONES 668").Cells(FindRowNumberOMP, 1).End(xlUp).Row 'La nueva fila de OMP será la primera celda no vacía
                                Set myRangeOMP = wrdDoc.Content 'Rango del Word para buscar la misión 668
                                With myRangeOMP.Find 'Busca si la misión 668 ya existe en el Word
                                    .Text = Sheets("MISIONES 668").Cells(OMPRow, 1)
                                    .MatchCase = True
                                    .MatchWholeWord = True
                                    OMPfound = .Execute
                                End With
                                If OMPfound = True Then 'Si la misión 668 YA existe
                                    'No hacer nada
                                Else 'Si la misión 668 NO existe
                                    wrdDoc.Bookmarks("Misiones").Range.Text = Sheets("MISIONES 668").Cells(OMPRow, 1).Value & ", $" 'Se mete el ", $" para poner el marcador en caso de más de una misión
                                    MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                                    wrdDoc.Bookmarks.Add Name:="Misiones", Range:=MyRange
                                End If
                            Else 'Si la fila de la OMP no está vacía
                                wrdDoc.Bookmarks("Misiones").Range.Text = Sheets("MISIONES 668").Cells(FindRowNumberOMP, 1).Value & ", $"
                                MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="Misiones", Range:=MyRange
                            End If
                            Set FindRowOMP = Sheets("MISIONES 668").Range("B:B").FindNext(FindRowOMP) 'Busca la siguiente coincidencia del SAG en la columna SAG de las misiones
                            If FindRowOMP.Row = FirstRowOMP Then 'La fila actual de la misión coincide con la primera misión (ya que el loop es infinito y recorre las misiones infinitamente)
                                Exit Do 'Sale del loop cuando vuelve a la primera misión econtrada para este SAG
                            End If
                        Loop
                    End If
                End If
            Next s
            MyRange.SetRange Start:=MyRange.Start - 2, End:=MyRange.End 'Mueve el rango del marcador para borrar el sobrante ", $"
            wrdDoc.Bookmarks.Add Name:="Misiones", Range:=MyRange
            wrdDoc.Bookmarks("Misiones").Range.Text = "." 'Borra el sobrante de los SAGs (", $") y pone un punto "."
        End With

        'Objeto de adquisición del expediente
        If Repuesto = "" Then 'El repuesto sería Otros
            wrdDoc.Bookmarks("Objeto").Delete
        Else
            If countSAG = 1 Then 'Si sólo hay un SAG
                wrdDoc.Bookmarks("Objeto").Range.Text = "la adquisición y el suministro de repuestos " & Repuesto & " incluidos en los catálogos ilustrados de piezas asociados al " & SistArm & ""
            Else 'Si hay más de 1 SAG
                wrdDoc.Bookmarks("Objeto").Range.Text = "la adquisición y el suministro de repuestos " & Repuesto & " incluidos en los catálogos ilustrados de piezas asociados a los " & SistArm & ""
            End If
        End If
        With Sheets("SAG")
            If countSAG > 3 Then 'Si hay más de 3 SAGs con porcentaje mayor de 0, entonces el texto de nuestros marcadores es distinto
                wrdDoc.Bookmarks("SAG").Range.Text = "" 'Borra el SAG ya que viene incluido en el marcador anterior
            Else 'Si hay 3 o menos de 3 SAGs entonces el texto en los marcadores llevará escrito el nombre de cada SAG
                For l = 14 To LastColSAG - 1 'El -1 es para que no coja la última columna donde vienen los 100%
                    If Not .Cells(FindRowNumberSAG, l).Value = 0 Then 'Todas las aeronaves
                        If .Cells(FindRowNumberSAG, l).Address = NoCeroDir Then 'Si la celda del SAG es el penúltimo SAG, entonces el marcador será "y", en vez de una coma ","
                            Esp = " y @"
                        Else 'Si no es la penúltima, seguirá con una coma
                            Esp = ", @"
                        End If
                        If .Cells(1, l).Value = "AE.9" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "F-5 (A.9)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "C.15" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "F-18 HORNET (C.15)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "C.16" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "EUROFIGHTER TYPHOON 2000 (C.16)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "H.21" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "SUPERPUMA (HD/HT.21)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "H.24" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "SIKORSKY (H.24)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "H.25" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "COLIBRÍ (H.25)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "H.27" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "COUGAR (HT.27)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "H.29" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "NH-90 (H.29)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "E.24" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "BECHCRAFT 33C (E.24)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "E.25" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-101 (E.25)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "E.26" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "T-35 C TAMIZ (E.26)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "E.27" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "PILATUS PC-21 (E.27)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "T.11" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "FALCON 20 (T.11)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "T.12" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-212 (T.12)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "T.18" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "FALCON 900 (T.18)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "T.19" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-235 (T.19)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "TR.20" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "CESSNA CITATION (TR.20)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "T.21" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-295 (T.21)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "T.22" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "A310 (T.22)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "T.23" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "A400M (T.23)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "P.3" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "P-3 ORION" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "UD.13 / UD. 14" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "CL-215 (UD.13) y CL-415 (UD.14)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "U.22" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "BEECHCRAFT C 90 -KING AIR (U.22)" & Esp 'Introducimos el "@" para poder crear un marcador y seguir uniendo SAGs
                            myRange2.Find.Execute FindText:="@", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                        ElseIf .Cells(1, l).Value = "XLAN" Then
                            'No hacer nada
                        ElseIf .Cells(1, l).Value = "XEPS" Then
                            'No hacer nada
                        End If
                    End If
                Next l
                myRange2.SetRange Start:=myRange2.Start - 2, End:=myRange2.End
                wrdDoc.Bookmarks.Add Name:="SAG", Range:=myRange2
                wrdDoc.Bookmarks("SAG").Range.Text = "" 'Borra el sobrante de los SAGs (", @")
            End If
        End With

        'Coronel Jefe OAD
        wrdDoc.Bookmarks("CoronelJef").Range.Text = CoronelJef
        If CheckBoxCorAcc.Value = True Then
            wrdDoc.Bookmarks("CoronelJef2").Range.Font.ColorIndex = wdRed
            wrdDoc.Bookmarks("CoronelJef3").Range.Font.ColorIndex = wdRed
            wrdDoc.Bookmarks("CoronelJef2").Delete
            wrdDoc.Bookmarks("CoronelJef3").Delete
        Else
            wrdDoc.Bookmarks("CoronelJef2").Delete
            wrdDoc.Bookmarks("CoronelJef3").Delete
        End If

        'Guardamos el documento en Word y PDF
        wrdDoc.SaveAs (sFolderPathForSave & "\" & rEXP & " INFORME 668.docx") 'Guarda el Word
'        wrdDoc.ExportAsFixedFormat OutputFileName:=sFolderPathForSave & "\" & _
'            rEXP & " INFORME 668.pdf", ExportFormat:=wdExportFormatPDF 'Guarda el PDF

        'Cerramos el Word
        wrdApp.Quit
        Set wrdApp = Nothing

    Else
        'Continúa con el código
    End If
End Sub

Sub Informe()
    'INFORME
    If CheckBoxIn.Value = True Then

        'Elegimos si es Ley 9-2017 o LCSDPS
        If ComboBoxLey.Text = "Ley 9-2017" Then 'Ley 9-2017
            strPath = sFolderPathForLoad & "\PLANTILLA INFORME LEY 9-2017.dotx"
        ElseIf ComboBoxLey.Text = "LCSPDS" Then 'LCSPDS
            strPath = sFolderPathForLoad & "\PLANTILLA INFORME LCSPDS.dotx"
        End If

        'Abrimos la plantilla
        Set wrdApp = CreateObject("Word.Application")
        Set wrdDoc = wrdApp.Documents.Add(Template:=strPath, NewTemplate:=False, DocumentType:=0) 'Abrir Plantilla en ruta
        'wrdApp.Visible = True

        'Rango por si fuera necesario
        Set MyRange = wrdDoc.Content

        'Rellenamos el Número de Expediente
        wrdDoc.Bookmarks("NumExp").Range.Text = rEXP 'Cogemos el nºexp que se ha tecleado

        'Título del Expediente
        wrdDoc.Bookmarks("Titulo").Range.Text = Titulo 'Cogemos el título ya calculado previamente

        'Repuestos a adquirir por el Expediente
        If Repuesto = "" Then
            wrdDoc.Bookmarks("Repuesto").Delete
        Else
            wrdDoc.Bookmarks("Repuesto").Range.Text = Repuesto
        End If
        
        'Sistemas de Armas y SAGs
        If countSAG = 1 Then 'Si sólo hay un SAG
            wrdDoc.Bookmarks("SistArm").Range.Text = "del " & SistArm
        Else 'Si hay más de 1 SAG
            wrdDoc.Bookmarks("SistArm").Range.Text = "de los " & SistArm
        End If
        With Sheets("SAG")
            If countSAG > 3 Then 'Si hay más de 3 SAGs con porcentaje mayor de 0, entonces el texto de nuestros marcadores es distinto
                wrdDoc.Bookmarks("SAG").Range.Text = "" 'Borra el SAG ya que viene incluido en el marcador anterior
            Else 'Si hay 3 o menos de 3 SAGs entonces el texto en los marcadores llevará escrito el nombre de cada SAG
                For l = 14 To LastColSAG - 1 'El -1 es para que no coja la última columna donde vienen los 100%
                    If Not .Cells(FindRowNumberSAG, l).Value = 0 Then 'Todas las aeronaves
                        If .Cells(FindRowNumberSAG, l).Address = NoCeroDir Then 'Si la celda del SAG es el penúltimo SAG, entonces el marcador será "y", en vez de una coma ","
                            Esp = " y $"
                        Else 'Si no es la penúltima, seguirá con una coma
                            Esp = ", $"
                        End If
                        If .Cells(1, l).Value = "AE.9" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "F-5 (A.9)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "C.15" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "F-18 HORNET (C.15)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "C.16" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "EUROFIGHTER TYPHOON 2000 (C.16)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "H.21" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "SUPERPUMA (HD/HT.21)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "H.24" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "SIKORSKY (H.24)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "H.25" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "COLIBRÍ (H.25)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "H.27" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "COUGAR (HT.27)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "H.29" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "NH-90 (H.29)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "E.24" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "BECHCRAFT 33C (E.24)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "E.25" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-101 (E.25)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "E.26" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "T-35 C TAMIZ (E.26)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "E.27" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "PILATUS PC-21 (E.27)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.11" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "FALCON 20 (T.11)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.12" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-212 (T.12)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.18" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "FALCON 900 (T.18)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.19" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-235 (T.19)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "TR.20" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "CESSNA CITATION (TR.20)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.21" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-295 (T.21)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.22" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "A310 (T.22)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.23" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "A400M (T.23)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "P.3" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "P-3 ORION" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "UD.13 / UD. 14" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "CL-215 (UD.13) y CL-415 (UD.14)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "U.22" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "BEECHCRAFT C 90 -KING AIR (U.22)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "XLAN" Then
                            'No hacer nada
                        ElseIf .Cells(1, l).Value = "XEPS" Then
                            'No hacer nada
                        End If
                    End If
                Next l
                MyRange.SetRange Start:=MyRange.Start - 2, End:=MyRange.End
                wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                wrdDoc.Bookmarks("SAG").Range.Text = " " 'Borra el sobrante de los SAGs (", $")
            End If
        End With
        
        'Lotes en el Expediente
        If CheckBoxLotes.Value = True Then 'Si hay Lotes
            Set MyRange = wrdDoc.Bookmarks("Lotes").Range 'Borra el marcador y el espacio que dejaría moviendo su rango
            With MyRange
                .MoveStart Unit:=wdCharacter, Count:=-1
                .Select
            End With
            MyRange = ""
        Else 'No hay lotes
            wrdDoc.Bookmarks("Lotes").Range.Font.ColorIndex = wdBlack 'Deja el parrafo al no haber Lotes
            wrdDoc.Bookmarks("Lotes").Delete
        End If

        'Presupuesto Base de Licitación (PBL)
        wrdDoc.Bookmarks("PBL").Range.Text = PBL

        'Tabla Anualidades
        Set myTable = wrdDoc.Tables(1)
        For j = 1 To NumAnual - 1
            myTable.Rows.Add 'Añade fila por anualidad
        Next j
        With Sheets("PAAD")
            If NumAnual = 1 Then
                myTable.Cell(2, 1).Range.Text = Right(.Cells(2, 22).Value, 4) 'Año anualidad primera
                myTable.Cell(2, 2).Range.Text = FormatCurrency(.Cells(FindRowNumber, 22).Value, 0, , , vbTrue) 'Dinero anualidad
            Else
                For k = 1 To NumAnual - 1
                    myTable.Cell(2, 1).Range.Text = Right(.Cells(2, 22).Value, 4) 'Año anualidad primera
                    myTable.Cell(2 + k, 1).Range.Text = Right(.Cells(2, 22 + k).Value, 4) 'Demás años anualidades
                    myTable.Cell(2, 2).Range.Text = FormatCurrency(.Cells(FindRowNumber, 22).Value, 0, , , vbTrue) 'Dinero anualidad
                    myTable.Cell(2 + k, 2).Range.Text = FormatCurrency(.Cells(FindRowNumber, 22 + k).Value, 0, , , vbTrue) 'Demás dineros anualidades
                Next k
            End If
        End With

        'Anualidades
        If NumAnual = 3 Then '3 anualidades
            wrdDoc.Bookmarks("Anualidades").Range.Text = "tres anualidades"
        ElseIf NumAnual = 2 Then '2 anualidades
            wrdDoc.Bookmarks("Anualidades").Range.Text = "dos anualidades"
        Else '1 anualidad
            wrdDoc.Bookmarks("Anualidades").Range.Text = "una anualidad"
        End If

        'Coronel o Comandante Acctal.
        wrdDoc.Bookmarks("Coronel").Range.Text = Coronel

        'Coronel Jefe OAD
        wrdDoc.Bookmarks("CoronelJef").Range.Text = CoronelJef

        'Guardamos el documento en Word y PDF
        wrdDoc.SaveAs (sFolderPathForSave & "\" & rEXP & " INFORME.docx") 'Guarda el Word
'        wrdDoc.ExportAsFixedFormat OutputFileName:=sFolderPathForSave & "\" & _
'            rEXP & " INFORME.pdf", ExportFormat:=wdExportFormatPDF 'Guarda el PDF

        'Cerramos el Word
        wrdApp.Quit
        Set wrdApp = Nothing

    Else
        'Continúa con el código
    End If
End Sub

Sub ApendiceAdicional()
    'APÉNDICE ADICIONAL
    If CheckBoxApen.Value = True Then
        strPath = sFolderPathForLoad & "\XXXXXX Apéndice adicional.pdf"
        FileCopy strPath, sFolderPathForSave & "\" & rEXP & " Apéndice adicional.pdf"
    Else
        'Continúa con el código
    End If
End Sub

Sub MemoriaCriteriosAdjudicacion()
    'MEMORIA CRITERIOS ADJUDICACIÓN
    If CheckBoxCritAdj.Value = True Then

        strPath = sFolderPathForLoad & "\PLANTILLA MEMORIA CRITERIOS ADJUDICACION.dotx"
        'Abrimos la plantilla
        Set wrdApp = CreateObject("Word.Application")
        Set wrdDoc = wrdApp.Documents.Add(Template:=strPath, NewTemplate:=False, DocumentType:=0) 'Abrir Plantilla en ruta
        'wrdApp.Visible = True

        'Rango por si fuera necesario
        Set MyRange = wrdDoc.Content

        'Rellenamos los Números de Expediente
        wrdDoc.Bookmarks("NumExp").Range.Text = rEXP 'Cogemos el nºexp que se ha tecleado
        wrdDoc.Bookmarks("NumExp2").Range.Text = rEXP 'Cogemos el nºexp que se ha tecleado

        'Título del Expediente
        wrdDoc.Bookmarks("Titulo").Range.Text = Titulo 'Cogemos el título ya calculado previamente

        'Repuestos a adquirir por el Expediente
        If Repuesto = "" Then
            wrdDoc.Bookmarks("Repuesto").Delete
        Else
            wrdDoc.Bookmarks("Repuesto").Range.Text = Repuesto
        End If
            
        'Sistemas de Armas y SAGs
        If countSAG = 1 Then 'Si sólo hay un SAG
            wrdDoc.Bookmarks("SistArm").Range.Text = " al " & SistArm
        Else 'Si hay más de 1 SAG
            wrdDoc.Bookmarks("SistArm").Range.Text = " a los " & SistArm
        End If
        With Sheets("SAG")
            If countSAG > 3 Then 'Si hay más de 3 SAGs con porcentaje mayor de 0, entonces el texto de nuestros marcadores es distinto
                wrdDoc.Bookmarks("SAG").Range.Text = "" 'Borra el SAG ya que viene incluido en el marcador anterior
            Else 'Si hay 3 o menos de 3 SAGs entonces el texto en los marcadores llevará escrito el nombre de cada SAG
                For l = 14 To LastColSAG - 1 'El -1 es para que no coja la última columna donde vienen los 100%
                    If Not .Cells(FindRowNumberSAG, l).Value = 0 Then 'Todas las aeronaves
                        If .Cells(FindRowNumberSAG, l).Address = NoCeroDir Then 'Si la celda del SAG es el penúltimo SAG, entonces el marcador será "y", en vez de una coma ","
                            Esp = " y $"
                        Else 'Si no es la penúltima, seguirá con una coma
                            Esp = ", $"
                        End If
                        If .Cells(1, l).Value = "AE.9" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "F-5 (A.9)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "C.15" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "F-18 HORNET (C.15)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "C.16" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "EUROFIGHTER TYPHOON 2000 (C.16)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "H.21" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "SUPERPUMA (HD/HT.21)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "H.24" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "SIKORSKY (H.24)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "H.25" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "COLIBRÍ (H.25)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "H.27" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "COUGAR (HT.27)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "H.29" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "NH-90 (H.29)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "E.24" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "BECHCRAFT 33C (E.24)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "E.25" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-101 (E.25)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "E.26" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "T-35 C TAMIZ (E.26)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "E.27" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "PILATUS PC-21 (E.27)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.11" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "FALCON 20 (T.11)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.12" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-212 (T.12)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.18" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "FALCON 900 (T.18)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.19" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-235 (T.19)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "TR.20" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "CESSNA CITATION (TR.20)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.21" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-295 (T.21)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.22" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "A310 (T.22)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.23" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "A400M (T.23)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "P.3" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "P-3 ORION" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "UD.13 / UD. 14" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "CL-215 (UD.13) y CL-415 (UD.14)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "U.22" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "BEECHCRAFT C 90 -KING AIR (U.22)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "XLAN" Then
                            'No hacer nada
                        ElseIf .Cells(1, l).Value = "XEPS" Then
                            'No hacer nada
                        End If
                    End If
                Next l
                MyRange.SetRange Start:=MyRange.Start - 2, End:=MyRange.End
                wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                wrdDoc.Bookmarks("SAG").Range.Text = " " 'Borra el sobrante de los SAGs (", $")
            End If
        End With
        
        'Flota
        wrdDoc.Bookmarks("Flota").Range.Text = Flota
        
        'Presupuesto Base de Licitación (PBL)
        wrdDoc.Bookmarks("PBL").Range.Text = PBL

        'Valor de Referencia
        wrdDoc.Bookmarks("VR").Range.Text = FormatCurrency(ValRef, 0, , , vbTrue)

        'Valor Ecónomico del Punto
        With Sheets("PAAD")
            EcoPunt = (.Cells(FindRowNumber, 25).Value - ValRefPre) / 80 'Cálculo del valor económico del punto
            wrdDoc.Bookmarks("EcoPunt").Range.Text = FormatCurrency(EcoPunt, 2, , , vbTrue) & "/punto"
        End With

        'Reducción
        With Sheets("PAAD")
            VCC = (60 * .Cells(FindRowNumber, 25).Value + 20 * ValRefPre) / 80
            Redu = .Cells(FindRowNumber, 25).Value - VCC 'Cálculo de la reducción sobre el PBL
            wrdDoc.Bookmarks("Redu").Range.Text = FormatCurrency(Redu, 0, , , vbTrue)
        End With

        'Guardamos el documento en Word y PDF
        wrdDoc.SaveAs (sFolderPathForSave & "\" & rEXP & " MEMORIA CRITERIOS ADJUDICACION.docx") 'Guarda el Word
'        wrdDoc.ExportAsFixedFormat OutputFileName:=sFolderPathForSave & "\" & _
'            rEXP & " MEMORIA CRITERIOS ADJUDICACION.pdf", ExportFormat:=wdExportFormatPDF 'Guarda el PDF

        'Cerramos el Word
        wrdApp.Quit
        Set wrdApp = Nothing

    Else
        'Continúa con el código
    End If
End Sub

Sub ModeloCumplimentacionFactura()
    'MODELO DE CUMPLIMENTACION DE FACTURA
    If CheckBoxCumpFact.Value = True Then

        strPath = sFolderPathForLoad & "\PLANTILLA MODELO DE CUMPLIMENTACION DE FACTURA.dotx"
        'Abrimos la plantilla
        Set wrdApp = CreateObject("Word.Application")
        Set wrdDoc = wrdApp.Documents.Add(Template:=strPath, NewTemplate:=False, DocumentType:=0) 'Abrir Plantilla en ruta
        'wrdApp.Visible = True

        'Rango por si fuera necesario
        Set MyRange = wrdDoc.Content

        'Rellenamos los Números de Expediente
        wrdDoc.Bookmarks("NumExp").Range.Text = rEXP 'Cogemos el nºexp que se ha tecleado

        'Título del Expediente
        wrdDoc.Bookmarks("Titulo").Range.Text = Titulo 'Cogemos el título ya calculado previamente

        'Guardamos el documento en Word y PDF
        wrdDoc.SaveAs (sFolderPathForSave & "\" & rEXP & " MODELO DE CUMPLIMENTACIÓN DE FACTURA.docx") 'Guarda el Word
        wrdDoc.ExportAsFixedFormat OutputFileName:=sFolderPathForSave & "\" & _
            rEXP & " MODELO DE CUMPLIMENTACIÓN DE FACTURA.pdf", ExportFormat:=wdExportFormatPDF 'Guarda el PDF

        'Cerramos el Word
        wrdApp.Quit
        Set wrdApp = Nothing

    Else
        'Continúa con el código
    End If
End Sub

Sub ModeloCumplimentacionOferta()
    'MODELO DE CUMPLIMENTACION DE OFERTA
    If CheckBoxCumpOfer.Value = True Then

        strPath = sFolderPathForLoad & "\PLANTILLA MODELO DE CUMPLIMENTACION DE OFERTA.dotx"
        'Abrimos la plantilla
        Set wrdApp = CreateObject("Word.Application")
        Set wrdDoc = wrdApp.Documents.Add(Template:=strPath, NewTemplate:=False, DocumentType:=0) 'Abrir Plantilla en ruta
        'wrdApp.Visible = True

        'Rango por si fuera necesario
        Set MyRange = wrdDoc.Content

        'Rellenamos los Números de Expediente
        wrdDoc.Bookmarks("NumExp").Range.Text = rEXP 'Cogemos el nºexp que se ha tecleado

        'Título del Expediente
        wrdDoc.Bookmarks("Titulo").Range.Text = Titulo 'Cogemos el título ya calculado previamente

        'Guardamos el documento en Word y PDF
        wrdDoc.SaveAs (sFolderPathForSave & "\" & rEXP & " MODELO DE CUMPLIMENTACION DE OFERTA.docx") 'Guarda el Word
        wrdDoc.ExportAsFixedFormat OutputFileName:=sFolderPathForSave & "\" & _
            rEXP & " MODELO DE CUMPLIMENTACION DE OFERTA.pdf", ExportFormat:=wdExportFormatPDF 'Guarda el PDF

        'Cerramos el Word
        wrdApp.Quit
        Set wrdApp = Nothing

    Else
        'Continúa con el código
    End If
End Sub

Sub ModeloSeguimientoPedidos()
    'MODELO DE SEGUIMIENTO DE PEDIDOS
    If CheckBoxSegPed.Value = True Then

        strPath = sFolderPathForLoad & "\PLANTILLA MODELO DE SEGUIMIENTO DE PEDIDOS.dotx"
        'Abrimos la plantilla
        Set wrdApp = CreateObject("Word.Application")
        Set wrdDoc = wrdApp.Documents.Add(Template:=strPath, NewTemplate:=False, DocumentType:=0) 'Abrir Plantilla en ruta
        'wrdApp.Visible = True

        'Rango por si fuera necesario
        Set MyRange = wrdDoc.Content

        'Rellenamos los Números de Expediente
        wrdDoc.Bookmarks("NumExp").Range.Text = rEXP 'Cogemos el nºexp que se ha tecleado

        'Título del Expediente
        wrdDoc.Bookmarks("Titulo").Range.Text = Titulo 'Cogemos el título ya calculado previamente

        'Guardamos el documento en Word y PDF
        wrdDoc.SaveAs (sFolderPathForSave & "\" & rEXP & " MODELO DE SEGUIMIENTO DE PEDIDOS.docx") 'Guarda el Word
        wrdDoc.ExportAsFixedFormat OutputFileName:=sFolderPathForSave & "\" & _
            rEXP & " MODELO DE SEGUIMIENTO DE PEDIDOS.pdf", ExportFormat:=wdExportFormatPDF 'Guarda el PDF

        'Cerramos el Word
        wrdApp.Quit
        Set wrdApp = Nothing

    Else
        'Continúa con el código
    End If
End Sub

Sub PDC()
    'PEDIDO DE CONTRATACIÓN (PDC)
    If CheckBoxPDC.Value = True Then

        strPath = sFolderPathForLoad & "\PLANTILLA PDC.dotx"
        'Abrimos la plantilla
        Set wrdApp = CreateObject("Word.Application")
        Set wrdDoc = wrdApp.Documents.Add(Template:=strPath, NewTemplate:=False, DocumentType:=0) 'Abrir Plantilla en ruta
        'wrdApp.Visible = True

        'Rango por si fuera necesario
        Set MyRange = wrdDoc.Content

        'Tipo de Negocio Jurídico
        If OptionButtonDAM.Value = True Then 'DAM
            wrdDoc.Bookmarks("Negocio").Range.Text = "CONTRATO BASADO EN AM"
        ElseIf OptionButtonNCP.Value = True Then 'NCP
            wrdDoc.Bookmarks("Negocio").Range.Text = "CONTRATO"
        ElseIf OptionButtonNSP.Value = True Then 'NSP
            wrdDoc.Bookmarks("Negocio").Range.Text = "CONTRATO"
        End If

        'Rellenamos los Números de Expediente
        wrdDoc.Bookmarks("NumExp").Range.Text = rEXP 'Cogemos el nºexp que se ha tecleado
        wrdDoc.Bookmarks("NumExp2").Range.Text = rEXP 'Cogemos el nºexp que se ha tecleado
        wrdDoc.Bookmarks("NumExp3").Range.Text = rEXP 'Cogemos el nºexp que se ha tecleado
        wrdDoc.Bookmarks("NumExp4").Range.Text = rEXP 'Cogemos el nºexp que se ha tecleado

        'Título del Expediente
        wrdDoc.Bookmarks("Titulo").Range.Text = Titulo 'Cogemos el título ya calculado previamente

        'Proceso Adjudicatario
        If OptionButtonDAM.Value = True Then 'DAM
            wrdDoc.Bookmarks("ProcAdj").Range.Text = "NEGOCIADO SIN PUBLICIDAD"
        ElseIf OptionButtonNCP.Value = True Then 'NCP
            wrdDoc.Bookmarks("ProcAdj").Range.Text = "NEGOCIADO CON PUBLICIDAD"
        ElseIf OptionButtonNSP.Value = True Then 'NSP
            wrdDoc.Bookmarks("ProcAdj").Range.Text = "NEGOCIADO SIN PUBLICIDAD"
        End If

        'Tipo de Expediente
        If ComboBoxTipExp.Text = "Nacional" Then 'Nacional
            wrdDoc.Bookmarks("TipExp").Range.Text = "NACIONAL"
        ElseIf ComboBoxTipExp.Text = "Extranjero" Then 'Extranjero
            wrdDoc.Bookmarks("TipExp").Range.Text = "EXT"
        End If

        'Elegimos si es Ley 9-2017 o LCSDPS
        If ComboBoxLey.Text = "Ley 9-2017" Then 'Ley 9-2017
            wrdDoc.Bookmarks("Ley").Range.Text = "LCSP"
        ElseIf ComboBoxLey.Text = "LCSPDS" Then 'LCSPDS
            wrdDoc.Bookmarks("Ley").Range.Text = "LCSPDS"
        End If

        'Fecha de plazo de ejecución
        With Sheets("PAAD")
            If Not .Cells(FindRowNumber, 24).Value = 0 Then 'Tercera anualidad
                wrdDoc.Bookmarks("Fecha").Range.Text = "30/11/" & Right(.Cells(2, 24).Value, 4)
            ElseIf Not .Cells(FindRowNumber, 23).Value = 0 Then 'Segunda anualidad
                wrdDoc.Bookmarks("Fecha").Range.Text = "30/11/" & Right(.Cells(2, 23).Value, 4)
            Else 'Primera anualidad
                wrdDoc.Bookmarks("Fecha").Range.Text = "30/11/" & Right(.Cells(2, 22).Value, 4)
            End If
        End With

        'Sistema de Armas
        With Sheets("SAG")
            Set FindRowSAG = .Range("C:C").Find(What:=rEXP, LookIn:=xlValues) 'Busca la fila del NºExp y SAG que necesitamos
            FindRowNumberSAG = FindRowSAG.Row 'Busca el número de la fila del Expediente y SAG
            If .Cells(FindRowNumberSAG, 2) = "C" Then
                Set FindRowSAG = .Range("C:C").Find(What:=rEXP, After:=.Cells(FindRowNumberSAG, 3), LookIn:=xlValues) 'Busca la fila del NºExp y SAG que necesitamos
                FindRowNumberSAG = FindRowSAG.Row 'Busca el número de la fila del Expediente y SAG
            End If
            Set MyRange = wrdDoc.Content
            LastCol = .Cells(FindRowNumberSAG, .Columns.Count).End(xlToLeft).Column 'Encontrará la última columna usada en la fila 1
            For l = 14 To LastCol - 1 'El -1 es para que no coja la última columna donde vienen los 100%
                If Not .Cells(FindRowNumberSAG, l).Value = 0 Then
                    wrdDoc.Bookmarks("SAG").Range.Text = .Cells(1, l).Value & " / $" 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                    MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                    wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                End If
            Next l
            MyRange.SetRange Start:=MyRange.Start - 3, End:=MyRange.End
            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
            wrdDoc.Bookmarks("SAG").Range.Text = "" 'Borra el sobrante de los SAGs (" / $")
        End With

        'OTAN/Grupo Material
        With Sheets("SAG")
            Set FindRowSAG = .Range("C:C").Find(What:=rEXP, LookIn:=xlValues) 'Busca la fila del NºExp y SAG que necesitamos
            FindRowNumberSAG = FindRowSAG.Row 'Busca el número de la fila del Expediente y SAG
            If .Cells(FindRowNumberSAG, 2) = "C" Then
                Set FindRowSAG = .Range("C:C").Find(What:=rEXP, After:=.Cells(FindRowNumberSAG, 3), LookIn:=xlValues) 'Busca la fila del NºExp y SAG que necesitamos
                FindRowNumberSAG = FindRowSAG.Row 'Busca el número de la fila del Expediente y SAG
            End If
            Set MyRange = wrdDoc.Content
            Set myRange2 = wrdDoc.Content
            LastCol = .Cells(FindRowNumberSAG, .Columns.Count).End(xlToLeft).Column 'Encontrará la última columna usada en la fila 1
            For m = 14 To LastCol - 1 'El -1 es para que no coja la última columna donde vienen los 100%
                If Not .Cells(FindRowNumberSAG, m).Value = 0 Then
                    If .Cells(1, m).Value = "AE.9" Then
                        wrdDoc.Bookmarks("GSAG").Range.Text = "1600A9" & " / $" 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                        MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                        wrdDoc.Bookmarks.Add Name:="GSAG", Range:=MyRange
                    ElseIf .Cells(1, m).Value = "TR.20" Then
                        wrdDoc.Bookmarks("GSAG").Range.Text = "1600T20" & " / $" 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                        MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                        wrdDoc.Bookmarks.Add Name:="GSAG", Range:=MyRange
                    ElseIf .Cells(1, m).Value = "UD.13 / UD. 14" Then
                        wrdDoc.Bookmarks("GSAG").Range.Text = "1600U13 / 1600U14" & " / $" 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                        MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                        wrdDoc.Bookmarks.Add Name:="GSAG", Range:=MyRange
                    ElseIf .Cells(1, m).Value = "XEPS" Then
                        wrdDoc.Bookmarks("GSAG").Range.Text = "8475XEPS" & " / $" 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                        MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                        wrdDoc.Bookmarks.Add Name:="GSAG", Range:=MyRange
                    ElseIf .Cells(1, m).Value = "XLAN" Then
                        wrdDoc.Bookmarks("GSAG").Range.Text = "1750PARA" & " / $" 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                        MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                        wrdDoc.Bookmarks.Add Name:="GSAG", Range:=MyRange
                    Else
                        wrdDoc.Bookmarks("GSAG").Range.Text = "1600" & .Cells(1, m).Value & " / $" 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                        myRange2.Find.Execute FindText:="1600", MatchCase:=True, MatchWholeWord:=False 'Busca el punto "." de los SAGs
                        wrdDoc.Bookmarks.Add Name:="GSAG2", Range:=myRange2
                        myRange2.SetRange Start:=myRange2.Start + 5, End:=myRange2.End + 2
                        wrdDoc.Bookmarks.Add Name:="GSAG2", Range:=myRange2
                        wrdDoc.Bookmarks("GSAG2").Range.Text = "" 'Borra ese punto "."
                        MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                        wrdDoc.Bookmarks.Add Name:="GSAG", Range:=MyRange
                    End If
                End If
            Next m
            MyRange.SetRange Start:=MyRange.Start - 3, End:=MyRange.End
            wrdDoc.Bookmarks.Add Name:="GSAG", Range:=MyRange
            wrdDoc.Bookmarks("GSAG").Range.Text = "" 'Borra el sobrante de los SAGs (" / $")
        End With

        'Porcentaje SAGs
        With Sheets("SAG")
            Set FindRowSAG = .Range("C:C").Find(What:=rEXP, LookIn:=xlValues) 'Busca la fila del NºExp y SAG que necesitamos
            FindRowNumberSAG = FindRowSAG.Row 'Busca el número de la fila del Expediente y SAG
            If .Cells(FindRowNumberSAG, 2) = "C" Then
                Set FindRowSAG = .Range("C:C").Find(What:=rEXP, After:=.Cells(FindRowNumberSAG, 3), LookIn:=xlValues) 'Busca la fila del NºExp y SAG que necesitamos
                FindRowNumberSAG = FindRowSAG.Row 'Busca el número de la fila del Expediente y SAG
            End If
            Set MyRange = wrdDoc.Content
            LastCol = .Cells(FindRowNumberSAG, .Columns.Count).End(xlToLeft).Column 'Encontrará la última columna usada en la fila 1
            For n = 14 To LastCol - 1 'El -1 es para que no coja la última columna donde vienen los 100%
                If Not .Cells(FindRowNumberSAG, n).Value = 0 Then
                    wrdDoc.Bookmarks("Porcen").Range.Text = FormatPercent(.Cells(FindRowNumberSAG, n).Value, 2, , , vbTrue) & " / $" 'Introducimos el "$" para poder crear un marcador y seguir uniendo porcentajes
                    MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                    wrdDoc.Bookmarks.Add Name:="Porcen", Range:=MyRange
                End If
            Next n
            MyRange.SetRange Start:=MyRange.Start - 3, End:=MyRange.End
            wrdDoc.Bookmarks.Add Name:="Porcen", Range:=MyRange
            wrdDoc.Bookmarks("Porcen").Range.Text = "" 'Borra el sobrante de los porcentajes de los SAGs (" / $")
        End With

        'Presupuesto Base de Licitación (PBL)
        wrdDoc.Bookmarks("PBL").Range.Text = PBL
        'Prórroga (igual al PBL)
        wrdDoc.Bookmarks("Prorr").Range.Text = Prorroga
        'Modificación (20% del PBL)
        wrdDoc.Bookmarks("Mod").Range.Text = Modif

        'Tabla Importe Anualidades
        Set myTable = wrdDoc.Tables(6)
        For o = 1 To NumAnual - 1
            myTable.Rows.Add 'Añade fila por anualidad
        Next o
        With Sheets("PAAD")
            If NumAnual = 1 Then
                myTable.Cell(3, 1).Range.Text = Right(.Cells(2, 22).Value, 4) 'Año anualidad primera
                myTable.Cell(3, 2).Range.Text = FormatCurrency(.Cells(FindRowNumber, 22).Value, 0, , , vbTrue) 'Dinero anualidad
            Else
                For p = 1 To NumAnual - 1
                    myTable.Cell(3, 1).Range.Text = Right(.Cells(2, 22).Value, 4) 'Año anualidad primera
                    myTable.Cell(3 + p, 1).Range.Text = Right(.Cells(2, 22 + p).Value, 4) 'Demás años anualidades
                    myTable.Cell(3, 2).Range.Text = FormatCurrency(.Cells(FindRowNumber, 22).Value, 0, , , vbTrue) 'Dinero anualidad
                    myTable.Cell(3 + p, 2).Range.Text = FormatCurrency(.Cells(FindRowNumber, 22 + p).Value, 0, , , vbTrue) 'Demás dineros anualidades
                Next p
            End If
        End With

        'Tabla datos de financiación -sección 14-
        Set myTable = wrdDoc.Tables(7)
        For q = 1 To NumAnual - 1
            myTable.Rows.Add 'Añade fila por anualidad
        Next q
        With Sheets("PAAD")
            If NumAnual = 1 Then
                myTable.Cell(3, 1).Range.Text = Right(.Cells(2, 22).Value, 4) 'Año anualidad primera
                myTable.Cell(3, 2).Range.Text = AppPres 'Aplicación presupuestaria primer año
                myTable.Cell(3, 3).Range.Text = FormatCurrency(.Cells(FindRowNumber, 22).Value, 0, , , vbTrue) 'Dinero anualidad
            Else
                For r = 1 To NumAnual - 1
                    myTable.Cell(3, 1).Range.Text = Right(.Cells(2, 22).Value, 4) 'Año anualidad primera
                    myTable.Cell(3 + r, 1).Range.Text = Right(.Cells(2, 22 + r).Value, 4) 'Demás años anualidades
                    myTable.Cell(3, 2).Range.Text = AppPres 'Aplicación presupuestaria primer año
                    myTable.Cell(3 + r, 2).Range.Text = AppPres 'Aplicación presupuestaria demás años
                    myTable.Cell(3, 3).Range.Text = FormatCurrency(.Cells(FindRowNumber, 22).Value, 0, , , vbTrue) 'Dinero anualidad
                    myTable.Cell(3 + r, 3).Range.Text = FormatCurrency(.Cells(FindRowNumber, 22 + r).Value, 0, , , vbTrue) 'Demás dineros anualidades
                Next r
            End If
        End With

        'Casillas marcadas (CheckBox)
        'NCP, NSP y DAM
        If OptionButtonDAM.Value = True Then 'DAM
            Set objCC = wrdDoc.SelectContentControlsByTag("NSP").Item(1)
            objCC.Checked = True
        ElseIf OptionButtonNCP.Value = True Then 'NCP
            Set objCC = wrdDoc.SelectContentControlsByTag("NCP").Item(1)
            objCC.Checked = True
        ElseIf OptionButtonNSP.Value = True Then 'NSP
            Set objCC = wrdDoc.SelectContentControlsByTag("NSP").Item(1)
            objCC.Checked = True
        End If
        'LCSP y LCSPDS
        If ComboBoxLey.Text = "Ley 9-2017" Then 'LCSP
            Set objCC = wrdDoc.SelectContentControlsByTag("LCSP").Item(1)
            objCC.Checked = True
        ElseIf ComboBoxLey.Text = "LCSPDS" Then 'LCSPDS
            Set objCC = wrdDoc.SelectContentControlsByTag("LCSPDS").Item(1)
            objCC.Checked = True
        End If
        'Lotes o no lotes
        If CheckBoxLotes.Value = True Then 'Lotes
            Set objCC = wrdDoc.SelectContentControlsByTag("LOTES").Item(1)
            objCC.Checked = True
        Else 'No lotes
            Set objCC = wrdDoc.SelectContentControlsByTag("TOTAL").Item(1)
            objCC.Checked = True
        End If

        'Justificación del NSP
        If OptionButtonDAM.Value = True Then 'DAM
            If EmpresaAM = "Otra empresa" Then
                wrdDoc.Bookmarks("JustNSP").Range.Text = _
                "El presente contrato derivado es consecuencia del Acuerdo Marco 20" & AcuerdoM & " " & LoteAM & " adjudicado a la empresa (RELLENAR EMPRESA)."
            Else
                wrdDoc.Bookmarks("JustNSP").Range.Text = _
                "El presente contrato derivado es consecuencia del Acuerdo Marco 20" & AcuerdoM & " " & LoteAM & " adjudicado a la empresa " & EmpresaAM & "."
            End If
        ElseIf OptionButtonNCP.Value = True Then 'NCP
            wrdDoc.Bookmarks("JustNSP").Range.Delete
        ElseIf OptionButtonNSP.Value = True Then 'NSP
            wrdDoc.Bookmarks("JustNSP").Range.Text = _
                "La contratación con el fabricante y responsable de diseño de los repuestos de célula de los aviones garantiza el suministro de la totalidad del material objeto de este expediente, la incorporación de las últimas mejoras en el material, la asistencia postventa que se precise y el aseguramiento de la calidad en la fabricación. En consecuencia se propone como sistema de adjudicación el Negociado sin Publicidad."
        End If

        'Empresa para el DAM o NSP
        Set MyRange = wrdDoc.Content
        If OptionButtonDAM.Value = True Then 'DAM
            wrdDoc.Bookmarks("Empresa").Range.Text = EmpresaAM
            MyRange.Find.Execute FindText:=EmpresaAM, MatchCase:=True, MatchWholeWord:=True, Forward:=False 'Busca el nombre de la Empresa desde el final
            If EmpresaAM = "Otra empresa" Then
                MyRange.Font.ColorIndex = wdRed 'Pone en rojo el nombre de la Empresa para que lo sustituyamos
            End If
            MyRange.ListFormat.ApplyBulletDefault 'Pone un Bullet Point al nombre de la Empresa
        ElseIf OptionButtonNCP.Value = True Then 'NCP
            wrdDoc.Bookmarks("Empresa").Range.Delete
        ElseIf OptionButtonNSP.Value = True Then 'NSP
            wrdDoc.Bookmarks("Empresa").Range.Text = EmpresaAM
            MyRange.Find.Execute FindText:=EmpresaAM, MatchCase:=True, MatchWholeWord:=True, Forward:=False 'Busca el nombre de la Empresa desde el final
            If EmpresaAM = "Otra empresa" Then
                MyRange.Font.ColorIndex = wdRed 'Pone en rojo el nombre de la Empresa para que lo sustituyamos
            End If
            MyRange.ListFormat.ApplyBulletDefault 'Pone un Bullet Point al nombre de la Empresa
        End If

        'Razones del NSP
        If OptionButtonDAM.Value = True Then 'DAM
            wrdDoc.Bookmarks("RazonesNSP").Range.Delete
        ElseIf OptionButtonNCP.Value = True Then 'NCP
            wrdDoc.Bookmarks("RazonesNSP").Range.Delete
        ElseIf OptionButtonNSP.Value = True Then 'NSP
            wrdDoc.Bookmarks("RazonesNSP").Range.Text = "Las expuestas en la Memoria Técnica."
        End If

        'Lugar de entrega y organismo receptor (Maestranzas)
        With Sheets("RECs")
            Set FindRowREC = .Range("A:A").Find(What:=rEXP, LookIn:=xlValues) 'Busca la fila del NºExp y REC que necesitamos
            FindRowNumberREC = FindRowREC.Row 'Busca el número de la fila del Expediente y REC
            If .Cells(FindRowNumberREC, 3) = "C" Then
                Set FindRowREC = .Range("A:A").Find(What:=rEXP, After:=.Cells(FindRowNumberREC, 3), LookIn:=xlValues) 'Busca la fila del NºExp y REC que necesitamos
                FindRowNumberREC = FindRowREC.Row 'Busca el número de la fila del Expediente y REC
            End If
            If .Cells(FindRowNumberREC, 9).Value = "MAESMA" Then
                wrdDoc.Bookmarks("Maestranza").Range.Text = "Madrid"
                wrdDoc.Bookmarks("Maestranza2").Range.Text = "Madrid"
            ElseIf .Cells(FindRowNumberREC, 9).Value = "MAESAL" Then
                wrdDoc.Bookmarks("Maestranza").Range.Text = "Albacete"
                wrdDoc.Bookmarks("Maestranza2").Range.Text = "Albacete"
            ElseIf .Cells(FindRowNumberREC, 9).Value = "MAESE" Then
                wrdDoc.Bookmarks("Maestranza").Range.Text = "Sevilla"
                wrdDoc.Bookmarks("Maestranza2").Range.Text = "Sevilla"
            ElseIf .Cells(FindRowNumberREC, 9).Value = "SEGEF" Then
                'Puede variar
                wrdDoc.Bookmarks("Maestranza").Delete
                wrdDoc.Bookmarks("Maestranza2").Delete
            ElseIf .Cells(FindRowNumberREC, 9).Value = "CLOTRA" Then
                'Puede variar
                wrdDoc.Bookmarks("Maestranza").Delete
                wrdDoc.Bookmarks("Maestranza2").Delete
            ElseIf .Cells(FindRowNumberREC, 9).Value = "SEMOT" Then
                'Puede variar
                wrdDoc.Bookmarks("Maestranza").Delete
                wrdDoc.Bookmarks("Maestranza2").Delete
            Else
                'Puede variar
                wrdDoc.Bookmarks("Maestranza").Delete
                wrdDoc.Bookmarks("Maestranza2").Delete
            End If

            'Responsable del Contrato (REC)
            If .Cells(FindRowNumberREC, 8).Value = "" Then
                wrdDoc.Bookmarks("REC").Delete
            Else
                str = .Cells(FindRowNumberREC, 8).Value 'Celda donde se encuentra el nombre del REC
                openPos = InStr(str, ".") 'Punto "." después del rango y antes del nombre del REC
                closePos = InStr(str, "(") 'Paréntesis "(" después del nombre del REC y antes de la Maestranza
                If closePos = 0 Then
                    wrdDoc.Bookmarks("REC").Delete
                ElseIf openPos = 0 Then
                    wrdDoc.Bookmarks("REC").Delete
                Else
                    midBit = Mid(str, openPos + 1, closePos - openPos - 1) 'El nombre del REC
                    wrdDoc.Bookmarks("REC").Range.Text = midBit
                End If
            End If


            'Graduación (Rango) del REC
            If .Cells(FindRowNumberREC, 8).Value = "" Then
                wrdDoc.Bookmarks("Graduacion").Delete
            Else
                RangStr = .Cells(FindRowNumberREC, 8).Value 'Celda donde se encuentra el rango del REC
                PosFin = InStr(RangStr, ".") 'Punto "." después del rango del REC
                If PosFin = 0 Then
                    wrdDoc.Bookmarks("Graduacion").Delete
                Else
                    Rang = Mid(RangStr, 1, PosFin - 1) 'El rango del REC
                    If Rang = "Col" Or Rang = "Co" Then 'Coronel
                        wrdDoc.Bookmarks("Graduacion").Range.Text = "CORONEL"
                    ElseIf Rang = "Tcol" Or Rang = "Tco" Then 'Teniente Coronel
                        wrdDoc.Bookmarks("Graduacion").Range.Text = "TENIENTE CORONEL"
                    ElseIf Rang = "Cte" Then 'Comandante
                        wrdDoc.Bookmarks("Graduacion").Range.Text = "COMANDANTE"
                    ElseIf Rang = "Cap" Then 'Capitán
                        wrdDoc.Bookmarks("Graduacion").Range.Text = "CAPITÁN"
                    ElseIf Rang = "Tte" Then 'Teniente
                        wrdDoc.Bookmarks("Graduacion").Range.Text = "TENIENTE"
                    ElseIf Rang = "Alfz" Then 'Alférez
                        wrdDoc.Bookmarks("Graduacion").Range.Text = "ALFÉREZ"
                    Else
                        wrdDoc.Bookmarks("Graduacion").Delete
                    End If
                End If
            End If

            'DestinoREC
            If .Cells(FindRowNumberREC, 10).Value = "" Then
                wrdDoc.Bookmarks("DestinoREC").Delete
            Else
                If .Cells(FindRowNumberREC, 9).Value = "CLOTRA" Then
                    'Puede variar
                ElseIf .Cells(FindRowNumberREC, 9).Value = "SEMOT" Then
                    'Puede variar
                Else
                    wrdDoc.Bookmarks("DestinoREC").Range.Text = .Cells(FindRowNumberREC, 10).Value
                End If
            End If
        End With

        'Coronel o Comandante Acctal.
        wrdDoc.Bookmarks("Coronel").Range.Text = Coronel

        'Coronel Jefe OAD
        wrdDoc.Bookmarks("CoronelJef").Range.Text = CoronelJef

        'Guardamos el documento en Word y PDF
        wrdDoc.SaveAs (sFolderPathForSave & "\" & rEXP & " PDC.docx") 'Guarda el Word
        'wrdDoc.ExportAsFixedFormat OutputFileName:=sFolderPathForSave & "\" & _
            'rEXP & " PDC.pdf", ExportFormat:=wdExportFormatPDF 'Guarda el PDF

        'Cerramos el Word
        wrdApp.Quit
        Set wrdApp = Nothing

    Else
        'Continúa con el código
    End If
End Sub

Sub Presupuesto()
    'PRESUPUESTO
    If CheckBoxPresup.Value = True Then

        strPath = sFolderPathForLoad & "\PLANTILLA PRESUPUESTO.dotx"
        'Abrimos la plantilla
        Set wrdApp = CreateObject("Word.Application")
        Set wrdDoc = wrdApp.Documents.Add(Template:=strPath, NewTemplate:=False, DocumentType:=0) 'Abrir Plantilla en ruta
        'wrdApp.Visible = True

        'Rango por si fuera necesario
        Set MyRange = wrdDoc.Content

        'Rellenamos el Número de Expediente
        wrdDoc.Bookmarks("NumExp").Range.Text = rEXP 'Cogemos el nºexp que se ha tecleado

        'Rellenamos el Número AM
        If OptionButtonDAM.Value = True Then
            wrdDoc.Bookmarks("NumAM").Range.Text = "20" & AcuerdoM
        Else 'NCP y NSP no tienen AM
            Set MyRange = wrdDoc.Bookmarks("NumAM").Range 'Borra el marcador y el espacio que dejaría moviendo su rango
                With MyRange
                    .MoveStart Unit:=wdCharacter, Count:=-3
                    .Select
                End With
                MyRange = ""
        End If

        'Rellenamos el LOTE AM
        If OptionButtonDAM.Value = True Then
            If Not LoteAM = "" Then
                wrdDoc.Bookmarks("LoteAM").Range.Text = LoteAM 'Lote AM
            Else 'Resto de AM no tiene LOTES
                Set MyRange = wrdDoc.Bookmarks("LoteAM").Range 'Borra el marcador y el espacio que dejaría moviendo su rango
                With MyRange
                    .MoveStart Unit:=wdCharacter, Count:=-1
                    .Select
                End With
                MyRange = ""
            End If
        Else 'NCP y NSP no tienen LOTES
            Set MyRange = wrdDoc.Bookmarks("LoteAM").Range 'Borra el marcador y el espacio que dejaría moviendo su rango
            With MyRange
                .MoveStart Unit:=wdCharacter, Count:=-1
                .Select
            End With
            MyRange = ""
        End If

        'Título del Expediente
        wrdDoc.Bookmarks("Titulo").Range.Text = Titulo 'Cogemos el título ya calculado previamente

        'Presupuesto Base de Licitación (PBL)
        wrdDoc.Bookmarks("PBL").Range.Text = PBL
        wrdDoc.Bookmarks("PBL2").Range.Text = PBL
        wrdDoc.Bookmarks("PBLTot").Range.Text = PBL
        wrdDoc.Bookmarks("PBLIVA").Range.Text = PBL
        'Prórroga (igual al PBL)
        wrdDoc.Bookmarks("Prorr").Range.Text = Prorroga
        wrdDoc.Bookmarks("ProrrTot").Range.Text = Prorroga
        'Modificación (20% del PBL)
        wrdDoc.Bookmarks("Mod").Range.Text = Modif
        wrdDoc.Bookmarks("ModTot").Range.Text = Modif

        'Cálculo del Valor Estimado del Contrato
        If OptionButtonDAM.Value = True Then 'DAM (No hay Prórrogas ni Modificaciones)
            wrdDoc.Bookmarks("ValEst").Range.Text = FormatCurrency((PBLSinForm + ProrrogaSinForm + ModifSinForm), 0, , , vbTrue) 'Pone formato € sin decimales
            wrdDoc.Bookmarks("ValEstTot").Range.Text = FormatCurrency((PBLSinForm + ProrrogaSinForm + ModifSinForm), 0, , , vbTrue) 'Pone formato € sin decimales
        ElseIf OptionButtonNCP.Value = True Then 'NCP
            wrdDoc.Bookmarks("ValEst").Range.Text = FormatCurrency((PBLSinForm + ProrrogaSinForm + ModifSinForm), 0, , , vbTrue) 'Pone formato € sin decimales
            wrdDoc.Bookmarks("ValEstTot").Range.Text = FormatCurrency((PBLSinForm + ProrrogaSinForm + ModifSinForm), 0, , , vbTrue) 'Pone formato € sin decimales
        ElseIf OptionButtonNSP.Value = True Then 'NSP
            wrdDoc.Bookmarks("ValEst").Range.Text = FormatCurrency((PBLSinForm + ProrrogaSinForm + ModifSinForm), 0, , , vbTrue) 'Pone formato € sin decimales
            wrdDoc.Bookmarks("ValEstTot").Range.Text = FormatCurrency((PBLSinForm + ProrrogaSinForm + ModifSinForm), 0, , , vbTrue) 'Pone formato € sin decimales
        End If

        'Tabla del Análisis del Valor Estimado del Contrato
        Set myTable = wrdDoc.Tables(3)
        If ComboBoxLey.Text = "Ley 9-2017" Then 'Ley 9-2017
            myTable.Cell(2, 2).Range.Text = PBL 'PBL
            myTable.Cell(2, 3).Range.Text = FormatCurrency((PBLSinForm + ProrrogaSinForm + ModifSinForm), 0, , , vbTrue) 'Valor Estimado
            myTable.Cell(2, 4).Range.Text = "100 %" 'Porcentajes (PBLtipo / PBLtodoelcontrato)
        ElseIf ComboBoxLey.Text = "LCSPDS" Then 'LCSPDS
            myTable.Cell(3, 2).Range.Text = PBL 'PBL
            myTable.Cell(3, 3).Range.Text = FormatCurrency((PBLSinForm + ProrrogaSinForm + ModifSinForm), 0, , , vbTrue) 'Valor Estimado
            myTable.Cell(3, 4).Range.Text = "100 %" 'Porcentajes (PBLtipo / PBLtodoelcontrato)
        End If
        myTable.Cell(5, 2).Range.Text = PBL 'PBL
        myTable.Cell(5, 3).Range.Text = FormatCurrency((PBLSinForm + ProrrogaSinForm + ModifSinForm), 0, , , vbTrue) 'Valor Estimado
        myTable.Cell(5, 4).Range.Text = "100 %" 'Porcentajes (PBLtipo / PBLtodoelcontrato)

        'Tabla del Análisis del Valor Estimado del Contrato 2
        Set myTable = wrdDoc.Tables(4)
        myTable.Cell(3, 2).Range.Text = PBL 'PBL Suministro
        myTable.Cell(3, 3).Range.Text = FormatCurrency((PBLSinForm + ProrrogaSinForm + ModifSinForm), 0, , , vbTrue) 'Valor Estimado Suministro
        myTable.Cell(3, 4).Range.Text = "100 %" 'Porcentajes (PBLtipo / PBLtodoelcontrato) Suministro
        myTable.Cell(7, 2).Range.Text = PBL 'PBL
        myTable.Cell(7, 3).Range.Text = FormatCurrency((PBLSinForm + ProrrogaSinForm + ModifSinForm), 0, , , vbTrue) 'Valor Estimado
        myTable.Cell(7, 4).Range.Text = "100 %" 'Porcentajes (PBLtipo / PBLtodoelcontrato)

        'Distribución Temporal del Gasto Contractual
        Set myTable = wrdDoc.Tables(5)
        For t = 1 To NumAnual - 1
            myTable.Rows.Add 'Añade fila por anualidad
        Next t
        With Sheets("PAAD")
            If NumAnual = 1 Then
                myTable.Cell(2, 1).Range.Text = Right(.Cells(2, 22).Value, 4) 'Año anualidad primera
                myTable.Cell(2, 2).Range.Text = FormatCurrency(.Cells(FindRowNumber, 22).Value, 0, , , vbTrue) 'Dinero anualidad primera
                myTable.Cell(2, 4).Range.Text = Prorroga 'Dinero prórroga
                myTable.Cell(2, 5).Range.Text = Modif 'Dinero modificación
                myTable.Cell(2, 8).Range.Text = FormatCurrency((.Cells(FindRowNumber, 22).Value + ProrrogaSinForm + ModifSinForm), 0, , , vbTrue) 'Dinero total
            ElseIf NumAnual = 2 Then
                myTable.Cell(2, 1).Range.Text = Right(.Cells(2, 22).Value, 4) 'Año anualidad primera
                myTable.Cell(3, 1).Range.Text = Right(.Cells(2, 23).Value, 4) 'Año anualidad segunda
                myTable.Cell(2, 2).Range.Text = FormatCurrency(.Cells(FindRowNumber, 22).Value, 0, , , vbTrue) 'Dinero anualidad primera
                myTable.Cell(3, 2).Range.Text = FormatCurrency(.Cells(FindRowNumber, 23).Value, 0, , , vbTrue) 'Dinero anualidad segunda
                myTable.Cell(2, 4).Range.Text = FormatCurrency(0, 0, , , vbTrue) 'Dinero prórroga primera anualidad
                myTable.Cell(3, 4).Range.Text = Prorroga 'Dinero prórroga segunda anualidad
                myTable.Cell(2, 5).Range.Text = FormatCurrency(0, 0, , , vbTrue) 'Dinero modificación primera anualidad
                myTable.Cell(3, 5).Range.Text = Modif 'Dinero modificación segunda anualidad
                myTable.Cell(2, 8).Range.Text = FormatCurrency(.Cells(FindRowNumber, 22).Value, 0, , , vbTrue) 'Dinero total anualidad primera
                myTable.Cell(3, 8).Range.Text = FormatCurrency((.Cells(FindRowNumber, 23).Value + ProrrogaSinForm + ModifSinForm), 0, , , vbTrue) 'Dinero total anualidad segunda
            Else
                myTable.Cell(2, 1).Range.Text = Right(.Cells(2, 22).Value, 4) 'Año anualidad primera
                myTable.Cell(3, 1).Range.Text = Right(.Cells(2, 23).Value, 4) 'Año anualidad segunda
                myTable.Cell(4, 1).Range.Text = Right(.Cells(2, 24).Value, 4) 'Año anualidad tercera
                myTable.Cell(2, 2).Range.Text = FormatCurrency(.Cells(FindRowNumber, 22).Value, 0, , , vbTrue) 'Dinero anualidad primera
                myTable.Cell(3, 2).Range.Text = FormatCurrency(.Cells(FindRowNumber, 23).Value, 0, , , vbTrue) 'Dinero anualidad segunda
                myTable.Cell(4, 2).Range.Text = FormatCurrency(.Cells(FindRowNumber, 24).Value, 0, , , vbTrue) 'Dinero anualidad tercera
                myTable.Cell(2, 4).Range.Text = FormatCurrency(0, 0, , , vbTrue) 'Dinero prórroga primera anualidad
                myTable.Cell(3, 4).Range.Text = FormatCurrency(0, 0, , , vbTrue) 'Dinero prórroga segunda anualidad
                myTable.Cell(4, 4).Range.Text = Prorroga 'Dinero prórroga tercera anualidad
                myTable.Cell(2, 5).Range.Text = FormatCurrency(0, 0, , , vbTrue) 'Dinero modificación primera anualidad
                myTable.Cell(3, 5).Range.Text = FormatCurrency((ModifSinForm / (NumAnual - 1)), 0, , , vbTrue) 'Dinero modificación segunda anualidad
                myTable.Cell(4, 5).Range.Text = FormatCurrency((ModifSinForm / (NumAnual - 1)), 0, , , vbTrue) 'Dinero modificación tercera anualidad
                myTable.Cell(2, 8).Range.Text = FormatCurrency(.Cells(FindRowNumber, 22).Value, 0, , , vbTrue) 'Dinero total anualidad primera
                myTable.Cell(3, 8).Range.Text = FormatCurrency((.Cells(FindRowNumber, 23).Value + (ModifSinForm / (NumAnual - 1))), 0, , , vbTrue) 'Dinero total anualidad segunda
                myTable.Cell(4, 8).Range.Text = FormatCurrency((.Cells(FindRowNumber, 24).Value + ProrrogaSinForm + (ModifSinForm / (NumAnual - 1))), 0, , , vbTrue) 'Dinero total anualidad tercera
            End If
        End With

        'Párrafo final
        If OptionButtonDAM.Value = True Then 'DAM
            wrdDoc.Bookmarks("Final").Range.Text = "Se anexa al presente documento el anexo de detalle del presupuesto que presenta una relación valorada de cada artículo que se estima adquirir, determinada conforme los precios de adjudicación del Acuerdo Marco del que deriva el presente contrato."
        ElseIf OptionButtonNCP.Value = True Then 'NCP
            wrdDoc.Bookmarks("Final").Range.Text = "Se adjunta al presente documento el detalle del presupuesto que presenta una relación valorada de cada artículo que se estima adquirir, determinada atendiendo al precio general de mercado en función de cotizaciones de expedientes predecesores y ofertas históricas de diferentes licitadores."
        ElseIf OptionButtonNSP.Value = True Then 'NSP
            wrdDoc.Bookmarks("Final").Range.Text = "Se adjunta al presente documento el detalle del presupuesto que presenta una relación valorada de cada artículo que se estima adquirir, determinada atendiendo al precio general de mercado en función de cotizaciones de expedientes predecesores y ofertas históricas de diferentes licitaciones."
        End If

        'Coronel o Comandante Acctal.
        wrdDoc.Bookmarks("Coronel").Range.Text = Coronel

        'Guardamos el documento en Word y PDF
        wrdDoc.SaveAs (sFolderPathForSave & "\" & rEXP & " PRESUPUESTO.docx") 'Guarda el Word
        'wrdDoc.ExportAsFixedFormat OutputFileName:=sFolderPathForSave & "\" & _
            'rEXP & " PRESUPUESTO.pdf", ExportFormat:=wdExportFormatPDF 'Guarda el PDF

        'Cerramos el Word
        wrdApp.Quit
        Set wrdApp = Nothing

    Else
        'Continúa con el código
    End If
End Sub

Sub PropuestaInclusionPCAP()
    'PROPUESTA INCLUSION PCAP
    If CheckBoxPCAP.Value = True Then

        'Elegimos si es NCP o NSP (DAM no tiene Propuestas al PCAP)
        If OptionButtonDAM.Value = True Then 'DAM no tiene propuestas al PCAP
            strPath = sFolderPathForLoad & "\PLANTILLA PROPUESTA INCLUSION PCAP NSP.dotx"
        ElseIf OptionButtonNCP.Value = True Then 'NCP
            strPath = sFolderPathForLoad & "\PLANTILLA PROPUESTA INCLUSION PCAP NCP.dotx"
        ElseIf OptionButtonNSP.Value = True Then 'NSP
            strPath = sFolderPathForLoad & "\PLANTILLA PROPUESTA INCLUSION PCAP NSP.dotx"
        End If

        'Abrimos la plantilla
        Set wrdApp = CreateObject("Word.Application")
        Set wrdDoc = wrdApp.Documents.Add(Template:=strPath, NewTemplate:=False, DocumentType:=0) 'Abrir Plantilla en ruta
        'wrdApp.Visible = True

        'Rango por si fuera necesario
        Set MyRange = wrdDoc.Content

        'Rellenamos el Número de Expediente
        wrdDoc.Bookmarks("NumExp").Range.Text = rEXP 'Cogemos el nºexp que se ha tecleado

        'Título del Expediente
        wrdDoc.Bookmarks("Titulo").Range.Text = Titulo 'Cogemos el título ya calculado previamente

        'Tipo de Negocio Jurídico
        wrdDoc.Bookmarks("TipExp").Range.Text = "Contrato administrativo de suministro" 'Cuando es NCP o NSP

        'Elegimos si es Ley 9-2017 o LCSDPS
        If ComboBoxLey.Text = "Ley 9-2017" Then 'Ley 9-2017
            wrdDoc.Bookmarks("Ley").Range.Delete
            If OptionButtonNCP.Value = True Then 'NCP
                wrdDoc.Bookmarks("Ley2").Range.Text = "artículo 166 de la LCSP," 'Sólo en NCP
            End If
        ElseIf ComboBoxLey.Text = "LCSPDS" Then 'LCSPDS
            If OptionButtonNCP.Value = True Then 'NCP
                wrdDoc.Bookmarks("Ley").Range.Text = "Así mismo, se considera dentro del ámbito de aplicación objetivo de la Ley de contratos del sector público en los ámbitos de la defensa y la seguridad (LCSPDS) según el artículo 2 de la citada norma legal." & vbCrLf
                wrdDoc.Paragraphs(9).OutlineLevel = wdOutlineLevelBodyText 'Texto independiente (depende del apartado 1)
                wrdDoc.Paragraphs(9).Range.ListFormat.RemoveNumbers 'Elimina la numeración del párrafo
                'wrdDoc.Paragraphs(9).LeftIndent = CentimetersToPoints(0.63) 'Sangría izquierda a 0.63 cm
                wrdDoc.Paragraphs(9).LeftIndent = 17.8605 'Sangría izquierda a 0.63 cm
                wrdDoc.Paragraphs(9).SpaceBefore = 6 'Espacio antes del párrafo de 6 cm
                wrdDoc.Paragraphs(9).SpaceAfter = 6 'Espacio después del párrafo de 6 cm
                wrdDoc.Bookmarks("Ley2").Range.Text = "artículo 43 de la LCSPDS," 'Sólo en NCP
            Else 'NSP
                wrdDoc.Bookmarks("Ley").Range.Text = "Así mismo, se considera dentro del ámbito de aplicación objetivo de la Ley de contratos del sector público en los ámbitos de la defensa y la seguridad (LCSPDS) según el artículo 2 de la citada norma legal." & vbCrLf
                wrdDoc.Paragraphs(8).OutlineLevel = wdOutlineLevelBodyText 'Texto independiente (depende del apartado 1)
                wrdDoc.Paragraphs(8).Range.ListFormat.RemoveNumbers 'Elimina la numeración del párrafo
                wrdDoc.Paragraphs(8).LeftIndent = 17.8605 'Sangría izquierda a 0.63 cm
                wrdDoc.Paragraphs(8).SpaceBefore = 6 'Espacio antes del párrafo de 6 cm
                wrdDoc.Paragraphs(8).SpaceAfter = 6 'Espacio después del párrafo de 6 cm
            End If
        End If
        
        'Repuestos a adquirir por el Expediente
        If Repuesto = "" Then
            wrdDoc.Bookmarks("Repuesto").Delete
        Else
            wrdDoc.Bookmarks("Repuesto").Range.Text = Repuesto
        End If
            
        'Sistemas de Armas y SAGs
        If countSAG = 1 Then 'Si sólo hay un SAG
            wrdDoc.Bookmarks("SistArm").Range.Text = " al " & SistArm
        Else 'Si hay más de 1 SAG
            wrdDoc.Bookmarks("SistArm").Range.Text = " a los " & SistArm
        End If
        With Sheets("SAG")
            If countSAG > 3 Then 'Si hay más de 3 SAGs con porcentaje mayor de 0, entonces el texto de nuestros marcadores es distinto
                wrdDoc.Bookmarks("SAG").Range.Text = "" 'Borra el SAG ya que viene incluido en el marcador anterior
            Else 'Si hay 3 o menos de 3 SAGs entonces el texto en los marcadores llevará escrito el nombre de cada SAG
                For l = 14 To LastColSAG - 1 'El -1 es para que no coja la última columna donde vienen los 100%
                    If Not .Cells(FindRowNumberSAG, l).Value = 0 Then 'Todas las aeronaves
                        If .Cells(FindRowNumberSAG, l).Address = NoCeroDir Then 'Si la celda del SAG es el penúltimo SAG, entonces el marcador será "y", en vez de una coma ","
                            Esp = " y $"
                        Else 'Si no es la penúltima, seguirá con una coma
                            Esp = ", $"
                        End If
                        If .Cells(1, l).Value = "AE.9" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "F-5 (A.9)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "C.15" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "F-18 HORNET (C.15)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "C.16" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "EUROFIGHTER TYPHOON 2000 (C.16)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "H.21" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "SUPERPUMA (HD/HT.21)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "H.24" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "SIKORSKY (H.24)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "H.25" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "COLIBRÍ (H.25)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "H.27" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "COUGAR (HT.27)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "H.29" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "NH-90 (H.29)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "E.24" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "BECHCRAFT 33C (E.24)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "E.25" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-101 (E.25)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "E.26" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "T-35 C TAMIZ (E.26)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "E.27" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "PILATUS PC-21 (E.27)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.11" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "FALCON 20 (T.11)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.12" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-212 (T.12)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.18" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "FALCON 900 (T.18)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.19" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-235 (T.19)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "TR.20" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "CESSNA CITATION (TR.20)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.21" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-295 (T.21)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.22" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "A310 (T.22)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "T.23" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "A400M (T.23)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "P.3" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "P-3 ORION" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "UD.13 / UD. 14" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "CL-215 (UD.13) y CL-415 (UD.14)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "U.22" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "BEECHCRAFT C 90 -KING AIR (U.22)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                        ElseIf .Cells(1, l).Value = "XLAN" Then
                            'No hacer nada
                        ElseIf .Cells(1, l).Value = "XEPS" Then
                            'No hacer nada
                        End If
                    End If
                Next l
                MyRange.SetRange Start:=MyRange.Start - 2, End:=MyRange.End
                wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                wrdDoc.Bookmarks("SAG").Range.Text = " " 'Borra el sobrante de los SAGs (", $")
            End If
        End With
        
        'Flota
        wrdDoc.Bookmarks("Flota").Range.Text = Flota

        'Lotes en el Expediente
        If CheckBoxLotes.Value = True Then 'Sí hay Lotes
            wrdDoc.Bookmarks("Lotes").Delete 'Borra el marcador y deja el párrafo
            wrdDoc.Bookmarks("NoLotes").Range = "" 'Borra el párrafo
            If OptionButtonNCP.Value = True Then 'NCP
                wrdDoc.Bookmarks("Lotes2").Delete 'Borra el marcador y deja los val. referencia de los lotes (Sólo en NCP)
            End If
        Else 'No hay lotes
            wrdDoc.Bookmarks("NoLotes").Delete 'Borra el marcador y deja el párrafo
            wrdDoc.Bookmarks("Lotes").Range = "" 'Borra el párrafo
            If OptionButtonNCP.Value = True Then 'NCP
                wrdDoc.Bookmarks("Lotes2").Range = "" 'Borra los val. referencia de los lotes (Sólo en NCP)
            End If
        End If

        'Tramitación Anticipada u Ordinaria
        If CheckBoxTrAnt.Value = True Then 'Anticipada
            wrdDoc.Bookmarks("Tramitacion").Range.Text = "ANTICIPADA"
        Else 'Ordinaria
            wrdDoc.Bookmarks("Tramitacion").Range.Text = "ORDINARIA"
        End If

        'Valor de Referencia (NCP)
        If OptionButtonNCP.Value = True Then 'NCP
            wrdDoc.Bookmarks("VR").Range.Text = FormatCurrency(ValRef, 0, , , vbTrue) 'Sólo en NCP
        End If

        'Fecha de finalización del expediente
        With Sheets("PAAD")
            If Not .Cells(FindRowNumber, 24).Value = 0 Then 'Tercera anualidad
                wrdDoc.Bookmarks("Fecha").Range.Text = "30 de noviembre de " & Right(.Cells(2, 24).Value, 4)
            ElseIf Not .Cells(FindRowNumber, 23).Value = 0 Then 'Segunda anualidad
                wrdDoc.Bookmarks("Fecha").Range.Text = "30 de noviembre de " & Right(.Cells(2, 23).Value, 4)
            Else 'Primera anualidad
                wrdDoc.Bookmarks("Fecha").Range.Text = "30 de noviembre de " & Right(.Cells(2, 22).Value, 4)
            End If
        End With

        'Coronel o Comandante Acctal.
        wrdDoc.Bookmarks("Coronel").Range.Text = Coronel

        'Coronel Jefe OAD
        wrdDoc.Bookmarks("CoronelJef").Range.Text = CoronelJef

        'Guardamos el documento en Word y PDF
        wrdDoc.SaveAs (sFolderPathForSave & "\" & rEXP & " PROPUESTA INCLUSION PCAP.docx") 'Guarda el Word
        'wrdDoc.ExportAsFixedFormat OutputFileName:=sFolderPathForSave & "\" & _
            'rEXP & " PROPUESTA INCLUSION PCAP.pdf", ExportFormat:=wdExportFormatPDF 'Guarda el PDF

        'Cerramos el Word
        wrdApp.Quit
        Set wrdApp = Nothing

    Else
        'Continúa con el código
    End If
End Sub

Sub PPT()
    'PLIEGO PRESCRIPCIONES TÉCNICAS (PPT)
    If CheckBoxPPT.Value = True Then
    
        'Elegimos si es NCP o NSP (DAM no tiene Propuestas al PCAP)
        If OptionButtonDAM.Value = True Then 'DAM no tiene propuestas al PCAP
            strPath = sFolderPathForLoad & "\PLANTILLA PPT DERIVADO AM.dotx"
        ElseIf OptionButtonNCP.Value = True Then 'NCP
            strPath = sFolderPathForLoad & "\PLANTILLA PPT REPUESTOS ABIERTO.dotx"
        ElseIf OptionButtonNSP.Value = True Then 'NSP
            strPath = sFolderPathForLoad & "\PLANTILLA PPT REPUESTOS ABIERTO.dotx"
        End If

        'Abrimos la plantilla
        Set wrdApp = CreateObject("Word.Application")
        Set wrdDoc = wrdApp.Documents.Add(Template:=strPath, NewTemplate:=False, DocumentType:=0) 'Abrir Plantilla en ruta
        'wrdApp.Visible = True

        'Rango por si fuera necesario
        Set MyRange = wrdDoc.Content
        Set myRange2 = wrdDoc.Content

        'Rellenamos el Número de Expediente
        wrdDoc.Bookmarks("NumExp").Range.Text = rEXP 'Cogemos el nºexp que se ha tecleado

        'Título del Expediente
        wrdDoc.Bookmarks("Titulo").Range.Text = Titulo 'Cogemos el título ya calculado previamente
        
        'Rellenamos el Número AM
        If OptionButtonDAM.Value = True Then 'DAM
            wrdDoc.Bookmarks("NumAM").Range.Text = AcuerdoM
            wrdDoc.Bookmarks("NumAM2").Range.Text = AcuerdoM
        End If

        'Rellenamos el LOTE AM
        If OptionButtonDAM.Value = True Then 'DAM
            If Not LoteAM = "" Then
                wrdDoc.Bookmarks("LoteAM").Range.Text = "(Lote " & LoteAM & ")." 'Lote AM
            Else 'Resto de AM no tiene LOTES
                Set MyRange = wrdDoc.Bookmarks("LoteAM").Range 'Borra el marcador y el espacio que dejaría moviendo su rango
                With MyRange
                    .MoveStart Unit:=wdCharacter, Count:=-1
                    .Select
                End With
                MyRange = ""
            End If
        End If
        
        'Repuestos a adquirir por el Expediente
        If Repuesto = "" Then 'Si se marcó "Otros repuestos"
            wrdDoc.Bookmarks("Repuesto").Delete
        Else
            wrdDoc.Bookmarks("Repuesto").Range.Text = Repuesto
        End If
        If OptionButtonDAM.Value = False Then 'No DAM
            If Repuesto = "" Then 'Si se marcó "Otros repuestos"
                wrdDoc.Bookmarks("Repuesto2").Delete
            Else
                wrdDoc.Bookmarks("Repuesto2").Range.Text = Repuesto
            End If
        End If
                
        'Sistemas de Armas y SAGs
        If countSAG = 1 Then 'Si sólo hay un SAG
            If OptionButtonDAM.Value = True Then 'DAM
                wrdDoc.Bookmarks("SistArm").Range.Text = " al " & SistArm
            Else 'NCP 'NSP
                wrdDoc.Bookmarks("SistArm").Range.Text = " al " & SistArm
                wrdDoc.Bookmarks("SistArm2").Range.Text = " al " & SistArm
                wrdDoc.Bookmarks("SistArm3").Range.Text = "dicho sistema"
            End If
        Else 'Si hay más de 1 SAG
            If OptionButtonDAM.Value = True Then 'DAM
                wrdDoc.Bookmarks("SistArm").Range.Text = " a los " & SistArm
            Else 'NCP 'NSP
                wrdDoc.Bookmarks("SistArm").Range.Text = " a los " & SistArm
                wrdDoc.Bookmarks("SistArm2").Range.Text = " a los " & SistArm
                wrdDoc.Bookmarks("SistArm3").Range.Text = "dichos sistemas"
            End If
        End If
        With Sheets("SAG")
            If countSAG > 3 Then 'Si hay más de 3 SAGs con porcentaje mayor de 0, entonces el texto de nuestros marcadores es distinto
                wrdDoc.Bookmarks("SAG").Range.Text = "" 'Borra el SAG ya que viene incluido en el marcador anterior
                If OptionButtonDAM.Value = False Then 'No DAM
                    wrdDoc.Bookmarks("SAG2").Range.Text = "" 'Borra el SAG ya que viene incluido en el marcador anterior
                End If
            Else 'Si hay 3 o menos de 3 SAGs entonces el texto en los marcadores llevará escrito el nombre de cada SAG
                For l = 14 To LastColSAG - 1 'El -1 es para que no coja la última columna donde vienen los 100%
                    If Not .Cells(FindRowNumberSAG, l).Value = 0 Then 'Todas las aeronaves
                        If .Cells(FindRowNumberSAG, l).Address = NoCeroDir Then 'Si la celda del SAG es el penúltimo SAG, entonces el marcador será "y", en vez de una coma ","
                            Esp = " y $"
                        Else 'Si no es la penúltima, seguirá con una coma
                            Esp = ", $"
                        End If
                        If .Cells(FindRowNumberSAG, l).Address = NoCeroDir Then 'Si la celda del SAG es el penúltimo SAG, entonces el marcador será "y", en vez de una coma ","
                            Esp2 = " y €"
                        Else 'Si no es la penúltima, seguirá con una coma
                            Esp2 = ", €"
                        End If
                        If .Cells(1, l).Value = "AE.9" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "F-5 (A.9)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "F-5 (A.9)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "C.15" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "F-18 HORNET (C.15)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "F-18 HORNET (C.15)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "C.16" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "EUROFIGHTER TYPHOON 2000 (C.16)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "EUROFIGHTER TYPHOON 2000 (C.16)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "H.21" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "SUPERPUMA (HD/HT.21)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "SUPERPUMA (HD/HT.21)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "H.24" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "SIKORSKY (H.24)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "SIKORSKY (H.24)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "H.25" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "COLIBRÍ (H.25)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "COLIBRÍ (H.25)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "H.27" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "COUGAR (HT.27)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "COUGAR (HT.27)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "H.29" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "NH-90 (H.29)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "NH-90 (H.29)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "E.24" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "BECHCRAFT 33C (E.24)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "BECHCRAFT 33C (E.24)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "E.25" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-101 (E.25)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "C-101 (E.25)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "E.26" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "T-35 C TAMIZ (E.26)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "T-35 C TAMIZ (E.26)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "E.27" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "PILATUS PC-21 (E.27)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "PILATUS PC-21 (E.27)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "T.11" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "FALCON 20 (T.11)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "FALCON 20 (T.11)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "T.12" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-212 (T.12)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "C-212 (T.12)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "T.18" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "FALCON 900 (T.18)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "FALCON 900 (T.18)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "T.19" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-235 (T.19)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "C-235 (T.19)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "TR.20" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "CESSNA CITATION (TR.20)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "CESSNA CITATION (TR.20)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "T.21" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "C-295 (T.21)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "C-295 (T.21)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                myRange2.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "T.22" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "A310 (T.22)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "A310 (T.22)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                MyRange.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "T.23" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "A400M (T.23)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "A400M (T.23)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                MyRange.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "P.3" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "P-3 ORION" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "P-3 ORION" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                MyRange.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "UD.13 / UD. 14" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "CL-215 (UD.13) y CL-415 (UD.14)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "CL-215 (UD.13) y CL-415 (UD.14)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                MyRange.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "U.22" Then
                            wrdDoc.Bookmarks("SAG").Range.Text = "BEECHCRAFT C 90 -KING AIR (U.22)" & Esp 'Introducimos el "$" para poder crear un marcador y seguir uniendo SAGs
                            MyRange.Find.Execute FindText:="$", MatchCase:=True, MatchWholeWord:=True
                            wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                            If OptionButtonDAM.Value = False Then 'No DAM
                                wrdDoc.Bookmarks("SAG2").Range.Text = "BEECHCRAFT C 90 -KING AIR (U.22)" & Esp2 'Introducimos el "€" para poder crear un marcador y seguir uniendo SAGs
                                MyRange.Find.Execute FindText:="€", MatchCase:=True, MatchWholeWord:=True
                                wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                            End If
                        ElseIf .Cells(1, l).Value = "XLAN" Then
                            'No hacer nada
                        ElseIf .Cells(1, l).Value = "XEPS" Then
                            'No hacer nada
                        End If
                    End If
                Next l
                MyRange.SetRange Start:=MyRange.Start - 2, End:=MyRange.End
                wrdDoc.Bookmarks.Add Name:="SAG", Range:=MyRange
                wrdDoc.Bookmarks("SAG").Range.Text = " " 'Borra el sobrante de los SAGs (", $")
                If OptionButtonDAM.Value = False Then 'No DAM
                    myRange2.SetRange Start:=myRange2.Start - 2, End:=myRange2.End
                    wrdDoc.Bookmarks.Add Name:="SAG2", Range:=myRange2
                    wrdDoc.Bookmarks("SAG2").Range.Text = " " 'Borra el sobrante de los SAGs (", €")
                End If
            End If
        End With
            
        'Flota
        If OptionButtonDAM.Value = True Then 'DAM
            'No hay flota
        Else 'NCP 'NSP
            wrdDoc.Bookmarks("Flota").Range.Text = Flota
            wrdDoc.Bookmarks("Flota2").Range.Text = Flota
        End If
        
        'Fecha de plazo de ejecución
        With Sheets("PAAD")
            If Not .Cells(FindRowNumber, 24).Value = 0 Then 'Tercera anualidad
                wrdDoc.Bookmarks("Fecha").Range.Text = "30/11/" & Right(.Cells(2, 24).Value, 4)
            ElseIf Not .Cells(FindRowNumber, 23).Value = 0 Then 'Segunda anualidad
                wrdDoc.Bookmarks("Fecha").Range.Text = "30/11/" & Right(.Cells(2, 23).Value, 4)
            Else 'Primera anualidad
                wrdDoc.Bookmarks("Fecha").Range.Text = "30/11/" & Right(.Cells(2, 22).Value, 4)
            End If
        End With
        
        'Fechas de entrega
        With Sheets("PAAD")
            wrdDoc.Bookmarks("Fecha2").Range.Text = "30/08/" & Right(.Cells(2, 22).Value, 4) 'Primera anualidad
            wrdDoc.Bookmarks("Fecha3").Range.Text = "15/12/" & Right(.Cells(2, 22).Value, 4) 'Primera anualidad fecha límite de entrega
        End With
        
        'Tabla Anualidades
        Set myTable = wrdDoc.Tables(1)
        For u = 1 To NumAnual - 1
            myTable.Rows.Add 'Añade fila por anualidad
        Next u
        With Sheets("PAAD")
            If NumAnual = 1 Then
                myTable.Cell(2, 1).Range.Text = Right(.Cells(2, 22).Value, 4) 'Año anualidad primera
                myTable.Cell(2, 2).Range.Text = "30/11/" & Right(.Cells(2, 22).Value, 4) 'Fecha límite ejecución
            Else
                For v = 1 To NumAnual - 1
                    myTable.Cell(2, 1).Range.Text = Right(.Cells(2, 22).Value, 4) 'Año anualidad primera
                    myTable.Cell(2 + v, 1).Range.Text = Right(.Cells(2, 22 + v).Value, 4) 'Demás años anualidades
                    myTable.Cell(2, 2).Range.Text = "30/11/" & Right(.Cells(2, 22).Value, 4) 'Fecha límite ejecución primera
                    myTable.Cell(2 + v, 2).Range.Text = "30/11/" & Right(.Cells(2, 22 + v).Value, 4) 'Demás fechas límite ejecución
                Next v
            End If
        End With
        
        'Lugar de entrega y destino final (Maestranzas)
        With Sheets("RECs")
            Set FindRowREC = .Range("A:A").Find(What:=rEXP, LookIn:=xlValues) 'Busca la fila del NºExp y REC que necesitamos
            FindRowNumberREC = FindRowREC.Row 'Busca el número de la fila del Expediente y REC
            If .Cells(FindRowNumberREC, 3) = "C" Then
                Set FindRowREC = .Range("A:A").Find(What:=rEXP, After:=.Cells(FindRowNumberREC, 3), LookIn:=xlValues) 'Busca la fila del NºExp y REC que necesitamos
                FindRowNumberREC = FindRowREC.Row 'Busca el número de la fila del Expediente y REC
            End If
            If .Cells(FindRowNumberREC, 9).Value = "MAESMA" Then
                wrdDoc.Bookmarks("Maestranza").Range.Text = "MADRID"
                wrdDoc.Bookmarks("Maestranza2").Range.Text = "MADRID"
            ElseIf .Cells(FindRowNumberREC, 9).Value = "MAESAL" Then
                wrdDoc.Bookmarks("Maestranza").Range.Text = "ALBACETE"
                wrdDoc.Bookmarks("Maestranza2").Range.Text = "ALBACETE"
            ElseIf .Cells(FindRowNumberREC, 9).Value = "MAESE" Then
                wrdDoc.Bookmarks("Maestranza").Range.Text = "SEVILLA"
                wrdDoc.Bookmarks("Maestranza2").Range.Text = "SEVILLA"
            ElseIf .Cells(FindRowNumberREC, 9).Value = "SEGEF" Then
                'Puede variar
                wrdDoc.Bookmarks("Maestranza").Delete
                wrdDoc.Bookmarks("Maestranza2").Delete
            ElseIf .Cells(FindRowNumberREC, 9).Value = "CLOTRA" Then
                'Puede variar
                wrdDoc.Bookmarks("Maestranza").Delete
                wrdDoc.Bookmarks("Maestranza2").Delete
            ElseIf .Cells(FindRowNumberREC, 9).Value = "SEMOT" Then
                'Puede variar
                wrdDoc.Bookmarks("Maestranza").Delete
                wrdDoc.Bookmarks("Maestranza2").Delete
            Else
                'Puede variar
                wrdDoc.Bookmarks("Maestranza").Delete
                wrdDoc.Bookmarks("Maestranza2").Delete
            End If
        End With
    
        'Guardamos el documento en Word y PDF
        wrdDoc.SaveAs (sFolderPathForSave & "\" & rEXP & " PPT.docx") 'Guarda el Word
        'wrdDoc.ExportAsFixedFormat OutputFileName:=sFolderPathForSave & "\" & _
            'rEXP & " PPT.pdf", ExportFormat:=wdExportFormatPDF 'Guarda el PDF

        'Cerramos el Word
        wrdApp.Quit
        Set wrdApp = Nothing

    Else
        'Continúa con el código
    End If
End Sub

Sub ArchivoLicitacionCARPETA()
    'ARCHIVO LICITACION (Carpeta)
    If CheckBoxArcLic.Value = True Then
        
        strPath = sFolderPathForSave & "\ARCHIVO LICITACION" 'Ruta donde guarda la carpeta ARCHIVO LICITACION
        MkDir strPath 'Crea la carpeta y la guarda
    
    Else
        'Continúa con el código
    End If
End Sub

Sub ModeloFacturaElectronicaCARPETA()
    'MODELO FACTURA ELECTRONICA (Carpeta y archivo)
    If CheckBoxFactElec.Value = True Then
        
        strPath = sFolderPathForSave & "\MODELO FACTURA ELECTRONICA" 'Ruta donde guarda la carpeta MODELO FACTURA ELECTRONICA
        MkDir strPath 'Crea la carpeta y la guarda
    
        'Crear archivo
        strPathSub = sFolderPathForLoad & "\XXXXXX MODELO FACTURA ELECTRONICA.xlsx" 'Ruta de la subcarpeta MODELO FACTURA ELECTRONICA
        FileCopy strPathSub, strPath & "\" & rEXP & " MODELO FACTURA ELECTRONICA.xlsx" 'Crea el archivo y lo guarda
    
    Else
        'Continúa con el código
    End If
End Sub

Sub ModeloSeguimientoPedidosCARPETA()
    'MODELO SEGUIMIENTO PEDIDOS (Carpeta y archivo)
    If CheckBoxSeguPed.Value = True Then
        
        strPath = sFolderPathForSave & "\MODELO SEGUIMIENTO PEDIDOS" 'Ruta donde guarda la carpeta MODELO SEGUIMIENTO PEDIDOS
        MkDir strPath 'Crea la carpeta y la guarda
    
        'Crear archivo
        strPathSub = sFolderPathForLoad & "\XXXXXX MODELO SEGUIMIENTO PEDIDOS.xlsx" 'Ruta de la subcarpeta MODELO SEGUIMIENTO PEDIDOS
        FileCopy strPathSub, strPath & "\" & rEXP & " MODELO SEGUIMIENTO PEDIDOS.xlsx" 'Crea el archivo y lo guarda
    
    Else
        'Continúa con el código
    End If
End Sub

Private Sub OptionButtonDAM_Click()
    MultiPage1.Pages(2).Visible = True 'Muestra la página AM
    MultiPage1.Pages(3).Visible = False 'Esconde la página NSP
    'Activar controles documentación
    CheckBoxApen.Enabled = True
    CheckBoxArcLic.Enabled = True
    CheckBoxCl9.Enabled = True
    CheckBoxCl15.Enabled = True
    CheckBoxComAcct.Enabled = True
    CheckBoxCorAcc.Enabled = True
    CheckBoxCritAdj.Enabled = True
    CheckBoxCumpFact.Enabled = True
    CheckBoxCumpOfer.Enabled = True
    CheckBoxFactElec.Enabled = True
    CheckBoxIn.Enabled = True
    CheckBoxIn668.Enabled = True
    CheckBoxLotes.Enabled = True
    CheckBoxPCAP.Enabled = True
    CheckBoxPDC.Enabled = True
    CheckBoxPPT.Enabled = True
    CheckBoxPresup.Enabled = True
    CheckBoxSegPed.Enabled = True
    CheckBoxSeguPed.Enabled = True
    CheckBoxTrAnt.Enabled = True
    ComboBoxFinan.Enabled = True
    ComboBoxLey.Enabled = True
    ComboBoxTipExp.Enabled = True
    lblFinan.Enabled = True
    lblLey.Enabled = True
    lblTipExp.Enabled = True
    btnCrearExpediente.Enabled = True
    'Activar botones repuestos
    OptionButtonAsien.Enabled = True
    OptionButtonBal.Enabled = True
    OptionButtonBat.Enabled = True
    OptionButtonCel.Enabled = True
    OptionButtonElec.Enabled = True
    OptionButtonEst.Enabled = True
    OptionButtonFren.Enabled = True
    OptionButtonHel.Enabled = True
    OptionButtonIlu.Enabled = True
    OptionButtonMot.Enabled = True
    OptionButtonNeu.Enabled = True
    OptionButtonOtros.Enabled = True
    OptionButtonSeg.Enabled = True
    OptionButtonTren.Enabled = True
    OptionButtonTub.Enabled = True
    OptionButtonUni.Enabled = True
End Sub

Private Sub OptionButtonNCP_Click()
    MultiPage1.Pages(2).Visible = False 'Esconde la página AM
    MultiPage1.Pages(3).Visible = False 'Esconde la página NSP
    'Activar controles documentación
    CheckBoxApen.Enabled = True
    CheckBoxArcLic.Enabled = True
    CheckBoxCl9.Enabled = True
    CheckBoxCl15.Enabled = True
    CheckBoxComAcct.Enabled = True
    CheckBoxCorAcc.Enabled = True
    CheckBoxCritAdj.Enabled = True
    CheckBoxCumpFact.Enabled = True
    CheckBoxCumpOfer.Enabled = True
    CheckBoxFactElec.Enabled = True
    CheckBoxIn.Enabled = True
    CheckBoxIn668.Enabled = True
    CheckBoxLotes.Enabled = True
    CheckBoxPCAP.Enabled = True
    CheckBoxPDC.Enabled = True
    CheckBoxPPT.Enabled = True
    CheckBoxPresup.Enabled = True
    CheckBoxSegPed.Enabled = True
    CheckBoxSeguPed.Enabled = True
    CheckBoxTrAnt.Enabled = True
    ComboBoxFinan.Enabled = True
    ComboBoxLey.Enabled = True
    ComboBoxTipExp.Enabled = True
    lblFinan.Enabled = True
    lblLey.Enabled = True
    lblTipExp.Enabled = True
    btnCrearExpediente.Enabled = True
    'Activar botones repuestos
    OptionButtonAsien.Enabled = True
    OptionButtonBal.Enabled = True
    OptionButtonBat.Enabled = True
    OptionButtonCel.Enabled = True
    OptionButtonElec.Enabled = True
    OptionButtonEst.Enabled = True
    OptionButtonFren.Enabled = True
    OptionButtonHel.Enabled = True
    OptionButtonIlu.Enabled = True
    OptionButtonMot.Enabled = True
    OptionButtonNeu.Enabled = True
    OptionButtonOtros.Enabled = True
    OptionButtonSeg.Enabled = True
    OptionButtonTren.Enabled = True
    OptionButtonTub.Enabled = True
    OptionButtonUni.Enabled = True
End Sub

Private Sub OptionButtonNSP_Click()
    MultiPage1.Pages(2).Visible = False 'Esconde la página AM
    MultiPage1.Pages(3).Visible = True 'Muestra la página NSP
    'Activar controles documentación
    CheckBoxApen.Enabled = True
    CheckBoxArcLic.Enabled = True
    CheckBoxCl9.Enabled = True
    CheckBoxCl15.Enabled = True
    CheckBoxComAcct.Enabled = True
    CheckBoxCorAcc.Enabled = True
    CheckBoxCritAdj.Enabled = True
    CheckBoxCumpFact.Enabled = True
    CheckBoxCumpOfer.Enabled = True
    CheckBoxFactElec.Enabled = True
    CheckBoxIn.Enabled = True
    CheckBoxIn668.Enabled = True
    CheckBoxLotes.Enabled = True
    CheckBoxPCAP.Enabled = True
    CheckBoxPDC.Enabled = True
    CheckBoxPPT.Enabled = True
    CheckBoxPresup.Enabled = True
    CheckBoxSegPed.Enabled = True
    CheckBoxSeguPed.Enabled = True
    CheckBoxTrAnt.Enabled = True
    ComboBoxFinan.Enabled = True
    ComboBoxLey.Enabled = True
    ComboBoxTipExp.Enabled = True
    lblFinan.Enabled = True
    lblLey.Enabled = True
    lblTipExp.Enabled = True
    btnCrearExpediente.Enabled = True
    'Activar botones repuestos
    OptionButtonAsien.Enabled = True
    OptionButtonBal.Enabled = True
    OptionButtonBat.Enabled = True
    OptionButtonCel.Enabled = True
    OptionButtonElec.Enabled = True
    OptionButtonEst.Enabled = True
    OptionButtonFren.Enabled = True
    OptionButtonHel.Enabled = True
    OptionButtonIlu.Enabled = True
    OptionButtonMot.Enabled = True
    OptionButtonNeu.Enabled = True
    OptionButtonOtros.Enabled = True
    OptionButtonSeg.Enabled = True
    OptionButtonTren.Enabled = True
    OptionButtonTub.Enabled = True
    OptionButtonUni.Enabled = True
End Sub

Private Sub UserForm_Initialize()
    ComboBoxLey.AddItem "Ley 9-2017" 'Ley 9-2017
    ComboBoxLey.AddItem "LCSPDS" 'LCSPDS
    ComboBoxFinan.AddItem "660" '660
    ComboBoxFinan.AddItem "668" '668
    ComboBoxTipExp.AddItem "Nacional" 'Nacional
    ComboBoxTipExp.AddItem "Extranjero" 'Extranjero
    'Empresas DAM
    ComboBoxEmp.AddItem "EUROPAVIA ESPAÑA S.A."
    ComboBoxEmp.AddItem "AIRBUS DEFENCE AND SPACE"
    ComboBoxEmp.AddItem "DILLERS S.A."
    ComboBoxEmp.AddItem "SAFRAN HELICOPTER ENGINES"
    ComboBoxEmp.AddItem "MARTIN BAKER AIRCRAFT CO LTD"
    ComboBoxEmp.AddItem "VIKING LIMITED"
    ComboBoxEmp.AddItem "GECI ESPAÑOLA, S.A."
    ComboBoxEmp.AddItem "CLIA SISTEMAS, S.L."
    ComboBoxEmp.AddItem "AIRBUS HELICOPTERS"
    ComboBoxEmp.AddItem "AERO PRECISION INDUSTRIES LLC"
    ComboBoxEmp.AddItem "ITP AERO"
    ComboBoxEmp.AddItem "AEROTECNIC METALLIC"
    ComboBoxEmp.AddItem "BOEING"
    ComboBoxEmp.AddItem "Otra empresa"
    'Empresas NSP
    ComboBoxEmp2.AddItem "EUROPAVIA ESPAÑA S.A."
    ComboBoxEmp2.AddItem "AIRBUS DEFENCE AND SPACE"
    ComboBoxEmp2.AddItem "DILLERS S.A."
    ComboBoxEmp2.AddItem "SAFRAN HELICOPTER ENGINES"
    ComboBoxEmp2.AddItem "MARTIN BAKER AIRCRAFT CO LTD"
    ComboBoxEmp2.AddItem "VIKING LIMITED"
    ComboBoxEmp2.AddItem "GECI ESPAÑOLA, S.A."
    ComboBoxEmp2.AddItem "CLIA SISTEMAS, S.L."
    ComboBoxEmp2.AddItem "AIRBUS HELICOPTERS"
    ComboBoxEmp2.AddItem "AERO PRECISION INDUSTRIES LLC"
    ComboBoxEmp2.AddItem "ITP AERO"
    ComboBoxEmp2.AddItem "AEROTECNIC METALLIC"
    ComboBoxEmp2.AddItem "BOEING"
    ComboBoxEmp2.AddItem "Otra empresa"
End Sub
