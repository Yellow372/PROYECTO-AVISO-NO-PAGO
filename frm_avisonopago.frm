VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_avisonopago 
   Caption         =   "Gestión de correos aviso no pago"
   ClientHeight    =   8865.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5070
   OleObjectBlob   =   "frm_avisonopago.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FRM_avisonopago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------
'Formulario para la gestión de coreos de aviso no pago
'----------------------------------------------------------------

'Variables globales

Dim wbDatos As Workbook

'----------------------------------------------------------------
'Seccion: Listas del formulario
'----------------------------------------------------------------

'Lista de compañias para reconocer la compañia que pertenece el correo
Private Sub UserForm_Initialize() 'Inicia el formulario
    With Me.cmb_cia 'listbox de compañias disponibles
        .Clear
        .AddItem "LA POSITIVA"
        .AddItem "PACIFICO"
        .AddItem "OHIO"
        .AddItem "QUALITAS"
        .AddItem "MAPFRE"
        .AddItem "INTERSEGURO"
        .AddItem "RIMAC"
        .ListIndex = -1 'Esto es para que inicie vacio
    End With
End Sub

'----------------------------------------------------------------
'Seccion: Boton buscar archivos
'----------------------------------------------------------------

'Boton "Buscar archivos"
Private Sub btn_bd_Click()
    Dim fd As FileDialog 'Sirve como interfaz
    Dim rutaBase As String
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Selecciona archivo Excel de origen"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm" 'Tipo de archhivos
        .AllowMultiSelect = False
        If .Show = -1 Then
            rutaBase = .SelectedItems(1) 'Abrimos el archivo para sacar los datos (solución: si mantenemos abierto, se puede extraer los datos)
            Set wbDatos = Workbooks.Open(rutaBase, ReadOnly:=True)
            MsgBox "Archivo abierto: " & wbDatos.Name, vbInformation 'mensaje de MsgBox
        End If
    End With
End Sub

'----------------------------------------------------------------
'Seccion: Boton registrar datos
'----------------------------------------------------------------

Private Sub btn_registrar_Click()
    Dim wsDestino As Worksheet
    Dim wsOrigen As Worksheet
    Dim lastRowDestino As Long, lastRowOrigen As Long
    Dim i As Long
    Dim respuesta As VbMsgBoxResult
    Dim cia As String

    If wbDatos Is Nothing Then
        MsgBox "Primero debes seleccionar un archivo con el botón 'Buscar Archivos'.", vbExclamation
        Exit Sub
    End If

    If Me.cmb_cia.ListIndex = -1 Then
        MsgBox "Selecciona una compañía antes de continuar.", vbExclamation
        Exit Sub
    End If

    ' Confirmación
    cia = Me.cmb_cia.Value
    respuesta = MsgBox("¿Confirma que el archivo cargado pertenece a la compañía: " & cia & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Compañía")

    If respuesta = vbNo Then Exit Sub

    On Error GoTo ErrHandler

    Set wsDestino = ThisWorkbook.Sheets("REPORTE")
    Set wsOrigen = wbDatos.Sheets(1)

    lastRowDestino = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row + 1
    lastRowOrigen = wsOrigen.Cells(wsOrigen.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False

    Select Case LCase(cia)
        Case "pacifico"
            For i = 2 To lastRowOrigen
                With wsDestino
                    .Cells(lastRowDestino, "A").Value = wsOrigen.Cells(i, "A").Value
                    .Cells(lastRowDestino, "C").Value = ""
                    .Cells(lastRowDestino, "E").Value = cia
                    .Cells(lastRowDestino, "H").Value = wsOrigen.Cells(i, "B").Value
                    .Cells(lastRowDestino, "I").Value = wsOrigen.Cells(i, "C").Value
                    .Cells(lastRowDestino, "J").Value = wsOrigen.Cells(i, "D").Value
                    .Cells(lastRowDestino, "K").Value = wsOrigen.Cells(i, "E").Value
                    .Cells(lastRowDestino, "L").Value = wsOrigen.Cells(i, "F").Value
                    .Cells(lastRowDestino, "P").Value = ""
                    .Cells(lastRowDestino, "Q").Value = ""
                    .Cells(lastRowDestino, "AB").Value = ""
                    .Cells(lastRowDestino, "AD").Value = ""
                    .Cells(lastRowDestino, "AE").Value = ""
                    .Cells(lastRowDestino, "AF").Value = ""
                    .Cells(lastRowDestino, "AG").Value = ""
                    .Cells(lastRowDestino, "AH").Value = ""
                    .Cells(lastRowDestino, "AI").Value = wsOrigen.Cells(i, "G").Value
                End With
                lastRowDestino = lastRowDestino + 1
            Next i

        Case "rimac"
            For i = 2 To lastRowOrigen
                With wsDestino
                    .Cells(lastRowDestino, "A").Value = wsOrigen.Cells(i, "A").Value
                    .Cells(lastRowDestino, "C").Value = ""
                    .Cells(lastRowDestino, "E").Value = cia
                    .Cells(lastRowDestino, "H").Value = wsOrigen.Cells(i, "B").Value
                    .Cells(lastRowDestino, "I").Value = wsOrigen.Cells(i, "C").Value
                    .Cells(lastRowDestino, "J").Value = wsOrigen.Cells(i, "D").Value
                    .Cells(lastRowDestino, "K").Value = wsOrigen.Cells(i, "E").Value
                    .Cells(lastRowDestino, "L").Value = wsOrigen.Cells(i, "F").Value
                    .Cells(lastRowDestino, "P").Value = wsOrigen.Cells(i, "G").Value
                    .Cells(lastRowDestino, "Q").Value = wsOrigen.Cells(i, "H").Value
                    .Cells(lastRowDestino, "AB").Value = wsOrigen.Cells(i, "I").Value
                    .Cells(lastRowDestino, "AD").Value = wsOrigen.Cells(i, "J").Value
                    .Cells(lastRowDestino, "AE").Value = wsOrigen.Cells(i, "K").Value
                    .Cells(lastRowDestino, "AF").Value = wsOrigen.Cells(i, "L").Value
                    .Cells(lastRowDestino, "AG").Value = wsOrigen.Cells(i, "M").Value
                    .Cells(lastRowDestino, "AH").Value = wsOrigen.Cells(i, "N").Value
                    .Cells(lastRowDestino, "AI").Value = wsOrigen.Cells(i, "O").Value
                End With
                lastRowDestino = lastRowDestino + 1
            Next i

        Case Else
            MsgBox "La compañía seleccionada no tiene lógica de extracción implementada aún.", vbExclamation
            Exit Sub
    End Select

    wbDatos.Close SaveChanges:=False
    Set wbDatos = Nothing

    MsgBox "Datos registrados exitosamente para la compañía: " & cia, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error en registrar: " & Err.Description, vbCritical
    Application.ScreenUpdating = True
End Sub

