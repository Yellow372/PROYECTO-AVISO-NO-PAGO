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
        .ListIndex = -1
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
'Seccion: Boton registrar información
'----------------------------------------------------------------

'Boton "Registrar información"

Private Sub btn_registrar_click()

    Dim wsDestino As Worksheet
    Dim wsOrigen As Worksheet
    Dim lastRowDestino As Long
    Dim lastRowOrigen As Long
    Dim i As Long
    Dim respuesta As VbMsgBoxResult
    Dim cia As String
    
    If wbDatos Is Nothing Then
        MsgBox "Primero debes seleccionar un archivo.", vbExclamation
        Exit Sub
    End If

    If Me.cmb_cia.ListIndex = 0 Then
        MsgBox "Selecciona una compañia antes de continuar.", vbExclamation
        Exit Sub
    End If
    
    ' Se implementará una confirmación antes de darle a registrar, esto evitará futuras confusiones entre los diferentes formatos de las cia's
    
    cia = Me.cmb_cia.Value
        respuesta = MsgBox("¿Confirma que el archivo cargado pertenece a la compañia: " & cia & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Compañia")
    
    If respuesta = vbNo Then Exit Sub 'Manejo de error en caso la respuesta sea No
    On Error GoTo ErrHandler
    
    Set wsDestino = ThisWorkbook.Sheets("REPORTE")
    Set wsOrigen = wbDatos.Sheets(1)
    
    lastRowDestino = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row + 1 'Se busca agregar datos al final de la hoja destino "REPORTES".
    lastRowOrigen = wsOrigen.Cells(wsOrigen.Rows.Count, "A").End(xlUp).Row 'Busca la última fila con datos en la hoja de origen que se copiara los datos
    
    Application.ScreenUpdating = False
    
    'Manejo de error para el caso de las fechas
    If Trim(txt_frecepcion.Value) = "" Or Trim(txt_fenvioaviso.Value) = "" Then 'en caso no sea necesario la fecha de envio de aviso vamos a quitarlo del codigo
        MsgBox "Debes ingresar la fecha de recepción y de aviso.", vbExclamation
        Exit Sub
        End If
        
    Select Case LCase(cia) 'Al tener diferentes formatos, lo que haremos es crear una lista de cada uno de ellos. (Luego buscaremos como reducir este código)
    Case "pacifico"
         For i = 2 To lastRowOrigen
         With wsDestino
         .Cells(lastRowDestino, "A").NumberFormat = "dd/mm/yyyy": .Cells(lastRowDestino, "A").Value = CDate(Me.txt_frecepcion.Value) 'Fecha de recepción
         .Cells(lastRowDestino, "B").NumberFormat = "dd/mm/yyyy": .Cells(lastRowDestino, "B").Value = CDate(Me.txt_fenvioaviso.Value) 'Fecha de envio de aviso
         .Cells(lastRowDestino, "E").Value = cia 'En cada formato el valor de la columan E siempre sera la cia que seleccionaremos en el combo box
         .Cells(lastRowDestino, "H").Value = wsOrigen.Cells(i, "A").Value 'FALTA DE PAGO, SUSPENSIÓN, RESOLUCIÓN, ANULACIÓN
         .Cells(lastRowDestino, "I").Value = wsOrigen.Cells(i, "B").Value 'POLIZA
         .Cells(lastRowDestino, "J").Value = wsOrigen.Cells(i, "D").Value 'TIPO DE DOCUMENTO DE IDENTIDAD
         .Cells(lastRowDestino, "K").Value = wsOrigen.Cells(i, "E").Value 'ID CLIENTE
         .Cells(lastRowDestino, "L").Value = wsOrigen.Cells(i, "F").Value 'CLIENTE
         'A partir de ahora, en cada formato no consideraremos las columnas en las que tengamos datos en blanco
         End With
         lastRowDestino = lastRowDestino + 1
         
        Next i
        
        Case "rimac"
             For i = 2 To lastRowOrigen
             With wsDestino
             .Cells(lastRowDestino, "A").NumberFormat = "dd/mm/yyyy": .Cells(lastRowDestino, "A").Value = Date
             .Cells(lastRowDestino, "E").Value = cia
             .Cells(lastRowDestino, "L").Value = wsOrigen.Cells(i, "B").Value 'CLIENTE
             .Cells(lastRowDestino, "P").Value = wsOrigen.Cells(i, "G").Value 'RIESGO
             .Cells(lastRowDestino, "I").Value = wsOrigen.Cells(i, "H").Value 'POLIZA
             .Cells(lastRowDestino, "K").Value = wsOrigen.Cells(i, "W").Value 'IDCLIENTE
             End With
             lastRowDestino = lastRowDestino + 1
             
        Next i
        
        Case Else 'Esto en caso se seleccione una compañia que no tiene un formato todavia definido
             MsgBox "La compañia seleccionada no tiene un formato implementado aún.", vbExclamation
             Exit Sub
    End Select
    
    wbDatos.Close SaveChanges:=False
    Set wbDatos = Nothing
    
    MsgBox "Datos registrados exitosamente para la compañia: " & cia, vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "Error en registrar: " & Err.Description, vbCritical
    Application.ScreenUpdating = True
End Sub
    
'----------------------------------------------------------------
'Seccion: Boton borrar y cerrar
'----------------------------------------------------------------

'Boton borrar

Private Sub btn_borrar_Click()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        Select Case TypeName(ctrl)
            Case "TextBox"
                ctrl.Value = ""
            Case "ComboBox"
                ctrl.ListIndex = -1
        End Select
    Next ctrl
End Sub

'Boton cerrar

Private Sub btn_cerrar_Click()
    Unload Me
End Sub

        
