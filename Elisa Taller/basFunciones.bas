Attribute VB_Name = "basFunciones"
Option Explicit
Dim PadLength As Integer
Dim x As Integer

Public Sub ActivaDesactivaBotonesListas(Lista As ListView, Formulario As Form, IndiceBotones As Double)

If Lista.ListItems.Count = 0 Then
    Formulario.tlbBotones(IndiceBotones).Buttons(1).Enabled = False
    Formulario.tlbBotones(IndiceBotones).Buttons(2).Enabled = False
    Exit Sub
End If

If SituacionLista(Lista).LLENA = True Then
    Formulario.tlbBotones(IndiceBotones).Buttons(1).Enabled = False
    Formulario.tlbBotones(IndiceBotones).Buttons(2).Enabled = True
ElseIf SituacionLista(Lista).MEDIA = True Then
    Formulario.tlbBotones(IndiceBotones).Buttons(1).Enabled = True
    Formulario.tlbBotones(IndiceBotones).Buttons(2).Enabled = True
ElseIf SituacionLista(Lista).VACIA = True Then
    Formulario.tlbBotones(IndiceBotones).Buttons(1).Enabled = True
    Formulario.tlbBotones(IndiceBotones).Buttons(2).Enabled = False
End If

End Sub

Public Sub SeleccionarTodo(Lista As ListView)
Dim ldblCont As Double

For ldblCont = 1 To Lista.ListItems.Count
    Lista.ListItems(ldblCont).Checked = True
Next ldblCont
End Sub

Public Function SituacionLista(Lista As ListView) As EstadoDeLista
Dim ldblCont As Double
Dim ldblSeleccionados As Double
Dim ldblNoSeleccionados As Double

Screen.MousePointer = vbHourglass

If Lista.ListItems.Count = 0 Then
    SituacionLista.LLENA = False
    SituacionLista.MEDIA = False
    SituacionLista.VACIA = True
    Screen.MousePointer = vbDefault
    Exit Function
End If

SituacionLista.LLENA = True
SituacionLista.MEDIA = True
SituacionLista.VACIA = True

ldblSeleccionados = 0
ldblNoSeleccionados = 0

For ldblCont = 1 To Lista.ListItems.Count
    If Lista.ListItems(ldblCont).Checked = True Then
        ldblSeleccionados = ldblSeleccionados + 1
    End If
    If Lista.ListItems(ldblCont).Checked = False Then
        ldblNoSeleccionados = ldblNoSeleccionados + 1
    End If
Next ldblCont

If ldblSeleccionados = Lista.ListItems.Count And ldblNoSeleccionados = 0 Then
    SituacionLista.LLENA = True
    SituacionLista.MEDIA = False
    SituacionLista.VACIA = False
ElseIf ldblNoSeleccionados = Lista.ListItems.Count And ldblSeleccionados = 0 Then
    SituacionLista.LLENA = False
    SituacionLista.MEDIA = False
    SituacionLista.VACIA = True
ElseIf ldblNoSeleccionados <> 0 And ldblSeleccionados <> 0 Then
    SituacionLista.LLENA = False
    SituacionLista.MEDIA = True
    SituacionLista.VACIA = False
End If

Screen.MousePointer = vbDefault

End Function

Public Sub DesmarcarTodo(Lista As ListView)
Dim ldblCont As Double

For ldblCont = 1 To Lista.ListItems.Count
    Lista.ListItems(ldblCont).Checked = False
Next ldblCont

End Sub

Public Sub MarcarTodo(Lista As ListView)
Dim ldblCont As Double

For ldblCont = 1 To Lista.ListItems.Count
    Lista.ListItems(ldblCont).Checked = True
Next ldblCont

End Sub

Public Function Iniciales(strTexto As String) As String
Dim intLargo As Integer, intIndice As Integer
Dim strCaracter As String * 1
Dim strInicial As String

intLargo = Len(strTexto)
For intIndice = 1 To intLargo
    If intIndice = 1 Then
        strInicial = Mid(strTexto, 1, 1)
    Else
        If Mid(strTexto, intIndice, 1) = " " And intIndice < intLargo Then
            If Mid(strTexto, intIndice + 1, 1) <> " " Then
                strInicial = strInicial & Mid(strTexto, intIndice + 1, 1)
            End If
        End If
    End If
Next
Iniciales = strInicial
End Function

Function ExportarDatos(ByRef lvwLista As ListView, cmdDialog As CommonDialog, hwndForm As Long)
    Dim i As Integer
    Dim j As Integer
    Dim strLinea As String
    Dim intCanal As Integer
    Dim retval As Long
    Dim strDato As String, Indice As Integer
    Err.Clear
    On Error GoTo ControlErrores
    If lvwLista.ListItems.Count <= 0 Then
        MsgBox "No hay elementos en la lista...", vbOKOnly + vbInformation, "Advertencia"
        Exit Function
    End If
  
    Screen.MousePointer = vbHourglass
    With cmdDialog
        .Flags = cdlOFNOverwritePrompt Or cdlOFNLongNames
        .CancelError = True
        .FileName = "*.XLS"
        .Filter = "Excel (*.XLS)|"
        .DefaultExt = "XLS"
        .ShowSave
    End With
    Screen.MousePointer = vbHourglass
    intCanal = FreeFile
    If cmdDialog.FileName = "" Then
        Exit Function
    End If
    Open cmdDialog.FileName For Output As #intCanal
    strLinea = ""
    For j = 1 To lvwLista.ColumnHeaders.Count
        If lvwLista.ColumnHeaders(j).Width > 0 Then
            strLinea = strLinea & lvwLista.ColumnHeaders(j).Text & Chr(9)
        End If
    Next
    strLinea = Left(strLinea, Len(strLinea) - 1)
    Print #intCanal, strLinea
    
    For i = 1 To lvwLista.ListItems.Count
        strLinea = ""
        For j = 1 To lvwLista.ColumnHeaders.Count
            If lvwLista.ColumnHeaders(j).Width > 0 Then
                If j = 1 Then
                    strDato = lvwLista.ListItems(i)
                Else
                    Indice = j - 1
                    strDato = lvwLista.ListItems(i).SubItems(Indice)
                End If
                'kjcv 15.08.16
                If IsNumeric(strDato) Then
                    strLinea = strLinea & strDato & Chr(9)
                ElseIf IsDate(strDato) Then
                    If Hour(strDato) Then
                        strLinea = strLinea & Format(strDato, "HH:mm") & Chr(9)
                    Else
                        strLinea = strLinea & Format(DateValue(strDato), "dd/mmm/yyyy") & Chr(9)
                    End If
                Else
                    strLinea = strLinea & strDato & Chr(9)
                End If
                
'                If IsDate(strDato) Then
'                    If Hour(strDato) Then
'                        strLinea = strLinea & Format(strDato, "HH:mm") & Chr(9)
'                    Else
'                        strLinea = strLinea & Format(DateValue(strDato), "dd/mmm/yyyy") & Chr(9)
'                    End If
'                Else
'                    strLinea = strLinea & strDato & Chr(9)
'                End If
                
            End If
        Next
        strLinea = Left(strLinea, Len(strLinea) - 1)
        Print #intCanal, strLinea
    Next
    Close #intCanal
    
    retval = ShellExecute(hwndForm, "open", cmdDialog.FileName, "", "", SW_RESTORE)
    Screen.MousePointer = vbDefault
    Exit Function
ControlErrores:
    Select Case Err.Number
    Case 70
        MsgBox "Existe una planilla Excel activa..." & Chr(13) & "Cierre la planilla activa" & Chr(13) & "y vuelva a exportar los datos...", vbOKOnly + vbInformation, "Advertencia"
    Case Else
        'MsgBox "Se produjo el siguiente error: " & Err.Description, vbOKOnly + vbInformation, "Advertencia"
    End Select
    Err.Clear
    Screen.MousePointer = vbDefault
End Function

Function ExportarDatosGrid(Cabecera As String, rsDatos As Recordset, cmdDialog As CommonDialog, hwndForm As Long)
    Dim i As Integer
    Dim j As Integer
    Dim strLinea As String
    Dim intCanal As Integer
    Dim retval As Long
    Dim strDato As String, Indice As Integer
    
  
    Err.Clear
    On Error GoTo ControlErrores
    If rsDatos.RecordCount <= 0 Then
        MsgBox "No hay elementos en la lista...", vbOKOnly + vbInformation, "Advertencia"
        Exit Function
    End If
  
    Screen.MousePointer = vbHourglass
    With cmdDialog
        .Flags = cdlOFNOverwritePrompt Or cdlOFNLongNames
        .CancelError = True
        .FileName = "*.XLS"
        .Filter = "Excel (*.XLS)|"
        .DefaultExt = "XLS"
        .ShowSave
    End With
    Screen.MousePointer = vbHourglass
    intCanal = FreeFile
    If cmdDialog.FileName = "" Then
        Exit Function
    End If
    Open cmdDialog.FileName For Output As #intCanal
    strLinea = ""

    strLinea = Cabecera
    strLinea = Left(strLinea, Len(strLinea) - 1)
    Print #intCanal, strLinea
        
    While Not rsDatos.EOF
    
        strLinea = ""
        For i = 0 To rsDatos.Fields.Count - 1
            strDato = IIf(IsNull(rsDatos.Fields(i).Value), "", rsDatos.Fields(i).Value)
            strDato = LTrim(RTrim(strDato))
            strDato = Replace(strDato, vbCrLf, "") ' Quitar los enters
            strDato = Replace(strDato, vbTab, "") ' quitar los tabs
            strDato = Replace(strDato, vbNewLine, "") ' quitar los new line
            strDato = Replace(strDato, vbCr, "") ' quitar los line break
            
        
            If IsNumeric(strDato) Then
                strLinea = strLinea & strDato & Chr(9)
            ElseIf IsDate(strDato) Then
                If Hour(strDato) Then
                    strLinea = strLinea & Format(strDato, "HH:mm") & Chr(9)
                Else
                    strLinea = strLinea & Format(DateValue(strDato), "dd/mmm/yyyy") & Chr(9)
                End If
            Else
                strLinea = strLinea & strDato & Chr(9)
            End If
           
        Next i
        strLinea = Left(strLinea, Len(strLinea) - 1)
        Print #intCanal, strLinea

        rsDatos.MoveNext
    
    Wend
    
    
    Close #intCanal
    
    retval = ShellExecute(hwndForm, "open", cmdDialog.FileName, "", "", SW_RESTORE)
    Screen.MousePointer = vbDefault
    Exit Function
ControlErrores:
    Select Case Err.Number
    Case 70
        MsgBox "Existe una planilla Excel activa..." & Chr(13) & "Cierre la planilla activa" & Chr(13) & "y vuelva a exportar los datos...", vbOKOnly + vbInformation, "Advertencia"
    Case Else
    End Select
    Err.Clear
    Screen.MousePointer = vbDefault
End Function

Public Function FamMateriales(strTexto As String) As String
gstrSql = "SELECT Id_Familia From Glbl_Familia WHERE Descripcion LIKE '%" & strTexto & "%'"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
With gadoPrincipal
    If Not .BOF And Not .EOF Then
        FamMateriales = !Id_Familia
    Else
        FamMateriales = ""
    End If
End With
End If
End Function

Public Function ImprimirDocumento(Origen As gInforme) As Boolean
Dim fintContador As Integer
Dim arrInventario(30) As String
Dim intAux As Integer
Dim strRut As String

 Dim Cant As Integer
        Dim linea As String
        Dim x As Integer

With frmRecepcion
            
On Error Resume Next

.cdImpresora.Flags = &H80000 Or &H40000 Or &H1
.cdImpresora.CancelError = True
.cdImpresora.Action = 5

If Origen = gRecepcion And Err.Number = 0 Then
    
        '.cdImpresora.CancelError = True
        '.cdImpresora.ShowPrinter
        
       
        Printer.Copies = .cdImpresora.Copies
        
        '/////////////////////ENCABEZADO DE LA RECEPCION
        ImprimeObjeto 500, 10000, ValorNulo(CStr(Val(Mid(.lblNroRecepcion, 6, 15)))), 12, True '//////////  NRO RECEPCION
        ImprimeObjeto 900, 10000, ValorNulo(.txtNroCono), 12 '//////////  NRO CONO
        ImprimeObjeto 1500, 1800, ValorNulo(.lblCliente), 8  '//////////  NBE CLIENTE
        If Not IsNull(.txtRut) Then
            If Trim(.txtRut) <> "" Then
                If gstrEditaRut = "S" Then
                    ImprimeObjeto 1500, 9200, Format(Mid(.txtRut, 1, Len(.txtRut) - 1), "00000000"), 8 '//////////  RUT CLIENTE
                    ImprimeObjeto 1500, 10000, "-" & ValorNulo(Mid(.txtRut, Len(.txtRut), 1)), 8
                Else
                    ImprimeObjeto 1500, 9200, .txtRut, 8 '//////////  RUT CLIENTE
                End If
            Else
                ImprimeObjeto 1500, 9200, "SIN DNI", 8   '//////////  RUT CLIENTE
            End If
        Else
'            ImprimeObjeto 1500, 9200, "SIN DNI, 8   '//////////  RUT CLIENTE", 8, False, False
            ImprimeObjeto 1500, 9200, "SIN DNI", 8   '//////////  RUT CLIENTE
            'ImprimeObjeto 1500, 10000, ValorNulo(Mid(strRut, Len(strRut), 1)), 8
        End If

         ImprimeObjeto 2000, 1800, ValorNulo(.txtDir), 8 '//////////  DIRECCION CLIENTE
        ImprimeObjeto 2000, 9500, ValorNulo(.txtComuna), 8 '//////////  COMUNA CLIENTE
        ImprimeObjeto 2450, 3000, ValorNulo(.txtSolicita), 8 '//////////  SOLICITADO POR
        ImprimeObjeto 2450, 7500, ValorNulo(.lblFono), 8 '//////////  FONOS
        ImprimeObjeto 3000, 1800, ValorNulo(.lblModelo), 8 '//////////  MODELO
        ImprimeObjeto 3000, 4800, ValorNulo(.txtAño), 8 '//////////  AÑO
        ImprimeObjeto 3000, 6500, ValorNulo(.lblColorE), 8  '//////////  COLOR EXT / INT
        ImprimeObjeto 3000, 9800, ValorNulo(.txtPatente), 8 '//////////  PATENTE
        ImprimeObjeto 3500, 1800, ValorNulo(.lblVin), 8 '//////////  VIN
        ImprimeObjeto 3500, 7200, ValorNulo(Format(.pckFecVta, "dd/mm/yyyy")), 8 '//////////  FECHA VENTA
        ImprimeObjeto 3500, 9800, ValorNulo(.txtKilAct), 8 '//////////  KILOMETROS
        ImprimeObjeto 3900, 3000, ValorNulo(.txtFolioGarantia), 8 '//////////  FOLIO GARANTIA

'        ImprimeObjeto 3900, 8000, ValorNulo(Format$(Time, "hh:mm")), 8
        ImprimeObjeto 3900, 8600, ValorNulo(Format$(.pckFechaAtencion.Value, "dd")), 8   '//////////  FECHA RECEPCION
        ImprimeObjeto 3900, 9100, ValorNulo(Format$(.pckFechaAtencion.Value, "MMMM")), 8
        ImprimeObjeto 3900, 10200, ValorNulo(Format$(.pckFechaAtencion.Value, "yyyy")), 8

        ImprimeObjeto 4350, 8600, ValorNulo(Format$(.pckFechaEntrega.Value, "dd")), 8 '////////// FECHA ENTREGA
        ImprimeObjeto 4350, 9100, ValorNulo(Format$(.pckFechaEntrega.Value, "MMMM")), 8 '////////// FECHA ENTREGA
        ImprimeObjeto 4350, 10200, ValorNulo(Format$(.pckFechaEntrega.Value, "yyyy")), 8 '////////// FECHA ENTREGA

'        ImprimeObjeto 4850, 7200, ValorNulo(Iniciales(.dtcRecepcionista.Text)), 8  '////////// RECEPCIONISTA
        ImprimeObjeto 4850, 7100, ValorNulo(.dtcRecepcionista.Text), 8 '////////// RECEPCIONISTA
        ImprimeObjeto 4850, 9800, ValorNulo(.cboHora.Text), 8 '//////////  HORA
                
        Dim nroMov As String
        nroMov = "   " & .nroMovilXId(.dtcRecepcionista.BoundText)
        
        ImprimeObjeto 4850, 8600, ValorNulo(nroMov), 8 '////////// wcs 23/03/20 Recepcionista tel.
        
'        ImprimeObjeto 8800, 3800, ValorNulo(Trim(.txtComentario)), 8
        
        
        ImprimirComentario (.txtComentario)
        
        
        

        '//////////////////////////////////////////REVISION
        Dim intX As Integer
        Dim strRevision As String

        With .lvwServiciosMecanica
            If .ListItems.Count > 0 Then
                For intX = 1 To .ListItems.Count
                    Set .SelectedItem = .ListItems(intX)
                    If Mid(.SelectedItem, 1, 2) = "RV" Then
                       strRevision = .SelectedItem.SubItems(1)
                    End If
                Next
                If strRevision <> "" Then
                    strRevision = Mid(strRevision, 9, Len(strRevision) - 11)
                    ImprimeObjeto 7000, 8000, strRevision, 8, True
                End If
            End If

        End With
        '////////////////////////// IMPRIME 9 SERVICIOS
        If .lvwOtrosServicios.ListItems.Count > 0 Then
            If .lvwOtrosServicios.ListItems.Count <= gintNumeroLineasRecepcion Then
                For intX = 1 To .lvwOtrosServicios.ListItems.Count
                    Set .lvwOtrosServicios.SelectedItem = .lvwOtrosServicios.ListItems(intX)
                    With .lvwOtrosServicios
                        ImprimeObjeto 8900 + (intX * 300), 2400, .SelectedItem.SubItems(1), 8
                    End With
                Next
            Else
                For intX = 1 To gintNumeroLineasRecepcion
                    Set .lvwOtrosServicios.SelectedItem = .lvwOtrosServicios.ListItems(intX)
                    With .lvwOtrosServicios
                        ImprimeObjeto 9000 + (intX * 300), 2400, .SelectedItem.SubItems(1), 8
                    End With
                Next
            End If
        End If
        
        
        '/////////////////////////////////////////////////
        Printer.EndDoc
 '   End With
ElseIf Origen = gPresupuesto Then

ElseIf Origen = gOT Then

Else
    MsgBox "Impresión Cancelada.", vbInformation, "Advertencia"
    Err.Clear
    Exit Function
End If
End With
End Function

Private Sub ImprimirComentario(Comentario As String)
Dim LARGO As Integer
Dim linea As String
Dim i As Integer
Dim x As Integer
Dim Y As Integer

Dim Lineas As Variant


'j = 20 'caracteres x linea
LARGO = Len(Comentario)
i = 1
Y = 5800
x = 2000

'Dim i As Integer

Lineas = Split(Comentario, vbCrLf)

For i = LBound(Lineas) To UBound(Lineas)
    Printer.CurrentX = x
    Printer.CurrentY = Y
    Printer.Print Lineas(i)
    Y = Y + 300
Next


End Sub

Public Function ImprimirDocumentoRecepcion(Origen As gInforme) As Boolean
Dim adoTemp As New ADODB.Recordset
Dim strSql As String
Dim fintContador As Integer
Dim arrInventario(30) As String
Dim intAux As Integer
Dim strRut As String
Dim i, x As Integer
Dim ContadorPagina As Integer

On Error Resume Next

frmRecepcion.cdImpresora.Flags = &H80000 Or &H40000 Or &H1
frmRecepcion.cdImpresora.CancelError = True
frmRecepcion.cdImpresora.Action = 5


If Origen = gRecepcion And Err.Number = 0 Then
    With frmRecepcion
        '/////////////////////ENCABEZADO DE LA RECEPCION
       ' Err.Clear
       ' On Error GoTo Error_Impresion
        
        ContadorPagina = 1
        
        Printer.Copies = .cdImpresora.Copies
      
        Printer.Font.Name = "Courier New"

        ImprimirEncabezadoRecepcion
        
        ImprimeObjeto 1700, 100, "Señores", 12, True
        ImprimeObjeto 2000, 100, "Nombre    : " & ValorNulo(.lblCliente), 8
        If Not IsNull(.txtRut) Then
            If Trim(.txtRut) <> "" Then
                If gstrEditaRut = "S" Then
                    ImprimeObjeto 2200, 100, gstrNombreRut & "       : " & Format(Mid(.txtRut, 1, Len(.txtRut) - 1), "00000000") & "-" & ValorNulo(Mid(.txtRut, Len(.txtRut), 1)), 8  '//////////  RUT CLIENTE
                    'ImprimeObjeto 2200, 2250, "-" & ValorNulo(Mid(.txtRut, Len(.txtRut), 1)), 8
                Else
                    ImprimeObjeto 2200, 100, gstrNombreRut & "       : " & Format(.txtRut, "000000000"), 8  '//////////  RUT CLIENTE
                End If
            Else
                ImprimeObjeto 2200, 10, gstrNombreRut & "       : " & "SIN " & gstrNombreRut, 8    '//////////  RUT CLIENTE"
            End If
        Else
            ImprimeObjeto 2200, 100, gstrNombreRut & "       : " & "SIN " & gstrNombreRut, 8  '//////////  RUT CLIENTE
        End If
        ImprimeObjeto 2400, 100, "Dirección : " & ValorNulo(.txtDir), 8 '//////////  DIRECCION CLIENTE
        ImprimeObjeto 2600, 100, "Teléfono  : " & ValorNulo(.lblFono), 8 '//////////  FONOS


       ' ImprimeObjeto 2000, 9500, ValorNulo(.txtComuna), 8 '//////////  COMUNA CLIENTE
       ' ImprimeObjeto 2450, 3000, ValorNulo(.txtSolicita), 8 '//////////  SOLICITADO POR

        ImprimeObjeto 3200, 100, "Vehículo", 12, True
         ImprimeObjeto 3500, 100, gstrNombrePatente & "      : " & ValorNulo(.txtPatente), 10  '//////////  PATENTE
        ImprimeObjeto 3500, 6500, "Año          : " & ValorNulo(.txtAño), 10 '//////////  AÑO
         ImprimeObjeto 3700, 100, "Marca        : " & ValorNulo(.lblMarca), 10 '//////////  MARCA
        ImprimeObjeto 3700, 6500, "Color        : " & ValorNulo(.lblColorE), 10  '//////////  COLOR EXT / INT
         ImprimeObjeto 3900, 100, "Modelo       : " & ValorNulo(.lblModelo), 10 '//////////  MODELO
        ImprimeObjeto 3900, 6500, "kms          : " & ValorNulo(.txtKilAct), 10 '//////////  KILOMETROS
         ImprimeObjeto 4100, 100, "Chasis       : " & ValorNulo(.lblChasis), 10 '///////////   Chasis
        ImprimeObjeto 4100, 6500, "F.Atención   : " & ValorNulo(Format$(.pckFechaAtencion.Value, "dd/mm/yyyy")), 10   '//////////  FECHA RECEPCION
         ImprimeObjeto 4300, 100, "Nº Motor     : " & ValorNulo(.lblMotor), 10 '//// nº motor
        ImprimeObjeto 4300, 6500, "F.Entrega    : " & ValorNulo(Format$(.pckFechaEntrega.Value, "dd/mm/yyyy")), 10 '////////// FECHA ENTREGA
         ImprimeObjeto 4500, 100, "Nº Vin       : " & ValorNulo(.lblVin), 10
        ImprimeObjeto 4500, 6500, "Hora Entrega : " & ValorNulo(.cboHora.Text), 10 '//////////  HORA
        ImprimeObjeto 4700, 100, "Fec.Venta    : " & ValorNulo(Format(.pckFecVta, "dd/mm/yyyy")), 10 '///// fecha venta
         
        
       ' ImprimeObjeto 3900, 3000, ValorNulo(.txtFolioGarantia), 8 '//////////  FOLIO GARANTIA
       ' ImprimeObjeto 3900, 8000, ValorNulo(Format$(Time, "hh:mm")), 8

        ImprimeObjeto 5000, 100, "Recepcionista : " & ValorNulo(.dtcRecepcionista.Text), 10, True '////////// RECEPCIONISTA
        ImprimeObjeto 5200, 100, "Solicitante   : " & ValorNulo(.txtSolicita), 10, True '////////// SOLICITANTE


        '//////////////////////////////////////////REVISION
        Dim intX As Integer
        Dim strRevision As String
        Dim fila As Integer

        fila = 5600
        ImprimeObjeto fila, 100, "Detalle de Servicios", 12, True

        fila = fila + 400


        With .lvwServiciosMecanica
            If .ListItems.Count > 0 Then
                For intX = 1 To .ListItems.Count
                    ImprimeObjeto fila, 100, .SelectedItem.SubItems(1), 9
                    fila = fila + 200
                    
                Next
            End If

        End With

        If .lvwOtrosServicios.ListItems.Count > 0 Then
                For intX = 1 To .lvwOtrosServicios.ListItems.Count
                    Set .lvwOtrosServicios.SelectedItem = .lvwOtrosServicios.ListItems(intX)
                    With .lvwOtrosServicios
                        ImprimeObjeto fila, 100, .SelectedItem.SubItems(1), 9
                        fila = fila + 200
                        If fila > 9000 And ContadorPagina = 1 Then
                            ImprimirPieRecepcion
                            Printer.EndDoc
                            ImprimirEncabezadoRecepcion
                            ContadorPagina = 2
                            fila = 1700
                            ImprimeObjeto fila, 100, "Detalle de Servicios", 12, True
                            fila = fila + 400
                        ElseIf fila > 14000 And ContadorPagina > 1 Then
                            Printer.EndDoc
                            ImprimirEncabezadoRecepcion
                            ContadorPagina = 2
                            fila = 1700
                            ImprimeObjeto fila, 100, "Detalle de Servicios", 12, True
                            fila = fila + 400
                        End If
                    End With
                Next
        End If

        If .lvwServiciosCarroceria.ListItems.Count > 0 Then
                For intX = 1 To .lvwServiciosCarroceria.ListItems.Count
                    Set .lvwServiciosCarroceria.SelectedItem = .lvwServiciosCarroceria.ListItems(intX)
                    With .lvwServiciosCarroceria
                        ImprimeObjeto fila, 100, .SelectedItem.SubItems(2), 9
                        fila = fila + 200
                        If fila > 9000 And ContadorPagina = 1 Then
                            ImprimirPieRecepcion
                            Printer.EndDoc
                            ImprimirEncabezadoRecepcion
                            ContadorPagina = 2
                            fila = 1700
                            ImprimeObjeto fila, 100, "Detalle de Servicios", 12, True
                            fila = fila + 400
                        ElseIf fila > 14000 And ContadorPagina > 1 Then
                            Printer.EndDoc
                            ImprimirEncabezadoRecepcion
                            ContadorPagina = 2
                            fila = 1700
                            ImprimeObjeto fila, 100, "Detalle de Servicios", 12, True
                            fila = fila + 400
                        End If
                    End With
                Next
        End If

        If .lvwServiciosTerceros.ListItems.Count > 0 Then
                For intX = 1 To .lvwServiciosTerceros.ListItems.Count
                    Set .lvwServiciosTerceros.SelectedItem = .lvwServiciosTerceros.ListItems(intX)
                    With .lvwServiciosTerceros
                        ImprimeObjeto fila, 100, .SelectedItem.SubItems(3), 9
                        fila = fila + 200
                        If fila > 9000 And ContadorPagina = 1 Then
                            ImprimirPieRecepcion
                            Printer.EndDoc
                            ImprimirEncabezadoRecepcion
                            ContadorPagina = 2
                            fila = 1700
                            ImprimeObjeto fila, 100, "Detalle de Servicios", 12, True
                            fila = fila + 400
                        ElseIf fila > 14000 And ContadorPagina > 1 Then
                            Printer.EndDoc
                            ImprimirEncabezadoRecepcion
                            ContadorPagina = 2
                            fila = 1700
                            ImprimeObjeto fila, 100, "Detalle de Servicios", 12, True
                            fila = fila + 400
                        End If
                    End With
                Next
        End If

        ImprimeObjeto fila + 200, 100, "Comentario", 12, True
        ImprimeObjeto fila + 600, 100, .txtComentario, 7, False, False
        
        
        'ImprimeObjeto fila + 800, 100, Mid(.txtComentario, 76, 75), 7, False, False
        'ImprimeObjeto fila + 1000, 100, Mid(.txtComentario, 151, 75), 7, False, False
        'ImprimeObjeto fila + 1200, 100, Mid(.txtComentario, 226, 75), 7, False, False
        'ImprimeObjeto fila + 1400, 100, Mid(.txtComentario, 301, 75), 7, False, False
        
        If ContadorPagina = 1 Then
            ImprimirPieRecepcion
        End If


        '/////////////////////////////////////////////////
        Printer.EndDoc
    End With
ElseIf Origen = gPresupuesto Then

ElseIf Origen = gOT Then

Else
    MsgBox "Impresión Cancelada.", vbInformation, "Advertencia"
    Err.Clear

    Exit Function
End If
'Exit Function
'Error_Impresion:
'    If Err.Number <> 32755 Then
'        MsgBox "Se ha producido el siguiente error " & Err.Number & " " & Err.Description, vbInformation, "Advertencia"
'    End If
'    Err.Clear
End Function
Public Function ImprimirEncabezadoRecepcion()
        ImprimeObjeto 100, 100, gstrEmpresa, 9, False, False
        ImprimeObjeto 300, 100, gstrSucursal, 9, False, False
        ImprimeObjeto 500, 100, gstrDirSuc, 9, False, False
        ImprimeObjeto 700, 100, "Fono: " & gstrTelefono, 9
        ImprimeObjeto 300, 8000, Format(Now, "dddd, d mmm yyyy"), 9
        ImprimeObjeto 500, 8000, Format$(Time, "hh:mm:ss"), 9

        'ImprimeObjeto 800, 4300, "Recepción " & IIf(.optRecepcion(0) = True, "Mecanica", "Carrocería"), 12, True
        ImprimeObjeto 1200, 4000, "Orden de Trabajo Nº: " & ValorNulo(CStr(Val(Mid(frmRecepcion.lblNroRecepcion, 6, 15)))), 12, True
        ImprimeObjeto 1450, 2500, "Cono Nº: " & ValorNulo(frmRecepcion.txtNroCono) & "  Color: " & frmRecepcion.dtcTipoCono.Text, 10, True
        ImprimeObjeto 1450, 6500, "Tipo Orden: " & frmRecepcion.dtcGarantia.Text, 10, True

End Function
Public Function ImprimirPieRecepcion()
Dim fila As Integer
Dim i, x, j As Integer
Dim strSql As String
Dim adoTemp As New ADODB.Recordset


fila = 9400

    If gblnImprimeImagen = True Then
         '// imprime texto si no
         'fila = fila + 1600
         For i = 1 To 5
             ImprimeObjeto fila, x + 100, " SI NO", 6, True
             x = x + 2200
         Next
         
         '/// imprime inventario
         fila = fila + 200
         
         'strSql = "Select descripcion from Tllr_Estado_Recepcion Where vigencia='S'"
         'If Conexion.SendHost(strSql, AdoTemp, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
         '    With AdoTemp
         '        If Not .BOF And Not .EOF Then
                     i = 1    'cuenta las columnas (maxima 5)
                     x = 0    'suma espacio entre cada columna
                     'While Not .EOF
                     For j = 1 To frmRecepcion.lvwInventario.ListItems.Count
                        If frmRecepcion.lvwInventario.ListItems(j).Checked = True Then
                            ImprimeObjeto fila, 100 + x, " X  __ " & frmRecepcion.lvwInventario.ListItems(j).SubItems(1), 6, False
                        Else
                            ImprimeObjeto fila, 100 + x, " __  X " & frmRecepcion.lvwInventario.ListItems(j).SubItems(1), 6, False
                        End If
                         x = x + 2200
                         i = i + 1
                         If i = 6 Then
                             i = 1
                             x = 0
                             fila = fila + 200
                         End If
                         '.MoveNext
                     'Wend
                     Next j
         '        End If
         '    End With
         'End If
         
         '/// imprime imagen de vehículo
         fila = fila + 400
         Printer.PaintPicture frmRecepcion.Image1.Picture, 200, fila, 10200, 2500
         
         fila = fila + 2500
         ImprimeObjeto fila, 100, "NOTA", 10, True
         ImprimeObjeto fila + 300, 100, gstrNotaRecepcion, 7, False, False
         
         If frmRecepcion.cmbBencina.ListIndex = 0 Then
            ImprimeObjeto fila, 7500, "X------|------|------|------|", 10
            ImprimeObjeto fila + 300, 7500, "0     1/4    1/2    3/4     F", 10
        ElseIf frmRecepcion.cmbBencina.ListIndex = 1 Then
            ImprimeObjeto fila, 7500, "|------X------|------|------|", 10
            ImprimeObjeto fila + 300, 7500, "0     1/4    1/2    3/4     F", 10
        ElseIf frmRecepcion.cmbBencina.ListIndex = 2 Then
            ImprimeObjeto fila, 7500, "|------|------X------|------|", 10
            ImprimeObjeto fila + 300, 7500, "0     1/4    1/2    3/4     F", 10
        ElseIf frmRecepcion.cmbBencina.ListIndex = 3 Then
            ImprimeObjeto fila, 7500, "|------|------|------X------|", 10
            ImprimeObjeto fila + 300, 7500, "0     1/4    1/2    3/4     F", 10
        ElseIf frmRecepcion.cmbBencina.ListIndex = 4 Then
            ImprimeObjeto fila, 7500, "|------|------|------|------X", 10
            ImprimeObjeto fila + 300, 7500, "0     1/4    1/2    3/4     F", 10
        Else
            ImprimeObjeto fila, 7500, "|------|------|------|------|", 10
            ImprimeObjeto fila + 300, 7500, "0     1/4    1/2    3/4     F", 10
        End If
        
         
        fila = fila + 1200
        ImprimeObjeto fila, 2800, "__________________________", 10
        ImprimeObjeto fila, 6200, "__________________________", 10
        fila = fila + 300
        ImprimeObjeto fila, 3800, "Recepcionista", 8
        ImprimeObjeto fila, 7200, "Cliente", 8

         
    Else
         'fila = fila + 1600
         ImprimeObjeto fila, 100, "NOTA", 10, True
         ImprimeObjeto fila + 300, 100, gstrNotaRecepcion, 7, False, False
         
'            fila = fila + 1200
'            ImprimeObjeto fila, 2800, "__________________________", 10
'            ImprimeObjeto fila, 6200, "__________________________", 10
'            fila = fila + 300
'            ImprimeObjeto fila, 3800, "Recepcionista", 8
'            ImprimeObjeto fila, 7200, "Cliente", 8
         
    End If
        
End Function
Public Function ImprimirDocumentoPiamonte(Origen As gInforme) As Boolean
Dim fintContador As Integer
Dim arrInventario(30) As String
Dim intAux As Integer
Dim strRut As String
Dim lintNumeroLineasOcupadas As Integer
Dim lintInicioLineaImpresion As Integer
Dim i As Integer

If Origen = gRecepcion Then
    With frmRecepcion
       ' /////////////////////ENCABEZADO DE LA RECEPCION
        ImprimeObjeto 1050, 8300, ValorNulo(CStr(Val(Mid(.lblNroRecepcion, 6, 15)))), 10, True '//////////  NRO RECEPCION
        ImprimeObjeto 1100, 10400, ValorNulo(.txtNroCono), 12 '//////////  NRO CONO
        ImprimeObjeto 1650, 8000, ValorNulo(.dtcRecepcionista.Text), 7  '////////// RECEPCIONISTA
        ImprimeObjeto 2000, 7650, "(" & ValorNulo(.dtcGarantia.Text) & ")", 12  '////////// GARANTIA
        ImprimeObjeto 2600, 1400, ValorNulo(.lblCliente), 7  '//////////  NOMBRE CLIENTE
        If Not IsNull(.txtRut) Then
            If Trim(.txtRut) <> "" Then
                ImprimeObjeto 2600, 9300, Format(Mid(.txtRut, 1, Len(.txtRut) - 1), "00000000"), 7   '//////////  RUT CLIENTE
                ImprimeObjeto 2600, 10000, "-" & ValorNulo(Mid(.txtRut, Len(.txtRut), 1)), 7
            Else
                ImprimeObjeto 2600, 9300, "SIN RUT", 7   '//////////  RUT CLIENTE
            End If
        Else
            ImprimeObjeto 2600, 9300, "SIN RUT", 7   '//////////  RUT CLIENTE
        End If

        ImprimeObjeto 2850, 1400, ValorNulo(.txtDir), 7 '//////////  DIRECCION CLIENTE
        ImprimeObjeto 2850, 9300, ValorNulo(.lblFono), 7 '//////////  FONOS
        ImprimeObjeto 3070, 1500, ValorNulo(.txtSolicita), 7 '//////////  SOLICITADO POR
        ImprimeObjeto 3070, 9500, ValorNulo(.txtComuna), 7 '//////////  COMUNA CLIENTE

        ImprimeObjeto 3580, 1600, ValorNulo(.lblMarca) & " " & ValorNulo(.lblModelo), 7 '//////////  MARCA MODELO
        ImprimeObjeto 3580, 7000, ValorNulo(.txtPatente), 7 '//////////  PATENTE
        ImprimeObjeto 3580, 9300, ValorNulo(.lblColorE), 7  '//////////  COLOR EXT / INT
        ImprimeObjeto 3820, 1000, ValorNulo(.txtKilAct), 7 '//////////  KILOMETROS
        ImprimeObjeto 3820, 2800, ValorNulo(.txtAño), 7 '//////////  AÑO
        ImprimeObjeto 3820, 5600, ValorNulo(.lblChasis), 7 '//////////  CHASSIS
        ImprimeObjeto 3820, 9300, ValorNulo(.lblMotor), 7 '//////////  MOTOR
        ImprimeObjeto 4080, 1500, ValorNulo(.txtConcesionario), 7 '////////// CONSECIONARIO
        ImprimeObjeto 4080, 6000, ValorNulo(Format(.pckFecVta, "dd/mm/yyyy")), 7 '//////////  FECHA VENTA
        ImprimeObjeto 4080, 9300, ValorNulo(.lblVin), 7 '//////////  VIN
        ImprimeObjeto 4330, 1500, ValorNulo(.lblCompañia), 7 '//////////  CIA SEGURO
        ImprimeObjeto 4330, 5800, ValorNulo(.txtLiquidador), 7 '//////////  LIQUIDADOR
        ImprimeObjeto 4580, 1500, ValorNulo(.txtNroSiniestro), 7 '//////////  SINIESTRO
        ImprimeObjeto 4580, 5800, ValorNulo(Format(.txtDeduciblePesos, "###,###,##0")), 7 '////////// DEDUCIBLE PESOS

        ImprimeObjeto 4810, 1400, ValorNulo(Format(.pckFechaAtencion, "dd/mm/yyyy")), 7 '//////////  FECHA RECEPCION
        ImprimeObjeto 4810, 4000, ValorNulo(Format$(Time, "hh:mm")), 7
        ImprimeObjeto 4810, 8000, ValorNulo(Format(.pckFechaEntrega, "dd/mm/yyyy")), 7 '//////////  FECHA ENTREGA
        ImprimeObjeto 4810, 10100, ValorNulo(.cboHora.Text), 7 '//////////  HORA


      '  //////////////////////////////////////////REVISION
        Dim intX As Integer
        Dim strRevision As String

        With .lvwServiciosMecanica
            If .ListItems.Count > 0 Then
                For intX = 1 To .ListItems.Count
                    Set .SelectedItem = .ListItems(intX)
                    If Mid(.SelectedItem, 1, 2) = "RV" Then
                       strRevision = .SelectedItem.SubItems(1)
                    End If
                Next
                If strRevision <> "" Then
                    strRevision = Mid(strRevision, 9, Len(strRevision) - 11)
                    ImprimeObjeto 5740, 5750, strRevision, 6, True
                End If
            End If

        End With
     '   ////////////////////////// IMPRIME 9(PARAMETRO) SERVICIOS
        If .lvwOtrosServicios.ListItems.Count > 0 Then  '// OTROS SERVICIOS
            If .lvwOtrosServicios.ListItems.Count <= gintNumeroLineasRecepcion Then
                For intX = 1 To .lvwOtrosServicios.ListItems.Count
                    Set .lvwOtrosServicios.SelectedItem = .lvwOtrosServicios.ListItems(intX)
                    With .lvwOtrosServicios
                        ImprimeObjeto 7300 + (intX * 250), 2200, .SelectedItem.SubItems(1), 7
                    End With
                Next
            Else
                For intX = 1 To gintNumeroLineasRecepcion   'numero de lineas a imprimir
                    Set .lvwOtrosServicios.SelectedItem = .lvwOtrosServicios.ListItems(intX)
                    With .lvwOtrosServicios
                        ImprimeObjeto 7400 + (intX * 250), 2200, .SelectedItem.SubItems(1), 7
                    End With
                Next
            End If
        End If

        If .lvwServiciosTerceros.ListItems.Count > 0 Then  '// SERVICIO DE TERCEROS
            If .lvwServiciosTerceros.ListItems.Count <= gintNumeroLineasRecepcion - .lvwOtrosServicios.ListItems.Count Then
                For intX = .lvwOtrosServicios.ListItems.Count + 1 To .lvwOtrosServicios.ListItems.Count + .lvwServiciosTerceros.ListItems.Count
                    Set .lvwServiciosTerceros.SelectedItem = .lvwServiciosTerceros.ListItems(intX - .lvwOtrosServicios.ListItems.Count)
                    With .lvwServiciosTerceros
                        ImprimeObjeto 7300 + (intX * 250), 2200, .SelectedItem.SubItems(3), 7
                    End With
                Next
            Else
                For intX = 1 To gintNumeroLineasRecepcion   'numero de lineas a imprimir
                    Set .lvwServiciosTerceros.SelectedItem = .lvwServiciosTerceros.ListItems(intX)
                    With .lvwServiciosTerceros
                        ImprimeObjeto 7400 + (intX * 250), 2200, .SelectedItem.SubItems(3), 7
                    End With
                Next
            End If
        End If

        If .lvwServiciosCarroceria.ListItems.Count > 0 Then  '// SERVICIO DE CARROCERIA
            If .lvwServiciosCarroceria.ListItems.Count <= gintNumeroLineasRecepcion - (.lvwOtrosServicios.ListItems.Count - .lvwServiciosCarroceria.ListItems.Count) Then
                For intX = .lvwOtrosServicios.ListItems.Count + .lvwServiciosTerceros.ListItems.Count + 1 To .lvwOtrosServicios.ListItems.Count + .lvwServiciosTerceros.ListItems.Count + .lvwServiciosCarroceria.ListItems.Count
                    Set .lvwServiciosCarroceria.SelectedItem = .lvwServiciosCarroceria.ListItems(intX - (.lvwOtrosServicios.ListItems.Count + .lvwServiciosTerceros.ListItems.Count))
                    With .lvwServiciosCarroceria
                        ImprimeObjeto 7300 + (intX * 250), 2200, .SelectedItem.SubItems(2), 7
                    End With
                Next
            Else
                For intX = 1 To gintNumeroLineasRecepcion   'numero de lineas a imprimir
                    Set .lvwServiciosCarroceria.SelectedItem = .lvwServiciosCarroceria.ListItems(intX)
                    With .lvwServiciosCarroceria
                        ImprimeObjeto 7400 + (intX * 250), 2200, .SelectedItem.SubItems(2), 7
                    End With
                Next
            End If
        End If

        ImprimeObjeto 13400, 2200, Mid(.txtComentario, 1, 75), 7, False, False
        ImprimeObjeto 13600, 2200, Mid(.txtComentario, 76, 75), 7, False, False
        ImprimeObjeto 13800, 2200, Mid(.txtComentario, 151, 75), 7, False, False
        ImprimeObjeto 14000, 2200, Mid(.txtComentario, 226, 75), 7, False, False
        ImprimeObjeto 14200, 2200, Mid(.txtComentario, 301, 75), 7, False, False
        
        lintNumeroLineasOcupadas = .lvwOtrosServicios.ListItems.Count + .lvwServiciosTerceros.ListItems.Count + .lvwServiciosCarroceria.ListItems.Count
        lintInicioLineaImpresion = 7300 + (intX * 250) + 250
        'repuestos
        If .lvwRepuestos.ListItems.Count > 0 Then
            For i = 1 To .lvwRepuestos.ListItems.Count
                If lintNumeroLineasOcupadas < gintNumeroLineasRecepcion Then
                    ImprimeObjeto lintInicioLineaImpresion, 2200, .lvwRepuestos.ListItems(i).SubItems(1), 7
                    lintNumeroLineasOcupadas = lintNumeroLineasOcupadas + 1
                    lintInicioLineaImpresion = lintInicioLineaImpresion + 240
                Else
                    lintNumeroLineasOcupadas = 1
                    lintInicioLineaImpresion = 7300 + 240
                    Printer.EndDoc
                    ImprimeObjeto lintInicioLineaImpresion, 2200, .lvwRepuestos.ListItems(i).SubItems(1), 7
                    lintInicioLineaImpresion = lintInicioLineaImpresion + 240
                End If
            Next
        End If
        
    '    /////////////////////////////////////////////////
        Printer.EndDoc
    End With
ElseIf Origen = gPresupuesto Then

ElseIf Origen = gOT Then

Else
    Exit Function
End If
End Function
Public Function ImprimirDocumentoKlassik(Origen As gInforme) As Boolean
Dim fintContador As Integer
Dim arrInventario(30) As String
Dim intAux As Integer
Dim strRut As String
Dim lintNumeroLineasOcupadas As Integer
Dim lintInicioLineaImpresion As Integer
Dim i As Integer

If Origen = gRecepcion Then
    With frmRecepcion
       ' /////////////////////ENCABEZADO DE LA RECEPCION
        ImprimeObjeto 1000, 2000, ValorNulo(CStr(Val(Mid(.lblNroRecepcion, 6, 15)))), 14, True '//////////  NRO RECEPCION
        'ImprimeObjeto 2000, 7650, "(" & ValorNulo(.dtcGarantia.Text) & ")", 12  '////////// GARANTIA
        ImprimeObjeto 1000, 3500, ValorNulo(.lblCliente), 7  '//////////  NOMBRE CLIENTE
        ImprimeObjeto 1500, 4000, ValorNulo(.txtDir), 7 '//////////  DIRECCION CLIENTE
        If Not IsNull(.txtRut) Then
            If Trim(.txtRut) <> "" Then
                ImprimeObjeto 2000, 3500, Format(Mid(.txtRut, 1, Len(.txtRut) - 1), "00000000"), 7   '//////////  RUT CLIENTE
                ImprimeObjeto 2000, 4200, "-" & ValorNulo(Mid(.txtRut, Len(.txtRut), 1)), 7
            Else
                ImprimeObjeto 2000, 3500, "SIN DNI", 7   '//////////  RUT CLIENTE
            End If
        Else
            ImprimeObjeto 2000, 3500, "SIN DNI", 7   '//////////  RUT CLIENTE
        End If

        
        ImprimeObjeto 2000, 8000, ValorNulo(.lblFono), 7 '//////////  FONOS
        'ImprimeObjeto 3070, 1500, ValorNulo(.txtSolicita), 7 '//////////  SOLICITADO POR
        'ImprimeObjeto 3070, 9500, ValorNulo(.txtComuna), 7 '//////////  COMUNA CLIENTE

        ImprimeObjeto 2500, 3000, ValorNulo(.lblModelo), 7 '//////////  MARCA MODELO
        ImprimeObjeto 2500, 5700, ValorNulo(.txtAño), 7 '//////////  AÑO
        ImprimeObjeto 2500, 7000, ValorNulo(.lblColorE), 7  '//////////  COLOR EXT / INT
        ImprimeObjeto 2500, 9500, ValorNulo(.txtPatente), 7 '//////////  PATENTE
        ImprimeObjeto 3000, 3000, ValorNulo(.dtcRecepcionista.Text), 7  '////////// RECEPCIONISTA
        ImprimeObjeto 3000, 5700, ValorNulo(.txtKilAct), 7 '//////////  KILOMETROS
        ImprimeObjeto 3050, 7300, ValorNulo(Format(.pckFechaAtencion, "dd/mm/yyyy")), 7 '//////////  FECHA RECEPCION
        ImprimeObjeto 3050, 9500, ValorNulo(Format(.pckFechaEntrega, "dd/mm/yyyy")), 7 '//////////  FECHA ENTREGA
        
        ImprimeObjeto 4500, 1000, ValorNulo(.lblChasis), 7 '//////////  CHASSIS
        ImprimeObjeto 5000, 1000, ValorNulo(.lblMotor), 7 '//////////  MOTOR
        'ImprimeObjeto 4080, 1500, ValorNulo(.txtConcesionario), 7 '////////// CONSECIONARIO
        'ImprimeObjeto 4080, 6000, ValorNulo(Format(.pckFecVta, "dd/mm/yyyy")), 7 '//////////  FECHA VENTA
        'ImprimeObjeto 4080, 9300, ValorNulo(.lblVin), 7 '//////////  VIN
        'ImprimeObjeto 4330, 1500, ValorNulo(.lblCompañia), 7 '//////////  CIA SEGURO
        'ImprimeObjeto 4330, 5800, ValorNulo(.txtLiquidador), 7 '//////////  LIQUIDADOR
        'ImprimeObjeto 4580, 1500, ValorNulo(.txtNroSiniestro), 7 '//////////  SINIESTRO
        'ImprimeObjeto 4580, 5800, ValorNulo(Format(.txtDeduciblePesos, "###,###,##0")), 7 '////////// DEDUCIBLE PESOS

        
        'ImprimeObjeto 4810, 4000, ValorNulo(Format$(Time, "hh:mm")), 7
        
        'ImprimeObjeto 4810, 10100, ValorNulo(.cboHora.Text), 7 '//////////  HORA


      '  //////////////////////////////////////////REVISION
        Dim intX As Integer
        Dim strRevision As String

        With .lvwServiciosMecanica
            If .ListItems.Count > 0 Then
                For intX = 1 To .ListItems.Count
                    Set .SelectedItem = .ListItems(intX)
                    If Mid(.SelectedItem, 1, 2) = "RV" Then
                       strRevision = .SelectedItem.SubItems(1)
                    End If
                Next
                If strRevision <> "" Then
                    strRevision = Mid(strRevision, 9, Len(strRevision) - 11)
                    ImprimeObjeto 5740, 5750, strRevision, 6, True
                End If
            End If

        End With
     '   ////////////////////////// IMPRIME SERVICIOS
        If .lvwOtrosServicios.ListItems.Count > 0 Then  '// OTROS SERVICIOS
            If .lvwOtrosServicios.ListItems.Count <= gintNumeroLineasRecepcion Then
                For intX = 1 To .lvwOtrosServicios.ListItems.Count
                    Set .lvwOtrosServicios.SelectedItem = .lvwOtrosServicios.ListItems(intX)
                    With .lvwOtrosServicios
                        ImprimeObjeto 8500 + (intX * 250), 2700, .SelectedItem.SubItems(1), 7
                    End With
                Next
            Else
                For intX = 1 To gintNumeroLineasRecepcion   'numero de lineas a imprimir
                    Set .lvwOtrosServicios.SelectedItem = .lvwOtrosServicios.ListItems(intX)
                    With .lvwOtrosServicios
                        ImprimeObjeto 8600 + (intX * 250), 2700, .SelectedItem.SubItems(1), 7
                    End With
                Next
            End If
        End If

        If .lvwServiciosTerceros.ListItems.Count > 0 Then  '// SERVICIO DE TERCEROS
            If .lvwServiciosTerceros.ListItems.Count <= gintNumeroLineasRecepcion - .lvwOtrosServicios.ListItems.Count Then
                For intX = .lvwOtrosServicios.ListItems.Count + 1 To .lvwOtrosServicios.ListItems.Count + .lvwServiciosTerceros.ListItems.Count
                    Set .lvwServiciosTerceros.SelectedItem = .lvwServiciosTerceros.ListItems(intX - .lvwOtrosServicios.ListItems.Count)
                    With .lvwServiciosTerceros
                        ImprimeObjeto 8500 + (intX * 250), 2700, .SelectedItem.SubItems(3), 7
                    End With
                Next
            Else
                For intX = 1 To gintNumeroLineasRecepcion   'numero de lineas a imprimir
                    Set .lvwServiciosTerceros.SelectedItem = .lvwServiciosTerceros.ListItems(intX)
                    With .lvwServiciosTerceros
                        ImprimeObjeto 8600 + (intX * 250), 2700, .SelectedItem.SubItems(3), 7
                    End With
                Next
            End If
        End If

        If .lvwServiciosCarroceria.ListItems.Count > 0 Then  '// SERVICIO DE CARROCERIA
            If .lvwServiciosCarroceria.ListItems.Count <= gintNumeroLineasRecepcion - (.lvwOtrosServicios.ListItems.Count - .lvwServiciosCarroceria.ListItems.Count) Then
                For intX = .lvwOtrosServicios.ListItems.Count + .lvwServiciosTerceros.ListItems.Count + 1 To .lvwOtrosServicios.ListItems.Count + .lvwServiciosTerceros.ListItems.Count + .lvwServiciosCarroceria.ListItems.Count
                    Set .lvwServiciosCarroceria.SelectedItem = .lvwServiciosCarroceria.ListItems(intX - (.lvwOtrosServicios.ListItems.Count + .lvwServiciosTerceros.ListItems.Count))
                    With .lvwServiciosCarroceria
                        ImprimeObjeto 8500 + (intX * 250), 2700, .SelectedItem.SubItems(2), 7
                    End With
                Next
            Else
                For intX = 1 To gintNumeroLineasRecepcion   'numero de lineas a imprimir
                    Set .lvwServiciosCarroceria.SelectedItem = .lvwServiciosCarroceria.ListItems(intX)
                    With .lvwServiciosCarroceria
                        ImprimeObjeto 8600 + (intX * 250), 2700, .SelectedItem.SubItems(2), 7
                    End With
                Next
            End If
        End If

        'ImprimeObjeto 13400, 2200, Mid(.txtComentario, 1, 75), 7, False, False
        'ImprimeObjeto 13600, 2200, Mid(.txtComentario, 76, 75), 7, False, False
        'ImprimeObjeto 13800, 2200, Mid(.txtComentario, 151, 75), 7, False, False
        'ImprimeObjeto 14000, 2200, Mid(.txtComentario, 226, 75), 7, False, False
        'ImprimeObjeto 14200, 2200, Mid(.txtComentario, 301, 75), 7, False, False
        
        lintNumeroLineasOcupadas = .lvwOtrosServicios.ListItems.Count + .lvwServiciosTerceros.ListItems.Count + .lvwServiciosCarroceria.ListItems.Count
        lintInicioLineaImpresion = 8500 + (intX * 250) + 250
        'repuestos
        If .lvwRepuestos.ListItems.Count > 0 Then
            For i = 1 To .lvwRepuestos.ListItems.Count
                If lintNumeroLineasOcupadas < gintNumeroLineasRecepcion Then
                    ImprimeObjeto lintInicioLineaImpresion, 2700, .lvwRepuestos.ListItems(i).SubItems(1), 7
                    lintNumeroLineasOcupadas = lintNumeroLineasOcupadas + 1
                    lintInicioLineaImpresion = lintInicioLineaImpresion + 240
                Else
                    lintNumeroLineasOcupadas = 1
                    lintInicioLineaImpresion = 8500 + 240
                    Printer.EndDoc
                    ImprimeObjeto lintInicioLineaImpresion, 2700, .lvwRepuestos.ListItems(i).SubItems(1), 7
                    lintInicioLineaImpresion = lintInicioLineaImpresion + 240
                End If
            Next
        End If
        
    '    /////////////////////////////////////////////////
        Printer.EndDoc
    End With
Else
    Exit Function
End If
End Function


Public Function InicializaCorrelativos() As Boolean
Dim strSql As String

strSql = "UPDATE TLLR_PARAMETROS SET NroRecMec=1,NroRecCar=1,NroPreMec=1,NroPreCar=1,NroOTMec=1,NroOTCar=1"
If Conexion.SendHost(strSql, , , gcTiempoEspera) = apOk Then
    InicializaCorrelativos = True
Else
    InicializaCorrelativos = False
End If

End Function
Public Function NroDiasHabiles(pdteInicio As Date, pdteFinal As Date) As Long
Dim lngDifDias As Long, lngCuenta As Long, lngDias As Long
Dim dteFecha As Date

lngDifDias = DateDiff("d", pdteInicio, pdteFinal)
lngDias = 0
For lngCuenta = 0 To lngDifDias
    dteFecha = DateAdd("d", lngCuenta, pdteInicio)
'    MsgBox dteFecha & "     " & Weekday(dteFecha)
    If Weekday(dteFecha) > 1 And Weekday(dteFecha) < 7 Then
        lngDias = lngDias + 1
    End If
Next

NroDiasHabiles = lngDias

End Function

Public Function TraeCorrelativo(gParametro As gcParametro, strEmpresa As String, strSucursal As String, strSeccion As String) As String
Dim recAux As New ADODB.Recordset
Dim strSql As String
Dim lngNro As Long

strSql = ""
If gParametro = gcOrdenTrabajo Then
    strSql = "SELECT max(ID_OT) AS PARAMETRO FROM TLLR_OT"
ElseIf gParametro = gcPresupuesto Then
    strSql = "SELECT max(ID_PRESUPUESTO) AS PARAMETRO FROM TLLR_OT "
End If

strSql = strSql & " Where id_garantia<>'PRE' and estado <> 'R' and estado <> 'P' and Id_Empresa='" & strEmpresa & "' And Id_Sucursal='" & strSucursal & "' AND SECCION_OT = '" & strSeccion & "'"
If Conexion.SendHost(strSql, recAux, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        lngNro = CLng(Val(Mid(IIf(Not IsNull(recAux!parametro), recAux!parametro, 0), 6, 10)))
        lngNro = lngNro + 1
        TraeCorrelativo = CStr(Year(Now)) & "-" & Format(CStr(lngNro), "0000000000")
    End If
End If
End Function
Public Function TraeCorrelativoReserva(strEmpresa As String, strSucursal As String) As String
Dim recAux As New ADODB.Recordset
Dim strSql As String
Dim lngNro As Long

strSql = "SELECT max(ID_Reserva) AS PARAMETRO FROM TLLR_ReservaHora "

strSql = strSql & " Where Id_Empresa='" & strEmpresa & "'"
If Conexion.SendHost(strSql, recAux, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        lngNro = CLng(Val(IIf(Not IsNull(recAux!parametro), recAux!parametro, 0)))
        lngNro = lngNro + 1
        TraeCorrelativoReserva = Format(CStr(lngNro), "00000")
    End If
End If
End Function
'kjcv 19.04.18
Public Function Lpad(MyValue$, MyPadCharacter$, MyPaddedLength%)

PadLength = MyPaddedLength - Len(MyValue)
Dim PadString As String
For x = 1 To PadLength
   PadString = PadString & MyPadCharacter
Next
Lpad = PadString + MyValue

End Function
Public Function TraeCorrelativoPresupuesto(strEmpresa As String, strSucursal As String, strSeccion As String) As String
Dim recAux As New ADODB.Recordset
Dim strSql As String
Dim lngNro As Long

strSql = "SELECT NroPreCar AS PARAMETRO FROM TLLR_PARAMETRO "

strSql = strSql & " Where Id_Empresa='" & strEmpresa & "' And Id_Sucursal='" & strSucursal & "'"
If Conexion.SendHost(strSql, recAux, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        lngNro = CLng(Val(IIf(Not IsNull(recAux!parametro), recAux!parametro, 0)))
        lngNro = lngNro + 1
        TraeCorrelativoPresupuesto = Format(CStr(lngNro), "000000")
    End If
End If

'actualiza numero de presupuesto en parametros
strSql = "Update Tllr_Parametro Set NroPreCar = " & lngNro
strSql = strSql & " Where Id_Empresa='" & strEmpresa & "' And Id_Sucursal = '" & strSucursal & "'"
If Conexion.SendHost(strSql, , , , gcTiempoEspera) = apAbort Then
    MsgBox "Problemas para Actualizar numero de presupuesto en Parametros", vbExclamation, "Actualizando"
End If

End Function


Public Function FormatOT(pstrOT As String) As String
'FormatOT = CStr(Year(Now)) & "-" & Format(pstrOT, "0000000000")
'kjcv 02.01.13
FormatOT = "" & "-" & Format(pstrOT, "0000000000")
End Function
Public Function FormatPresupuesto(pstrOT As String) As String
FormatPresupuesto = "P-" & Format(pstrOT, "000000")
End Function

Public Function FormatTime(pstrTime As String) As String
Dim strHora As String
strHora = Format(Mid(pstrTime, 1, 2), "00")
FormatTime = strHora & ":" & "00"
End Function

Public Function Atributos(strPrefijoSistema As String, strOpcion As String, ByRef AccesoCrear As Boolean, ByRef AccesoEditar As Boolean, ByRef AccesoBorrar As Boolean, ByRef AccesoImprimir As Boolean) As Boolean
    Dim adoTemp As New ADODB.Recordset
    Dim lstrSQL As String
    Dim lstrPerfil As String

    Atributos = False
    AccesoCrear = False
    AccesoEditar = False
    AccesoBorrar = False
    AccesoImprimir = False
    If Left(UCase(Trim(gstrIdUsuario)), 7) = "SERINFO" Then
        Atributos = True
        AccesoCrear = True
        AccesoEditar = True
        AccesoBorrar = True
        AccesoImprimir = True
    End If
    lstrSQL = "select id_perfil from " & strPrefijoSistema & "_usuario where id_user='" & gstrIdUsuario & "'"
    If Conexion.SendHost(lstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
        'kjcv 08.02.16
            gstrIdPerfil = adoTemp!id_perfil
            lstrSQL = "select * from " & strPrefijoSistema & "_Perfil_Opcion where id_perfil='" & adoTemp!id_perfil & "' and id_opcion='" & strOpcion & "'"
            Conexion.CloseHost adoTemp
            If Conexion.SendHost(lstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
                If Not adoTemp.BOF And Not adoTemp.EOF Then
                    Atributos = IIf(UCase(adoTemp!Opc_Acceso) = "S", True, False)
                    AccesoCrear = IIf(UCase(adoTemp!Opc_Crear) = "S", True, False)
                    AccesoEditar = IIf(UCase(adoTemp!Opc_Editar) = "S", True, False)
                    AccesoBorrar = IIf(UCase(adoTemp!Opc_Borrar) = "S", True, False)
                    AccesoImprimir = IIf(UCase(adoTemp!Opc_Imprimir) = "S", True, False)
                Else
                    Conexion.CloseHost adoTemp
                    Exit Function
                End If
                Conexion.CloseHost adoTemp
            Else
                Conexion.CloseHost adoTemp
                Exit Function
            End If
        Else
            Conexion.CloseHost adoTemp
            Exit Function
        End If
    Else
        Exit Function
    End If
End Function

Public Function TraeCargo(strIdGarantia As String) As String
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT Id_Tipo_Cargo FROM Tllr_Garantias WHERE Id_empresa='" & gstrIdEmpresa & "' and  Id_Garantia = '" & strIdGarantia & "'"
If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        TraeCargo = ValorNulo(recAux!Id_Tipo_Cargo)
    End If
End If

End Function
'kjcv  06.09.16
Public Function TraeDireccion(strIdCliente As String) As String
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT Direccion FROM Glbl_Cliente_Proveedor WHERE Id_Cliente_Proveedor = '" & strIdCliente & "'"
If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        TraeDireccion = ValorNulo(recAux!Direccion)
    End If
End If

End Function
Public Function TraeCargoOT(strIdOt As String, SeccionOT As String) As String
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT Tllr_Garantias.Id_Tipo_Cargo "
strSql = strSql & "FROM Tllr_OT INNER JOIN "
strSql = strSql & "Tllr_Garantias ON "
strSql = strSql & "Tllr_OT.Id_Garantia = Tllr_Garantias.Id_Garantia "
strSql = strSql & "WHERE Tllr_OT.Id_OT='" & strIdOt & "' And "
strSql = strSql & "Tllr_OT.Id_Empresa='" & gstrIdEmpresa & "' AND "
strSql = strSql & "Tllr_OT.Id_Sucursal='" & gstrIdSucursal & "' AND "
strSql = strSql & "Tllr_OT.Seccion_OT='" & SeccionOT & "'"

If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        TraeCargoOT = recAux!Id_Tipo_Cargo
    End If
End If

End Function

Public Function TraeTipoOT(strIdGarantia As String) As String
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT Descripcion FROM Tllr_Garantias WHERE Id_Garantia = '" & strIdGarantia & "'"
If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        TraeTipoOT = recAux!Descripcion
    End If
End If

End Function


Public Function ColorExtDes(strIdColExt As String)
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT Descripcion FROM Glbl_Color_Exterior WHERE Id_Color_Exterior = '" & strIdColExt & "'"
If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        ColorExtDes = recAux!Descripcion
    End If
End If
End Function

Public Function ColorIntDes(strIdColInt As String)
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT Descripcion FROM Glbl_Color_Interior WHERE Id_Color_Interior = '" & strIdColInt & "'"
If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        ColorIntDes = recAux!Descripcion
    End If
End If
End Function

Public Function ConcesionarioDes(strIdCon As String)
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT Razon_Social FROM Glbl_Concesionarios WHERE Id_Concesionario = '" & strIdCon & "'"
If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        ConcesionarioDes = recAux!Razon_Social
    End If
End If
End Function

Public Function ClienteDes(strIdCiente As String)
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT Razon_Social FROM Glbl_Cliente_Proveedor WHERE Id_Cliente_Proveedor = '" & strIdCiente & "'"
If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        ClienteDes = recAux!Razon_Social
    End If
End If
End Function

Public Function CiaSegDes(strIdCiaSeg As String)
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT Nombre FROM Tllr_Compañia_Seguro WHERE Id_Compañia_Seguro = '" & strIdCiaSeg & "'"
If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        CiaSegDes = recAux!Nombre
    End If
End If
End Function

Public Function TraeCargoDes(strIdCargo As String) As String
Dim recAux As New ADODB.Recordset
Dim strSql As String

Set recAux = New ADODB.Recordset

strSql = "SELECT Descripcion FROM Tllr_Tipo_Cargo WHERE Id_tipo_Cargo = '" & strIdCargo & "'"
If Conexion.SendHost(strSql, recAux, adOpenDynamic, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        TraeCargoDes = Trim(recAux!Descripcion)
    End If
End If
Conexion.CloseHost recAux

End Function


Public Function TraeNombreMecanico(strIdMecanico As String) As String
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT Nombre FROM Tllr_Mecanicos WHERE Id_Mecanico = '" & strIdMecanico & "'"
If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        TraeNombreMecanico = recAux!Nombre
    End If
End If
Conexion.CloseHost recAux
End Function



Public Function ProveedorS(strIdProveedor As String) As String
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT TOP 1 Razon_Social From Glbl_Cliente_Proveedor WHERE Id_Cliente_Proveedor = '" & strIdProveedor & "' ORDER BY ID_Cliente_Proveedor"
If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        ProveedorS = recAux!Razon_Social
    End If
End If
Conexion.CloseHost recAux
End Function

Public Function NombreRecepcionista(strIdRecepcionista As String) As String
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "select Nombre from tllr_mecanicos where Es_Recepcionista='S' and Vigencia='S' and Id_Mecanico='" & strIdRecepcionista & "'"
If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        NombreRecepcionista = recAux!Nombre
    End If
End If
Conexion.CloseHost recAux
End Function
Public Function ModeloD(stridMarca As String, stridModelo As String) As String
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT TOP 1 Descripcion FROM GLBL_Modelo WHERE ID_MARCA = '" & stridMarca & "' and ID_Modelo = '" & stridModelo & "' ORDER BY ID_MARCA,id_modelo"
If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        ModeloD = recAux!Descripcion
    End If
End If
Conexion.CloseHost recAux
End Function
Public Function MarcaD(stridMarca As String) As String
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT TOP 1 Descripcion FROM GLBL_Marca WHERE ID_MARCA = '" & stridMarca & "' ORDER BY ID_MARCA"
If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        MarcaD = recAux!Descripcion
    End If
End If
End Function

Public Function ClienteD(strIdCliente As String) As String
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT TOP 1 Razon_Social FROM GLBL_Cliente_Proveedor WHERE ID_Cliente_Proveedor= '" & strIdCliente & "' ORDER BY ID_Cliente_Proveedor"
If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        ClienteD = recAux!Razon_Social
    End If
End If
End Function

Public Function MecanicoD(strIdMecanico As String) As String
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT Nombre From Tllr_Mecanicos WHERE (Id_Mecanico = '" & strIdMecanico & "') ORDER BY Id_Mecanico"

If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        MecanicoD = recAux!Nombre
    Else
        MecanicoD = "(Ninguno)"
    End If
End If
End Function
Public Function IVA(strEmpresa As String, strSucursal As String, gciMode As gcIva) As Double
Dim strSql As String
Dim recAux As New ADODB.Recordset

strSql = "Select IVA as Parametro from Tllr_Parametro"
strSql = strSql & " Where Id_Empresa='" & strEmpresa & "' And Id_Sucursal='" & strSucursal & "' AND ID=1"
If Conexion.SendHost(strSql, recAux, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    With recAux
        If Not .BOF And Not .EOF Then
            IVA = IIf(gciMode = gcIvaUnoPto, 1 + ((!parametro) / 100), ((!parametro) / 100))
        End If
        .Close
    End With
End If
Set recAux = Nothing
End Function


Public Function FamiliaRep(strIdRep As String) As String
Dim strSql As String
Dim recAux As New ADODB.Recordset

strSql = "SELECT Id_Familia FROM Stck_Item "
strSql = strSql & " Where Id_Item='" & strIdRep & "' "
If Conexion.SendHost(strSql, recAux, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    With recAux
        If Not .BOF And Not .EOF Then
            FamiliaRep = IIf(Not IsNull(!Id_Familia), !Id_Familia, "999")
        End If
        .Close
    End With
End If
Set recAux = Nothing
End Function
Public Function TraeIndiceOtrosServicio(strEmpresa As String, strSucursal As String) As String
Dim strSql As String
Dim recAux As New ADODB.Recordset

strSql = "Select CorrelativoOtrosServicios as Parametro from Tllr_Parametro"
strSql = strSql & " Where Id_Empresa='" & strEmpresa & "' And Id_Sucursal='" & strSucursal & "' AND ID=1"
If Conexion.SendHost(strSql, recAux, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    With recAux
        If Not .BOF And Not .EOF Then
            TraeIndiceOtrosServicio = CStr(IIf(IsNull(!parametro), "1", !parametro))
        End If
        .Close
    End With
End If
Set recAux = Nothing
End Function

Public Function TraeIndiceTrabajosTerceros(strEmpresa As String, strSucursal As String) As String
Dim strSql As String
Dim recAux As New ADODB.Recordset

strSql = "Select CorrelativoTrabajoTercero as Parametro from Tllr_Parametro"
strSql = strSql & " Where Id_Empresa='" & strEmpresa & "' And Id_Sucursal='" & strSucursal & "' AND ID=1"
If Conexion.SendHost(strSql, recAux, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    With recAux
        If Not .BOF And Not .EOF Then
            TraeIndiceTrabajosTerceros = CStr(!parametro)
        End If
        .Close
    End With
End If
Set recAux = Nothing
End Function


Public Function ValorPorcentaje(Total As Double, Porcentaje As Single) As Double
ValorPorcentaje = Round((Total * Porcentaje) / 100, gintDecimalesMoneda)
End Function

Public Function PorcentajeMonto(Total As Double, Valor As Single) As Double
If Total <> 0 Then
    PorcentajeMonto = Round((Valor * 100) / Total, 2)
Else
    PorcentajeMonto = 0
End If
End Function

Public Function ValorHora(strIdEmpresa As String, strIdSucursal As String) As Double
Dim recAux As New ADODB.Recordset
Dim strSql As String

strSql = "SELECT PrecioManoObra, PrecioManoObraGarantia From Tllr_Parametro WHERE (Id_Sucursal = '" & strIdSucursal & "') AND (Id_Empresa = '" & strIdEmpresa & "')"
If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        If gstrProcedencia = "Movimientos" Then
            If gblnPreciosMarca = False Then
                ValorHora = IIf(frmRecepcion.dtcGarantia.BoundText <> "GFB", recAux!PrecioManoObra, recAux!PrecioManoObraGarantia)
            Else
                strSql = "SELECT VentaManoObra, VentaMOGarantia From Tllr_Marca_Precios_MO WHERE (Id_Marca = '" & frmRecepcion.lblIdMarca & "')"
                If Conexion.SendHost(strSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                    If Not recAux.BOF And Not recAux.EOF Then
                        ValorHora = IIf(frmRecepcion.dtcGarantia.BoundText = "GFB", recAux!VentaMOGarantia, recAux!VentaManoObra)
                    End If
                End If
            End If
        Else
            ValorHora = recAux!PrecioManoObra
        End If
    End If
End If

End Function

Public Function NombreEmpresa(strCodigo As String) As String
Dim strSql As String
Dim adoTemp As New ADODB.Recordset
strSql = "SELECT Razon_Social as Nombre  From Glbl_Empresa WHERE Id_Empresa = '" & strCodigo & "'"
If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoTemp
        If Not .BOF And Not .EOF Then
            NombreEmpresa = !Nombre
        Else
            NombreEmpresa = "Sin Empresa"
        End If
    End With
End If
End Function

'kjcv 03.07.14
Public Function CodigoEmpleado(strCodigo As String) As String
Dim strSql As String
Dim adoTemp As New ADODB.Recordset

strSql = "SELECT id_Empleado from glbl_usuario where Id_User='" & strCodigo & "' and Id_Empresa='" & gstrIdEmpresa & "'"
   If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoTemp
        If Not .BOF And Not .EOF Then
            CodigoEmpleado = !id_Empleado
'        Else
'            CodigoEmpleado = "Sin Empleado"
        End If
    End With
End If

End Function

Public Function CodigoMecanico(strCodigo As String, strEmpresa As String, strSucursal As String) As String
Dim strSql As String
Dim adoTemp As New ADODB.Recordset

    strSql = " SELECT Id_Mecanico FROM Tllr_Mecanicos WHERE Es_Recepcionista = 'S' AND ID_EMPRESA='" & strEmpresa & "'"
    strSql = strSql & " AND ID_SUCURSAL='" & strSucursal & "' And vigencia='S'"
    strSql = strSql & " AND Rut_Mecanico='" & strCodigo & "'"
    If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With adoTemp
            If Not .BOF And Not .EOF Then
                CodigoMecanico = !Id_Mecanico
            End If
    
        End With
    End If
End Function

Public Function NombreSucursal(strCodigoEmpresa As String, strCodigoSucursal As String) As String
Dim strSql As String
Dim adoTemp As New ADODB.Recordset
strSql = "SELECT Descripcion as Nombre  From Glbl_Sucursal WHERE Id_Empresa = '" & strCodigoEmpresa & "' and Id_sucursal= '" & strCodigoSucursal & "' "
If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoTemp
        If Not .BOF And Not .EOF Then
            NombreSucursal = !Nombre
        Else
            NombreSucursal = "Sin Sucursal"
        End If
    End With
End If
End Function

Public Function NombreCiaSeg(strIdCiaSeg As String) As String
Dim strSql As String
Dim adoTemp As New ADODB.Recordset

strSql = "SELECT Nombre FROM Tllr_Compañia_Seguro WHERE ID_Compañia_Seguro='" & strIdCiaSeg & "' "
If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoTemp
        If Not .BOF And Not .EOF Then
            NombreCiaSeg = !Nombre
        Else
            NombreCiaSeg = "Sin Sucursal"
        End If
    End With
End If
End Function
Public Function DireccionSucursal(strCodigoEmpresa As String, strCodigoSucursal As String) As String
Dim strSql As String
Dim adoTemp As New ADODB.Recordset
strSql = "SELECT Direccion,telefono,Fax  From Glbl_Sucursal WHERE Id_Empresa = '" & strCodigoEmpresa & "' and Id_sucursal= '" & strCodigoSucursal & "' "
If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoTemp
        If Not .BOF And Not .EOF Then
            DireccionSucursal = !Direccion
            gstrTelefono = !Telefono
            gstrFax = IIf(IsNull(!fax), "", !fax)
        Else
            DireccionSucursal = "."
        End If
    End With
End If
End Function
Public Function FormatoValor(vntNumero As Variant, strSigla As String, intDecimal As Integer) As String
    Dim lstrFormato As String

    If intDecimal > 0 Then
        lstrFormato = "." & String(intDecimal, "0")
    End If
    lstrFormato = " #,##0" & lstrFormato & " ;(#,##0" & lstrFormato & "); #,##0" & lstrFormato & " "
    vntNumero = Round(Val(vntNumero), intDecimal)
    FormatoValor = strSigla & Format(vntNumero, lstrFormato)
End Function

Public Function ValorNulo(Valor As Variant) As Variant
    If IsNull(Valor) Then
        ValorNulo = ""
    Else
        ValorNulo = Valor
    End If
End Function
Public Function ParametrosDefecto(strEmpresa As String, strSucursal As String) As Boolean
Dim mstrSQL As String
Dim AdoPrincipal As New ADODB.Recordset

mstrSQL = "SELECT TOP 1 * From Tllr_Parametro"
mstrSQL = mstrSQL & "  WHERE (Id_Empresa = '" & strEmpresa & "') AND (Id_Sucursal = '" & strSucursal & "') AND (Id = 1)"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            gcurPrecioManoObra = IIf(Not IsNull(!PrecioManoObra), !PrecioManoObra, 0)
            gcurCostoManoObra = IIf(Not IsNull(!VALOR_MANO_COSTO), !VALOR_MANO_COSTO, 0)
            gcurInsumo = IIf(Not IsNull(!Insumo), !Insumo, 0)
            gcurSeguroTaller = IIf(Not IsNull(!SeguroTaller), !SeguroTaller, 0)
            gintNroRecDefectoQry = IIf(Not IsNull(!NroRecDefectoQry), !NroRecDefectoQry, 10)
            gstrIdCargoDefecto = IIf(Not IsNull(!IdCargoDefecto), !IdCargoDefecto, "")
            gstrIdTipoOtDefecto = IIf(Not IsNull(!IdTipoOtDefecto), !IdTipoOtDefecto, "")
            'kjcv 21.07.16
            gstrIdCargoInterno = IIf(Not IsNull(!CargoInterno), !CargoInterno, "")
'            gdblIva = IIf(Not IsNull(!IVA), !IVA, 18)
            gstrMecanicoDefectoSecMec = IIf(Not IsNull(!MecanicoDefectoSecMec), !MecanicoDefectoSecMec, "")
            gstrMecanicoDefectoSecCar = IIf(Not IsNull(!MecanicoDefectoSecCar), !MecanicoDefectoSecCar, "")
            gstrMecanicoDefectoSecDes = IIf(Not IsNull(!MecanicoDefectoSecDes), !MecanicoDefectoSecDes, "")
            gstrMecanicoDefectoSecPin = IIf(Not IsNull(!MecanicoDefectoSecPin), !MecanicoDefectoSecPin, "")
            gdblNroHorOblg = IIf(Not IsNull(!NroHorasTrabajo), !NroHorasTrabajo, 8)
            gdblValorExistencia = IIf(Not IsNull(!Valor_Existencia), !Valor_Existencia, 2000000)
            gintNumeroLineasRecepcion = IIf(Not IsNull(!LineasRecepcion), !LineasRecepcion, 9)
            'kjcv 09.07.15
            gstrPorPrecioGtia = IIf(Not IsNull(!porc_precio_gtia), !porc_precio_gtia, 0)
            If Not IsNull(!Estadoprodmecanico) Then
                If !Estadoprodmecanico = "F" Then
                    gstrEstadoProdMecanico = !Estadoprodmecanico ' "IN('B','F')"
                ElseIf !Estadoprodmecanico = "L" Then
                    gstrEstadoProdMecanico = !Estadoprodmecanico '"IN('L','C')"
                Else
                    gstrEstadoProdMecanico = !Estadoprodmecanico '"IN('B','F','L','C')"
                End If
            Else
                gstrEstadoProdMecanico = "A"   '"IN('B','F','L','C')"
            End If
            gblnEnviaMailBodega = IIf(!EnviaMailBodega = "S", True, False)
            gintHoraInicio = IIf(Not IsNull(!HoraInicio), !HoraInicio, 8)
            gintHoratermino = IIf(Not IsNull(!HoraTermino), !HoraTermino, 20)
            gintIntervaloMinutos = IIf(Not IsNull(!IntervaloMinutos), !IntervaloMinutos, 30)
            gstrMecanicoDiasHabiles = IIf(Not IsNull(!MecanicoDiasHabiles), !MecanicoDiasHabiles, "")
            gblnTraspasaRepuestos = IIf(!TraspasaRepuestos = "S", True, False)
            gintDescuentoMaximo = IIf(Not IsNull(!DescuentoMaximo), !DescuentoMaximo, 15)
            'kjcv 13.03.17
            gintDescuentoMaximoCIA = IIf(Not IsNull(!DsctMaxCiaSeg), !DsctMaxCiaSeg, 15)
            gstrNotaRecepcion = IIf(Not IsNull(!NotaRecepcion), !NotaRecepcion, "")
            gstrNotaPresupuesto = IIf(Not IsNull(!NotaPresupuesto), !NotaPresupuesto, "")
            gcurMaterialesMO = IIf(Not IsNull(!MaterialesMO), !MaterialesMO, 0)
            'gcurMaterialesPesos = IIf(Not IsNull(!MaterialesPesos), !MaterialesPesos, 0)
            gstrMailRepuestosFallidos = IIf(Not IsNull(!MailRepuestosFallidos), !MailRepuestosFallidos, "")
            gstrCodigoLubricantes = IIf(Not IsNull(!CodFamiliaLubricantes), !CodFamiliaLubricantes, "0")
            gstrCodigoMateriales = IIf(Not IsNull(!CodFamiliaMateriales), !CodFamiliaMateriales, "0")
            gstrCodigoInsumos = IIf(Not IsNull(!CodFamiliaInsumos), !CodFamiliaInsumos, "0")
            gblnImprimeImagen = IIf(!ImprimeImagen = "S", True, False)
            gblnValidaCostoRepuestos = IIf(!ValidaCostoRepuestos = "S", True, False)
            gstrMonedaLocal = IIf(Not IsNull(!Id_Moneda_Local), Retorna_Valor_General("Select Sigla from glbl_moneda where id_moneda='" & !Id_Moneda_Local & "'", gcdynamic), "$")
            gintDecimalesMoneda = IIf(Not IsNull(!DecimalesMoneda), !DecimalesMoneda, 0)
            gblnPreciosMarca = IIf(!PreciosMarca = "S", True, False)
            gblnBloqueaSubtotalRep = IIf(!BloqueaSubtotalRep = "S", True, False)
            gblnValidaServiciosCero = IIf(!ValidaServiciosCero = "S", True, False)
            gstrServiciosMarca = IIf(UCase(!ServiciosMarca) = "S", "S", "N")
            gstrCargoDeducibleMas = IIf(Not IsNull(!CargoDeducibleMas), !CargoDeducibleMas, "")
            gstrCargoDeducibleMenos = IIf(Not IsNull(!CargoDeducibleMenos), !CargoDeducibleMenos, "")
            gstrAsignaRecursos = IIf(UCase(!AsignaRecursos) = "S", "S", "N")
            gstrCargoGtiaFabrica = IIf(Not IsNull(!CargoGarantiaFabrica), !CargoGarantiaFabrica, "GFB")
            ParametrosDefecto = True
        Else
            ParametrosDefecto = False
        End If
    End With
Else
    ParametrosDefecto = False
End If
End Function

Public Function FormatoRut(strCodigo As String) As String
    If Len(Trim(strCodigo)) > 1 Then
        If UCase(Trim(gstrEditaRut)) = "S" Then
            FormatoRut = Format(Left(Trim(strCodigo), Len(Trim(strCodigo)) - 1), "#,###") & "-" & Right(Trim(strCodigo), 1)
        Else
            FormatoRut = Trim(strCodigo)
        End If
    Else
        FormatoRut = Trim(strCodigo)
    End If
End Function
'kjcv 10.06.14
Public Function MarcaxDefault()
Dim Sql As String
Dim adoMarca As New ADODB.Recordset

Sql = "Select * from Elisa_Parametros where Id_Empresa='" & gstrIdEmpresa & "' and Id_Sucursal='" & gstrIdSucursal & "'"
If Conexion.SendHost(Sql, adoMarca, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not adoMarca.BOF And Not adoMarca.EOF Then
        adoMarca.MoveFirst
        strIdMarcaDefecto = adoMarca!Id_Marca_Defecto
    End If
End If

End Function

Public Function ParametrosInternacionales(strEmpresa As String) As Boolean
Dim mstrSQL As String
Dim AdoPrincipal As New ADODB.Recordset

mstrSQL = "SELECT * From Glbl_Parametros_Internacionales"
mstrSQL = mstrSQL & "  WHERE Id_Empresa = '" & strEmpresa & "'"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            gstrNombreRut = IIf(Not IsNull(!Nombre_Rut), !Nombre_Rut, "DNI")
            gstrValidaRut = IIf(Not IsNull(!Valida_Rut), !Valida_Rut, "S")
            gstrEditaRut = IIf(Not IsNull(!Edita_Rut), !Valida_Rut, "S")
            gstrNombrePatente = IIf(Not IsNull(!Nombre_Patente), !Nombre_Patente, "Placa")
            gstrValidaPatente = IIf(Not IsNull(!Valida_Patente), !Valida_Patente, "S")
            gstrNombreIva = IIf(Not IsNull(!Nombre_Iva), !Nombre_Iva, "Iva")
            gstrNombreDP = IIf(Not IsNull(!Nombre_DP), !Nombre_DP, "Desabolladura")
            gstrNombreComuna = IIf(Not IsNull(!Nombre_Comuna), !Nombre_Comuna, "Distrito")
            gstrNombreCiudad = IIf(Not IsNull(!Nombre_Ciudad), !Nombre_Ciudad, "Provincia")
            gstrNombreSucursal = IIf(Not IsNull(!Nombre_Sucursal), !Nombre_Sucursal, "Sucursal")
            gstrNombreBodega = IIf(Not IsNull(!Nombre_Bodega), !Nombre_Bodega, "Bodega")
            ParametrosInternacionales = True
        Else
            ParametrosInternacionales = False
        End If
    End With
Else
    ParametrosInternacionales = False
End If
End Function

Public Function SacarFormatoValor(strValor As String, strSigla As String) As String
    Dim lintI As Integer, lintJ As Integer
    Dim strValor1 As String

    If strValor = "" Then
        SacarFormatoValor = "0"   'strValor
        Exit Function
    End If
    If strSigla <> "" Then
        Do
            lintI = InStr(1, strValor, strSigla)
            If lintI <> 0 Then
                strValor1 = ""
                For lintJ = 1 To Len(strValor)
                    If Not (lintJ >= lintI And lintJ <= (lintI + Len(strSigla)) - 1) Then
                        strValor1 = strValor1 & Mid(strValor, lintJ, 1)
                    End If
                Next
                strValor = strValor1
            Else
              Exit Do
            End If
        Loop
    End If
    SacarFormatoValor = CDbl(strValor)
End Function
Public Function ValorServicioCarroceria()

End Function
Public Function TipoConcepto(strIdConcepto As String) As String
Dim mstrSQL As String
Dim AdoPrincipal As New ADODB.Recordset

mstrSQL = "SELECT TOP 1 D_P AS TIPO FROM Tllr_Concepto WHERE ID_CONCEPTO='" & strIdConcepto & "'"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
        If Not .BOF And Not .EOF Then
            TipoConcepto = !Tipo
        Else
            TipoConcepto = "N"
        End If
    End With
End If
End Function

Public Function ConsultaVehiculo(strVehiculo As String) As Boolean
gstrSql = "SELECT COUNT(*) AS EXISTE FROM Tllr_Vehiculo_Cliente WHERE (Patente = '" & strVehiculo & "')"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            If !existe > 0 Then
                ConsultaVehiculo = True
            Else
                ConsultaVehiculo = False
            End If
        End If
    End With
End If
Conexion.CloseHost gadoPrincipal
End Function
'kjcv 15.11.13
Public Function ConsultaPatente(strPatente As String) As Boolean
gstrSql = "SELECT Patente FROM Tllr_Vehiculo_Cliente WHERE (Patente = '" & strPatente & "') AND Cliente_Problema='S'"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            If !Patente > 0 Then
                ConsultaPatente = True
            Else
                ConsultaPatente = False
            End If
        End If
    End With
End If
Conexion.CloseHost gadoPrincipal
End Function
'kjcv 30.10.15
Public Function ConsultaCliente(strCodigo As String) As Boolean
gstrSql = "SELECT Id_Cliente_Proveedor as Codigo FROM Glbl_Cliente_Proveedor WHERE Id_Cliente_Proveedor='" & strCodigo & "' AND Cliente_Problema='S' "
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            If !Codigo > 0 Then
                ConsultaCliente = True
            Else
                ConsultaCliente = False
            End If
        End If
    End With
End If
Conexion.CloseHost gadoPrincipal
End Function
Public Function ConsultaVehiculoPropio(strVehiculo As String) As Boolean
gstrSql = "SELECT COUNT(*) AS EXISTE FROM Tllr_Vehiculo_Propio WHERE (id_vehiculo = '" & strVehiculo & "')"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            If !existe > 0 Then
                ConsultaVehiculoPropio = True
            Else
                ConsultaVehiculoPropio = False
            End If
        End If
    End With
End If
Conexion.CloseHost gadoPrincipal
End Function

Public Function ConsultaVinExistencia(strVehiculo As String) As Boolean

'Dim RescataValorExistencia As Double
'Dim Parametros As TIPO_PARAMETROS_CONTABLES
'
''//// RESCATA CONDICION VEHICULO Y TIPO VEHICULO DE STOCK
'
'gstrSql = "SELECT AUTO_STOCK.ID_CONDICION_VEHICULO, AUTO_STOCK.ID_MARCA, GLBL_MODELO.ID_TIPOVEHICULO FROM AUTO_STOCK"
'gstrSql = gstrSql & " INNER JOIN GLBL_MODELO ON GLBL_MODELO.ID_MARCA = AUTO_STOCK.ID_MARCA AND GLBL_MODELO.ID_MODELO = AUTO_STOCK.ID_MODELO"
'gstrSql = gstrSql & " WHERE AUTO_STOCK.VIN LIKE '%" & strVehiculo & "'"
'
'If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
'    With gadoPrincipal
'        If Not .BOF And Not .EOF Then
'          '//// VERIFICA SI SE ENCUENTRA EN EXISTENCIA (CONT)
'          TraeParametrosContables Parametros, gstrIdSucursal, gstrIdEmpresa
'          RescataValorExistencia = TraeCostoContableVehiculos(TraeCuentaCostoVehiculoDesde(gstrIdEmpresa, gstrIdSucursal, ValorNulo(!Id_Marca), ValorNulo(!id_tipovehiculo), ValorNulo(!ID_CONDICION_VEHICULO)), TraeCuentaCostoVehiculoHasta(gstrIdEmpresa, gstrIdSucursal, ValorNulo(!Id_Marca), ValorNulo(!id_tipovehiculo), ValorNulo(!ID_CONDICION_VEHICULO)), APPatente.VintoRut(Right(ValorNulo(strVehiculo), 7)), gstrIdEmpresa, Parametros)
'          If RescataValorExistencia < gdblValorExistencia Then
'            MsgBox "El VIN Ingresado No Esta en Existencia", vbInformation, "Vin en Stock"
'            ConsultaVinExistencia = False
'          Else
'            ConsultaVinExistencia = True
'          End If
'        Else
'            MsgBox "El VIN Ingresado No existe en Stock", vbInformation, "Vin en Stock"
'            ConsultaVinExistencia = False
'        End If
'    End With
'End If
'Conexion.CloseHost gadoPrincipal
End Function
Function TraeCostoContableVehiculos(CuentaDesde As String, CuentaHasta As String, rut As String, strIdEmpresa As String, Parametros As TIPO_PARAMETROS_CONTABLES) As Double
    Dim adoTem As New ADODB.Recordset
    Dim lstrSQL As String
    
    Set adoTem = New ADODB.Recordset
    
    lstrSQL = "SELECT SUM(Cont_Comprobante_Contable_Detalle.DEBE)"
    lstrSQL = lstrSQL & " - SUM(Cont_Comprobante_Contable_Detalle.HABER)"
    lstrSQL = lstrSQL & " AS Saldo"
    lstrSQL = lstrSQL & " FROM Cont_Comprobante_Contable LEFT OUTER JOIN"
    lstrSQL = lstrSQL & " Cont_Comprobante_Contable_Detalle ON"
    lstrSQL = lstrSQL & " Cont_Comprobante_Contable.id_Tipo_Comprobante = Cont_Comprobante_Contable_Detalle.id_Tipo_Comprobante"
    lstrSQL = lstrSQL & " AND"
    lstrSQL = lstrSQL & " Cont_Comprobante_Contable.id_Folio = Cont_Comprobante_Contable_Detalle.id_Folio"
    lstrSQL = lstrSQL & " AND"
    lstrSQL = lstrSQL & " Cont_Comprobante_Contable.id_Empresa = Cont_Comprobante_Contable_Detalle.id_Empresa"
    lstrSQL = lstrSQL & " WHERE (Cont_Comprobante_Contable.ESTADO = '1') AND"
    lstrSQL = lstrSQL & " (Cont_Comprobante_Contable_Detalle.CUENTA BETWEEN"
    lstrSQL = lstrSQL & " '" & CuentaDesde & "' and '" & CuentaHasta & "' ) AND"
    lstrSQL = lstrSQL & " (Cont_Comprobante_Contable_Detalle.AUXILIAR =  '" & SacarFormatoRut(rut) & "')"
'    lstrSql = lstrSql & " AND"
'    lstrSql = lstrSql & " ("
'    lstrSql = lstrSql & " (Cont_Comprobante_Contable_Detalle.id_Tipo_Comprobante"
'    lstrSql = lstrSql & " = '" & Parametros.Cont_id_Tipo_Comprobante_Diario & "' OR Cont_Comprobante_Contable_Detalle.id_Tipo_Comprobante"
'    lstrSql = lstrSql & " = '" & Parametros.Cont_id_Tipo_Comprobante_Mes_Anterior & "'))"
    TraeCostoContableVehiculos = 0
    If Conexion.SendHost(lstrSQL, adoTem, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        If Not adoTem.BOF And Not adoTem.EOF Then
            TraeCostoContableVehiculos = IIf(IsNull(adoTem!Saldo), 0, adoTem!Saldo)
        End If
    End If
    Conexion.CloseHost adoTem
End Function
Function TraeCuentaCostoVehiculoDesde(strIdEmpresa As String, strIdSucursal As String, stridMarca As String, strIdTipoVehiculo As String, stridCondicionVehiculo As String) As String
    Dim strSql As String
    Dim adoTemp As New ADODB.Recordset
    
    strSql = "SELECT Glbl_Tipo_Vehiculo_Cta_Existencia.Cuenta_Costo_Desde FROM Glbl_Tipo_Vehiculo_Cta_Existencia WHERE id_Empresa = '" & strIdEmpresa & "' and id_Sucursal = '" & strIdSucursal & "' and id_Marca = '" & stridMarca & "' and Id_Tipo_Vehiculo = '" & strIdTipoVehiculo & "' and Id_Condicion_Vehiculo = '" & stridCondicionVehiculo & "'"
    
    If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            TraeCuentaCostoVehiculoDesde = IIf(IsNull(adoTemp!Cuenta_Costo_Desde), "", adoTemp!Cuenta_Costo_Desde)
        End If
    End If
    Conexion.CloseHost adoTemp
End Function
Function TraeCuentaCostoVehiculoHasta(strIdEmpresa As String, strIdSucursal As String, stridMarca As String, strIdTipoVehiculo As String, stridCondicionVehiculo As String) As String
    Dim strSql As String
    Dim adoTemp As New ADODB.Recordset
    
    strSql = "SELECT Glbl_Tipo_Vehiculo_Cta_Existencia.Cuenta_Costo_Hasta FROM Glbl_Tipo_Vehiculo_Cta_Existencia WHERE id_Empresa = '" & strIdEmpresa & "' and id_Sucursal = '" & strIdSucursal & "' and id_Marca = '" & stridMarca & "' and Id_Tipo_Vehiculo = '" & strIdTipoVehiculo & "' and Id_Condicion_Vehiculo = '" & stridCondicionVehiculo & "'"
    
    If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            TraeCuentaCostoVehiculoHasta = IIf(IsNull(adoTemp!Cuenta_Costo_Hasta), "", adoTemp!Cuenta_Costo_Hasta)
        End If
    End If
    Conexion.CloseHost adoTemp
End Function
Function TraeParametrosContables(Tabla As TIPO_PARAMETROS_CONTABLES, strIdSucursal As String, strIdEmpresa As String) As Boolean
    Dim strSql As String
    Dim adoTemp As New ADODB.Recordset
    
    strSql = "SELECT * FROM Auto_Parametros WHERE Id_Empresa = '" & strIdEmpresa & "' and Id_Sucursal = '" & strIdSucursal & "'"
    
    If Conexion.SendHost(strSql, adoTemp, adOpenForwardOnly, adLockReadOnly, 10) = apOk Then
        Tabla.Cont_id_Tipo_Comprobante_Diario = IIf(IsNull(adoTemp!Cont_id_Tipo_Comprobante_Diario), "", adoTemp!Cont_id_Tipo_Comprobante_Diario)
        Tabla.Cont_id_Tipo_Comprobante_Mes_Anterior = IIf(IsNull(adoTemp!Cont_id_Tipo_Comprobante_Mes_Anterior), "", adoTemp!Cont_id_Tipo_Comprobante_Mes_Anterior)
        Tabla.Cont_id_Tipo_Docto = IIf(IsNull(adoTemp!Cont_id_Tipo_Docto), "", adoTemp!Cont_id_Tipo_Docto)
        Tabla.Cont_Iva = IIf(IsNull(adoTemp!Cont_Iva), "", adoTemp!Cont_Iva)
        Tabla.Cont_Proveedor = IIf(IsNull(adoTemp!Cont_Proveedor), "", adoTemp!Cont_Proveedor)
        Tabla.Cont_id_Tipo_Auxiliar = IIf(IsNull(adoTemp!Cont_id_Tipo_Auxiliar), "", adoTemp!Cont_id_Tipo_Auxiliar)
'        ParamGlob.strIdMarcaDefecto = adoTemp!CodigoMarcaVehiculo
        TraeParametrosContables = True
    Else
        TraeParametrosContables = False
    End If
    
    Conexion.CloseHost adoTemp
End Function

Public Function SacarFormatoRut(txtCodigo As String) As String
    If txtCodigo = "" Then
        SacarFormatoRut = txtCodigo
        Exit Function
    End If
    
    Dim i As Integer
    SacarFormatoRut = ""
    For i = 1 To Len(txtCodigo)
        If Mid(txtCodigo, i, 1) <> "-" And Mid(txtCodigo, i, 1) <> "," And Mid(txtCodigo, i, 1) <> "." Then
            SacarFormatoRut = SacarFormatoRut & Mid(txtCodigo, i, 1)
        End If
    Next
End Function

Public Function TraeTipoCargoAsociado(strGarantia As String) As String
gstrSql = "SELECT Id_Tipo_Cargo From Tllr_Garantias WHERE Id_Garantia = '" & strGarantia & "'"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        If Not .BOF And Not .EOF Then
            TraeTipoCargoAsociado = !Id_Tipo_Cargo
        End If
    End With
End If
Conexion.CloseHost gadoPrincipal
End Function

Public Function TraeHorasDefinidas(strCiaSeg As String, strConcepto As String, strPartePieza As String) As Currency

gstrSql = "SELECT Horas FROM Tllr_CiaSeguro_Concepto_Parte_Pieza"
gstrSql = gstrSql & " WHERE Id_Compañia_Seguro = '" & strCiaSeg & "' AND Id_Concepto = '" & strConcepto & "' AND Id_Parte_Pieza = '" & strPartePieza & "'"

If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        If Not .BOF And Not .EOF Then
            TraeHorasDefinidas = IIf(Not IsNull(!Horas), !Horas, 0)
        Else
            TraeHorasDefinidas = 0
        End If
    End With
End If
End Function
Public Function TraeValorDefinido(strCiaSeg As String, strConcepto As String, strPartePieza As String) As Currency

gstrSql = "SELECT Valor FROM Tllr_CiaSeguro_Concepto_Parte_Pieza"
gstrSql = gstrSql & " WHERE Id_Compañia_Seguro = '" & strCiaSeg & "' AND Id_Concepto = '" & strConcepto & "' AND Id_Parte_Pieza = '" & strPartePieza & "'"

If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        If Not .BOF And Not .EOF Then
            TraeValorDefinido = IIf(Not IsNull(!Valor), !Valor, 0)
        Else
            TraeValorDefinido = 0
        End If
    End With
End If
End Function
'kjcv 16.07.13 Tipo de Cambio de Compañia de Seguro
Public Function traeParidadMonedaMesCS(codMoneda As String, dFecha As Date, CodComp As String, Id_Empresa As String) As Double
Dim tablaMONEDA As New ADODB.Recordset
Dim Sql As String

Sql = ""
Sql = " SELECT Paridad FROM Glbl_Moneda_Tipo_Cambio_Mes_CS "
Sql = Sql & " WHERE id_Moneda='" & codMoneda & "' "
'sql = sql & " AND Año='" & Format(dFecha, "YYYY") & "' "
Sql = Sql & " AND Anio='" & Format(dFecha, "YYYY") & "' "
Sql = Sql & " AND Mes='" & Format(dFecha, "MM") & "'"
Sql = Sql & "AND id_Compañia_Seguro='" & CodComp & "' and Id_empresa='" & Id_Empresa & "'"
If Conexion.SendHost(Sql, tablaMONEDA, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then
        If tablaMONEDA.RecordCount <> 0 Then
            traeParidadMonedaMesCS = tablaMONEDA!Paridad
            Exit Function
        End If
End If
Conexion.CloseHost tablaMONEDA
End Function
'kjcv 08.09.16
'selecciona Valor Hora de Compañia Seguros
Public Function traeValorHoraCS(codCS As String, Id_Empresa As String) As Double
Dim tablaValorHora As New ADODB.Recordset
Dim Sql As String

Sql = ""
Sql = " SELECT isnull(Valor_Hora_Defecto,0) as ValorHora FROM Tllr_Compañia_Seguro "
Sql = Sql & " WHERE Id_Empresa='" & Id_Empresa & "' and Id_Compañia_Seguro='" & codCS & "' "
If Conexion.SendHost(Sql, tablaValorHora, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then

        If tablaValorHora.RecordCount <> 0 Then
            traeValorHoraCS = tablaValorHora!ValorHora
            Exit Function
        End If

End If
Conexion.CloseHost tablaValorHora


End Function

'kjcv 13.08.12
Public Function traeParidadMonedaMes(codMoneda As String, dFecha As Date) As Double
Dim tablaMONEDA As New ADODB.Recordset
Dim Sql As String

Sql = ""
Sql = " SELECT Paridad FROM Glbl_Moneda_Tipo_Cambio_Mes "
Sql = Sql & " WHERE id_Moneda='" & codMoneda & "' "
'sql = sql & " AND Año='" & Format(dFecha, "YYYY") & "' "
Sql = Sql & " AND Anio='" & Format(dFecha, "YYYY") & "' "
Sql = Sql & " AND Mes='" & Format(dFecha, "MM") & "'"
If Conexion.SendHost(Sql, tablaMONEDA, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then

        If tablaMONEDA.RecordCount <> 0 Then
            traeParidadMonedaMes = tablaMONEDA!Paridad
            Exit Function
        End If

End If
Conexion.CloseHost tablaMONEDA


End Function
'kjcv 21.10.15 Tipo de Cambio de Garantia Fabrica
Public Function traeParidadMonedaMesGarantia(codMoneda As String, dFecha As Date) As Double
Dim tablaMONEDA As New ADODB.Recordset
Dim Sql As String

Sql = ""
Sql = " SELECT isnull(Paridad,0) as Paridad  FROM Glbl_Moneda_Tipo_Cambio_Mes_Garantia "
Sql = Sql & " WHERE id_Moneda='" & codMoneda & "' "
Sql = Sql & " AND Anio='" & Format(dFecha, "YYYY") & "' "
Sql = Sql & " AND Mes='" & Format(dFecha, "MM") & "'"
If Conexion.SendHost(Sql, tablaMONEDA, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then
        If tablaMONEDA.RecordCount <> 0 Then
            traeParidadMonedaMesGarantia = tablaMONEDA!Paridad
            Exit Function
        End If
End If
Conexion.CloseHost tablaMONEDA
End Function

'kjcv 06.11.12
Public Function traeParidadMoneda(codMoneda As String) As Double
Dim tablaMONEDA As New ADODB.Recordset
Dim Sql As String

Sql = ""
Sql = " SELECT Paridad FROM Glbl_Moneda "
Sql = Sql & " WHERE id_Moneda='" & codMoneda & "'  AND vigencia='S'"
If Conexion.SendHost(Sql, tablaMONEDA, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then

        If tablaMONEDA.RecordCount <> 0 Then
            traeParidadMoneda = tablaMONEDA!Paridad
            Exit Function
        End If

End If
Conexion.CloseHost tablaMONEDA


End Function

Public Function VerificaRepuesto(pstrIdRepuesto As String, pstrNroOT As String, pstrSeccion As String, pstrNombreTabla As String) As Boolean

gstrSql = "SELECT COUNT(*) AS CUANTOS"
gstrSql = gstrSql & " From " & pstrNombreTabla
gstrSql = gstrSql & " WHERE (Id_Empresa = '" & gstrIdEmpresa & "') AND (Id_Sucursal = '" & gstrIdSucursal & "') AND"
gstrSql = gstrSql & " (Id_OT = '" & pstrNroOT & "') AND (Seccion_OT = '" & pstrSeccion & "') AND (Id_Item = '" & pstrIdRepuesto & "')"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
With gadoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        If !CUANTOS > 0 Then
            VerificaRepuesto = True
        Else
            VerificaRepuesto = False
        End If
    End If
End With
End If
End Function
Function ValidaServicioMecanica(pstrIdOT As String, pstrSeccion As String, pstrIdMarca As String, pstrIdModelo As String, pstrIdServicio As String) As Boolean

gstrSql = "SELECT COUNT(*) AS CUANTOS"
gstrSql = gstrSql & " From Tllr_Mecanica_Ot "
gstrSql = gstrSql & " WHERE (Id_Empresa = '" & gstrIdEmpresa & "') AND (Id_Sucursal = '" & gstrIdSucursal & "') AND"
gstrSql = gstrSql & " (Id_OT = '" & pstrIdOT & "') AND (Seccion_OT = '" & pstrSeccion & "') AND (Id_Marca = '" & pstrIdMarca & "') And"
gstrSql = gstrSql & " (Id_Modelo = '" & pstrIdModelo & "') And (Id_Servicio = '" & pstrIdServicio & "')"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
With gadoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        If !CUANTOS > 0 Then
            ValidaServicioMecanica = True
        Else
            ValidaServicioMecanica = False
        End If
    End If
End With
End If
End Function
Public Function VerificaMarcaConcesionario(pstrIdConcesionario As String, pstrIdMarca As String) As Boolean

gstrSql = "SELECT COUNT(*) AS CUANTOS"
gstrSql = gstrSql & " From Glbl_Concesionarios_Vs_Marca"
gstrSql = gstrSql & " WHERE (Id_Concesionario = '" & pstrIdConcesionario & "') AND (Id_Marca = '" & pstrIdMarca & "')"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
With gadoPrincipal
    If Not .BOF And Not .EOF Then
        .MoveFirst
        If !CUANTOS > 0 Then
            VerificaMarcaConcesionario = True
        Else
            VerificaMarcaConcesionario = False
        End If
    End If
End With
End If
End Function

Public Function BOM(Fecha As Date) As Date  '//Bof of Month
    BOM = DateValue("01/" & Month(Fecha) & "/" & Year(Fecha))
End Function
Public Function EOM(Fecha As Date) As Date '//End of Month
    EOM = DateValue("01/" & IIf(Month(Fecha) = 12, 1, Month(Fecha) + 1) & "/" & IIf(Month(Fecha) = 12, Year(Fecha) + 1, Year(Fecha))) - 1
End Function
Public Sub ReOrdenaLista(ByRef Lista As ListView, ByVal Cabecera As MSComctlLib.ColumnHeader)
    If Lista.SortKey = Cabecera.Index - 1 Then
        If Lista.SortOrder = lvwAscending Then
            Lista.SortOrder = lvwDescending
        Else
            Lista.SortOrder = lvwAscending
        End If
    Else
        Lista.Sorted = False
        Lista.SortKey = Cabecera.Index - 1
        Lista.SortOrder = lvwAscending
        Lista.Sorted = True
    End If
End Sub
Public Sub ReOrdenaListaNumero(ByRef Lista As ListView, Cabecera As Integer)
    If Lista.SortKey = Cabecera Then
        If Lista.SortOrder = lvwAscending Then
            Lista.SortOrder = lvwDescending
        Else
            Lista.SortOrder = lvwAscending
        End If
    Else
        Lista.Sorted = False
        Lista.SortKey = Cabecera
        Lista.SortOrder = lvwAscending
        Lista.Sorted = True
    End If
End Sub


Public Function VeriLiq() As Boolean
Dim strAux As String
gflag = False

Screen.MousePointer = 1
frmPermiso.Show 1

If UCase(gstrVerificacion) = gstrPassWordLiquidador Then
'If NoEsLaPassword(gstrVerificacion, gstrVerificaMecanico) Then
    VeriLiq = True
    gflag = True
Else
    VeriLiq = False
    gflag = False
End If

End Function

Public Function NoEsLaPassword(pstrpassword As Integer, pstrIdMecanico As String)

    gstrSql = "SELECT passwordliquidador FROM Tllr_mecanicos"
    gstrSql = gstrSql & " WHERE Vigencia='S' and Id_mecanico = '" & pstrIdMecanico & "' And id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With gadoPrincipal
            If Not .BOF And Not .EOF Then
                If pstrpassword = !PasswordLiquidador Then
                    NoEsLaPassword = True
                Else
                    NoEsLaPassword = False
                End If
            Else
                NoEsLaPassword = False
            End If
        End With
    End If
    
    Conexion.CloseHost gadoPrincipal
End Function


Public Function CorrelativoOrdenCompra(pstrEmpresa As String, pstrSucursal As String) As Long

gstrSql = "SELECT MAX(ID_ORDEN) AS MAXIMO FROM TLLR_ORDEN_COMPRA WHERE ID_EMPRESA = '" & pstrEmpresa & "' AND ID_SUCURSAL = '" & pstrSucursal & "'"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            CorrelativoOrdenCompra = IIf(Not IsNull(!MAXIMO), !MAXIMO + 1, 1)
        Else
            CorrelativoOrdenCompra = 1
        End If

    End With
End If
Conexion.CloseHost gadoPrincipal

End Function
Function TipoImpresion() As String

    Dim tbRegistros As New ADODB.Recordset
    Dim lstrSQL As String
    lstrSQL = "SELECT * FROM Tllr_Parametro WHERE id_sucursal = '" & gstrIdSucursal & "' and id_empresa = '" & gstrIdEmpresa & "'"
    If Conexion.SendHost(lstrSQL, tbRegistros, adOpenKeyset, adLockReadOnly, 10) = apOk Then
        If tbRegistros.RecordCount > 0 Then
            TipoImpresion = IIf(IsNull(UCase$(tbRegistros!TipoImpresion)), "C", UCase$(tbRegistros!TipoImpresion))
        Else
            TipoImpresion = "C"
        End If
    End If
    Conexion.CloseHost tbRegistros
End Function

Function VeriDeducible(dStrOT, dStrSeccion) As String

    Dim tbRegistros As New ADODB.Recordset
    Dim lstrSQL As String
    lstrSQL = "SELECT deducible_facturado FROM Tllr_OT WHERE id_sucursal = '" & gstrIdSucursal & "' and id_empresa = '" & gstrIdEmpresa & "' and id_ot='" & dStrOT & "' and seccion_ot='" & dStrSeccion & "'"
    If Conexion.SendHost(lstrSQL, tbRegistros, adOpenKeyset, adLockReadOnly, 10) = apOk Then
        If tbRegistros.RecordCount > 0 Then
            If tbRegistros!deducible_facturado = "S" Then
                VeriDeducible = True
            Else
                VeriDeducible = False
            End If
        End If
    End If
    Conexion.CloseHost tbRegistros
End Function
Function VerificaClienteFacturado(dStrOT, dStrSeccion, dStrCargo) As Boolean

    Dim tbRegistros As New ADODB.Recordset
    Dim lstrSQL As String
    
    VerificaClienteFacturado = False
    lstrSQL = "SELECT Estado FROM Tllr_Facturacion WHERE id_sucursal = '" & gstrIdSucursal & "' and id_empresa = '" & gstrIdEmpresa & "' and id_ot='" & dStrOT & "' and seccion_ot='" & dStrSeccion & "' And Id_Cargo='" & dStrCargo & "'"
    If Conexion.SendHost(lstrSQL, tbRegistros, adOpenKeyset, adLockReadOnly, 10) = apOk Then
        If tbRegistros.RecordCount > 0 Then
            If tbRegistros!estado = "V" Then
                VerificaClienteFacturado = False
            Else
                VerificaClienteFacturado = True
            End If
        End If
    End If
    Conexion.CloseHost tbRegistros
End Function

Function VerificaLetraCarroceria(caracter As Integer) As Integer
    If caracter = 65 Then
        VerificaLetraCarroceria = caracter
    ElseIf caracter = 68 Then
        VerificaLetraCarroceria = caracter
    ElseIf caracter = 80 Then
        VerificaLetraCarroceria = caracter
    Else
        VerificaLetraCarroceria = 0
    End If
End Function

Function Retorna_Valor_General(strSql, Optional Apertura As gcApertura)
Dim AdoPaso As New ADODB.Recordset
'Esta funcion me retorna una valor solicitado desde una consulta SQL
'a la tabla General
    If IsMissing(Apertura) = True Or Apertura = 0 Then 'Si falta el valor o es 1 por defecto es dynamico...
        If Not Conexion.SendHost(strSql, AdoPaso, adOpenDynamic, adLockOptimistic, gcTiempoEspera) = apOk Then
            MsgBox "Error en Conexion con el Host...", vbCritical, "Stock Pro"
            End
        End If
    End If
    
    If Apertura = 3 Then 'Si falta el valor o es 1 por defecto es dynamico...
        If Not Conexion.SendHost(strSql, AdoPaso, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then
            MsgBox "Error en Conexion con el Host...", vbCritical, "Stock Pro"
            End
        End If
    End If
    
    If Apertura = 1 Then 'Si falta el valor o es 1 por defecto es dynamico...
        If Not Conexion.SendHost(strSql, AdoPaso, adOpenStatic, adLockOptimistic, gcTiempoEspera) = apOk Then
            MsgBox "Error en Conexion con el Host...", vbCritical, "Stock Pro"
            End
        End If
    End If
    
    If Apertura = 2 Then 'Si falta el valor o es 1 por defecto es dynamico...
        If Not Conexion.SendHost(strSql, AdoPaso, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            MsgBox "Error en Conexion con el Host...", vbCritical, "Stock Pro"
            End
        End If
    End If
    If Not (AdoPaso.EOF = True And AdoPaso.BOF = True) Then
        Do Until AdoPaso.EOF
            Retorna_Valor_General = ValorNulo(AdoPaso.Fields(0))
            AdoPaso.MoveNext
        Loop
    End If
    AdoPaso.Close
End Function

Public Sub Actualiza_Saldos(cantidad, operacion, empresa, Sucursal, bodega, ubicacion, pieza)
Dim adoSaldo As Recordset
Dim lstrSQL As String
Dim Contador As Double
Dim Cn2 As New ADODB.Connection

'Primero vee si existe el registro si no, lo crea con valores 0
lstrSQL = "Select * From Stck_Saldos Where Id_Item = '" & pieza & "' and Id_Empresa = '" & empresa & "' and Id_sucursal = '" & Sucursal & "' and Id_Bodega = '" & bodega & "' and Id_Ubicacion='" & ubicacion & "'"
 Contador = 0
    If Not Conexion.SendHost(lstrSQL, adoSaldo, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then
        MsgBox "Error en Conexion con el Host...", vbCritical, "Stock Pro"
        End
    Else
        If Not (adoSaldo.EOF = True And adoSaldo.BOF = True) Then
            Do Until adoSaldo.EOF
                  Contador = Contador + 1
                adoSaldo.MoveNext
            Loop
        End If
    Conexion.CloseHost adoSaldo
    End If
If Contador = 0 Then  'Crea el registro
    lstrSQL = "Insert  INTO Stck_Saldos (Id_Item,Id_Empresa,Id_Sucursal,Id_Bodega,Id_Ubicacion,saldo, Usr_Id, Usr_Fecha, Entrada, Salida) Values ('" & pieza & "','" & empresa & "','" & Sucursal & "','" & bodega & "','" & ubicacion & "',0,'" & gstrIdUsuario & "','" & Format(Date, "DD/MM/YYYY") + " " + Format(Time, "HH:MM:SS") & "',0,0)"
    If Conexion.SendHost(lstrSQL, , , , gcTiempoEspera) = apOk Then
    End If
End If

'Luego actualiza
Set adoSaldo = cnnAux.Execute("exec Stck_Actualiza_Saldos_En_Linea " & cantidad & ",'" & operacion & "','" & empresa & "','" & Sucursal & "','" & bodega & "','" & ubicacion & "','" & pieza & "'")
End Sub

Public Sub Actualiza_Saldos_VS_Detalle(operacion, strSql)
Dim adoSaldo As New ADODB.Recordset

    If Not Conexion.SendHost(strSql, adoSaldo, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then
        MsgBox "Error en Conexion con el Host...", vbCritical, "Stock Pro"
        End
    Else
        If Not (adoSaldo.EOF = True And adoSaldo.BOF = True) Then
            Do Until adoSaldo.EOF
                  Call Actualiza_Saldos(adoSaldo.Fields(0) * -1, operacion, adoSaldo.Fields(1), adoSaldo.Fields(2), adoSaldo.Fields(3), adoSaldo.Fields(4), adoSaldo.Fields(5))
                adoSaldo.MoveNext
            Loop
        End If
    Conexion.CloseHost adoSaldo
    End If
End Sub

Function TraeNumeroDocumento(pstrSeccionOT As String, pstrNumeroOt, pstrCargo) As String
Dim mstrCodigoSeccion As String

    If pstrSeccionOT = "M" Then
        mstrCodigoSeccion = Retorna_Valor_General("Select Cod_Taller from Vpro_Parametros_Globales Where Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal = '" & gstrIdSucursal & "'", gcdynamic)
    Else
        mstrCodigoSeccion = Retorna_Valor_General("Select Cod_DyP from Vpro_Parametros_Globales Where Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal = '" & gstrIdSucursal & "'", gcdynamic)
    End If
    If pstrCargo = "" Then
        TraeNumeroDocumento = Retorna_Valor_General("Select Numero_Documento from Vpro_Facturacion Where Id_Tipo_Rescate='" & mstrCodigoSeccion & "' And Numero_Rescate = '" & pstrNumeroOt & "' And Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal = '" & gstrIdSucursal & "'", gcdynamic)
    Else
        TraeNumeroDocumento = Retorna_Valor_General("Select Numero_Documento from Vpro_Facturacion Where Id_Tipo_Rescate='" & mstrCodigoSeccion & "' And Numero_Rescate = '" & pstrNumeroOt & "' And Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal = '" & gstrIdSucursal & "' And Id_Tipo_Cargo='" & pstrCargo & "'", gcdynamic)
    End If
End Function
Public Sub LoginConsola()

    Dim strSql As String
    Dim adoTemp As New ADODB.Recordset
    strSql = "SELECT Glbl_Usuario.Id_Empresa, Glbl_Empresa.Razon_Social, "
    strSql = strSql & " Glbl_Usuario.Id_Sucursal, Glbl_Sucursal.Descripcion, "
    strSql = strSql & " Glbl_Usuario.Id_User , Glbl_Usuario.LOGIN FROM Glbl_Usuario INNER JOIN"
    strSql = strSql & " Glbl_Empresa ON Glbl_Usuario.Id_Empresa = Glbl_Empresa.Id_Empresa INNER JOIN"
    strSql = strSql & " Glbl_Sucursal ON Glbl_Usuario.Id_Empresa = Glbl_Sucursal.Id_Empresa AND"
    strSql = strSql & " Glbl_Usuario.Id_Sucursal = Glbl_Sucursal.Id_Sucursal"
    strSql = strSql & " where Glbl_Usuario.Id_Empresa='" & gstrIdEmpresa & "' and Glbl_Usuario.Id_Sucursal='" & gstrIdSucursal & "' and Glbl_Usuario.Id_User='" & gstrIdUsuario & "'"
    If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            DatosCliente.CodigoEmpresa = adoTemp!Id_Empresa
            DatosCliente.NombreEmpresa = adoTemp!Razon_Social
            DatosCliente.CodigoSucursal = adoTemp!Id_Sucursal
            DatosCliente.NombreSucursal = adoTemp!Descripcion
            DatosCliente.CodigoUsuario = adoTemp!id_user
            DatosCliente.NombreUsuario = adoTemp!Login
            DatosCliente.Modulo = "ElisaTaller"
            DatosCliente.hwnd = frmMain.hwnd
            DatosCliente.ArchivoINI = Command()
        Else
            End
        End If
    Else
        End
    End If
    Conexion.CloseHost adoTemp



    Conexion.CloseHost adoTemp
    DatosCliente.Comando = "LOGIN"
    frmMain.txtComando = DatosCliente.Comando & DatosCliente.CodigoEmpresa & DatosCliente.NombreEmpresa & DatosCliente.CodigoSucursal & DatosCliente.NombreSucursal & _
                                    DatosCliente.CodigoUsuario & DatosCliente.NombreUsuario & DatosCliente.Modulo & DatosCliente.hwnd & DatosCliente.ArchivoINI
    
    If ExisteArchivo(gstrRutaApclient & "\apclient.exe") Then
        Shell gstrRutaApclient & "\apclient.exe"
    End If
End Sub

Public Function ExisteArchivo(Archivo As String)
Dim i%
    On Error Resume Next
    i = Len(Dir$(Archivo))
    If Err Or i = 0 Then
        ExisteArchivo = False
    Else
        ExisteArchivo = True
    End If
End Function
Function Valores_Consumo_Repuesto(LStrNumeroOt As String, lstrTipoCargo As String, LstrFamilia As gcFamilia)
Dim AdoPaso As New ADODB.Recordset
Dim LDblSumatoria As Double
Dim strSql As String
Dim StrFamilia As String

If LstrFamilia = gcLubricantes Then
    StrFamilia = " and Stck_Item.Id_Familia = '" & gstrCodigoLubricantes & "'"  '90'
End If
If LstrFamilia = gcMateriales Then
    StrFamilia = " and Stck_Item.Id_Familia = '" & gstrCodigoMateriales & "'" '85'
End If
If LstrFamilia = gcRepuesto Then
    StrFamilia = " and Stck_Item.Id_Familia <> '" & gstrCodigoLubricantes & "' and Stck_Item.Id_Familia <> '" & gstrCodigoMateriales & "'" '85'"
End If
If LstrFamilia = gcTodos Then
    StrFamilia = ""
End If


    strSql = "SELECT ROUND(SUM(ISNULL(Stck_Mayor_Auxiliar.Cantidad, 0) "
    strSql = strSql & "* ISNULL(Stck_Mayor_Auxiliar.Costo_Unitario_Promedio, 0)), gintDecimalesMoneda) "
    strSql = strSql & "AS Subtotal FROM Stck_Consumo_Taller_Detalle INNER JOIN "
    strSql = strSql & "Stck_Consumo_Taller ON "
    strSql = strSql & "Stck_Consumo_Taller_Detalle.Id_Empresa = Stck_Consumo_Taller.Id_Empresa "
    strSql = strSql & "AND Stck_Consumo_Taller_Detalle.Id_Empresa = Stck_Consumo_Taller.Id_Empresa "
    strSql = strSql & "AND Stck_Consumo_Taller_Detalle.Id_Sucursal = Stck_Consumo_Taller.Id_Sucursal "
    strSql = strSql & "AND Stck_Consumo_Taller_Detalle.Id_Sucursal = Stck_Consumo_Taller.Id_Sucursal "
    strSql = strSql & "AND Stck_Consumo_Taller_Detalle.Id_Consumo = Stck_Consumo_Taller.Id_Consumo "
    strSql = strSql & "AND Stck_Consumo_Taller_Detalle.Id_Consumo = Stck_Consumo_Taller.Id_Consumo "
    strSql = strSql & "Inner Join Stck_Mayor_Auxiliar ON "
    strSql = strSql & "Stck_Consumo_Taller_Detalle.Id_Empresa = Stck_Mayor_Auxiliar.Id_Empresa "
    strSql = strSql & "AND Stck_Consumo_Taller_Detalle.Id_Sucursal = Stck_Mayor_Auxiliar.Id_Sucursal "
    strSql = strSql & "AND Stck_Consumo_Taller_Detalle.Id_Consumo = Stck_Mayor_Auxiliar.Numero_Docto "
    strSql = strSql & "AND Stck_Consumo_Taller_Detalle.id_Linea = Stck_Mayor_Auxiliar.linea "
    strSql = strSql & "AND Stck_Consumo_Taller_Detalle.Id_Bodega = Stck_Mayor_Auxiliar.Id_Bodega "
    strSql = strSql & "AND Stck_Consumo_Taller_Detalle.Id_Ubicacion = Stck_Mayor_Auxiliar.Id_Ubicacion "
    strSql = strSql & "AND Stck_Consumo_Taller_Detalle.Id_Item = Stck_Mayor_Auxiliar.Id_Item Inner Join Stck_Item On Stck_Item.Id_Item = Stck_Mayor_Auxiliar.Id_Item "
    strSql = strSql & "WHERE (Stck_Mayor_Auxiliar.Id_Tipo_Docto = 'CT') AND "
    strSql = strSql & "(Stck_Consumo_Taller.Id_OT = '" & LStrNumeroOt & "') And Stck_Mayor_Auxiliar.TipoCargo='" & lstrTipoCargo & "'" & StrFamilia

    If Not Conexion.SendHost(strSql, AdoPaso, adOpenDynamic, adLockOptimistic, gcTiempoEspera) = apOk Then
            MsgBox "Error en Conexion con el Host...", vbCritical, "Taller Pro"
            End
    End If

    If Not (AdoPaso.EOF = True And AdoPaso.BOF = True) Then
        Do Until AdoPaso.EOF
            LDblSumatoria = IIf(IsNull(AdoPaso!SubTotal), 0, AdoPaso!SubTotal)
            AdoPaso.MoveNext
        Loop
    End If
    AdoPaso.Close
    
    strSql = "SELECT ROUND(SUM(ISNULL(Stck_Mayor_Auxiliar.Cantidad, 0) "
    strSql = strSql & "* ISNULL(Stck_Mayor_Auxiliar.Costo_Unitario_Promedio, 0)), gintDecimalesMoneda) "
    strSql = strSql & "AS Subtotal FROM Stck_devolucion_Taller_Detalle INNER JOIN "
    strSql = strSql & "Stck_devolucion_Taller ON "
    strSql = strSql & "Stck_devolucion_Taller_Detalle.Id_Empresa = Stck_devolucion_Taller.Id_Empresa "
    strSql = strSql & "AND Stck_devolucion_Taller_Detalle.Id_Empresa = Stck_devolucion_Taller.Id_Empresa "
    strSql = strSql & "AND Stck_devolucion_Taller_Detalle.Id_Sucursal = Stck_devolucion_Taller.Id_Sucursal "
    strSql = strSql & "AND Stck_devolucion_Taller_Detalle.Id_Sucursal = Stck_devolucion_Taller.Id_Sucursal "
    strSql = strSql & "AND Stck_devolucion_Taller_Detalle.Id_devolucion = Stck_devolucion_Taller.id_devolucion "
    strSql = strSql & "AND Stck_devolucion_Taller_Detalle.Id_devolucion = Stck_devolucion_Taller.id_devolucion "
    strSql = strSql & "Inner Join Stck_Mayor_Auxiliar ON "
    strSql = strSql & "Stck_devolucion_Taller_Detalle.Id_Empresa = Stck_Mayor_Auxiliar.Id_Empresa "
    strSql = strSql & "AND Stck_devolucion_Taller_Detalle.Id_Sucursal = Stck_Mayor_Auxiliar.Id_Sucursal "
    strSql = strSql & "AND Stck_devolucion_Taller_Detalle.Id_devolucion = Stck_Mayor_Auxiliar.Numero_Docto "
    strSql = strSql & "AND Stck_devolucion_Taller_Detalle.id_Linea = Stck_Mayor_Auxiliar.linea "
    strSql = strSql & "AND Stck_devolucion_Taller_Detalle.Id_Bodega = Stck_Mayor_Auxiliar.Id_Bodega "
    strSql = strSql & "AND Stck_devolucion_Taller_Detalle.Id_Ubicacion = Stck_Mayor_Auxiliar.Id_Ubicacion "
    strSql = strSql & "AND Stck_devolucion_Taller_Detalle.Id_Item = Stck_Mayor_Auxiliar.Id_Item Inner Join Stck_Item On Stck_Item.Id_Item = Stck_Mayor_Auxiliar.Id_Item "
    strSql = strSql & "WHERE (Stck_Mayor_Auxiliar.Id_Tipo_Docto = 'DT') AND "
    strSql = strSql & "(Stck_devolucion_Taller.Id_OT = '" & LStrNumeroOt & "') And Stck_Mayor_Auxiliar.TipoCargo='" & lstrTipoCargo & "'" & StrFamilia

    If Not Conexion.SendHost(strSql, AdoPaso, adOpenDynamic, adLockOptimistic, gcTiempoEspera) = apOk Then
            MsgBox "Error en Conexion con el Host...", vbCritical, "Taller Pro"
            End
    End If

    If Not (AdoPaso.EOF = True And AdoPaso.BOF = True) Then
        Do Until AdoPaso.EOF
            LDblSumatoria = LDblSumatoria - IIf(IsNull(AdoPaso!SubTotal), 0, AdoPaso!SubTotal)
            AdoPaso.MoveNext
        Loop
    End If
    AdoPaso.Close
    
    
Valores_Consumo_Repuesto = LDblSumatoria
End Function
Function Valores_Venta_Repuestos(LStrNumeroOt As String, lstrTipoCargo As String, LstrFamilia As gcFamilia) As VentaRepuestos
Dim AdoPaso As New ADODB.Recordset
Dim strSql As String
Dim StrFamilia As String

    If LstrFamilia = gcLubricantes Then
        StrFamilia = " and Stck_Item.Id_Familia = '" & gstrCodigoLubricantes & "'" '90'
    End If
    If LstrFamilia = gcMateriales Then
        StrFamilia = " and Stck_Item.Id_Familia = '" & gstrCodigoMateriales & "'" '85'
    End If
    If LstrFamilia = gcRepuesto Then
        StrFamilia = " and Stck_Item.Id_Familia <> '" & gstrCodigoLubricantes & "' and Stck_Item.Id_Familia <> '" & gstrCodigoMateriales & "'"
    End If
    If LstrFamilia = gcTodos Then
        StrFamilia = ""
    End If
    
    strSql = "SELECT SUM(Tllr_Repuestos_OT.Subtotal) as Venta, SUM(isnull(Tllr_Repuestos_OT.Monto_Descuento,0)) as Mdr"
    strSql = strSql & " From Tllr_Repuestos_OT"
    strSql = strSql & " LEFT OUTER JOIN Stck_Item ON Tllr_Repuestos_OT.Id_Item = Stck_Item.Id_Item"
    strSql = strSql & " Where Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    strSql = strSql & " AND id_ot='" & LStrNumeroOt & "' and Id_Tipo_Cargo='" & lstrTipoCargo & "'"
    strSql = strSql & StrFamilia
    
    
    If Not Conexion.SendHost(strSql, AdoPaso, adOpenDynamic, adLockOptimistic, gcTiempoEspera) = apOk Then
            MsgBox "Error en Conexion con el Host...", vbCritical, "Taller Pro"
            End
    End If

    If Not (AdoPaso.EOF = True And AdoPaso.BOF = True) Then
        Do Until AdoPaso.EOF
            Valores_Venta_Repuestos.Repuestos = IIf(IsNull(AdoPaso!VENTA), 0, AdoPaso!VENTA)
            Valores_Venta_Repuestos.Descuentos = IIf(IsNull(AdoPaso!Mdr), 0, AdoPaso!Mdr)
            AdoPaso.MoveNext
        Loop
    End If
    AdoPaso.Close
End Function

Function Costo_Promedio_Repuesto(LStrNumeroOt As String, LStrCodigoRepuesto As String)
Dim AdoPaso As New ADODB.Recordset
Dim strSql As String
'kjcv 15.08.12 Se agrego paridad, para que calcule en soles (S/.)Costo y Subtotal
strSql = "SELECT isnull(Stck_Mayor_Auxiliar.Costo_Unitario_Promedio,0)* isnull(Stck_Consumo_Taller_Detalle.Paridad,0) as Costo, isnull(Stck_Mayor_Auxiliar.Cantidad,0) as Cantidad, " _
        & " isnull(Stck_Mayor_Auxiliar.Cantidad,0) * isnull(Stck_Mayor_Auxiliar.Costo_Unitario_Promedio,0) * isnull(Stck_Consumo_Taller_Detalle.Paridad,0) AS Subtotal, " _
        & " Stck_Consumo_Taller_Detalle.Precio_Unitario FROM Stck_Consumo_Taller_Detalle INNER JOIN Stck_Consumo_Taller ON " _
        & " Stck_Consumo_Taller_Detalle.Id_Empresa = Stck_Consumo_Taller.Id_Empresa AND Stck_Consumo_Taller_Detalle.Id_Empresa = Stck_Consumo_Taller.Id_Empresa AND " _
        & " Stck_Consumo_Taller_Detalle.Id_Sucursal = Stck_Consumo_Taller.Id_Sucursal AND Stck_Consumo_Taller_Detalle.Id_Sucursal = Stck_Consumo_Taller.Id_Sucursal AND Stck_Consumo_Taller_Detalle.Id_Consumo = Stck_Consumo_Taller.Id_Consumo " _
        & " AND Stck_Consumo_Taller_Detalle.Id_Consumo = Stck_Consumo_Taller.Id_Consumo Inner Join Stck_Mayor_Auxiliar ON Stck_Consumo_Taller_Detalle.Id_Empresa = Stck_Mayor_Auxiliar.Id_Empresa AND Stck_Consumo_Taller_Detalle.Id_Sucursal = Stck_Mayor_Auxiliar.Id_Sucursal AND " _
        & " Stck_Consumo_Taller_Detalle.Id_Consumo = Stck_Mayor_Auxiliar.Numero_Docto AND Stck_Consumo_Taller_Detalle.id_Linea = Stck_Mayor_Auxiliar.linea AND Stck_Consumo_Taller_Detalle.Id_Bodega = Stck_Mayor_Auxiliar.Id_Bodega AND Stck_Consumo_Taller_Detalle.Id_Ubicacion = Stck_Mayor_Auxiliar.Id_Ubicacion AND Stck_Consumo_Taller_Detalle.Id_Item = Stck_Mayor_Auxiliar.Id_Item " _
        & " WHERE (Stck_Mayor_Auxiliar.Id_Tipo_Docto = 'CT') AND (Stck_Consumo_Taller_Detalle.Id_Item = '" & LStrCodigoRepuesto & "') AND (Stck_Consumo_Taller.Id_OT = '" & LStrNumeroOt & "')"
   If Not Conexion.SendHost(strSql, AdoPaso, adOpenDynamic, adLockOptimistic, gcTiempoEspera) = apOk Then
            MsgBox "Error en Conexion con el Host...", vbCritical, "Stock Pro"
            End
    End If

   If Not (AdoPaso.EOF = True And AdoPaso.BOF = True) Then
        Do Until AdoPaso.EOF
            Costo_Promedio_Repuesto = IIf(IsNull(AdoPaso!SubTotal), 0, AdoPaso!SubTotal)
            AdoPaso.MoveNext
        Loop
    End If
    AdoPaso.Close
End Function

Function CostoRepuesto(lstrItem As String, lintCantidad As Double) As Double
    CostoRepuesto = Retorna_Valor_General("SELECT TOP 1 Costo_Unitario_Promedio From Stck_Mayor_Auxiliar WHERE (Id_Empresa = '" & gstrIdEmpresa & "') AND (Id_Item = '" & lstrItem & "'" & ") ORDER BY Fecha DESC", gcdynamic) '* lintCantidad
End Function
Public Function ImprimeMiReporte(lvLista As ListView, cmdImpresora As CommonDialog, strTituloReporte As String)
    Dim DatosImpresion As TIPO_DATOS_REPORTE
    DatosImpresion.DireccionEmpresa = gstrDirSuc
    DatosImpresion.iD_Usuario = gstrIdUsuario
    DatosImpresion.NombreEmpresa = gstrEmpresa
    DatosImpresion.RutEmpresa = gstrIdEmpresa
    DatosImpresion.TituloReporte = strTituloReporte
    ImprimirListas lvLista, cmdImpresora, DatosImpresion
End Function
Public Sub ImprimirListas(lvLista As ListView, cmdImpresora As CommonDialog, DatosReporte As TIPO_DATOS_REPORTE)
    Dim i As Integer
    Dim K As Integer
    Dim lngFil As Long
    Dim lngCol As Long
    Dim j As Integer
    Dim intPagina As Integer
    Dim strLinea As String
    Dim intMayor As Integer
On Error GoTo Errores
    'seleccionar impresora
    Err.Clear
    cmdImpresora.CancelError = True
    cmdImpresora.Flags = &H100000 Or &H8 Or &H4&
    cmdImpresora.ShowPrinter
    
    Screen.MousePointer = vbHourglass
    intPagina = 1
    lngFil = 0
    lngCol = 0
    
    
    
    lngFil = ImpresionCabecera(lngFil, 1, cmdImpresora, DatosReporte, lvLista)

    
    For i = 1 To lvLista.ListItems.Count
        lngCol = 0
        For j = 1 To lvLista.ColumnHeaders.Count
            For K = 1 To lvLista.ColumnHeaders.Count
                If lvLista.ColumnHeaders(K).Position = j Then
                    Exit For
                End If
            Next
            If lvLista.ColumnHeaders.Item(K).Width <> 0 Then
                If K = 1 Then
                    strLinea = lvLista.ListItems(i).Text
                Else
                    strLinea = lvLista.ListItems(i).SubItems(K - 1)
                End If
                ImprimirLinea AjustaTexto(strLinea, lvLista.ColumnHeaders.Item(K).Tag), lngFil, False, lngCol, False, cmdImpresora.FontName, cmdImpresora.FontSize, cmdImpresora.FontBold, cmdImpresora.FontItalic, cmdImpresora.FontUnderline
                lngCol = lngCol + lvLista.ColumnHeaders.Item(K).Tag + 150
            End If
        Next
        lngFil = lngFil + Printer.TextHeight("A")
    Next
    Printer.EndDoc
Exit Sub
Errores:
    MsgBox "Impresión cancelada..."
End Sub
Public Sub ImprimirLinea(strLinea As String, ByRef x As Long, blnIncX As Boolean, ByRef Y As Long, blnIncY As Boolean, Optional strFont As String, Optional intFontSize As Integer, Optional blnFontBold As Boolean, Optional blnFontItalic As Boolean, Optional blnFontUnderline As Boolean)
    '//Parametros opcionales...
    If Not IsMissing(strFont) Then
        Printer.FontName = strFont
    End If
    If Not IsMissing(intFontSize) Then
        Printer.FontSize = intFontSize
    End If
    If Not IsMissing(blnFontBold) Then
        If blnFontBold Then
            Printer.FontBold = blnFontBold
        Else
            blnFontBold = Not Printer.FontBold
        End If
    End If
    If Not IsMissing(blnFontItalic) Then
        If blnFontItalic Then
            Printer.FontItalic = blnFontItalic
        Else
            blnFontItalic = Not Printer.FontItalic
        End If
    End If
    If Not IsMissing(blnFontUnderline) Then
        If blnFontUnderline Then
            Printer.FontUnderline = blnFontUnderline
        Else
            blnFontUnderline = Not Printer.FontUnderline
        End If
    End If
    '//Parametros Fijos...
    Printer.CurrentX = Y
    Printer.CurrentY = x
    Printer.Print strLinea
    '//Incrementa Valores....
    If blnIncX Then
        x = x + Printer.TextHeight(strLinea)
    End If
    If blnIncY Then
        Y = Y + Printer.TextWidth(strLinea)
    End If
    If Not IsMissing(blnFontBold) Then
        Printer.FontBold = Not blnFontBold
    End If
    If Not IsMissing(blnFontItalic) Then
        Printer.FontItalic = Not blnFontItalic
    End If
    If Not IsMissing(blnFontUnderline) Then
        Printer.FontUnderline = Not blnFontUnderline
    End If
End Sub
Private Function ImpresionCabecera(ByRef lngFila As Long, intPagina As Integer, cmdImpresora As CommonDialog, DatosReporte As TIPO_DATOS_REPORTE, ByRef lvLista As ListView) As Long
    Dim intMayor  As Integer
    Dim strLinea As String
    Dim lngCol As Long
    Dim intMargen As Integer
    Dim lngFil As Long
    Dim i As Integer
    Dim j As Integer
    Dim K As Integer
    If Printer.Orientation = vbPRORPortrait Then
        intMargen = 400
    Else
        intMargen = 1000
    End If
    Printer.FontSize = cmdImpresora.FontSize
    Printer.FontName = cmdImpresora.FontName

    '//////////////
    'Nombre Empresa
    ImprimirLinea DatosReporte.NombreEmpresa, lngFila, False, 0, False, cmdImpresora.FontName, cmdImpresora.FontSize, cmdImpresora.FontBold, cmdImpresora.FontItalic, cmdImpresora.FontUnderline
    '//
    strLinea = "PAGINA :       " & Format(intPagina, "0000")
    lngCol = Printer.Width - Printer.TextWidth(strLinea) - intMargen
    ImprimirLinea strLinea, lngFila, True, lngCol, False, cmdImpresora.FontName, cmdImpresora.FontSize, cmdImpresora.FontBold, cmdImpresora.FontItalic, cmdImpresora.FontUnderline
    
    'Rut
    ImprimirLinea DatosReporte.RutEmpresa, lngFila, False, 0, False, cmdImpresora.FontName, cmdImpresora.FontSize, cmdImpresora.FontBold, cmdImpresora.FontItalic, cmdImpresora.FontUnderline
    '//
    strLinea = "FECHA  : " & Format(Date, "DD/MM/YYYY")
    lngCol = Printer.Width - Printer.TextWidth(strLinea) - intMargen
    ImprimirLinea strLinea, lngFila, True, lngCol, False, cmdImpresora.FontName, cmdImpresora.FontSize, cmdImpresora.FontBold, cmdImpresora.FontItalic, cmdImpresora.FontUnderline
    
    
    'Direccion
    ImprimirLinea DatosReporte.DireccionEmpresa, lngFila, False, 0, False, cmdImpresora.FontName, cmdImpresora.FontSize, cmdImpresora.FontBold, cmdImpresora.FontItalic, cmdImpresora.FontUnderline
    Printer.FontSize = cmdImpresora.FontSize
    Printer.FontName = cmdImpresora.FontName
    '//
    strLinea = "HORA   :   " & Format(Time, "HH:MM:SS")
    lngCol = Printer.Width - Printer.TextWidth(strLinea) - intMargen
    ImprimirLinea strLinea, lngFila, True, lngCol, False, cmdImpresora.FontName, cmdImpresora.FontSize, cmdImpresora.FontBold, cmdImpresora.FontItalic, cmdImpresora.FontUnderline
    '//Usuario
    strLinea = "USUARIO:   " & DatosReporte.iD_Usuario
    lngCol = Printer.Width - Printer.TextWidth(strLinea) - intMargen
    ImprimirLinea strLinea, lngFila, True, lngCol, False, cmdImpresora.FontName, cmdImpresora.FontSize, cmdImpresora.FontBold, cmdImpresora.FontItalic, cmdImpresora.FontUnderline
    
    '//
    Printer.FontSize = cmdImpresora.FontSize
    '////
    'Titulo del Reporte
    strLinea = DatosReporte.TituloReporte
    lngCol = (Printer.Width - Printer.TextWidth(strLinea) - intMargen) / 2
    ImprimirLinea strLinea, lngFila, True, lngCol, False, cmdImpresora.FontName, cmdImpresora.FontSize, cmdImpresora.FontBold, cmdImpresora.FontItalic, cmdImpresora.FontUnderline
    
    '===
    '//Analisa Archo por columna...
    Printer.FontSize = cmdImpresora.FontSize
    lngFil = lngFila
    lngFil = lngFil + Printer.TextHeight("A")
    Printer.Print "-----"
    lngFil = lngFil + Printer.TextHeight("A")
    For i = 1 To lvLista.ColumnHeaders.Count
        For K = 1 To lvLista.ColumnHeaders.Count
            If lvLista.ColumnHeaders(K).Position = i Then
                Exit For
            End If
        Next
        If lvLista.ColumnHeaders.Item(K).Width <> 0 Then
            intMayor = 0
            For j = 1 To lvLista.ListItems.Count
                If K = 1 Then
                    strLinea = lvLista.ListItems(j).Text
                Else
                    strLinea = lvLista.ListItems(j).SubItems(K - 1)
                End If
                If Printer.TextWidth(strLinea) > intMayor Then
                    intMayor = Printer.TextWidth(strLinea)
                End If
            Next
            If intMayor < lvLista.ColumnHeaders.Item(K).Width Then
                lvLista.ColumnHeaders.Item(K).Tag = intMayor
            Else
                lvLista.ColumnHeaders.Item(K).Tag = lvLista.ColumnHeaders.Item(K).Width
            End If
        End If
    Next
    
    lngCol = 0
    For i = 1 To lvLista.ColumnHeaders.Count
        For K = 1 To lvLista.ColumnHeaders.Count
            If lvLista.ColumnHeaders(K).Position = i Then
                Exit For
            End If
        Next
        If lvLista.ColumnHeaders.Item(K).Width <> 0 Then
            strLinea = lvLista.ColumnHeaders.Item(K).Text
            If lvLista.ColumnHeaders.Item(K).Tag <> "" Then
                ImprimirLinea AjustaTexto(strLinea, lvLista.ColumnHeaders.Item(K).Tag), lngFil, False, lngCol, False, cmdImpresora.FontName, cmdImpresora.FontSize, cmdImpresora.FontBold, cmdImpresora.FontItalic, cmdImpresora.FontUnderline
                lngCol = lngCol + lvLista.ColumnHeaders.Item(K).Tag + 150
            End If
        End If
    Next
    lngFil = lngFil + Printer.TextHeight(strLinea)
    ImpresionCabecera = lngFil
End Function
Public Function AjustaTexto(strLinea As String, lngWidth As Long)
    Dim i As Integer
    AjustaTexto = ""
    For i = Len(strLinea) To 1 Step -1
        If Printer.TextWidth(Left(strLinea, i)) <= lngWidth Then
            AjustaTexto = Left(strLinea, i)
            Exit For
        End If
    Next
End Function
Public Function ObjHideColumnHeader(ByRef objListView As ListView) As Boolean
    Dim Item As ListItem
    Dim i As Integer
    Load frmVistaDatos
    Set gObjListView = objListView
    For i = 1 To objListView.ColumnHeaders.Count
        If objListView.ColumnHeaders(i).Tag <> "N" Then
            Set Item = frmVistaDatos.lvwVistas.ListItems.Add(, , objListView.ColumnHeaders(i).Text)
            If objListView.ColumnHeaders(i).Width = 0 Then
                Item.Checked = False
            Else
                Item.Checked = True
            End If
        End If
    Next i
    frmVistaDatos.Show vbModal
    
End Function


Public Sub GetData(ByVal Formulario As Form, strSql As String, intSheet As Integer)
Dim i As Integer

Screen.MousePointer = vbHourglass

Formulario.sprGrillaPrincipal.Sheet = intSheet

Formulario.sprGrillaPrincipal.Redraw = False

If strSql = "" Then
    Screen.MousePointer = 0
    Exit Sub
End If

' pasa el SQL as recordset
Formulario.datDatos.RecordSource = ""
Formulario.datDatos.RecordSource = strSql
Formulario.datDatos.ConnectionString = strConnect
Formulario.datDatos.CommandTimeout = 999
Formulario.datDatos.Refresh

' pone los titulos de columna
Formulario.sprGrillaPrincipal.Row = 0
Formulario.sprGrillaPrincipal.Col = 0
For i = 0 To Formulario.datDatos.Recordset.Fields.Count - 1
    Formulario.sprGrillaPrincipal.Col = Formulario.sprGrillaPrincipal.Col + 1
    Formulario.sprGrillaPrincipal.Text = Formulario.datDatos.Recordset.Fields(i).Name
Next i
 
If Formulario.datDatos.Recordset.RecordCount = 0 Then
    For i = 0 To Formulario.datDatos.Recordset.Fields.Count - 1
        Formulario.sprGrillaPrincipal.Row = 1
        Formulario.sprGrillaPrincipal.Col = i + 1
        Formulario.sprGrillaPrincipal.Text = " "
    Next i
    Exit Sub
End If

Formulario.datDatos.Recordset.MoveLast
Formulario.datDatos.Recordset.MoveFirst


Formulario.sprGrillaPrincipal.Row = 1
' pone el contenido fila por fila
While Not Formulario.datDatos.Recordset.EOF
    Formulario.sprGrillaPrincipal.Col = 0
    
    For i = 0 To Formulario.datDatos.Recordset.Fields.Count - 1
        Formulario.sprGrillaPrincipal.Col = Formulario.sprGrillaPrincipal.Col + 1
        If Not IsNull(Formulario.datDatos.Recordset(i)) Then
            Formulario.sprGrillaPrincipal.Text = Formulario.datDatos.Recordset(i)
            'If I > 0 Then
                'Formulario.sprGrillaPrincipal.ColMerge = MergeRestricted
            'End If
        End If
    Next i
        
    ' pasa a la siguiente fila
    Formulario.datDatos.Recordset.MoveNext
    Formulario.sprGrillaPrincipal.Row = Formulario.sprGrillaPrincipal.Row + 1

Wend

Formulario.sprGrillaPrincipal.Redraw = True

Screen.MousePointer = vbDefault

End Sub

Public Function TraeNumCol(ByVal ControlData As Adodc, strTxtColCampo As String) As Long
Dim fldTemp As ADODB.Field
Dim ldblCount As Double

TraeNumCol = 0

' busca las columnas adecuadas
ldblCount = 1
For Each fldTemp In ControlData.Recordset.Fields
    ' busca columna que contiene la nota (nota_int)
    If UCase$(fldTemp.Name) = UCase$(strTxtColCampo) Then
        TraeNumCol = ldblCount
        Exit For
    End If
    ldblCount = ldblCount + 1
Next fldTemp

End Function

Public Function TraeNumColSpread(ByVal varSpread As fpSpread, strTxtColCampo As String) As Long
Dim ldblCount As Double
Dim lvarTitulo As Variant

TraeNumColSpread = 0

For ldblCount = 1 To varSpread.MaxCols
    varSpread.GetText CLng(ldblCount), 0, lvarTitulo
    If lvarTitulo = strTxtColCampo Then
        TraeNumColSpread = CLng(ldblCount)
        Exit For
    End If
Next ldblCount

End Function

Public Sub GetDataVertical(ByVal Formulario As Form, strSql As String, intSheet As Integer)
Dim i As Integer

Screen.MousePointer = vbHourglass

Formulario.sprGrillaPrincipal.Sheet = intSheet

Formulario.sprGrillaPrincipal.Redraw = False

If strSql = "" Then
    Screen.MousePointer = 0
    Exit Sub
End If

' pasa el SQL as recordset
Formulario.datDatos.RecordSource = ""
Formulario.datDatos.RecordSource = strSql
Formulario.datDatos.Refresh

If Formulario.datDatos.Recordset.RecordCount = 0 Then Exit Sub

Formulario.datDatos.Recordset.MoveLast
Formulario.datDatos.Recordset.MoveFirst

' setea progess bar
Formulario.pb1.Min = 0
Formulario.pb1.Max = Formulario.datDatos.Recordset.Fields.Count
Formulario.pb1.Value = Formulario.pb1.Min
Formulario.pb1.Visible = True

' actualiza statusBar
Formulario.statusBar.Panels(1).Text = "Progreso: " & Formulario.pb1.Value & " de " & Formulario.datDatos.Recordset.RecordCount

' pone los titulos de columna
Formulario.sprGrillaPrincipal.Row = 0
Formulario.sprGrillaPrincipal.Col = 1
Formulario.sprGrillaPrincipal.Text = "CONCEPTO"
Formulario.sprGrillaPrincipal.Col = 2
Formulario.sprGrillaPrincipal.Text = "DESCRIPCION"

Formulario.sprGrillaPrincipal.Row = 1
' pone el contenido fila por fila
If Not Formulario.datDatos.Recordset.EOF Then
    Formulario.sprGrillaPrincipal.Col = 0
    
    For i = 0 To Formulario.datDatos.Recordset.Fields.Count - 1
        Formulario.sprGrillaPrincipal.Col = 1
        Formulario.sprGrillaPrincipal.Text = Formulario.datDatos.Recordset.Fields(i).Name
        Formulario.sprGrillaPrincipal.Col = 2
        Formulario.sprGrillaPrincipal.Text = Formulario.datDatos.Recordset(i)
        Formulario.sprGrillaPrincipal.Row = Formulario.sprGrillaPrincipal.Row + 1
    Next i
    
    ' incrementa el progress bar
    If Formulario.pb1.Value < Formulario.pb1.Max Then
        Formulario.pb1.Value = Formulario.pb1.Value + 1
    End If
    
    ' actualiza statusBar
    Formulario.statusBar.Panels(1).Text = "Progreso: " & Formulario.pb1.Value & " de " & Formulario.datDatos.Recordset.RecordCount
End If

Formulario.sprGrillaPrincipal.Redraw = True

' resetea progressBar
Formulario.pb1.Value = 0
Formulario.statusBar.Panels(1).Text = ""
Formulario.statusBar.Panels(2).Text = ""

Screen.MousePointer = vbDefault

End Sub

Public Sub SeteaSpreadPostSoloHoja(ByVal varSpread As fpSpread, lngColOcultarDerecha As Long)
Dim llngCol As Long
Dim ldblCont As Double
Dim ldblMejorAnchoCol As Double
Dim ldblMejorAnchoRow As Double
Dim lvarTemp As Variant

Screen.MousePointer = vbHourglass

varSpread.Redraw = False

' ajusta cantidad de filas y columnas
varSpread.MaxCols = varSpread.DataColCnt
varSpread.MaxRows = varSpread.DataRowCnt
'
For llngCol = 1 To varSpread.DataColCnt
    
    varSpread.Lock = True
    
    ' setea cada columna a su mejor ancho
    ldblMejorAnchoCol = 0
    ldblMejorAnchoCol = varSpread.MaxTextColWidth(llngCol) + 2
    If ldblMejorAnchoCol > 40 Then ldblMejorAnchoCol = 40 ' para no agranadrla demaciado
    varSpread.ColWidth(llngCol) = ldblMejorAnchoCol
    
    ' define el tipo o formato para cada columna
    varSpread.GetText llngCol, 0, lvarTemp
    varSpread.Col = llngCol
    varSpread.Row = -1
    ' formato predeterminado
    varSpread.CellType = CellTypeStaticText
    If UCase(lvarTemp) Like " *" Then
        varSpread.TypeHAlign = TypeHAlignCenter
    End If
    ' cambia el formato predeterminado por checkbox si...
    If UCase(lvarTemp) Like "* " Then
        With varSpread
            .CellType = CellTypeCheckBox
            .TypeCheckType = TypeCheckTypeThreeState
            .TypeCheckCenter = True
            .TypeVAlign = TypeHAlignCenter
            .TypeCheckPicture(0) = LoadPicture(App.Path & "\imagenes\chk_no.bmp")
            .TypeCheckPicture(1) = LoadPicture(App.Path & "\imagenes\chk_si.bmp")
            .TypeCheckPicture(2) = LoadPicture(App.Path & "\imagenes\chk_no_paso.bmp")
            .TypeCheckPicture(3) = LoadPicture(App.Path & "\imagenes\chk_no_paso.bmp")
            .TypeCheckPicture(4) = LoadPicture(App.Path & "\imagenes\chk_no_paso.bmp")
            .TypeCheckPicture(5) = LoadPicture(App.Path & "\imagenes\chk_no_paso.bmp")
        End With
    End If
    ' cambia el formato predeterminado por float sin decimales si...
    'Or UCase(lvarTemp) Like "*FOLIO*"
    If UCase(lvarTemp) Like "*TOT*" _
    Or UCase(lvarTemp) Like "*VALOR*" _
    Or UCase(lvarTemp) Like "*$*" _
    Or UCase(lvarTemp) Like "*MONTO*" _
    Or UCase(lvarTemp) Like "*KMS*" _
    Or UCase(lvarTemp) Like "*NETO*" _
    Or UCase(lvarTemp) Like "*PRECIO*" _
    Or UCase(lvarTemp) Like "*COSTO*" _
    Or UCase(lvarTemp) Like "*VENTA*" _
    Or UCase(lvarTemp) Like "*TRASLADO*" _
    Or UCase(lvarTemp) Like "*PROVIS*" _
    Or UCase(lvarTemp) Like "*BONO*" _
    Or UCase(lvarTemp) Like "*MARGEN*" _
    Or UCase(lvarTemp) Like "*EXENTO*" _
    Or UCase(lvarTemp) Like "*LINEA*" _
    Or UCase(lvarTemp) Like "*IVA*" _
    Or UCase(lvarTemp) Like "*IMPUESTO*" _
    Or UCase(lvarTemp) Like "*NUM*" _
    Or UCase(lvarTemp) Like "*DESC*TO*" _
    Or UCase(lvarTemp) Like "*DEDUCIB*" Then
        With varSpread
            .CellType = CellTypeNumber
            .TypeNumberDecimal = "."
            .TypeNumberLeadingZero = TypeLeadingZeroNo
            .TypeNumberNegStyle = TypeNumberNegStyle1
            .TypeNumberDecPlaces = 0
            .TypeNumberSeparator = ","
            .TypeNumberShowSep = True
        End With
    End If
    ' cambia el formato predeterminado por float con 2 decimales si...
    If UCase(lvarTemp) Like "*PORCENT*" _
    Or UCase(lvarTemp) Like "*%*" _
    Or UCase(lvarTemp) Like "*CANT*" _
    Or UCase(lvarTemp) Like "*HORA*" _
    Or UCase(lvarTemp) Like "*RECARG*" Then
        With varSpread
            .CellType = CellTypeNumber
            .TypeNumberDecimal = "."
            .TypeNumberLeadingZero = TypeLeadingZeroYes
            .TypeNumberNegStyle = TypeNumberNegStyle1
            .TypeNumberDecPlaces = 2
            .TypeNumberSeparator = ","
            .TypeNumberShowSep = True
        End With
    End If
    '
    ' setea las filas para la columna comentarios o similares...
    If InStr(1, UCase$(lvarTemp), "NOTA", vbTextCompare) <> 0 _
    Or InStr(1, UCase$(lvarTemp), "COMENTARIO", vbTextCompare) <> 0 _
    Or InStr(1, UCase$(lvarTemp), "OBSERVACIONES", vbTextCompare) <> 0 Then
        varSpread.Redraw = True
        For ldblCont = 1 To varSpread.DataRowCnt
                varSpread.Col = llngCol
                varSpread.Row = ldblCont
                If Trim$(varSpread.Text) <> "" And Len(varSpread.Text) > 20 Then
                    varSpread.TypeTextWordWrap = True
                    ldblMejorAnchoRow = 0
                    ldblMejorAnchoRow = varSpread.MaxTextRowHeight(ldblCont)
                    varSpread.RowHeight(ldblCont) = ldblMejorAnchoRow
                End If
        Next ldblCont
    End If

Next llngCol
'
'
' setea la fila de titulos a su mejor ancho
ldblMejorAnchoRow = 0
ldblMejorAnchoRow = varSpread.MaxTextRowHeight(0)
varSpread.RowHeight(0) = ldblMejorAnchoRow
If varSpread.RowHeight(0) > 20 Then varSpread.RowHeight(0) = 20
'
' oculta las ultimas dos columnas (id_empresa y id_sucursal)
If lngColOcultarDerecha < varSpread.MaxCols Then
    For llngCol = varSpread.MaxCols To (varSpread.MaxCols - lngColOcultarDerecha) + 1 Step -1
        varSpread.ColWidth(llngCol) = 0
    Next llngCol
End If

varSpread.GetText 1, 1, lvarTemp
If varSpread.DataRowCnt = 1 And lvarTemp = " " Then
    varSpread.DeleteRows 1, 1
    varSpread.MaxRows = 0
End If

varSpread.Redraw = True

Screen.MousePointer = vbDefault

End Sub

Public Sub SeteaSpreadPostSoloHoja_OLD(ByVal varSpread As fpSpread, lngColOcultarDerecha As Long)
Dim llngCol As Long
Dim llngCol2 As Long
Dim ldblMejorAnchoCol As Double
Dim ldblMejorAnchoRow As Double
Dim ldblMejorAnchoFilas As Double
Dim lvarTemp As Variant

Screen.MousePointer = vbHourglass

varSpread.Redraw = False

' ajusta cantidad de filas y columnas
varSpread.MaxCols = varSpread.DataColCnt
varSpread.MaxRows = varSpread.DataRowCnt
'
For llngCol = 1 To varSpread.DataColCnt
    
    varSpread.Lock = True
    
    ' setea cada columna a su mejor ancho
    ldblMejorAnchoCol = 0
    ldblMejorAnchoCol = varSpread.MaxTextColWidth(llngCol) + 2
    If ldblMejorAnchoCol > 25 Then ldblMejorAnchoCol = 25 ' para no agranadrla demaciado
    varSpread.ColWidth(llngCol) = ldblMejorAnchoCol
    
    ' define el tipo o formato para cada columna
    varSpread.GetText llngCol, 0, lvarTemp
    varSpread.Col = llngCol
    varSpread.Row = -1
    ' formato predeterminado
    varSpread.CellType = CellTypeStaticText
    If UCase(lvarTemp) Like " *" Then
        varSpread.TypeHAlign = TypeHAlignCenter
    End If
    ' cambia el formato predeterminado por checkbox si...
    If UCase(lvarTemp) Like "* " Then
        With varSpread
            .CellType = CellTypeCheckBox
            .TypeCheckType = TypeCheckTypeThreeState
            .TypeCheckCenter = True
            .TypeVAlign = TypeHAlignCenter
            .TypeCheckPicture(0) = LoadPicture(App.Path & "\imagenes\chk_no.bmp")
            .TypeCheckPicture(1) = LoadPicture(App.Path & "\imagenes\chk_si.bmp")
            .TypeCheckPicture(2) = LoadPicture(App.Path & "\imagenes\chk_no_paso.bmp")
            .TypeCheckPicture(3) = LoadPicture(App.Path & "\imagenes\chk_no_paso.bmp")
            .TypeCheckPicture(4) = LoadPicture(App.Path & "\imagenes\chk_no_paso.bmp")
            .TypeCheckPicture(5) = LoadPicture(App.Path & "\imagenes\chk_no_paso.bmp")
        End With
    End If
    ' cambia el formato predeterminado por float sin decimales si...
    If UCase(lvarTemp) Like "*TOTAL*" _
    Or UCase(lvarTemp) Like "*VALOR*" _
    Or UCase(lvarTemp) Like "*$*" _
    Or UCase(lvarTemp) Like "*MONTO*" _
    Or UCase(lvarTemp) Like "*KMS*" _
    Or UCase(lvarTemp) Like "*NETO*" _
    Or UCase(lvarTemp) Like "*PRECIO*" _
    Or UCase(lvarTemp) Like "*NUM*" _
    Or UCase(lvarTemp) Like "*DESC*TO*" _
    Or UCase(lvarTemp) Like "*DIA*" _
    Or UCase(lvarTemp) Like "*DS*" _
    Or UCase(lvarTemp) Like "*KILOM*" _
    Or UCase(lvarTemp) Like "*DUEÑO*" _
    Or UCase(lvarTemp) Like "*DEDUCIB*" Then
        With varSpread
            .CellType = CellTypeNumber
            .TypeNumberDecimal = "."
            .TypeNumberLeadingZero = TypeLeadingZeroNo
            .TypeNumberNegStyle = TypeNumberNegStyle1
            .TypeNumberDecPlaces = 0
            .TypeNumberSeparator = ","
            .TypeNumberShowSep = True
        End With
    End If
    ' cambia el formato predeterminado por float con 2 decimales si...
    If UCase(lvarTemp) Like "*PORCEN*" _
    Or UCase(lvarTemp) Like "*%*" _
    Or UCase(lvarTemp) Like "*CANT*" _
    Or UCase(lvarTemp) Like "*HORA*" _
    Or UCase(lvarTemp) Like "*RECARG*" Then
        With varSpread
            .CellType = CellTypeNumber
            .TypeNumberDecimal = "."
            .TypeNumberLeadingZero = TypeLeadingZeroYes
            .TypeNumberNegStyle = TypeNumberNegStyle1
            .TypeNumberDecPlaces = 2
            .TypeNumberSeparator = ","
            .TypeNumberShowSep = True
        End With
    End If
    ldblMejorAnchoFilas = varSpread.RowHeight(varSpread.Row)
    If UCase(lvarTemp) = "RESULTADO" Then
        For llngCol2 = 1 To varSpread.MaxRows
            varSpread.Col = llngCol
            varSpread.Row = llngCol2
            varSpread.CellType = CellTypeStaticText
            varSpread.FontName = "Verdana"
            varSpread.FontSize = 5
            varSpread.ForeColor = &HC0&
            varSpread.RowHeight(llngCol2) = ldblMejorAnchoFilas
        Next llngCol2
    End If
Next llngCol
'
' setea la fila de titulos a su mejor ancho
ldblMejorAnchoRow = 0
ldblMejorAnchoRow = varSpread.MaxTextRowHeight(0)
varSpread.RowHeight(0) = ldblMejorAnchoRow
'If varSpread.RowHeight(0) > 20 Then varSpread.RowHeight(0) = 20
'
' oculta las ultimas dos columnas (id_empresa y id_sucursal)
If lngColOcultarDerecha < varSpread.MaxCols Then
    For llngCol = varSpread.MaxCols To (varSpread.MaxCols - lngColOcultarDerecha) + 1 Step -1
        varSpread.ColWidth(llngCol) = 0
    Next llngCol
End If

varSpread.GetText 1, 1, lvarTemp
If varSpread.DataRowCnt = 1 And lvarTemp = " " Then
    varSpread.DeleteRows 1, 1
    varSpread.MaxRows = 0
End If

varSpread.Redraw = True

Screen.MousePointer = vbDefault

End Sub

Public Sub SeteaSpreadSoloHoja(ByVal varSpread As fpSpread)
Dim lstrNombreSheet As String

Screen.MousePointer = vbHourglass

varSpread.Redraw = False

With varSpread
    lstrNombreSheet = .SheetName ' guarda la sheet original
    .ResetSheet .ActiveSheet ' formatea la sheet
    .SheetName = lstrNombreSheet ' vuelve a poner el nombre de la sheet xq se borró con el formateo
    
    .MaxRows = 65000
    .VirtualMaxRows = 65000
    
    .Row = -1
    .Col = -1
    
    .AllowCellOverflow = True
    .BackColorStyle = BackColorStyleUnderGrid 'color debajo de las rayas
    .SetOddEvenRowColor &HFFFFFF, &H0&, &HEFEFEF, &H0& ' setea colores intermedios por fila
    .RowHeight(-1) = 10
    .GridColor = &H8000000F ' color de las lineas de la grilla
    .CursorStyle = CursorStyleArrow ' cursor tipo flechita
    .UserColAction = UserColActionDefault
    .UserResize = UserResizeColumns
    .UserResizeCol = UserResizeOn
    .UserResizeRow = UserResizeOff
    .EditEnterAction = EditEnterActionNext
    
    ' tool tip text de cada celdita
    .TextTipDelay = 200
    .TextTip = TextTipFixed
    .SetTextTipAppearance "Verdana", "8", False, False, &HFFFF&, &H800000
    
    ' el font
    .Font.Name = "Verdana"
    .Font.Size = 8
    .Font.Bold = False
    .Font.Strikethrough = False
    .Font.Underline = False
    .Font.Italic = False
    .FontName = "Verdana"
    .FontSize = 8
    .FontBold = False
    .FontStrikethru = False
    .FontUnderline = False
    .FontItalic = False

    
    ' el sroll
    .ScrollBars = ScrollBarsBoth
    .ScrollBarMaxAlign = True
    .ScrollBarShowMax = True
    .ScrollBarTrack = ScrollBarTrackBoth
    .ScrollBarHeight = -1
    .ScrollBarWidth = -1
    .ShowScrollTips = ShowScrollTipsBoth
    .ScrollBarExtMode = True
End With

varSpread.Redraw = True

Screen.MousePointer = vbDefault

End Sub

'kjcv 06.07.18
Public Function ValidaCliente(Codigo As String) As Boolean
Dim Sql As String
Dim VEstado As String * 1
Dim Tabla As New ADODB.Recordset
Dim TGen1_Conexion As Variant
Dim gstrMensajeClienteBloqueado As String

VEstado = "V"
gstrMensajeClienteBloqueado = "CLIENTE BLOQUEADO!!! NO ATENDER..."

'Valida Cliente No vigente
Sql = "SELECT Direccion, Vigencia, Id_Comuna, Id_Ciudad, Id_Pais FROM Glbl_Cliente_Proveedor WHERE Id_Cliente_Proveedor='" & Codigo & "'"
If Conexion.SendHost(Sql, Tabla, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Tabla.RecordCount <> 0 Then
        
        If Not IsNull(Tabla!vigencia) Then VEstado = Trim$(Tabla!vigencia)
    Else
        MsgBox "Elisa no puede encontrar el Cliente." & Chr(13) & "Por favor, intente nuevamente.", vbExclamation, "Elisa"
        TGen1_Conexion.CloseHost Tabla
        Exit Function
    End If
End If
Conexion.CloseHost Tabla

If Trim$(VEstado) <> "S" Then
    MsgBox gstrMensajeClienteBloqueado, vbExclamation, "CLIENTE BLOQUEADO"
    ValidaCliente = False
    Exit Function
Else
    ValidaCliente = True
    
End If

End Function

