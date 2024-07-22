Attribute VB_Name = "basProcedimientos"
Option Explicit
Public Sub EliminaRegistros(pstrIdEmpresa As String, _
                            pstrIdSucursal As String, _
                            pstrIdOT As String, _
                            pstrSeccionOT As String)

gstrSql = "DELETE FROM TLLR_FACTURACION "
gstrSql = gstrSql & " WHERE ID_EMPRESA='" & pstrIdEmpresa & "' "
gstrSql = gstrSql & " AND ID_SUCURSAL='" & pstrIdSucursal & "' "
gstrSql = gstrSql & " AND ID_OT='" & pstrIdOT & "' "
gstrSql = gstrSql & " AND SECCION_OT='" & pstrSeccionOT & "' And Estado = 'V'"
Conexion.SendHost gstrSql, , , , gcTiempoEspera

End Sub

Public Sub FillPartePieza(dtcObjeto As DataCombo, datObjeto As Adodc)
Dim mstrSQL As String
Dim AdoPrincipal As New ADODB.Recordset

mstrSQL = "SELECT Id_Parte_Pieza AS CODIGO, Descripcion As NOMBRE"
mstrSQL = mstrSQL & " From Tllr_Parte_Pieza order by Descripcion"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With datObjeto
            Set .Recordset = AdoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcObjeto.ListField = "Nombre"
                dtcObjeto.BoundColumn = "Codigo"
            End If
        End With
End If ' por el otro
Set AdoPrincipal = New ADODB.Recordset
Conexion.CloseHost AdoPrincipal
End Sub

Public Sub NormalizaConfiguracion()
    Dim ConfiguracionRegionalObligatoria As apCFGRegionalSTR
    Dim blnCambios As Boolean
    
    '//Solo deben inicializar los parametros que quieren ser chequeados...
    With ConfiguracionRegionalObligatoria
        '// Numeros
        .NumSimboloDecimal = "."
        .NumDigitosDecimales = "2"
        .NumSimboloSeparacionMiles = ","
        .NumDigitosGrupo = "3;0"
        .NumSimboloSignoNegativo = "-"
        '.NumFormatoNumeroNegativo = "1"
        '.NumMostrarCerosIzquierda = "1"
        '.NumSeperadorListas = ";"
        '.NumSistemaMedida = "0"
        
        '// Moneda
        .MonSimboloMoneda = gstrMonedaLocal '"$"
        '.MonFormatoMonedaPositivo = "2"
        '.MonFormatoMonedaNegativo = "14"
        .MonSimboloDecimal = "."
        .MonDigitosDecimales = "2"
        .MonSimboloSeparacionMiles = ","
        .MonDigitosGrupo = "3;0"
        
        '//Hora
        .FormatoHora = "HH:mm:ss"
        .SeperadoHora = ":"
        '.SimboloAM = "AM"
        '.SimboloPM = "PM"
        
        '//Fecha
        .FormatoFechaCorta = "dd/MM/yyyy"
        .SeparadorFecha = "/"
        '.FormatoFechaLarga = "dddd, dd' de 'MMMM' de 'yyyy"
    End With
    
    If NormalizaConfiguracionRegional(ConfiguracionRegionalObligatoria, blnCambios) Then
        If blnCambios Then
            End
        End If
    Else
        MsgBox "La configuración regional no esta normalizada imposible ejecutar este programa...", vbInformation, "Advertencia"
        End
    End If
End Sub
Public Sub ActualizaConfiguracionWindows()
Dim Valido As Integer

Valido = WritePrivateProfileString("intl", "iTime", "1", "win.ini")
Valido = WritePrivateProfileString("intl", "s1159", "a.m.", "win.ini")
Valido = WritePrivateProfileString("intl", "s2359", "p.m.", "win.ini")
Valido = WritePrivateProfileString("intl", "sCurrency", gstrMonedaLocal, "win.ini")
Valido = WritePrivateProfileString("intl", "sDate", "/", "win.ini")
Valido = WritePrivateProfileString("intl", "sDecimal", ".", "win.ini")
Valido = WritePrivateProfileString("intl", "sList", ";", "win.ini")
Valido = WritePrivateProfileString("intl", "sLongDate", "dddd, dd' de 'MMMM' de 'yyyy", "win.ini")
Valido = WritePrivateProfileString("intl", "sShortDate", "dd/MM/yyyy", "win.ini")
Valido = WritePrivateProfileString("intl", "sTime", ":", "win.ini")

End Sub
Public Sub MarcaTexto(ByRef Text As TextBox)
    Text.SelStart = 0
    Text.SelLength = Len(Text)
End Sub

Public Sub FillTipoCargo(dtcObjeto As DataCombo, datObjeto As Adodc)

gstrSql = "SELECT Id_Tipo_Cargo AS CODIGO, Descripcion  AS NOMBRE FROM Tllr_Tipo_Cargo WHERE Id_Empresa='" & gstrIdEmpresa & "' and Vigencia = N'S' ORDER BY Descripcion"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With datObjeto
        Set .Recordset = gadoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcObjeto.ListField = "Nombre"
            dtcObjeto.BoundColumn = "Codigo"
        End If
    End With
End If ' por el otro
Set gadoPrincipal = New ADODB.Recordset
Conexion.CloseHost gadoPrincipal
End Sub
Public Sub FillMecanicos(dtcObjeto As DataCombo, datObjeto As Adodc)

gstrSql = "SELECT Id_Mecanico AS CODIGO, Nombre FROM Tllr_Mecanicos WHERE Vigencia = 'S' AND ID_EMPRESA='" & gstrIdEmpresa & "' AND ID_SUCURSAL='" & gstrIdSucursal & "' ORDER BY Nombre "
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With datObjeto
        Set .Recordset = gadoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcObjeto.ListField = "Nombre"
            dtcObjeto.BoundColumn = "Codigo"
        End If
    End With
End If ' por el otro
Set gadoPrincipal = New ADODB.Recordset
Conexion.CloseHost gadoPrincipal
End Sub
Public Sub ImprimeObjeto(intFil As Integer, intCol As Integer, strTexto As String, intFontSize As Integer, Optional blnBold As Boolean, Optional blnItalic As Boolean)
Printer.CurrentY = intFil
Printer.CurrentX = intCol
Printer.FontSize = intFontSize
Printer.FontBold = IIf(Not IsMissing(blnBold), blnBold, False)
Printer.FontItalic = IIf(Not IsMissing(blnItalic), blnItalic, False)
Printer.Print strTexto
End Sub

Public Sub IncrementaCorrelativo(gParametro As gcParametro, strEmpresa As String, strSucursal As String)
Dim strSql As String

Select Case gParametro
Case Is = 1
    strSql = "Update Tllr_Parametro Set NroPreMec = NroPreMec + 1 "
Case Is = 2
    strSql = "Update Tllr_Parametro Set NroPreCar = NroPreCar + 1 "
Case Is = 3
    strSql = "Update Tllr_Parametro Set NroOTMec = NroOTMec + 1 "
Case Is = 4
    strSql = "Update Tllr_Parametro Set NroOTCar = NroOTCar + 1 "
Case Else
    Exit Sub
End Select

strSql = strSql & "Where Id_Empresa='" & strEmpresa & "' And Id_Sucursal='" & strSucursal & "' AND ID=1"

Conexion.SendHost strSql, , , , gcTiempoEspera

End Sub
Public Sub IncrementaCorrelativoOtrosServicios(strEmpresa As String, strSucursal As String)
Dim strSql As String
strSql = "Update Tllr_Parametro Set CorrelativoOtrosservicios = CorrelativoOtrosservicios + 1 "
strSql = strSql & "Where Id_Empresa='" & strEmpresa & "' And Id_Sucursal='" & strSucursal & "' AND ID=1"
Conexion.SendHost strSql, , , , gcTiempoEspera
End Sub
Public Sub IncrementaCorrelativoTrabajosTerceros(strEmpresa As String, strSucursal As String)
Dim strSql As String
strSql = "Update Tllr_Parametro Set CorrelativoTrabajoTercero = CorrelativoTrabajoTercero + 1 "
strSql = strSql & "Where Id_Empresa='" & strEmpresa & "' And Id_Sucursal='" & strSucursal & "' AND ID=1"
Conexion.SendHost strSql, , , , gcTiempoEspera
End Sub
Public Sub SelectingItem(lvwObjeto As ListView, Opcion As gopOpcionItem)
Dim intIndice As Integer

If lvwObjeto.ListItems.Count > 0 Then
    For intIndice = 1 To lvwObjeto.ListItems.Count
        Set lvwObjeto.SelectedItem = lvwObjeto.ListItems(intIndice)
        If Opcion = gcSelectAll Then
            lvwObjeto.SelectedItem.Checked = True
        ElseIf Opcion = gcUnSelectAll Then
            lvwObjeto.SelectedItem.Checked = False
        End If
    Next
End If
End Sub

Public Sub SetCheckOff(lvwObjeto As ListView)
Dim intX As Integer
With lvwObjeto
    If .ListItems.Count > 0 Then
        For intX = 1 To .ListItems.Count
            Set .SelectedItem = .ListItems(intX)
            If .SelectedItem.Checked = True Then .SelectedItem.Checked = False
        Next
    End If
End With
End Sub

Public Sub FillTipoCono(dtcObjeto As DataCombo, datObjeto As Adodc)
dtcObjeto.Enabled = True
gstrSql = "SELECT Id_Tipo_Cono as codigo, Color as nombre FROM Tllr_Tipo_Cono WHERE Vigencia = 'S' order by Color"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With datObjeto
        Set .Recordset = gadoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcObjeto.ListField = "Nombre"
            dtcObjeto.BoundColumn = "Codigo"
        End If
    End With
End If ' por el otro
Set gadoPrincipal = New ADODB.Recordset
Conexion.CloseHost gadoPrincipal
End Sub
Public Sub FillRecepcionista(dtcObjeto As DataCombo, datObjeto As Adodc)
    gstrSql = "SELECT Id_Mecanico AS CODIGO, Nombre FROM Tllr_Mecanicos WHERE Es_Recepcionista = 'S' AND ID_EMPRESA='" & gstrIdEmpresa & "' AND ID_SUCURSAL='" & gstrIdSucursal & "' And vigencia='S' ORDER BY Nombre"
    dtcObjeto.Enabled = True
    If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        With datObjeto
            Set .Recordset = gadoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcObjeto.ListField = "Nombre"
                dtcObjeto.BoundColumn = "Codigo"
            End If
        End With
    End If ' por el otro
    Set gadoPrincipal = New ADODB.Recordset
    Conexion.CloseHost gadoPrincipal
End Sub
Public Sub FillLiquidador(dtcObjeto As DataCombo, datObjeto As Adodc)
    gstrSql = "SELECT Id_Mecanico AS CODIGO, Nombre FROM Tllr_Mecanicos WHERE Es_Liquidador = 'S' AND Vigencia='S' AND ID_EMPRESA='" & gstrIdEmpresa & "' AND ID_SUCURSAL='" & gstrIdSucursal & "'"
    dtcObjeto.Enabled = True
    If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        With datObjeto
            Set .Recordset = gadoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcObjeto.ListField = "Nombre"
                dtcObjeto.BoundColumn = "Codigo"
            End If
        End With
    End If ' por el otro
    Set gadoPrincipal = New ADODB.Recordset
    Conexion.CloseHost gadoPrincipal
End Sub
Public Sub FillActivador(dtcObjeto As DataCombo, datObjeto As Adodc)
    gstrSql = "SELECT Id_Mecanico AS CODIGO, Nombre FROM Tllr_Mecanicos WHERE Es_Activador = 'S' AND ID_EMPRESA='" & gstrIdEmpresa & "' AND ID_SUCURSAL='" & gstrIdSucursal & "'"
    dtcObjeto.Enabled = True
    If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        With datObjeto
            Set .Recordset = gadoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcObjeto.ListField = "Nombre"
                dtcObjeto.BoundColumn = "Codigo"
            End If
        End With
    End If ' por el otro
    Set gadoPrincipal = New ADODB.Recordset
    Conexion.CloseHost gadoPrincipal
End Sub


Public Sub FillGarantia(dtcObjeto As DataCombo, datObjeto As Adodc, MuestraOtPresupuesto As Boolean)

If MuestraOtPresupuesto = False Then
    gstrSql = "SELECT Id_Garantia AS CODIGO, Descripcion AS NOMBRE FROM Tllr_Garantias WHERE Id_Empresa='" & gstrIdEmpresa & "' and VIGENCIA = 'S' AND ID_GARANTIA <> 'PRE' ORDER BY Descripcion"
Else
    gstrSql = "SELECT Id_Garantia AS CODIGO, Descripcion AS NOMBRE FROM Tllr_Garantias WHERE Id_Empresa='" & gstrIdEmpresa & "' and VIGENCIA = 'S' ORDER BY Descripcion"
End If
dtcObjeto.Enabled = True
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With datObjeto
        Set .Recordset = gadoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcObjeto.ListField = "Nombre"
            dtcObjeto.BoundColumn = "Codigo"
            If .Recordset.RecordCount < 2 Then
                dtcObjeto.BoundText = .Recordset!Codigo
                dtcObjeto.Enabled = False
            End If
        End If
    End With
End If ' por el otro
Set gadoPrincipal = New ADODB.Recordset
Conexion.CloseHost gadoPrincipal
End Sub
Public Sub FillPromocion(dtcObjeto As DataCombo, datObjeto As Adodc)
Dim sqlPromo As String

sqlPromo = "select Id_Promo as CODIGO, Descripcion as Nombre from Promocion where Id_Empresa='" & gstrIdEmpresa & "' and Id_Sucursal='" & gstrIdSucursal & "' and VIGENCIA = 'S'   "
dtcObjeto.Enabled = True
If Conexion.SendHost(sqlPromo, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With datObjeto
        Set .Recordset = gadoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcObjeto.ListField = "Nombre"
            dtcObjeto.BoundColumn = "Codigo"
        End If
    End With
Set gadoPrincipal = New ADODB.Recordset
Conexion.CloseHost gadoPrincipal
End If

End Sub
Public Sub FillTrabajos(dtcObjeto As DataCombo, datObjeto As Adodc)
Dim sqlTrabajo As String

sqlTrabajo = "select Id_Trabajo as CODIGO, Descripcion as Nombre from Tllr_Trabajo where Id_Empresa='" & gstrIdEmpresa & "' and VIGENCIA = 'S'   "
dtcObjeto.Enabled = True
If Conexion.SendHost(sqlTrabajo, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With datObjeto
        Set .Recordset = gadoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcObjeto.ListField = "Nombre"
            dtcObjeto.BoundColumn = "Codigo"
        End If
    End With
Set gadoPrincipal = New ADODB.Recordset
Conexion.CloseHost gadoPrincipal
End If

End Sub

Public Sub FillTime(intHraIni As Integer, intHraFin As Integer, cboObjeto As ComboBox)
Dim intHra As Integer, intMin As Integer

For intHra = intHraIni To intHraFin
    For intMin = 0 To 59 Step gintIntervaloMinutos
        cboObjeto.AddItem Format$(intHra, "00") & ":" & Format$(intMin, "00")
    Next
Next
End Sub

Sub FillConceptosVsCiaSeguro(dtcObjeto As DataCombo, datObjeto As Adodc, strCiaSeg As String)

gstrSql = "SELECT Tllr_CiaSeguro_Concepto.Id_Concepto as codigo, Tllr_Concepto.Descripcion as nombre"
gstrSql = gstrSql & " FROM Tllr_CiaSeguro_Concepto LEFT OUTER JOIN Tllr_Concepto ON Tllr_CiaSeguro_Concepto.Id_Concepto = Tllr_Concepto.Id_Concepto"
gstrSql = gstrSql & " WHERE Tllr_CiaSeguro_Concepto.Id_Compañia_Seguro = '" & strCiaSeg & "' "

If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With datObjeto
        Set .Recordset = gadoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcObjeto.ListField = "Nombre"
            dtcObjeto.BoundColumn = "Codigo"
'            If .Recordset.RecordCount < 2 Then
'                dtcObjeto.BoundText = .Recordset!codigo
'                dtcObjeto.Enabled = False
'            End If
        End If
    End With
End If ' por el otro
Set gadoPrincipal = New ADODB.Recordset
Conexion.CloseHost gadoPrincipal
End Sub



Public Sub GeneraRegistroFactura(pstrIdEmpresa As String, pstrIdSucursal As String, pstrIdOT As String, _
                                    pstrSeccionOT As String, pstrPatente As String, pstrMarca As String, _
                                    pstrModelo As String, pstrCliente As String, pcurInsumos As Currency, _
                                    pcurMateriales As Currency, pcurSeguroTaller As Currency, _
                                    pstrRutCliente As String, dteFechaLiquidacion As Date, _
                                    Optional pstrIdCompañia As String, Optional pstrCompañia As String, _
                                    Optional pcurDeducibleUF As Currency, Optional pcurDeduciblePss As Currency)
Dim curTotalMecanica As Currency
Dim curTotalOtros As Currency
Dim curTotalManoObra As Currency
Dim curTotalCarroceria As Currency
Dim curTotalTerceros As Currency
Dim curTotalRepuestos As Currency

Dim strCargoActual As String
Dim strAquienFactura As String
Dim recAux As New ADODB.Recordset
Dim curTotalLinea As Currency
Dim curDeducible As Currency
Dim strFacturable As String

'//////////////////////////////ENCABEZADO

gstrSql = "SELECT ID_TIPO_CARGO , DESCRIPCION AS CARGO, QUIEN_CANCELA, FACTURABLE FROM TLLR_TIPO_CARGO WHERE ID_EMPRESA='" & gstrIdEmpresa & "' ORDER BY ID_TIPO_CARGO"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With gadoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            While Not .EOF
                strCargoActual = !Id_Tipo_Cargo
                strFacturable = IIf(IsNull(!Facturable), "S", !Facturable)
                If strCargoActual = gstrCargoDeducibleMenos Then
                    strAquienFactura = frmRecepcion.lblCompañia.Tag
                Else
                    If Val(Mid(ValorNulo(!QUIEN_CAnCELA), 1, 8)) > 0 Then
                        strAquienFactura = !QUIEN_CAnCELA
                    Else
                        strAquienFactura = pstrRutCliente
                    End If
                End If
                
                '//////////////////mecanica
                gstrSql = "SELECT SUM(SubTotal) AS TOTALMECANICA "
                gstrSql = gstrSql & " From Tllr_Mecanica_OT"
                gstrSql = gstrSql & " WHERE ID_OT = '" & pstrIdOT & "' "
                gstrSql = gstrSql & " AND SECCION_OT = '" & pstrSeccionOT & "' "
                gstrSql = gstrSql & " AND ID_TIPO_CARGO='" & strCargoActual & "' "
                gstrSql = gstrSql & " AND ID_EMPRESA='" & pstrIdEmpresa & "' "
                gstrSql = gstrSql & " AND ID_SUCURSAL='" & pstrIdSucursal & "' "
                gstrSql = gstrSql & " AND FACTURADO='N'"
                gstrSql = gstrSql & " GROUP BY ID_EMPRESA,ID_SUCURSAL,Id_OT, Seccion_OT, Id_Tipo_Cargo"
                If Conexion.SendHost(gstrSql, recAux, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not recAux.BOF And Not recAux.EOF Then
                        recAux.MoveFirst
                        curTotalMecanica = recAux!TOTALMECANICA
                    Else
                        curTotalMecanica = 0
                    End If
                End If
                Conexion.CloseHost recAux
                '//////////////////carroceria
                gstrSql = "SELECT SUM(SubTotal) AS TOTALCARROCERIA "
                gstrSql = gstrSql & " From Tllr_Carroceria_OT"
                gstrSql = gstrSql & " WHERE ID_OT = '" & pstrIdOT & "' "
                gstrSql = gstrSql & " AND SECCION_OT = '" & pstrSeccionOT & "' "
                gstrSql = gstrSql & " AND ID_TIPO_CARGO='" & strCargoActual & "' "
                gstrSql = gstrSql & " AND ID_EMPRESA='" & pstrIdEmpresa & "' "
                gstrSql = gstrSql & " AND ID_SUCURSAL='" & pstrIdSucursal & "' "
                gstrSql = gstrSql & " AND FACTURADO='N'"
                gstrSql = gstrSql & " GROUP BY ID_EMPRESA,ID_SUCURSAL,Id_OT, Seccion_OT, Id_Tipo_Cargo"
                If Conexion.SendHost(gstrSql, recAux, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not recAux.BOF And Not recAux.EOF Then
                        recAux.MoveFirst
                        curTotalCarroceria = recAux!TOTALCARROCERIA
                    Else
                        curTotalCarroceria = 0
                    End If
                End If
                Conexion.CloseHost recAux
                '//////////////////otros servicios
                gstrSql = "SELECT SUM(SubTotal) AS TOTALOTROS "
                gstrSql = gstrSql & " From Tllr_OTRO_OT"
                gstrSql = gstrSql & " WHERE ID_OT = '" & pstrIdOT & "' "
                gstrSql = gstrSql & " AND SECCION_OT = '" & pstrSeccionOT & "' "
                gstrSql = gstrSql & " AND ID_TIPO_CARGO='" & strCargoActual & "' "
                gstrSql = gstrSql & " AND ID_EMPRESA='" & pstrIdEmpresa & "' "
                gstrSql = gstrSql & " AND ID_SUCURSAL='" & pstrIdSucursal & "' "
                gstrSql = gstrSql & " AND FACTURADO='N' "
                gstrSql = gstrSql & " GROUP BY ID_EMPRESA,ID_SUCURSAL,Id_OT, Seccion_OT, Id_Tipo_Cargo"
                If Conexion.SendHost(gstrSql, recAux, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not recAux.BOF And Not recAux.EOF Then
                        recAux.MoveFirst
                        curTotalOtros = recAux!TOTALOTROS
                    Else
                        curTotalOtros = 0
                    End If
                End If
                Conexion.CloseHost recAux
                '//////////////////Terceros
                gstrSql = "SELECT SUM(SubTotal) AS TOTALTERCEROS "
                gstrSql = gstrSql & " From Tllr_Terceros_OT"
                gstrSql = gstrSql & " WHERE ID_OT = '" & pstrIdOT & "' "
                gstrSql = gstrSql & " AND SECCION_OT = '" & pstrSeccionOT & "' "
                gstrSql = gstrSql & " AND ID_TIPO_CARGO='" & strCargoActual & "' "
                gstrSql = gstrSql & " AND ID_EMPRESA='" & pstrIdEmpresa & "' "
                gstrSql = gstrSql & " AND ID_SUCURSAL='" & pstrIdSucursal & "' "
                gstrSql = gstrSql & " AND FACTURADO='N' "
                gstrSql = gstrSql & " GROUP BY ID_EMPRESA,ID_SUCURSAL,Id_OT, Seccion_OT, Id_Tipo_Cargo"
                If Conexion.SendHost(gstrSql, recAux, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not recAux.BOF And Not recAux.EOF Then
                        recAux.MoveFirst
                        curTotalTerceros = recAux!TOTALTERCEROS
                    Else
                        curTotalTerceros = 0
                    End If
                End If
                Conexion.CloseHost recAux
                '//////////////////Repuestos
                gstrSql = "SELECT SUM(SubTotal) AS TOTALREPUESTOS "
                gstrSql = gstrSql & " From Tllr_Repuestos_OT"
                gstrSql = gstrSql & " WHERE ID_OT = '" & pstrIdOT & "' "
                gstrSql = gstrSql & " AND SECCION_OT = '" & pstrSeccionOT & "' "
                gstrSql = gstrSql & " AND ID_TIPO_CARGO='" & strCargoActual & "' "
                gstrSql = gstrSql & " AND ID_EMPRESA='" & pstrIdEmpresa & "' "
                gstrSql = gstrSql & " AND ID_SUCURSAL='" & pstrIdSucursal & "' "
                gstrSql = gstrSql & " AND FACTURADO='N' "
                gstrSql = gstrSql & " GROUP BY ID_EMPRESA,ID_SUCURSAL,Id_OT, Seccion_OT, Id_Tipo_Cargo"
                If Conexion.SendHost(gstrSql, recAux, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not recAux.BOF And Not recAux.EOF Then
                        recAux.MoveFirst
                        curTotalRepuestos = recAux!TOTALREPUESTOS
                    Else
                        curTotalRepuestos = 0
                    End If
                End If
                Conexion.CloseHost recAux
                '////////////////////////////////////////////////////////////////
                curTotalManoObra = curTotalMecanica + curTotalOtros
                
                'chequea si el cliente ya esta facturado, para no facturar los insumos de nuevo
                If VerificaClienteFacturado(pstrIdOT, pstrSeccionOT, strCargoActual) = True Then
                    curTotalLinea = curTotalManoObra + curTotalCarroceria + curTotalTerceros + curTotalRepuestos
                Else
                curTotalLinea = curTotalManoObra + curTotalCarroceria + curTotalTerceros + curTotalRepuestos + IIf(strCargoActual = gstrCargoDeducibleMas, pcurInsumos, 0) + IIf(strCargoActual = gstrCargoDeducibleMas, pcurSeguroTaller, 0)
'                    curTotalLinea = curTotalManoObra + curTotalCarroceria + curTotalTerceros + (curTotalRepuestos - pcurMateriales) + IIf(strCargoActual = "01", pcurInsumos, 0) + IIf(strCargoActual = "01", pcurMateriales, 0) + IIf(strCargoActual = "01", pcurSeguroTaller, 0)
                End If
                
                '//////////////////////////////////////////////////////////////
                'AQUI PREGUNTO POR EL DEDUCIBLE
                If strCargoActual = gstrCargoDeducibleMas Then  'AL CLIENTE LE SUMAMOS EL DEDUCIBLE
                   If Val(frmRecepcion.txtDeduciblePesos) <> 0 Or Val(frmRecepcion.txtDeducibleUF) <> 0 Then
                        If VeriDeducible(pstrIdOT, pstrSeccionOT) = False Then  'verifica si esta facturado
                            curTotalLinea = curTotalLinea + Val(frmRecepcion.txtDeduciblePesos)
                            'curTotalLinea = curTotalManoObra + curTotalCarroceria + curTotalTerceros + curTotalRepuestos + IIf(strCargoActual = gstrCargoDeducibleMas, pcurInsumos, 0) + IIf(strCargoActual = gstrCargoDeducibleMas, pcurSeguroTaller, 0) + Val(frmRecepcion.txtDeduciblePesos)
                            'curTotalLinea = curTotalManoObra + curTotalCarroceria + curTotalTerceros + (curTotalRepuestos - pcurMateriales) + IIf(strCargoActual = "01", pcurInsumos, 0) + IIf(strCargoActual = "01", pcurMateriales, 0) + IIf(strCargoActual = "01", pcurSeguroTaller, 0) + Val(frmRecepcion.txtDeduciblePesos)
                        End If
                    End If
                End If
                If strCargoActual = gstrCargoDeducibleMenos Then  'A LA CIA. SEGUROS LE RESTAMOS EL DEDUCIBLE
                    If Val(frmRecepcion.txtDeduciblePesos) <> 0 Or Val(frmRecepcion.txtDeducibleUF) <> 0 Then
                        curTotalLinea = curTotalLinea - Val(frmRecepcion.txtDeduciblePesos)
                        'curTotalLinea = (curTotalManoObra + curTotalCarroceria + curTotalTerceros + curTotalRepuestos + IIf(strCargoActual = gstrCargoDeducibleMas, pcurInsumos, 0) + IIf(strCargoActual = gstrCargoDeducibleMas, pcurSeguroTaller, 0)) - Val(frmRecepcion.txtDeduciblePesos)
                        'curTotalLinea = (curTotalManoObra + curTotalCarroceria + curTotalTerceros + (curTotalRepuestos - pcurMateriales) + IIf(strCargoActual = "01", pcurInsumos, 0) + IIf(strCargoActual = "01", pcurMateriales, 0) + IIf(strCargoActual = "01", pcurSeguroTaller, 0)) - Val(frmRecepcion.txtDeduciblePesos)
                    End If
                End If
 
                If curTotalLinea > 0 Then
                
                    gstrSql2 = "SELECT  * FROM TLLR_FACTURACION "
                    gstrSql2 = gstrSql2 & " WHERE ID_EMPRESA='" & pstrIdEmpresa & "' "
                    gstrSql2 = gstrSql2 & " AND ID_SUCURSAL='" & pstrIdSucursal & "' "
                    gstrSql2 = gstrSql2 & " AND ID_OT='" & pstrIdOT & "' "
                    gstrSql2 = gstrSql2 & " AND SECCION_OT='" & pstrSeccionOT & "' And Id_Cargo = '" & strCargoActual & "'"
                    If Conexion.SendHost(gstrSql2, recAux, adOpenForwardOnly, adLockOptimistic, gcTiempoEspera) = apOk Then
                            If Not recAux.BOF And Not recAux.EOF Then
                               'hay cargo insertado
                            Else
                                
'                            End If
'                    End If
'                    Conexion.CloseHost recAux
                    
                '////////////////////////////////////////////////////////////////
                '//////////////////AQUI SE CREA LOS REGISTROS PARA LOS CARGOS
                    gstrSql = "INSERT INTO TLLR_FACTURACION "
                    gstrSql = gstrSql & " (Id_Empresa, Id_Sucursal, Id_OT , Seccion_OT, Id_Cargo, "
                    gstrSql = gstrSql & " Patente, Marca, Modelo, Cliente, Rut, A_Quien_Factura,"
                    gstrSql = gstrSql & " Total_Mano_Obra, Total_Carroceria, Total_Terceros, Total_Repuestos, Insumos ,"
                    gstrSql = gstrSql & " Total_General, Porcentaje_Descuento, Monto_Descuento,  Total_Neto, "
                    gstrSql = gstrSql & " Valor_Afecto,  Valor_Exento,  Iva, TOTAL , "
'                    gstrSql = gstrSql & " Estado, Fecha_Liquidacion,  Fecha_Facturacion, Materiales, SeguroTaller)"
                    'kjcv 14.01.14 se inserta campo Deducible_Pesos
                    gstrSql = gstrSql & " Estado, Fecha_Liquidacion,  Fecha_Facturacion, Materiales, SeguroTaller,Deducible_Pesos)"
                    gstrSql = gstrSql & " VALUES ( '" & pstrIdEmpresa & "', '" & pstrIdSucursal & "', '" & pstrIdOT & "' , '" & pstrSeccionOT & "',  '" & strCargoActual & "', "
                    gstrSql = gstrSql & " '" & pstrPatente & "',  '" & pstrMarca & "', '" & pstrModelo & "', '" & pstrCliente & "', '" & pstrRutCliente & "', '" & strAquienFactura & "',"
                    gstrSql = gstrSql & " " & curTotalManoObra & ", " & curTotalCarroceria & ",  " & curTotalTerceros & ", " & curTotalRepuestos & ", " & IIf(strCargoActual = gstrCargoDeducibleMas, pcurInsumos, 0) & ", "
                    gstrSql = gstrSql & " " & curTotalLinea & ", 0, 0, " & curTotalLinea & ", " & curTotalLinea & ", "
                    gstrSql = gstrSql & " 0, 0, 0, 'V', '" & dteFechaLiquidacion & "', '" & dteFechaLiquidacion & "', "
'                    gstrSql = gstrSql & 0 & "," & IIf(strCargoActual = gstrCargoDeducibleMas, pcurSeguroTaller, 0) & ")"
                    'kjcv 14.01.14 campo Deducible_Pesos
                    gstrSql = gstrSql & 0 & "," & IIf(strCargoActual = gstrCargoDeducibleMas, pcurSeguroTaller, 0) & ""
                    gstrSql = gstrSql & "," & IIf(strCargoActual = gstrCargoDeducibleMas, frmRecepcion.txtDeduciblePesos, 0) & " )"
                    
                    Conexion.SendHost gstrSql, , adOpenKeyset, adLockOptimistic, gcTiempoEspera

                    curTotalMecanica = 0
                    curTotalOtros = 0
                    curTotalManoObra = 0
                    curTotalCarroceria = 0
                    curTotalTerceros = 0
                    curTotalRepuestos = 0
                    curTotalLinea = 0
                    gstrSql = ""
                    
                             End If
                    End If
                    Conexion.CloseHost recAux
                    
                End If
                .MoveNext
            Wend
        End If
    End With
End If

End Sub

