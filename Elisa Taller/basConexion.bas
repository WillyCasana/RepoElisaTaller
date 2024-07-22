Attribute VB_Name = "basConexion"
Option Explicit
Public Sub Main()
Dim lstrArchivoIni As String
lstrArchivoIni = Command()
gstrArchivoIni = Command()
gintProcedencia = 0
gapAccion = apninguno
'If apfLogin.apLogin(lstrArchivoIni, Conexion, adUseClient, cnnAux, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario, gstrIdEmpleado, LoadResString(4)) = True Then
If Libreria.Login("elisataller", gstrArchivoIni, Conexion, strConnect, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario) = True Then
    '//Tllr_00
    If Not Atributos("Glbl", "Tllr_00", True, True, True, True) Then
        MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
        Exit Sub
    End If
    
    gstrEmpresa = NombreEmpresa(gstrIdEmpresa)
    gstrSucursal = NombreSucursal(gstrIdEmpresa, gstrIdSucursal)
    gstrDirSuc = DireccionSucursal(gstrIdEmpresa, gstrIdSucursal)
    gstrUsuario = gstrIdUsuario
    gstrPathReporte = LetConnectionString("TLLR", "RPT", lstrArchivoIni, 256)
    gstrRutaApclient = LetConnectionString("APSERVER", "APCLIENT", lstrArchivoIni, 256)
'    gstrPassWordLiquidador = PWLiquidador(gstrIdEmpresa, gstrIdSucursal)
    gstrIdEmpleado = CodigoEmpleado(gstrIdUsuario)
    gstrPassWordLiquidador = PasswordLiquidador(gstrIdEmpresa, gstrIdSucursal, gstrIdEmpleado)
    With frmMain
        .stbMain.Panels(1).Text = gstrEmpresa
        .stbMain.Panels(1).Bevel = sbrInset
        .stbMain.Panels(2).Text = gstrSucursal '& "/" & gstrInicial
        .stbMain.Panels(2).Bevel = sbrInset
        .stbMain.Panels(3).Text = gstrUsuario
        .stbMain.Panels(3).Bevel = sbrInset
        .Show
    End With
MarcaxDefault
 If ParametrosDefecto(gstrIdEmpresa, gstrIdSucursal) = False Then
    MsgBox LoadResString(101), vbCritical + vbOKOnly, "ElisaTaller"
 End If
 
 If ParametrosInternacionales(gstrIdEmpresa) = False Then
    MsgBox "Parametros Internacionales Incompletos", vbCritical + vbOKOnly, "ElisaTaller"
 End If
 
gstrProcedencia = ""
'ActualizaConfiguracionWindows
'NormalizaConfiguracion

'MODIFICADO POR FDO DIAZ EL 02/01/2001 ACTIVA Y DESACTIVA OPCIONES SEGUN EMPRESA
If InStr(gstrEmpresa, "POMPEYO") = 1 Or InStr(gstrEmpresa, "SERINFO") Or gstrIdEmpresa = "969851806" Then
    frmMain.mnuServiteca.Visible = True
End If

frmMain.Show
End If
End Sub


Public Function PWLiquidador(strEmpresa As String, strSucursal As String) As String
Dim strSql As String
Dim recAux As New ADODB.Recordset

Set recAux = New ADODB.Recordset
strSql = "Select PasswordLiquidador as Parametro from Tllr_Parametro"
strSql = strSql & " Where Id_Empresa='" & strEmpresa & "' And Id_Sucursal='" & strSucursal & "' AND ID=1"
If Conexion.SendHost(strSql, recAux, adOpenStatic, adLockReadOnly, gcTiempoEspera) = apOk Then
    With recAux
        If Not .BOF And Not .EOF Then
            PWLiquidador = IIf(Not IsNull(!parametro), !parametro, "NULA")
        End If
        .Close
    End With
End If
Conexion.CloseHost recAux
Set recAux = Nothing

End Function

Public Function PasswordLiquidador(strEmpresa As String, strSucursal As String, pstrIdMecanico As String)
Dim strSql As String
Dim adoPwL As New ADODB.Recordset

Set adoPwL = New ADODB.Recordset

strSql = "SELECT passwordliquidador FROM Tllr_mecanicos"
strSql = strSql & " WHERE  Vigencia='S' and Rut_mecanico = '" & pstrIdMecanico & "' And id_empresa='" & strEmpresa & "' And Id_Sucursal='" & strSucursal & "'"
If Conexion.SendHost(strSql, adoPwL, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With adoPwL
            If Not .BOF And Not .EOF Then
                PasswordLiquidador = IIf(Not IsNull(!PasswordLiquidador), !PasswordLiquidador, "NULA")
            End If
            .Close
            
        End With
End If
Conexion.CloseHost adoPwL
Set adoPwL = Nothing
 
End Function







