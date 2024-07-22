Attribute VB_Name = "Serviteca"
Option Explicit
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const Guion As String = "*"

Public gblnNuevo As Boolean
Public gstrRutCliente As String
Public gstrNombreCliente As String
Public GStrAnexoBuscadorItem As String
Public gstrDiasProximoLLamado As String
Public Retorno As String
Public tam As Integer
Public Valido As Integer

Public apConexion As New APCONADO.ConnectionAdo
Public adoConexion As New ADODB.Connection

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function traeCLIENTE(RutCliente As String) As String
Dim tablaCLIENTE As New ADODB.Recordset
Dim sql As String

sql = ""
sql = "SELECT Razon_Social FROM Glbl_Cliente_Proveedor WHERE Id_Cliente_Proveedor ='" & Trim$(RutCliente) & "' AND Vigencia='S'"
If Conexion.SendHost(sql, tablaCLIENTE, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If tablaCLIENTE.EOF = False And tablaCLIENTE.BOF = False Then
        traeCLIENTE = UCase$(Trim$(tablaCLIENTE!Razon_Social))
    Else
        traeCLIENTE = "."
    End If
Else
    traeCLIENTE = "."
End If
Conexion.CloseHost tablaCLIENTE

End Function


Function ExisteCliente(ByVal RutCliente As String) As Boolean
Dim tablaCli As New ADODB.Recordset
Dim sql As String

ExisteCliente = False
sql = ""
sql = "SELECT Id_Cliente_Proveedor FROM Glbl_Cliente_Proveedor WHERE Rut ='" & Trim$(RutCliente) & "'"
If Conexion.SendHost(sql, tablaCli, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If tablaCli.RecordCount <> 0 Then
        ExisteCliente = True
    Else
        ExisteCliente = False
    End If
Else
    MsgBox "Problemas de apertura en Base de Datos de Clientes.", vbCritical, "Maestro de Clientes"
End If
Conexion.CloseHost tablaCli

End Function

Function RutValido(ByVal rut As String) As Boolean
Dim taum As Integer
Dim sp, xru, xid As String
Dim p, N, i As Integer
Dim digito As String
Dim LARGO As Long
Dim re As Integer

If gstrValidaRut = "S" Then
    If rut = "" Then
        RutValido = False
        Exit Function
    End If
    
    taum = 0
    LARGO = Len(rut)
    xid = Mid$(rut, LARGO, LARGO)
    LARGO = LARGO - 1
    xru = Mid$(rut, 1, LARGO)
    N = Len(Trim$(xru))
    i = 2
    While N > 0
      sp = Mid$(Trim$(xru), N, 1)
      p = Val(sp)
      taum = taum + (i * p)
      If i = 7 Then
       i = 2
      Else
       i = i + 1
      End If
      N = N - 1
    Wend
    re = Int(taum / 11)
    re = taum - (re * 11)
    re = Int(11 - re)
    
    Select Case re
       Case 10
         digito = "K"
       Case 11
         digito = "0"
       Case Else
         digito = Trim$(Str$(re))
    End Select
    
    If xid = digito Then
       RutValido = True
    Else
       RutValido = False
    End If
End If
End Function


Public Function traeMARCA(IdMarca As String) As String
Dim tablaMarca As New ADODB.Recordset
Dim sql As String

sql = ""
sql = "SELECT Descripcion FROM Glbl_Marca WHERE Id_Marca ='" & IdMarca & "' AND Vigencia='S'"
If Conexion.SendHost(sql, tablaMarca, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If tablaMarca.EOF = False And tablaMarca.BOF = False Then
        traeMARCA = UCase$(Trim$(tablaMarca!Descripcion))
    Else
        traeMARCA = "."
    End If
Else
    traeMARCA = "."
End If
Conexion.CloseHost tablaMarca

End Function


Public Function traeMODELO(IdMarca As String, IdModelo As String) As String
Dim tablaModelo As New ADODB.Recordset
Dim sql As String

sql = ""
sql = "SELECT Descripcion FROM Glbl_Modelo WHERE Id_Marca ='" & IdMarca & "' AND Id_Modelo='" & IdModelo & "' AND Vigencia='S'"
If Conexion.SendHost(sql, tablaModelo, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If tablaModelo.EOF = False And tablaModelo.BOF = False Then
        traeMODELO = UCase$(Trim$(tablaModelo!Descripcion))
    Else
        traeMODELO = "."
    End If
Else
    traeMODELO = "."
End If
Conexion.CloseHost tablaModelo

End Function


Public Function TraeNumOT() As Double
Dim sql As String
Dim tablaNumDoc As New ADODB.Recordset

TraeNumOT = 1
sql = ""
sql = sql & "SELECT Max(Ultimo_Numero) AS Numero "
sql = sql & "FROM Srvt_Correlativo_OT "
sql = sql & "WHERE Srvt_Correlativo_OT.Id_Empresa='" & gstrIdEmpresa & "' "
sql = sql & "AND Srvt_Correlativo_OT.Id_Sucursal='" & gstrIdSucursal & "'"
If Conexion.SendHost(sql, tablaNumDoc, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
    If tablaNumDoc.EOF = False And tablaNumDoc.BOF = False Then
        If Not IsNull(tablaNumDoc!NUMERO) Then
            TraeNumOT = CDbl(tablaNumDoc!NUMERO) + 1
        End If
    End If
End If
Conexion.CloseHost tablaNumDoc

End Function

Public Function ProcesoRegistros(Accion As gcProceso, Optional Porcentaje As Single)
    Select Case Accion
        Case gcInicioProceso
            frmProceso.Show
        Case gcAvanceProceso
            frmProceso.pbProceso.Value = Porcentaje
            frmProceso.lblProceso = Format(Porcentaje, "##0") & "%"
            frmProceso.Refresh
        Case gcFinProceso
            frmProceso.pbProceso.Value = 100
            frmProceso.lblProceso = Format(100, "##0") & "%"
            frmProceso.Refresh
            Unload frmProceso
    End Select
End Function

Public Function FijarFormulario(ByRef Formulario As Form)
    Dim x As Long, x1 As Long
    Dim y As Long, y1 As Long
    
    x = (Screen.Height - Formulario.Height) / 2
    y = (Screen.Width - Formulario.Width) / 2
    x1 = x + Formulario.Height
    y1 = y + Formulario.Width

    Call SetWindowPos(Formulario.hwnd, -1, x, y, x1, y1, SWP_SHOWWINDOW + SWP_NOSIZE + SWP_NOMOVE)
End Function

Public Function ValorNuloNum(Valor As Variant) As Variant
    If IsNull(Valor) Then
        ValorNuloNum = "0"
    Else
        ValorNuloNum = Valor
    End If
End Function

