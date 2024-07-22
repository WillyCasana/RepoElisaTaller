Attribute VB_Name = "cfgRegional"
Option Explicit

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Public Const HKEY_CURRENT_USER = &H80000001
Public Const KEY_READ = &H20019
Public Const KEY_WRITE = &H20006
Public Const KEY_SET_VALUE = &H2
Public Const REG_SZ = 1

Public Const EWX_FORCE = 4 'Force any applications to quit instead of prompting the user to close them.
Public Const EWX_LOGOFF = 0 'Log off the network.
Public Const EWX_POWEROFF = 8 'Shut down the system and, if possible, turn the computer off.
Public Const EWX_REBOOT = 2 'Perform a full reboot of the system.
Public Const EWX_SHUTDOWN = 1 'Shut down the system.



Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Public Type apCFGRegionalSTR
    NumSimboloDecimal As String
    NumDigitosDecimales As String
    NumSimboloSeparacionMiles As String
    NumDigitosGrupo As String
    NumSimboloSignoNegativo As String
    NumFormatoNumeroNegativo As String
    NumMostrarCerosIzquierda As String
    NumSeperadorListas As String
    NumSistemaMedida As String
    MonSimboloMoneda As String
    MonFormatoMonedaPositivo As String
    MonFormatoMonedaNegativo As String
    MonSimboloDecimal As String
    MonDigitosDecimales As String
    MonSimboloSeparacionMiles As String
    MonDigitosGrupo As String
    FormatoHora As String
    SeperadoHora As String
    SimboloAM As String
    SimboloPM As String
    FormatoFechaCorta As String
    SeparadorFecha As String
    FormatoFechaLarga As String
End Type
Function CargaConfiguracionRegional(ByRef ConfiguracionRegional As apCFGRegionalSTR) As Boolean
    Dim hKey As Long
    Dim subkey As String
    Dim retval As Long
    
    CargaConfiguracionRegional = False
    Err.Clear
    On Error GoTo Error

    subkey = "Control Panel\International"


    retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_READ, hKey)
    If retval <> 0 Then
        Exit Function
    End If

    With ConfiguracionRegional
        '// Numeros
        .NumSimboloDecimal = CargaConfiguracionRegionalDetalle("sDecimal", hKey)
        .NumDigitosDecimales = CargaConfiguracionRegionalDetalle("iDigits", hKey)
        .NumSimboloSeparacionMiles = CargaConfiguracionRegionalDetalle("sThousand", hKey)
        .NumDigitosGrupo = CargaConfiguracionRegionalDetalle("sGrouping", hKey)
        .NumSimboloSignoNegativo = CargaConfiguracionRegionalDetalle("sNegativeSign", hKey)
        .NumFormatoNumeroNegativo = CargaConfiguracionRegionalDetalle("iNegNumber", hKey)
        .NumMostrarCerosIzquierda = CargaConfiguracionRegionalDetalle("iLZero", hKey)
        .NumSeperadorListas = CargaConfiguracionRegionalDetalle("sList", hKey)
        .NumSistemaMedida = CargaConfiguracionRegionalDetalle("iMeasure", hKey)
        
        '// Moneda
        .MonSimboloMoneda = CargaConfiguracionRegionalDetalle("sCurrency", hKey)
        .MonFormatoMonedaPositivo = CargaConfiguracionRegionalDetalle("iCurrency", hKey)
        .MonFormatoMonedaNegativo = CargaConfiguracionRegionalDetalle("iNegCurr", hKey)
        .MonSimboloDecimal = CargaConfiguracionRegionalDetalle("sMonDecimalSep", hKey)
        .MonDigitosDecimales = CargaConfiguracionRegionalDetalle("iCurrDigits", hKey)
        .MonSimboloSeparacionMiles = CargaConfiguracionRegionalDetalle("sMonThousandSep", hKey)
        .MonDigitosGrupo = CargaConfiguracionRegionalDetalle("sMonGrouping", hKey)
        
        '//Hora
        .FormatoHora = CargaConfiguracionRegionalDetalle("sTimeFormat", hKey)
        .SeperadoHora = CargaConfiguracionRegionalDetalle("sTime", hKey)
        .SimboloAM = CargaConfiguracionRegionalDetalle("s1159", hKey)
        .SimboloPM = CargaConfiguracionRegionalDetalle("s2359", hKey)
        
        '//Fecha
        .FormatoFechaCorta = CargaConfiguracionRegionalDetalle("sShortDate", hKey)
        .SeparadorFecha = CargaConfiguracionRegionalDetalle("sDate", hKey)
        .FormatoFechaLarga = CargaConfiguracionRegionalDetalle("sLongDate", hKey)
    End With

    retval = RegCloseKey(hKey)
    CargaConfiguracionRegional = True
    Exit Function
Error:
    Err.Clear
    CargaConfiguracionRegional = False
End Function
Private Function CargaConfiguracionRegionalDetalle(strVariable As String, hKey As Long) As String
    Dim stringbuffer As String
    Dim datatype As Long
    Dim slength As Long
    Dim retval As Long

    CargaConfiguracionRegionalDetalle = ""
    stringbuffer = Space(255)
    slength = 255
    retval = RegQueryValueEx(hKey, strVariable, 0, datatype, ByVal stringbuffer, slength)
    If retval = 0 Then
        If datatype = REG_SZ Then
            stringbuffer = Left(stringbuffer, slength - 1)
            CargaConfiguracionRegionalDetalle = stringbuffer
        End If
    End If
End Function
Public Function NormalizaConfiguracionRegional(ByRef ConfiguracionRegionalObligatoria As apCFGRegionalSTR, ByRef blnCambios As Boolean) As Boolean
    Dim hKey As Long
    Dim subkey As String
    Dim retval As Long
    Dim secattr As SECURITY_ATTRIBUTES
    Dim neworused As Long
    Dim strMensaje As String
    Dim ConfiguracionRegional As apCFGRegionalSTR

    NormalizaConfiguracionRegional = False
    blnCambios = False
    Err.Clear
    On Error GoTo Error

    If Not CargaConfiguracionRegional(ConfiguracionRegional) Then
        NormalizaConfiguracionRegional = True
        Exit Function
    End If


    '//Analisa los cambios a realizar
    '//Numeros
    If ConfiguracionRegionalObligatoria.NumSimboloDecimal <> "" Then If ConfiguracionRegional.NumSimboloDecimal <> ConfiguracionRegionalObligatoria.NumSimboloDecimal Then strMensaje = strMensaje & "Número - Símbolo Decimal (" & ConfiguracionRegional.NumSimboloDecimal & ") sebe ser: " & ConfiguracionRegionalObligatoria.NumSimboloDecimal & Chr(13)
    If ConfiguracionRegionalObligatoria.NumDigitosDecimales <> "" Then If ConfiguracionRegional.NumDigitosDecimales <> ConfiguracionRegionalObligatoria.NumDigitosDecimales Then strMensaje = strMensaje & "Número - Número de Digitos Decimales (" & ConfiguracionRegional.NumDigitosDecimales & ") debe ser: " & ConfiguracionRegionalObligatoria.NumDigitosDecimales & Chr(13)
    If ConfiguracionRegionalObligatoria.NumSimboloSeparacionMiles <> "" Then If ConfiguracionRegional.NumSimboloSeparacionMiles <> ConfiguracionRegionalObligatoria.NumSimboloSeparacionMiles Then strMensaje = strMensaje & "Número - Símbolo de separación de miles (" & ConfiguracionRegional.NumSimboloSeparacionMiles & ") debe ser: " & ConfiguracionRegionalObligatoria.NumSimboloSeparacionMiles & Chr(13)
    If ConfiguracionRegionalObligatoria.NumDigitosGrupo <> "" Then If ConfiguracionRegional.NumDigitosGrupo <> ConfiguracionRegionalObligatoria.NumDigitosGrupo Then strMensaje = strMensaje & "Número - Número de dígitos en grupo (" & ConfiguracionRegional.NumDigitosGrupo & ") debe ser: " & ConfiguracionRegionalObligatoria.NumDigitosGrupo & Chr(13)
    If ConfiguracionRegionalObligatoria.NumSimboloSignoNegativo <> "" Then If ConfiguracionRegional.NumSimboloSignoNegativo <> ConfiguracionRegionalObligatoria.NumSimboloSignoNegativo Then strMensaje = strMensaje & "Número - Símbolo de signo negativo (" & ConfiguracionRegional.NumSimboloSignoNegativo & ") debe ser: " & ConfiguracionRegionalObligatoria.NumSimboloSignoNegativo & Chr(13)
    If ConfiguracionRegionalObligatoria.NumFormatoNumeroNegativo <> "" Then If ConfiguracionRegional.NumFormatoNumeroNegativo <> ConfiguracionRegionalObligatoria.NumFormatoNumeroNegativo Then strMensaje = strMensaje & "Número - Formato de número negativo (" & ConfiguracionRegional.NumFormatoNumeroNegativo & ") debe ser: " & ConfiguracionRegionalObligatoria.NumFormatoNumeroNegativo & Chr(13)
    If ConfiguracionRegionalObligatoria.NumMostrarCerosIzquierda <> "" Then If ConfiguracionRegional.NumMostrarCerosIzquierda <> ConfiguracionRegionalObligatoria.NumMostrarCerosIzquierda Then strMensaje = strMensaje & "Número - Mostrar ceros a la izquierda (" & ConfiguracionRegional.NumMostrarCerosIzquierda & ") debe ser: " & ConfiguracionRegionalObligatoria.NumMostrarCerosIzquierda & Chr(13)
    If ConfiguracionRegionalObligatoria.NumSeperadorListas <> "" Then If ConfiguracionRegional.NumSeperadorListas <> ConfiguracionRegionalObligatoria.NumSeperadorListas Then strMensaje = strMensaje & "Número - Separador de listas (" & ConfiguracionRegional.NumSeperadorListas & ") debe ser: " & ConfiguracionRegionalObligatoria.NumSeperadorListas & Chr(13)
    If ConfiguracionRegionalObligatoria.NumSistemaMedida <> "" Then If ConfiguracionRegional.NumSistemaMedida <> ConfiguracionRegionalObligatoria.NumSistemaMedida Then strMensaje = strMensaje & "Número - Sistema de medida (" & ConfiguracionRegional.NumSistemaMedida & ") debe ser: " & ConfiguracionRegionalObligatoria.NumSistemaMedida & Chr(13)
    '//Moneda
    If ConfiguracionRegionalObligatoria.MonSimboloMoneda <> "" Then If ConfiguracionRegional.MonSimboloMoneda <> ConfiguracionRegionalObligatoria.MonSimboloMoneda Then strMensaje = strMensaje & "Moneda - Símbolo de moneda (" & ConfiguracionRegional.MonSimboloMoneda & ") debe ser: " & ConfiguracionRegionalObligatoria.MonSimboloMoneda & Chr(13)
    If ConfiguracionRegionalObligatoria.MonFormatoMonedaPositivo <> "" Then If ConfiguracionRegional.MonFormatoMonedaPositivo <> ConfiguracionRegionalObligatoria.MonFormatoMonedaPositivo Then strMensaje = strMensaje & "Moneda - Formato de moneda positivo (" & ConfiguracionRegional.MonFormatoMonedaPositivo & ") debe ser: " & ConfiguracionRegionalObligatoria.MonFormatoMonedaPositivo & Chr(13)
    If ConfiguracionRegionalObligatoria.MonFormatoMonedaNegativo <> "" Then If ConfiguracionRegional.MonFormatoMonedaNegativo <> ConfiguracionRegionalObligatoria.MonFormatoMonedaNegativo Then strMensaje = strMensaje & "Moneda - Formato de moneda negativo (" & ConfiguracionRegional.MonFormatoMonedaNegativo & ") debe ser: " & ConfiguracionRegionalObligatoria.MonFormatoMonedaNegativo & Chr(13)
    If ConfiguracionRegionalObligatoria.MonSimboloDecimal <> "" Then If ConfiguracionRegional.MonSimboloDecimal <> ConfiguracionRegionalObligatoria.MonSimboloDecimal Then strMensaje = strMensaje & "Moneda - Símbolo decimal (" & ConfiguracionRegional.MonSimboloDecimal & ") debe ser: " & ConfiguracionRegionalObligatoria.MonSimboloDecimal & Chr(13)
    If ConfiguracionRegionalObligatoria.MonDigitosDecimales <> "" Then If ConfiguracionRegional.MonDigitosDecimales <> ConfiguracionRegionalObligatoria.MonDigitosDecimales Then strMensaje = strMensaje & "Moneda - Número de dígitos decimales(" & ConfiguracionRegional.MonDigitosDecimales & ") debe ser: " & ConfiguracionRegionalObligatoria.MonDigitosDecimales & Chr(13)
    If ConfiguracionRegionalObligatoria.MonSimboloSeparacionMiles <> "" Then If ConfiguracionRegional.MonSimboloSeparacionMiles <> ConfiguracionRegionalObligatoria.MonSimboloSeparacionMiles Then strMensaje = strMensaje & "Moneda - Símbolo de separación de miles (" & ConfiguracionRegional.MonSimboloSeparacionMiles & ") debe ser: " & ConfiguracionRegionalObligatoria.MonSimboloSeparacionMiles & Chr(13)
    If ConfiguracionRegionalObligatoria.MonDigitosGrupo <> "" Then If ConfiguracionRegional.MonDigitosGrupo <> ConfiguracionRegionalObligatoria.MonDigitosGrupo Then strMensaje = strMensaje & "Moneda - Número de digitos en grupo (" & ConfiguracionRegional.MonDigitosGrupo & ") debe ser: " & ConfiguracionRegionalObligatoria.MonDigitosGrupo & Chr(13)
    '//Hora
    If ConfiguracionRegionalObligatoria.FormatoHora <> "" Then If ConfiguracionRegional.FormatoHora <> ConfiguracionRegionalObligatoria.FormatoHora Then strMensaje = strMensaje & "Hora - Formato de hora (" & ConfiguracionRegional.FormatoHora & ") debe ser: " & ConfiguracionRegionalObligatoria.FormatoHora & Chr(13)
    If ConfiguracionRegionalObligatoria.SeperadoHora <> "" Then If ConfiguracionRegional.SeperadoHora <> ConfiguracionRegionalObligatoria.SeperadoHora Then strMensaje = strMensaje & "Hora - Separador de hora (" & ConfiguracionRegional.SeperadoHora & ") debe ser: " & ConfiguracionRegionalObligatoria.SeperadoHora & Chr(13)
    If ConfiguracionRegionalObligatoria.SimboloAM <> "" Then If ConfiguracionRegional.SimboloAM <> ConfiguracionRegionalObligatoria.SimboloAM Then strMensaje = strMensaje & "Hora - Símbolo a.m. (" & ConfiguracionRegional.SimboloAM & ") debe ser: " & ConfiguracionRegionalObligatoria.SimboloAM & Chr(13)
    If ConfiguracionRegionalObligatoria.SimboloPM <> "" Then If ConfiguracionRegional.SimboloPM <> ConfiguracionRegionalObligatoria.SimboloPM Then strMensaje = strMensaje & "Hora - Símbolo p.m. (" & ConfiguracionRegional.SimboloPM & ") debe ser: " & ConfiguracionRegionalObligatoria.SimboloPM & Chr(13)
    '//Fecha
    If ConfiguracionRegionalObligatoria.FormatoFechaCorta <> "" Then If ConfiguracionRegional.FormatoFechaCorta <> ConfiguracionRegionalObligatoria.FormatoFechaCorta Then strMensaje = strMensaje & "Fecha - Formato de fecha corta (" & ConfiguracionRegional.FormatoFechaCorta & ") debe ser: " & ConfiguracionRegionalObligatoria.FormatoFechaCorta & Chr(13)
    If ConfiguracionRegionalObligatoria.SeparadorFecha <> "" Then If ConfiguracionRegional.SeparadorFecha <> ConfiguracionRegionalObligatoria.SeparadorFecha Then strMensaje = strMensaje & "Fecha - Separador de fecha (" & ConfiguracionRegional.SeparadorFecha & ") debe ser: " & ConfiguracionRegionalObligatoria.SeparadorFecha & Chr(13)
    If ConfiguracionRegionalObligatoria.FormatoFechaLarga <> "" Then If ConfiguracionRegional.FormatoFechaLarga <> ConfiguracionRegionalObligatoria.FormatoFechaLarga Then strMensaje = strMensaje & "Fecha - Formato de fecha larga (" & ConfiguracionRegional.FormatoFechaLarga & ") debe ser: " & ConfiguracionRegionalObligatoria.FormatoFechaLarga & Chr(13)
    
    If strMensaje <> "" Then
        If MsgBox("Existen diferencias en la configuración regional en los siguientes parámetros:" & Chr(13) & strMensaje & Chr(13) & "¿ Desea normalizar estos parámetros ?", vbQuestion + vbYesNo + vbDefaultButton2, "Normalización configuración Regional") = vbNo Then
            Exit Function
        End If
    Else
        NormalizaConfiguracionRegional = True
        Exit Function
    End If

    subkey = "Control Panel\International"
    secattr.nLength = Len(secattr)
    secattr.lpSecurityDescriptor = 0
    secattr.bInheritHandle = True


    retval = RegCreateKeyEx(HKEY_CURRENT_USER, subkey, 0, "", 0, KEY_WRITE, secattr, hKey, neworused)
    If retval <> 0 Then
        Exit Function
    End If

    '//Realiza los cambios...
    '//Numeros
    If ConfiguracionRegionalObligatoria.NumSimboloDecimal <> "" Then If ConfiguracionRegional.NumSimboloDecimal <> ConfiguracionRegionalObligatoria.NumSimboloDecimal Then NormalizaConfiguracionRegionalDetalle "sDecimal", ConfiguracionRegionalObligatoria.NumSimboloDecimal, hKey
    If ConfiguracionRegionalObligatoria.NumDigitosDecimales <> "" Then If ConfiguracionRegional.NumDigitosDecimales <> ConfiguracionRegionalObligatoria.NumDigitosDecimales Then NormalizaConfiguracionRegionalDetalle "iDigits", ConfiguracionRegionalObligatoria.NumDigitosDecimales, hKey
    If ConfiguracionRegionalObligatoria.NumSimboloSeparacionMiles <> "" Then If ConfiguracionRegional.NumSimboloSeparacionMiles <> ConfiguracionRegionalObligatoria.NumSimboloSeparacionMiles Then NormalizaConfiguracionRegionalDetalle "sThousand", ConfiguracionRegionalObligatoria.NumSimboloSeparacionMiles, hKey
    If ConfiguracionRegionalObligatoria.NumDigitosGrupo <> "" Then If ConfiguracionRegional.NumDigitosGrupo <> ConfiguracionRegionalObligatoria.NumDigitosGrupo Then NormalizaConfiguracionRegionalDetalle "sGrouping", ConfiguracionRegionalObligatoria.NumDigitosGrupo, hKey
    If ConfiguracionRegionalObligatoria.NumSimboloSignoNegativo <> "" Then If ConfiguracionRegional.NumSimboloSignoNegativo <> ConfiguracionRegionalObligatoria.NumSimboloSignoNegativo Then NormalizaConfiguracionRegionalDetalle "sNegativeSign", ConfiguracionRegionalObligatoria.NumSimboloSignoNegativo, hKey
    If ConfiguracionRegionalObligatoria.NumFormatoNumeroNegativo <> "" Then If ConfiguracionRegional.NumFormatoNumeroNegativo <> ConfiguracionRegionalObligatoria.NumFormatoNumeroNegativo Then NormalizaConfiguracionRegionalDetalle "iNegNumber", ConfiguracionRegionalObligatoria.NumFormatoNumeroNegativo, hKey
    If ConfiguracionRegionalObligatoria.NumMostrarCerosIzquierda <> "" Then If ConfiguracionRegional.NumMostrarCerosIzquierda <> ConfiguracionRegionalObligatoria.NumMostrarCerosIzquierda Then NormalizaConfiguracionRegionalDetalle "iLZero", ConfiguracionRegionalObligatoria.NumMostrarCerosIzquierda, hKey
    If ConfiguracionRegionalObligatoria.NumSeperadorListas <> "" Then If ConfiguracionRegional.NumSeperadorListas <> ConfiguracionRegionalObligatoria.NumSeperadorListas Then NormalizaConfiguracionRegionalDetalle "sList", ConfiguracionRegionalObligatoria.NumSeperadorListas, hKey
    If ConfiguracionRegionalObligatoria.NumSistemaMedida <> "" Then If ConfiguracionRegional.NumSistemaMedida <> ConfiguracionRegionalObligatoria.NumSistemaMedida Then NormalizaConfiguracionRegionalDetalle "iMeasure", ConfiguracionRegionalObligatoria.NumSistemaMedida, hKey
    '//Moneda
    If ConfiguracionRegionalObligatoria.MonSimboloMoneda <> "" Then If ConfiguracionRegional.MonSimboloMoneda <> ConfiguracionRegionalObligatoria.MonSimboloMoneda Then NormalizaConfiguracionRegionalDetalle "sCurrency", ConfiguracionRegionalObligatoria.MonSimboloMoneda, hKey
    If ConfiguracionRegionalObligatoria.MonFormatoMonedaPositivo <> "" Then If ConfiguracionRegional.MonFormatoMonedaPositivo <> ConfiguracionRegionalObligatoria.MonFormatoMonedaPositivo Then NormalizaConfiguracionRegionalDetalle "iCurrency", ConfiguracionRegionalObligatoria.MonFormatoMonedaPositivo, hKey
    If ConfiguracionRegionalObligatoria.MonFormatoMonedaNegativo <> "" Then If ConfiguracionRegional.MonFormatoMonedaNegativo <> ConfiguracionRegionalObligatoria.MonFormatoMonedaNegativo Then NormalizaConfiguracionRegionalDetalle "iNegCurr", ConfiguracionRegionalObligatoria.MonFormatoMonedaNegativo, hKey
    If ConfiguracionRegionalObligatoria.MonSimboloDecimal <> "" Then If ConfiguracionRegional.MonSimboloDecimal <> ConfiguracionRegionalObligatoria.MonSimboloDecimal Then NormalizaConfiguracionRegionalDetalle "sMonDecimalSep", ConfiguracionRegionalObligatoria.MonSimboloDecimal, hKey
    If ConfiguracionRegionalObligatoria.MonDigitosDecimales <> "" Then If ConfiguracionRegional.MonDigitosDecimales <> ConfiguracionRegionalObligatoria.MonDigitosDecimales Then NormalizaConfiguracionRegionalDetalle "iCurrDigits", ConfiguracionRegionalObligatoria.MonDigitosDecimales, hKey
    If ConfiguracionRegionalObligatoria.MonSimboloSeparacionMiles <> "" Then If ConfiguracionRegional.MonSimboloSeparacionMiles <> ConfiguracionRegionalObligatoria.MonSimboloSeparacionMiles Then NormalizaConfiguracionRegionalDetalle "sMonThousandSep", ConfiguracionRegionalObligatoria.MonSimboloSeparacionMiles, hKey
    If ConfiguracionRegionalObligatoria.MonDigitosGrupo <> "" Then If ConfiguracionRegional.MonDigitosGrupo <> ConfiguracionRegionalObligatoria.MonDigitosGrupo Then NormalizaConfiguracionRegionalDetalle "sMonGrouping", ConfiguracionRegionalObligatoria.MonDigitosGrupo, hKey
    '//Hora
    If ConfiguracionRegionalObligatoria.FormatoHora <> "" Then If ConfiguracionRegional.FormatoHora <> ConfiguracionRegionalObligatoria.FormatoHora Then NormalizaConfiguracionRegionalDetalle "sTimeFormat", ConfiguracionRegionalObligatoria.FormatoHora, hKey
    If ConfiguracionRegionalObligatoria.SeperadoHora <> "" Then If ConfiguracionRegional.SeperadoHora <> ConfiguracionRegionalObligatoria.SeperadoHora Then NormalizaConfiguracionRegionalDetalle "sTime", ConfiguracionRegionalObligatoria.SeperadoHora, hKey
    If ConfiguracionRegionalObligatoria.SimboloAM <> "" Then If ConfiguracionRegional.SimboloAM <> ConfiguracionRegionalObligatoria.SimboloAM Then NormalizaConfiguracionRegionalDetalle "s1159", ConfiguracionRegionalObligatoria.SimboloAM, hKey
    If ConfiguracionRegionalObligatoria.SimboloPM <> "" Then If ConfiguracionRegional.SimboloPM <> ConfiguracionRegionalObligatoria.SimboloPM Then NormalizaConfiguracionRegionalDetalle "s2359", ConfiguracionRegionalObligatoria.SimboloPM, hKey
    '//Fecha
    If ConfiguracionRegionalObligatoria.FormatoFechaCorta <> "" Then If ConfiguracionRegional.FormatoFechaCorta <> ConfiguracionRegionalObligatoria.FormatoFechaCorta Then NormalizaConfiguracionRegionalDetalle "sShortDate", ConfiguracionRegionalObligatoria.FormatoFechaCorta, hKey
    If ConfiguracionRegionalObligatoria.SeparadorFecha <> "" Then If ConfiguracionRegional.SeparadorFecha <> ConfiguracionRegionalObligatoria.SeparadorFecha Then NormalizaConfiguracionRegionalDetalle "sDate", ConfiguracionRegionalObligatoria.SeparadorFecha, hKey
    If ConfiguracionRegionalObligatoria.FormatoFechaLarga <> "" Then If ConfiguracionRegional.FormatoFechaLarga <> ConfiguracionRegionalObligatoria.FormatoFechaLarga Then NormalizaConfiguracionRegionalDetalle "sLongDate", ConfiguracionRegionalObligatoria.FormatoFechaLarga, hKey

    retval = RegCloseKey(hKey)
    blnCambios = True

    MsgBox "Cierre todas sus aplicaciones y presione Aceptar para re-iniciar la sesión...", vbInformation, "Advertencia"
    retval = ExitWindowsEx(EWX_FORCE Or EWX_LOGOFF, 0)

    NormalizaConfiguracionRegional = True
    Exit Function

Error:
    Err.Clear
    NormalizaConfiguracionRegional = True
End Function
Private Function NormalizaConfiguracionRegionalDetalle(strVariable As String, strValor As String, hKey As Long) As Boolean
    Dim stringbuffer As String
    Dim retval As Long

    NormalizaConfiguracionRegionalDetalle = False
    stringbuffer = strValor & vbNullChar
    retval = RegSetValueEx(hKey, strVariable, 0, REG_SZ, ByVal stringbuffer, Len(stringbuffer))
    If retval = 0 Then
        NormalizaConfiguracionRegionalDetalle = True
    End If
End Function
