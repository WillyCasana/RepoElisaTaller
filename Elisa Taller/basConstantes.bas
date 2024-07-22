Attribute VB_Name = "basConstantes"
Option Explicit

Public Const SWP_SHOWWINDOW = &H40 'SWP_SHOWWINDOW +SWP_NOMOVE +SWP_NOSIZE
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_HIDEWINDOW = &H80
Public Const Ext_Width As Long = 6450
Public Const Ext_Height As Long = 3810
Public Const Vrp_Width As Long = 2715
Public Const Vrp_Height As Long = 2985
Public Const Dhr_Width As Long = 6270
Public Const Dhr_Height As Long = 3270

'//api ejecutar archivos...
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9

Public Const PROCESS_QUERY_INFORMATION = &H400

Public Const STILL_ACTIVE = &H103

Type EstadoDeLista
    LLENA As Boolean
    VACIA As Boolean
    MEDIA As Boolean
End Type
Global EstadoLista As EstadoDeLista
