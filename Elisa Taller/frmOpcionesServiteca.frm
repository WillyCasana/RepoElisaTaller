VERSION 5.00
Begin VB.Form frmOpcionesServiteca 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Opciones para Serviteca"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmOpcionesServiteca.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4455
      Begin VB.CommandButton cmdAceptar 
         Appearance      =   0  'Flat
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   200
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Appearance      =   0  'Flat
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   3120
         TabIndex        =   5
         Top             =   200
         Width           =   1215
      End
   End
   Begin VB.TextBox txtDiasLLamado 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   555
      Width           =   2415
   End
   Begin VB.TextBox txtNumSiguienteOT 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Días Próximo LLamado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Número Siguiente O.T."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   2055
   End
End
Attribute VB_Name = "frmOpcionesServiteca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
GuardeDatos
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Public Sub CargaDatos()
Me.txtNumSiguienteOT.Text = TraeNumOT

Retorno = Space$(128)
tam = Len(Retorno)
Valido = GetPrivateProfileString("TLLR", "PROXIMOLLAMADO", "", Retorno, tam, "AutoPro.ini")
gstrDiasProximoLLamado = Trim$(Left$(Retorno, Valido))

Me.txtDiasLLamado.Text = gstrDiasProximoLLamado

End Sub

Public Sub GuardeDatos()
Dim tablaCorrelativo As New ADODB.Recordset
Dim sql As String
Dim ldblNumero As Double

ldblNumero = CDbl(Me.txtNumSiguienteOT.Text) - 1

If ExisteNumeroOT(ldblNumero + 1) = False Then
    sql = ""
    sql = "INSERT INTO Srvt_Correlativo_OT"
    sql = sql & " (Id_Empresa, Id_Sucursal, Ultimo_Numero)"
    sql = sql & " Values ("
    sql = sql & "'" & gstrIdEmpresa & "', "
    sql = sql & "'" & gstrIdSucursal & "', "
    sql = sql & CDbl(ldblNumero) & ") "
    If Conexion.SendHost(sql, tablaCorrelativo, adOpenKeyset, adLockOptimistic, gcTiempoEspera) <> apOk Then
        MsgBox "No se ha guardado el número de documento."
        Exit Sub
    End If
    Conexion.CloseHost tablaCorrelativo
Else
    MsgBox "El número de Orden de Trabajo que intenta agregar ya existe." & Chr(13) & "Intente nuevamente.", vbExclamation, "ServiPro"
    Exit Sub
End If

Valido = WritePrivateProfileString("TLLR", "PROXIMOLLAMADO", Me.txtDiasLLamado.Text, "AutoPro.ini")

Unload Me

End Sub

Function ExisteNumeroOT(NumeroOT As Double) As Boolean
Dim Tabla As New ADODB.Recordset
Dim sql As String

ExisteNumeroOT = False

sql = ""
sql = "SELECT * FROM Srvt_OT WHERE Id_OT=" & NumeroOT
If Conexion.SendHost(sql, Tabla, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Tabla.EOF = False And Tabla.BOF = False Then
        ExisteNumeroOT = True
    Else
        ExisteNumeroOT = False
    End If
End If
Conexion.CloseHost Tabla

End Function

Private Sub Form_Load()
CargaDatos
End Sub

Private Sub txtNumSiguienteOT_LostFocus()
If Me.txtNumSiguienteOT.Text = "" Then Me.txtNumSiguienteOT.Text = "0"
If ExisteNumeroOT(Me.txtNumSiguienteOT.Text) = True Then
    MsgBox "El número de Orden de Trabajo que intenta agregar ya existe o no es válido." & Chr(13) & "Intente nuevamente con otro Numero de Orden de Trabajo.", vbExclamation, "ServiPro"
    Me.txtNumSiguienteOT.Text = TraeNumOT
End If
End Sub
