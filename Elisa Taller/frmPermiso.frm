VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPermiso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificación Liquidador"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmPermiso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   135
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   675
      Width           =   1920
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   3690
      TabIndex        =   2
      Top             =   675
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   360
      Left            =   3690
      TabIndex        =   1
      Top             =   135
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo dtcLiquidador 
      Bindings        =   "frmPermiso.frx":179A
      Height          =   315
      Left            =   135
      TabIndex        =   3
      Top             =   135
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ListField       =   "Nombre"
      BoundColumn     =   "Codigo"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc datLiquidador 
      Height          =   330
      Left            =   1845
      Top             =   135
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPermiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()


gstrVerificacion = IIf(Text1 = "", "0", Text1)


gstrVerificaMecanico = Me.dtcLiquidador.BoundText
If Act = 1 Then
    gUsr_Activacion = Me.dtcLiquidador.Text
End If
Unload Me
End Sub

Private Sub cmdCancelar_Click()
'Me.Tag = ""
gstrVerificacion = "0"
gflag = False
Unload Me
End Sub

Private Sub Form_Activate()
If Act = 0 Then
    FillLiquidador dtcLiquidador, datLiquidador
 ElseIf Act = 1 Then
    FillActivador dtcLiquidador, datLiquidador
End If
End Sub


Private Sub Form_Load()
gstrVerificacion = "0"
gstrVerificaMecanico = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, Text1, strDot)
End Sub
