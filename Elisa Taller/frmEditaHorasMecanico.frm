VERSION 5.00
Begin VB.Form frmEditaHorasMecanico 
   Caption         =   "Edición Horas Mecanicos"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   Icon            =   "frmEditaHorasMecanico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   4560
      TabIndex        =   11
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   3600
      TabIndex        =   10
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtHorasReales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2040
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtHorasCompradas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2040
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblMecanico 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2040
      TabIndex        =   7
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label lblAño 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2040
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblMes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2040
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Horas Reales         :"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Horas Compradas  :"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Mecánico              :"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Año                       :"
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
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Mes                       :"
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmEditaHorasMecanico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnSW As Boolean

Private Sub cmdAceptar_Click()
DescargaDatos
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If mblnSW Then
    CargaDatos
    mblnSW = False
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{tab}"
    End Select
End Sub

Private Sub Form_Load()
mblnSW = True
End Sub
Sub CargaDatos()
With frmMantenedorHorasMecanicos
    lblMes = .dtcMeses.Text
    lblAño = .txtAño
    lblMecanico = .lvwConceptos.SelectedItem.SubItems(1)
    txtHorasCompradas = SacarFormatoValor(.lvwConceptos.SelectedItem.SubItems(2), "")
    txtHorasReales = SacarFormatoValor(.lvwConceptos.SelectedItem.SubItems(3), "")
End With
End Sub
Sub DescargaDatos()
With frmMantenedorHorasMecanicos.lvwConceptos
    .SelectedItem.SubItems(2) = FormatoValor(txtHorasCompradas, "", 2)
    .SelectedItem.SubItems(3) = FormatoValor(txtHorasReales, "", 2)
End With
End Sub


Private Sub txtHorasCompradas_GotFocus()
MarcaTexto txtHorasCompradas
End Sub

Private Sub txtHorasCompradas_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtHorasCompradas, strDot)
End Sub

Private Sub txtHorasReales_GotFocus()
MarcaTexto txtHorasReales
End Sub

Private Sub txtHorasReales_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtHorasReales, strDot)
End Sub
