VERSION 5.00
Begin VB.Form frmPermisoDiasHabiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificación Clave"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmPermisoDiasHabiles.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDiasHabiles 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   3720
      TabIndex        =   3
      Top             =   675
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   3690
      TabIndex        =   2
      Top             =   135
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "N° Dias Habiles"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "PassWord"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmPermisoDiasHabiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
gstrVerificacion = IIf(txtPassword = "", "0", txtPassword)
gintDiasHabiles = IIf(txtDiasHabiles = "", 0, txtDiasHabiles)
Unload Me
End Sub

Private Sub cmdCancelar_Click()
'Me.Tag = ""
gstrVerificacion = "0"
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{tab}"
        Case vbKeyEscape
            KeyAscii = 0
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
If gblnDescuentoRepuesto = True Then
    Label2.Visible = False
    txtDiasHabiles.Visible = False
End If
End Sub

Private Sub txtDiasHabiles_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtDiasHabiles, strDot)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtPassword, strDot)
End Sub
