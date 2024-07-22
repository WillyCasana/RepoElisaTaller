VERSION 5.00
Begin VB.Form frmPatVinToRut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patente to Rut"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmPatVinToRut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.OptionButton optVin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "V.I.N."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   375
      Width           =   1110
   End
   Begin VB.OptionButton optPatente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Caption         =   "Placa"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   105
      TabIndex        =   3
      Top             =   90
      Width           =   1110
   End
   Begin VB.TextBox txtObjeto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   2
      Top             =   60
      Width           =   1830
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Cancel And Exit"
      Height          =   495
      Left            =   2490
      TabIndex        =   1
      Top             =   1185
      Width           =   2070
   End
   Begin VB.CommandButton cmdExe 
      Appearance      =   0  'Flat
      Caption         =   "Execute Convert"
      Height          =   495
      Left            =   135
      TabIndex        =   0
      Top             =   1200
      Width           =   2070
   End
   Begin VB.Label lblObjeto 
      Height          =   885
      Left            =   915
      TabIndex        =   5
      Top             =   1785
      Width           =   2940
   End
End
Attribute VB_Name = "frmPatVinToRut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExe_Click()
Dim strDig As String
Dim strRut As String

'If optPatente.Value = True Then
'    Call CheckPatente(txtObjeto, strDig, strRut)
'    lblObjeto.Caption = strRut
'ElseIf optVin.Value = True Then
'    lblObjeto.Caption = VintoRut(txtObjeto)
'End If

lblObjeto.Caption = strRut

End Sub

