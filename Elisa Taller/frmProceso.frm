VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesando Datos..."
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.Animation anProceso 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1296
      _Version        =   393216
      AutoPlay        =   -1  'True
      Center          =   -1  'True
      FullWidth       =   297
      FullHeight      =   49
   End
   Begin MSComctlLib.ProgressBar pbProceso 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Label lblProceso 
      Alignment       =   2  'Center
      Caption         =   "0%"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4335
   End
End
Attribute VB_Name = "frmProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    If Dir(App.Path & "\proceso.avi") <> "" Then
        anProceso.Open App.Path & "\proceso.avi"
    End If
    
    FijarFormulario Me
End Sub
