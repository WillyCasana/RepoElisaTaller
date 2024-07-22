VERSION 5.00
Begin VB.Form frmConfImprimirInventarioVehiculo 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Imprime inventario de vehículo"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   ClipControls    =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnNO 
      Caption         =   "NO"
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton btnSI 
      Caption         =   "SI"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   6375
   End
   Begin VB.Label lblmarcas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imprimirá el inventario de vehículo. Desea continuar?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
End
Attribute VB_Name = "frmConfImprimirInventarioVehiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    Screen.MousePointer = vbDefault
End Sub

Private Sub btnSI_Click()

    frmRecepcion.ConfirmarImprimirInventarioVehiculo = "S"
    
    Unload Me
End Sub

Private Sub btnNO_Click()
    frmRecepcion.ConfirmarImprimirInventarioVehiculo = "N"
    Unload Me
End Sub

