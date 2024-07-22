VERSION 5.00
Begin VB.Form frmEditaPreciosporMarca 
   Caption         =   "Edición Precios por Marca"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCostoMOGtia 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtVentaMOGtia 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2160
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   4200
      TabIndex        =   6
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   3120
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txtVentaMO 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtCostoMO 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Costo M.Obra Garantía :"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Venta M.Obra Garantía :"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblMarca 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "Venta Mano de Obra :"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Costo Mano de Obra :"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Marca :"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmEditaPreciosporMarca"
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
With frmPreciosporMarca
    Me.lblMarca = .lvDetalle.SelectedItem.SubItems(1)
    Me.txtCostoMO = SacarFormatoValor(.lvDetalle.SelectedItem.SubItems(2), "")
    Me.txtVentaMO = SacarFormatoValor(.lvDetalle.SelectedItem.SubItems(3), "")
    Me.txtCostoMOGtia = SacarFormatoValor(.lvDetalle.SelectedItem.SubItems(4), "")
    Me.txtVentaMOGtia = SacarFormatoValor(.lvDetalle.SelectedItem.SubItems(5), "")
End With
End Sub
Sub DescargaDatos()
With frmPreciosporMarca.lvDetalle
    .SelectedItem.SubItems(2) = FormatoValor(Me.txtCostoMO, gstrMonedaLocal, gintDecimalesMoneda)
    .SelectedItem.SubItems(3) = FormatoValor(Me.txtVentaMO, gstrMonedaLocal, gintDecimalesMoneda)
    .SelectedItem.SubItems(4) = FormatoValor(Me.txtCostoMOGtia, gstrMonedaLocal, gintDecimalesMoneda)
    .SelectedItem.SubItems(5) = FormatoValor(Me.txtVentaMOGtia, gstrMonedaLocal, gintDecimalesMoneda)
End With
End Sub

Private Sub txtCostoMO_GotFocus()
MarcaTexto txtCostoMO
End Sub

Private Sub txtCostoMO_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtCostoMO, strDot)
End Sub

Private Sub txtCostoMOGtia_GotFocus()
MarcaTexto txtCostoMOGtia
End Sub

Private Sub txtCostoMOGtia_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtCostoMOGtia, strDot)
End Sub

Private Sub txtVentaMO_GotFocus()
MarcaTexto txtVentaMO
End Sub

Private Sub txtVentaMO_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtVentaMO, strDot)
End Sub

Private Sub txtVentaMOGtia_GotFocus()
MarcaTexto txtVentaMOGtia
End Sub

Private Sub txtVentaMOGtia_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtVentaMOGtia, strDot)
End Sub
