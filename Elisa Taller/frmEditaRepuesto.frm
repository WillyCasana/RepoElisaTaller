VERSION 5.00
Begin VB.Form frmEditaRepuesto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Repuestos Solicitados"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   5055
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   3720
         TabIndex        =   16
         Top             =   200
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   2400
         TabIndex        =   15
         Top             =   200
         Width           =   1215
      End
   End
   Begin VB.TextBox txtTotal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0"
      Top             =   3000
      Width           =   3735
   End
   Begin VB.TextBox txtDescuento 
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Text            =   "0"
      Top             =   2520
      Width           =   3735
   End
   Begin VB.TextBox txtDespachado 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   2040
      Width           =   3735
   End
   Begin VB.TextBox txtSolicitado 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Text            =   "0"
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox txtValor 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox txtDescripcion 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "."
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox txtCodigoParte 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "."
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label7 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Descuento"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Despachado"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Solicitado"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Precio Unitario"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Descripción"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Código de Parte"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditaRepuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(1) = Me.txtCodigoParte.Text
frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(9) = Me.txtCodigoParte.Tag
frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(2) = Me.txtDescripcion.Text
frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(3) = Me.txtValor.Text
frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(4) = Me.txtSolicitado.Text
frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(5) = Me.txtDespachado.Text
frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(6) = Me.txtDescuento.Text
frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(7) = Me.txtTotal.Text
frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(8) = Me.txtDescripcion.Tag
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim ldblCont As Double

For ldblCont = 1 To frmOtServiteca.lsvRepuestos.ListItems.Count
    If ldblCont > frmOtServiteca.lsvRepuestos.ListItems.Count Then
        Exit For
    End If
    If frmOtServiteca.lsvRepuestos.ListItems(ldblCont).Selected = True Then
        Me.txtCodigoParte.Text = frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(1)
        Me.txtCodigoParte.Tag = frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(9)
        Me.txtDescripcion.Text = frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(2)
        Me.txtValor.Text = frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(3)
        Me.txtSolicitado.Text = frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(4)
        Me.txtDespachado.Text = frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(5)
        Me.txtDescuento.Text = frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(6)
        Me.txtTotal.Text = frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(7)
        Me.txtDescripcion.Tag = frmOtServiteca.lsvRepuestos.SelectedItem.SubItems(8)
    End If
Next ldblCont

End Sub

Private Sub txtDescuento_Change()

If Trim$(Me.txtDescuento.Text) = "" Then
    Me.txtDescuento.Text = "0"
End If
If Not IsNumeric(Me.txtDescuento.Text) Then
    MsgBox "El valor ningresado debe ser numérico.", vbExclamation, "ServiPro"
    Me.txtDescuento.Text = "0"
End If
If CDbl(Me.txtDescuento.Text) < 0 Or CDbl(Me.txtDescuento.Text) > 100 Then
    MsgBox "El valor del descuento debe ser entre 0 y 100.", vbExclamation, "ServiPro"
    Me.txtDescuento.Text = "0"
End If
If CDbl(Me.txtDescuento.Text) > 0 And CDbl(Me.txtDescuento.Text) <= 100 Then
    Me.txtTotal.Text = CDbl(SacarFormatoValor(Me.txtSolicitado.Text, gstrMonedaLocal)) * CDbl(SacarFormatoValor(Me.txtValor.Text, gstrMonedaLocal))
    Me.txtTotal.Text = CDbl(Me.txtTotal.Text) - (CDbl(Me.txtTotal.Text) * (CDbl(Me.txtDescuento.Text) / 100))
    Me.txtTotal.Text = FormatoValor(Me.txtTotal.Text, gstrMonedaLocal, gintDecimalesMoneda)
Else
    Me.txtDescuento.Text = "0"
    Me.txtTotal.Text = CDbl(SacarFormatoValor(Me.txtSolicitado.Text, gstrMonedaLocal)) * CDbl(SacarFormatoValor(Me.txtValor.Text, gstrMonedaLocal))
    Me.txtTotal.Text = FormatoValor(Me.txtTotal.Text, gstrMonedaLocal, gintDecimalesMoneda)
End If
End Sub

Private Sub txtDescuento_GotFocus()
Me.txtDescuento.SelStart = 0
Me.txtDescuento.SelLength = Len(Me.txtDescuento.Text)
End Sub

Private Sub txtDescuento_LostFocus()
If Not IsNumeric(Me.txtDescuento.Text) Then
    MsgBox "El valor ningresado debe ser numérico.", vbExclamation, "ServiPro"
    Me.txtDescuento.Text = "0"
End If
End Sub

Private Sub txtDespachado_Change()
If Trim$(Me.txtDespachado.Text) = "" Then
    Me.txtDespachado.Text = "0"
End If
If Not IsNumeric(Me.txtDespachado.Text) Then
    MsgBox "El valor ningresado debe ser numérico.", vbExclamation, "ServiPro"
    Me.txtDespachado.Text = "0"
End If
If CDbl(Me.txtDescuento.Text) > 0 And CDbl(Me.txtDescuento.Text) <= 100 Then
    Me.txtTotal.Text = CDbl(SacarFormatoValor(Me.txtSolicitado.Text, gstrMonedaLocal)) * CDbl(SacarFormatoValor(Me.txtValor.Text, gstrMonedaLocal))
    Me.txtTotal.Text = CDbl(Me.txtTotal.Text) - (CDbl(Me.txtTotal.Text) * (CDbl(Me.txtDescuento.Text) / 100))
    Me.txtTotal.Text = FormatoValor(Me.txtTotal.Text, gstrMonedaLocal, gintDecimalesMoneda)
Else
    Me.txtDescuento.Text = "0"
    Me.txtTotal.Text = CDbl(SacarFormatoValor(Me.txtSolicitado.Text, gstrMonedaLocal)) * CDbl(SacarFormatoValor(Me.txtValor.Text, gstrMonedaLocal))
    Me.txtTotal.Text = FormatoValor(Me.txtTotal.Text, gstrMonedaLocal, gintDecimalesMoneda)
End If
End Sub

Private Sub txtDespachado_GotFocus()
Me.txtDespachado.SelStart = 0
Me.txtDespachado.SelLength = Len(Me.txtDespachado.Text)
End Sub

Private Sub txtDespachado_LostFocus()
If Not IsNumeric(Me.txtDespachado.Text) Then
    MsgBox "El valor ningresado debe ser numérico.", vbExclamation, "ServiPro"
    Me.txtDespachado.Text = "0"
End If
End Sub

Private Sub txtSolicitado_Change()
If Trim$(Me.txtSolicitado.Text) = "" Then
    Me.txtSolicitado.Text = "0"
End If
If Not IsNumeric(Me.txtSolicitado.Text) Then
    MsgBox "El valor ningresado debe ser numérico.", vbExclamation, "ServiPro"
    Me.txtSolicitado.Text = "0"
End If
If CDbl(Me.txtDescuento.Text) > 0 And CDbl(Me.txtDescuento.Text) <= 100 Then
    Me.txtTotal.Text = CDbl(SacarFormatoValor(Me.txtSolicitado.Text, gstrMonedaLocal)) * CDbl(SacarFormatoValor(Me.txtValor.Text, gstrMonedaLocal))
    Me.txtTotal.Text = CDbl(Me.txtTotal.Text) - (CDbl(Me.txtTotal.Text) * (CDbl(Me.txtDescuento.Text) / 100))
    Me.txtTotal.Text = FormatoValor(Me.txtTotal.Text, gstrMonedaLocal, gintDecimalesMoneda)
Else
    Me.txtDescuento.Text = "0"
    Me.txtTotal.Text = CDbl(SacarFormatoValor(Me.txtSolicitado.Text, gstrMonedaLocal)) * CDbl(SacarFormatoValor(Me.txtValor.Text, gstrMonedaLocal))
    Me.txtTotal.Text = FormatoValor(Me.txtTotal.Text, gstrMonedaLocal, gintDecimalesMoneda)
End If
End Sub

Private Sub txtSolicitado_GotFocus()
Me.txtSolicitado.SelStart = 0
Me.txtSolicitado.SelLength = Len(Me.txtSolicitado.Text)
End Sub

Private Sub txtSolicitado_LostFocus()
If Not IsNumeric(Me.txtSolicitado.Text) Then
    MsgBox "El valor ningresado debe ser numérico.", vbExclamation, "ServiPro"
    Me.txtSolicitado.Text = "0"
End If
End Sub

Private Sub txtValor_Change()
'If Not IsNumeric(Me.txtValor.Text) Then
'    MsgBox "El valor ningresado debe ser numérico.", vbExclamation, "ServiPro"
'    Me.txtValor.Text = "0"
'End If
'If CDbl(Me.txtDescuento.Text) > 0 And CDbl(Me.txtDescuento.Text) <= 100 Then
'    Me.txtTotal.Text = CDbl(SacarFormatoValor(Me.txtSolicitado.Text, gstrMonedaLocal)) * CDbl(SacarFormatoValor(Me.txtValor.Text, gstrMonedaLocal))
'    Me.txtTotal.Text = CDbl(Me.txtTotal.Text) - (CDbl(Me.txtTotal.Text) * (CDbl(Me.txtDescuento.Text) / 100))
'    Me.txtTotal.Text = FormatoValor(Me.txtTotal.Text, gstrMonedaLocal, gintDecimalesMoneda)
'Else
'    Me.txtDescuento.Text = "0"
'    Me.txtTotal.Text = CDbl(SacarFormatoValor(Me.txtSolicitado.Text, gstrMonedaLocal)) * CDbl(SacarFormatoValor(Me.txtValor.Text, gstrMonedaLocal))
'    Me.txtTotal.Text = FormatoValor(Me.txtTotal.Text, gstrMonedaLocal, gintDecimalesMoneda)
'End If
End Sub

Private Sub txtValor_GotFocus()
'Me.txtValor.Text = SacarFormatoValor(Me.txtValor.Text, gstrMonedaLocal)
'Me.txtValor.SelStart = 0
'Me.txtValor.SelLength = Len(Me.txtValor.Text)
End Sub

Private Sub txtValor_LostFocus()
'If Not IsNumeric(Me.txtValor.Text) Then
'    MsgBox "El valor ningresado debe ser numérico.", vbExclamation, "ServiPro"
'    Me.txtValor.Text = "0"
'End If
'Me.txtValor.Text = FormatoValor(Me.txtValor.Text, gstrMonedaLocal, gintDecimalesMoneda)
End Sub
