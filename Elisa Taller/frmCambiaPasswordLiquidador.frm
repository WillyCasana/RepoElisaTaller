VERSION 5.00
Begin VB.Form frmCambiaPasswordLiquidador 
   Caption         =   "Cambia Password Liquidador"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   3090
   Icon            =   "frmCambiaPasswordLiquidador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   1755
      TabIndex        =   4
      Top             =   1980
      Width           =   1005
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   630
      TabIndex        =   3
      Top             =   1980
      Width           =   960
   End
   Begin VB.TextBox txtPassConfirma 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1665
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1260
      Width           =   1095
   End
   Begin VB.TextBox txtPassNueva 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1665
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   765
      Width           =   1095
   End
   Begin VB.TextBox txtPassAntigua 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1665
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   270
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Confirme Password"
      Height          =   240
      Left            =   135
      TabIndex        =   7
      Top             =   1350
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Password Nueva"
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   855
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Password Antigua"
      Height          =   240
      Left            =   135
      TabIndex        =   5
      Top             =   360
      Width           =   1320
   End
End
Attribute VB_Name = "frmCambiaPasswordLiquidador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    
    If Me.txtPassAntigua = "" Then
        MsgBox "Debe Ingresar la Password Antigua", vbInformation, "Advertencia"
        txtPassAntigua.SetFocus
        Exit Sub
    End If
    If Me.txtPassNueva = "2" Then
        MsgBox "Debe Ingresar la Password Nueva", vbInformation, "Advertencia"
        txtPassNueva.SetFocus
        Exit Sub
    End If
    If Me.txtPassConfirma = "2" Then
        MsgBox "Debe Confirmar la Password ", vbInformation, "Advertencia"
        txtPassConfirma.SetFocus
        Exit Sub
    End If
    
    mstrSql = "UPDATE Tllr_Mecanicos SET PasswordLiquidador = " & Me.txtPassNueva
    mstrSql = mstrSql & " WHERE Id_Mecanico = '" & frmMantenedorMecanicos.txtCodigo & "'"
    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
        MsgBox "La Password Fue Cambiada con exito", vbInformation, "Cambiar Password"
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
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

Private Sub txtPassAntigua_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtPassAntigua, strDot)
End Sub

Private Sub txtPassAntigua_LostFocus()
    If txtPassAntigua <> "" Then
        If NoEsLaPassword(txtPassAntigua, frmMantenedorMecanicos.txtCodigo) = False Then
              MsgBox "Password Mal Ingresada... Intente otra Vez", vbExclamation, "Cambiar Password"
              txtPassAntigua.SetFocus
        End If
    End If
End Sub

Private Sub txtPassConfirma_LostFocus()
If Me.txtPassNueva <> Me.txtPassConfirma Then
    MsgBox "La confirmación de la Password fue mal ingresada...", vbExclamation, "Cambiar Password"
    Me.txtPassConfirma.SetFocus
End If
End Sub

Private Sub txtPassNueva_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtPassNueva, strDot)
End Sub

