VERSION 5.00
Begin VB.Form frmAgregaFilaColTempario 
   Caption         =   "Agrega Columna Temparios de Carroceria"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   300
      Left            =   4080
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   3120
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.Frame frFilaCol 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtTipo 
         Height          =   285
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   3
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   2
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "A=Des/Mont  D=Desab.  P=Pintura"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   1120
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo                 :"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción     :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Código             :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAgregaFilaColTempario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    
    If txtCodigo = "" Then
        MsgBox "El código debe contener un valor...", vbInformation, "Advertencia"
        txtCodigo.SetFocus
        Exit Sub
    End If
    If txtDescripcion = "" Then
        MsgBox "La descripción debe contener un valor...", vbInformation, "Advertencia"
        txtDescripcion.SetFocus
        Exit Sub
    End If
    If Me.Tag = "Columna" Then
        If txtTipo = "" Then
            MsgBox "El tipo de Columna debe contener un valor...", vbInformation, "Advertencia"
            txtTipo.SetFocus
            Exit Sub
        End If
        If txtTipo <> "A" And txtTipo <> "D" And txtTipo <> "P" Then
            MsgBox "El tipo de Columna debe contener un valor A, D o P...", vbInformation, "Advertencia"
            txtTipo.SetFocus
            Exit Sub
        End If
    End If
    
    frmTemparios.lblCodigo = Me.txtCodigo
    frmTemparios.lblDescripcion = Me.txtDescripcion
    frmTemparios.lblTipo = Me.txtTipo
    
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    frmTemparios.lblCodigo = ""
    Unload Me
End Sub

Private Sub Form_Activate()
Screen.MousePointer = Default
If Me.Tag = "Fila" Then
    Me.Label3.Visible = False
    Me.txtTipo.Visible = False
    Me.Label4.Visible = False
    Me.Caption = "Agrega Fila Temparios de Carroceria"
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{tab}"
    End Select
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtTipo_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub
