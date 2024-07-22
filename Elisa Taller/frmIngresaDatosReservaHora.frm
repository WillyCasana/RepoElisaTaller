VERSION 5.00
Begin VB.Form frmIngresaDatosReservaHora 
   Caption         =   "Ingrese Datos Para la Reserva de Hora"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIngresaDatosReservaHora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtFono 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtVehiculo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   3
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "DNI/RUC"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Teléfono"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Vehiculo"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "frmIngresaDatosReservaHora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub DescargaDatos()
With frmReservadeHoras
    .lblModelo = txtVehiculo
    .lblCliente = txtNombre
    .lblFono = txtFono
    .optSinPatente.Value = True
End With
End Sub


Private Sub cmdAceptar_Click()
If Me.txtCodigo = "" Then
    MsgBox "El DNI/RUC debe contener un valor...", vbInformation, "Advertencia"
    txtCodigo.SetFocus
    Exit Sub
End If


If Me.txtNombre = "" Then
    MsgBox "El Nombre debe contener un valor...", vbInformation, "Advertencia"
    txtNombre.SetFocus
    Exit Sub
End If
If Me.txtFono = "" Then
    MsgBox "El Teléfono debe contener un valor...", vbInformation, "Advertencia"
    txtFono.SetFocus
    Exit Sub
End If
'kjcv 30.10.15
    If ConsultaCliente(txtCodigo) = True Then
        MsgBox "No hay Cupo en el Taller...", vbCritical, "Elisa Taller"
        Exit Sub
    End If
DescargaDatos
Unload Me
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

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    If txtCodigo <> "" Then
            If ConsultaCliente(txtCodigo) = True Then
                MsgBox "No hay Cupo en el Taller...", vbCritical, "Elisa Taller"
            End If
    End If
  End If
End Sub

Private Sub txtFono_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtVehiculo_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub
