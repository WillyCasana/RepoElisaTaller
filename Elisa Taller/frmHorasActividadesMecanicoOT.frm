VERSION 5.00
Begin VB.Form frmHorasActividadesMecanicoOT 
   Caption         =   "Ingrese Mecanicos por Actividad"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frFilaCol 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton cmdBuscaMecanico 
         Height          =   350
         Left            =   5040
         Picture         =   "frmHorasActividadesMecanicoOT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1300
         Width           =   375
      End
      Begin VB.TextBox txtHorasReales 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   2040
         TabIndex        =   0
         Text            =   "0"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtMecanicoAsignado 
         Height          =   330
         Left            =   2040
         TabIndex        =   3
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label lblHorasActividad 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblValor 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Horas Reales           :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Mecanico Asignado :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Horas Actividad       :"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Valor                        :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   300
      Left            =   4800
      TabIndex        =   2
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   3840
      TabIndex        =   1
      Top             =   2760
      Width           =   855
   End
End
Attribute VB_Name = "frmHorasActividadesMecanicoOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnSW As Boolean

Private Sub cmdAceptar_Click()

    If Me.txtHorasReales = "" Then
        MsgBox "Las Horas Reales debe contener un valor...", vbInformation, "Advertencia"
        Me.txtHorasReales.SetFocus
        Exit Sub
    End If
    If Me.txtMecanicoAsignado = "" Then
        MsgBox "El Mecánico debe Contener contener un valor...", vbInformation, "Advertencia"
        Me.txtMecanicoAsignado.SetFocus
        Exit Sub
    End If
    DescargaDatos
    Unload Me
End Sub

Private Sub cmdBuscaMecanico_Click()
    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "(select * from Tllr_Mecanicos where id_empresa='" & gstrIdEmpresa & "' and id_sucursal='" & gstrIdSucursal & "' And Vigencia='S') as Tllr_Mecanicos", "Id_Mecanico", "Nombre", "Buscar Mecánico")
    If gstrBusca <> "" Then
        txtMecanicoAsignado.Tag = gstrBusca
        txtMecanicoAsignado = TraeNombreMecanico(gstrBusca)
        txtHorasReales.SetFocus
    End If
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

Private Sub Form_Load()
mblnSW = True
End Sub
Sub CargaDatos()
If Me.Tag = "Mecanica" Then
    With frmIngresaHorasMecanicoActividades.lvDetalleActiv
        
        Me.lblHorasActividad = .SelectedItem.SubItems(2)
        Me.lblValor = SacarFormatoValor(.SelectedItem.SubItems(3), "")
        Me.txtHorasReales = SacarFormatoValor(.SelectedItem.SubItems(6), "")
        Me.txtMecanicoAsignado = .SelectedItem.SubItems(4)
        Me.txtMecanicoAsignado.Tag = .SelectedItem.SubItems(5)
    End With
Else
    With frmIngresaHorasMecanicoActividades.lvDetalleOtroS
        Me.lblHorasActividad = .SelectedItem.SubItems(2)
        Me.lblValor = SacarFormatoValor(.SelectedItem.SubItems(3), "")
        Me.txtHorasReales = SacarFormatoValor(.SelectedItem.SubItems(6), "")
        Me.txtMecanicoAsignado = .SelectedItem.SubItems(4)
        Me.txtMecanicoAsignado.Tag = .SelectedItem.SubItems(5)
    End With
End If
End Sub
Sub DescargaDatos()
If Me.Tag = "Mecanica" Then
    With frmIngresaHorasMecanicoActividades.lvDetalleActiv
        .SelectedItem.SubItems(4) = Me.txtMecanicoAsignado.Text
        .SelectedItem.SubItems(5) = Me.txtMecanicoAsignado.Tag
        .SelectedItem.SubItems(6) = FormatoValor(Me.txtHorasReales, "", 1)
    End With
Else
    With frmIngresaHorasMecanicoActividades.lvDetalleOtroS
        .SelectedItem.SubItems(4) = Me.txtMecanicoAsignado.Text
        .SelectedItem.SubItems(5) = Me.txtMecanicoAsignado.Tag
        .SelectedItem.SubItems(6) = FormatoValor(Me.txtHorasReales, "", 1)
    End With
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{tab}"
    End Select
End Sub

Private Sub txtHorasReales_GotFocus()
MarcaTexto txtHorasReales
End Sub

Private Sub txtHorasReales_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtHorasReales, strDot)
End Sub

Private Sub txtMecanicoAsignado_GotFocus()
MarcaTexto txtMecanicoAsignado
End Sub

Private Sub txtMecanicoAsignado_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtMecanicoAsignado_LostFocus()
Me.txtMecanicoAsignado.Tag = Me.txtMecanicoAsignado
Me.txtMecanicoAsignado = TraeNombreMecanico(Me.txtMecanicoAsignado.Tag)
End Sub
