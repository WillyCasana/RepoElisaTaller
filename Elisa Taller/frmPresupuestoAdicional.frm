VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPresupuestoAdicional 
   Caption         =   "Prespupuesto Adicional"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   Icon            =   "frmPresupuestoAdicional.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4815
      Begin VB.OptionButton optOTExistente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "OT Existente"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optOTNueva 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "OT Nueva"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txtOT 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   2
         Top             =   960
         Width           =   2295
      End
      Begin MSComctlLib.Toolbar tlbOT 
         Height          =   330
         Left            =   4320
         TabIndex        =   5
         Top             =   960
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar"
               ImageIndex      =   9
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "N° OT"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   240
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":179A
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":18AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":1D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":215C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":25B4
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":26C6
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":27D8
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":28EA
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":29FC
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":2B0E
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":2C20
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":2D32
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":2E44
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":2F56
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":3068
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":317A
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":328C
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":339E
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":34B0
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":35C2
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":3A14
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPresupuestoAdicional.frx":3E66
            Key             =   "Copiar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPresupuestoAdicional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
Dim mstrPatente As String

    If optOTExistente.Value = True Then
        If txtOT = "" Then
            MsgBox "El Número de OT debe contener un valor", vbExclamation, "Advertencia"
            Exit Sub
        End If
        If gstrEstadoOT <> "V" Then
            MsgBox "Esta Ot no se Encuentra Vigente...", vbExclamation, "Advertencia"
            Exit Sub
        End If
        mstrPatente = Retorna_Valor_General("Select Patente from Tllr_OT where Id_Ot='" & txtOT & "' And Seccion_Ot='" & gstrSeccion & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'", gcdynamic)
        If frmRecepcion.txtPatente <> mstrPatente Then
            MsgBox "La " & gstrNombrePatente & " no Coinciden...", vbExclamation, "Advertencia"
            Exit Sub
        End If
        
        gintOtExistente = 1
    Else
        gintOtExistente = 2
    End If
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
gintOtExistente = 0
Unload Me
End Sub

Private Sub Form_Load()
    gintOtExistente = 0
End Sub

Private Sub optOTExistente_Click()
    txtOT.Enabled = True
    tlbOT.Enabled = True
End Sub

Private Sub optOTNueva_Click()
    txtOT = ""
    txtOT.Enabled = False
    tlbOT.Enabled = False
End Sub

Private Sub tlbOT_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    frmBuscaOT.Show vbModal
    txtOT.Tag = gstrEstadoOT
    txtOT = gstrBusca
End If

End Sub
