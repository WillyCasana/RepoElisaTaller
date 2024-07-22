VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddOtrosServicios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trabajos Adicionales (Servicios no Temparizados)"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   Icon            =   "frmAddOtrosServicios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   1980
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":179A
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":18AC
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":19BE
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":1AD0
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":1BE2
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":1CF4
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":1E06
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":1F18
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":202A
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":213C
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":224E
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":2360
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":2472
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":2584
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":2696
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":27A8
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":28BA
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":2D0C
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":315E
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddOtrosServicios.frx":3270
            Key             =   "Salir"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   15
      TabIndex        =   6
      Top             =   330
      Width           =   6120
      Begin VB.TextBox txtSubTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1275
         Width           =   1170
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   1
         Top             =   570
         Width           =   4875
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   2
         Top             =   915
         Width           =   1170
      End
      Begin VB.TextBox txtTiempo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3345
         MaxLength       =   50
         TabIndex        =   3
         Top             =   915
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Sub - Total :"
         Height          =   315
         Index           =   4
         Left            =   75
         TabIndex        =   11
         Top             =   1275
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo:"
         Height          =   315
         Index           =   0
         Left            =   75
         TabIndex        =   10
         Top             =   255
         Width           =   915
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1065
         TabIndex        =   0
         Top             =   225
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción:"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   9
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Valor :"
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   8
         Top             =   945
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Tiempo :"
         Height          =   315
         Index           =   3
         Left            =   2370
         TabIndex        =   7
         Top             =   915
         Width           =   915
      End
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      DisabledImageList=   "ImgBarraHerramienta"
      HotImageList    =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Agregar"
            Object.ToolTipText     =   "Agregar a la Lista"
            ImageKey        =   "Editar"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar Modo Edición"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAddOtrosServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curSubTotal As Currency
Sub LimpiaCampos()
With Me
    .txtDescripcion = ""
    .txtTiempo = ""
    .txtSubTotal = ""
    '.txtValor = gcurPrecioManoObra
End With

End Sub
Sub DownLoadDataOS()

With frmRecepcion
    Set glsiItem = .lvwOtrosServicios.ListItems.Add(, , lblCodigo)
    glsiItem.SubItems(1) = IIf(txtDescripcion <> "", UCase(txtDescripcion), ".")
    glsiItem.SubItems(2) = IIf(txtTiempo <> "", FormatoValor(txtTiempo, "", 2), 0)
    glsiItem.SubItems(3) = IIf(txtValor <> "", FormatoValor(txtValor, "", gintDecimalesMoneda), 0)
    glsiItem.SubItems(4) = FormatoValor(0, "", 1)
    glsiItem.SubItems(5) = FormatoValor(0, "", 1)
    glsiItem.SubItems(6) = gstrIdCargo
    glsiItem.SubItems(7) = TraeCargoDes(gstrIdCargo)
    glsiItem.SubItems(8) = gstrMecanicoDefectoSecMec
    glsiItem.SubItems(9) = MecanicoD(gstrMecanicoDefectoSecMec)
    '//MODIFICADO POR FDO DIAZ EL 12/12/2000
    'glsiItem.SubItems(10) = FormatoValor(txtSubTotal, "", 0)
    glsiItem.SubItems(10) = FormatoValor(CCur(Val(txtValor) * Val(txtTiempo)), "", gintDecimalesMoneda)
    glsiItem.SubItems(11) = "N"
End With
IncrementaCorrelativoOtrosServicios gstrIdEmpresa, gstrIdSucursal
End Sub

Private Sub Form_Load()
lblCodigo = TraeIndiceOtrosServicio(gstrIdEmpresa, gstrIdSucursal)
txtValor = ValorHora(gstrIdEmpresa, gstrIdSucursal)
End Sub

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Nuevo"
        lblCodigo = TraeIndiceOtrosServicio(gstrIdEmpresa, gstrIdSucursal)
        LimpiaCampos
Case "Agregar"
    Set glsiItem = frmRecepcion.lvwOtrosServicios.FindItem(lblCodigo)
    If glsiItem Is Nothing Then
        DownLoadDataOS
        LimpiaCampos
        lblCodigo = TraeIndiceOtrosServicio(gstrIdEmpresa, gstrIdSucursal)
    Else
        MsgBox "El Item Que Intenta Agregar, ya Existe en la Lista, por favor Verifique"
    End If
Case "Cerrar"
    Unload Me
End Select

End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
If KeyAscii = 13 Then
    DownLoadDataOS
    LimpiaCampos
    lblCodigo = TraeIndiceOtrosServicio(gstrIdEmpresa, gstrIdSucursal)
End If
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub txtSubTotal_GotFocus()
MarcaTexto txtSubTotal
End Sub

Private Sub txtSubTotal_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    DownLoadDataOS
    LimpiaCampos
    lblCodigo = TraeIndiceOtrosServicio(gstrIdEmpresa, gstrIdSucursal)
End If
If KeyAscii = 27 Then
    Unload Me
End If
KeyAscii = CheckNumber(KeyAscii, txtSubTotal, strDot)
End Sub

Private Sub txtSubTotal_LostFocus()
    If txtValor <> "0" Then
        txtTiempo = CCur(Val(txtSubTotal) / Val(txtValor))  'gcurPrecioManoObra)
    End If
End Sub

Private Sub txtTiempo_GotFocus()
MarcaTexto txtTiempo
End Sub

Private Sub txtTiempo_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    DownLoadDataOS
    LimpiaCampos
    lblCodigo = TraeIndiceOtrosServicio(gstrIdEmpresa, gstrIdSucursal)
End If
If KeyAscii = 27 Then
    Unload Me
End If
KeyAscii = CheckNumber(KeyAscii, txtTiempo, strDot)
End Sub

Private Sub txtTiempo_LostFocus()
If Val(txtValor) > 0 Then
    curSubTotal = CCur(Val(txtValor) * Val(txtTiempo))
Else
    curSubTotal = 0
End If

txtSubTotal = CStr(curSubTotal)
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    DownLoadDataOS
    LimpiaCampos
    lblCodigo = TraeIndiceOtrosServicio(gstrIdEmpresa, gstrIdSucursal)
End If
If KeyAscii = 27 Then
    Unload Me
End If
KeyAscii = CheckNumber(KeyAscii, txtValor, strDot)
End Sub

Private Sub txtValor_LostFocus()

'If Val(txtTiempo) > 0 Then
'    curSubTotal = CCur(Val(txtValor) * Val(txtTiempo))
'Else
'    curSubTotal = 0
'End If
End Sub


