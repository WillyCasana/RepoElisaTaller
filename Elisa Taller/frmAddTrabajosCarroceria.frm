VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmAddTrabajosCarroceria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trabajos de Carroceria"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   Icon            =   "frmAddTrabajosCarroceria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3000
      Left            =   45
      TabIndex        =   14
      Top             =   315
      Width           =   6120
      Begin VB.TextBox txtTipoTrabajo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5490
         MaxLength       =   1
         TabIndex        =   2
         ToolTipText     =   "Digite Sólo las letras ""D"" ""P"" o ""A"""
         Top             =   855
         Width           =   330
      End
      Begin VB.TextBox txtProveedor 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1035
         TabIndex        =   28
         ToolTipText     =   "Puede Ingresar el Rut del Proveedor"
         Top             =   180
         Width           =   4155
      End
      Begin VB.TextBox txtTipoCargo 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1020
         TabIndex        =   12
         Top             =   2565
         Width           =   2850
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         MaxLength       =   70
         TabIndex        =   1
         Top             =   855
         Width           =   3975
      End
      Begin VB.TextBox txtMtoDcto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3630
         MaxLength       =   8
         TabIndex        =   10
         Top             =   1860
         Width           =   1320
      End
      Begin VB.TextBox txtPorcDcto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         MaxLength       =   4
         TabIndex        =   9
         Top             =   1845
         Width           =   795
      End
      Begin VB.TextBox txtSubTot 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3630
         MaxLength       =   9
         TabIndex        =   8
         Top             =   2205
         Width           =   1320
      End
      Begin VB.TextBox txtPreFin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   11
         Top             =   2190
         Width           =   1170
      End
      Begin VB.TextBox txtPorcRec 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1485
         Width           =   795
      End
      Begin VB.TextBox txtMtoRec 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3630
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1515
         Width           =   1320
      End
      Begin VB.TextBox txtPreUni 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3630
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1185
         Width           =   1320
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1170
         Width           =   795
      End
      Begin MSComctlLib.Toolbar tlbOpciones 
         Height          =   330
         Left            =   5325
         TabIndex        =   0
         Top             =   180
         Width           =   405
         _ExtentX        =   714
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
               Object.ToolTipText     =   "Busca Proveedor"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbCargo 
         Height          =   330
         Left            =   3975
         TabIndex        =   19
         Top             =   2595
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImgBarraHerramienta"
         DisabledImageList=   "ImgBarraHerramienta"
         HotImageList    =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "BuscarCargo"
               Object.ToolTipText     =   "Busca Cargo"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo "
         Height          =   195
         Left            =   5085
         TabIndex        =   29
         Top             =   900
         Width           =   330
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Dscto    :"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   27
         Top             =   1860
         Width           =   810
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dscto. (S/.)     :"
         Height          =   195
         Index           =   0
         Left            =   2535
         TabIndex        =   26
         Top             =   1875
         Width           =   1095
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   24
         Left            =   90
         TabIndex        =   25
         Top             =   900
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SubTotal   :"
         Height          =   195
         Index           =   6
         Left            =   2535
         TabIndex        =   24
         Top             =   2220
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Precio Final  :"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   23
         Top             =   2220
         Width           =   960
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Cargo :"
         Height          =   195
         Index           =   58
         Left            =   90
         TabIndex        =   22
         Top             =   2610
         Width           =   870
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recargo (S/.) :"
         Height          =   195
         Index           =   55
         Left            =   2535
         TabIndex        =   21
         Top             =   1575
         Width           =   1065
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Recargo :"
         Height          =   240
         Index           =   56
         Left            =   90
         TabIndex        =   20
         Top             =   1560
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor  :"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   18
         Top             =   210
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo       :"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   540
         Width           =   855
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1020
         TabIndex        =   3
         Top             =   540
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unitario (S/.)  :"
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   16
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad    :"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   15
         Top             =   1230
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   13
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
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar Modo Edición"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   5520
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
            Picture         =   "frmAddTrabajosCarroceria.frx":179A
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":18AC
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":19BE
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":1AD0
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":1BE2
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":1CF4
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":1E06
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":1F18
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":202A
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":213C
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":224E
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":2360
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":2472
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":2584
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":2696
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":27A8
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":28BA
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":2D0C
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":315E
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddTrabajosCarroceria.frx":3270
            Key             =   "Salir"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAddTrabajosCarroceria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curSubTotal  As Currency
Dim mstrnombre As String
Sub LimpiaCampos()
With Me
    'If MsgBox("Desea Conservar Proveedor", 4 + 32, "Trabajo de Tercero") = vbNo Then
    '    .lblProveedor.Tag = ""
    '    .lblProveedor.Caption = ""
    'End If
    .lblCodigo = ""
    .txtDescripcion = ""
    .txtCantidad = ""
    .txtPreUni = ""
    .txtPorcRec = ""
    .txtMtoRec = ""
    .txtPorcDcto = ""
    .txtMtoDcto = ""
    .txtPreFin = ""
    .txtSubTot = ""
    .txtDescripcion.SetFocus
    'If MsgBox("Desea Conservar Tipo Cargo", 4 + 32, "Trabajo de Tercero") = vbNo Then
    '    .lblTipoCargo = ""
    '    .lblTipoCargo.Tag = ""
    'End If
End With

End Sub
Sub DownLoadDataTer()
With frmRecepcion
    Set glsiItem = .lvwServiciosCarroceria.ListItems.Add(, , "Concepto Carroceria")
    glsiItem.SubItems(1) = "01"  'dtcConceptos.BoundText
    glsiItem.SubItems(2) = UCase(txtDescripcion)   'txtSeccion
    glsiItem.SubItems(3) = txtTipoTrabajo    '   "Parte / Pieza"   dtcPartePieza.Text
    glsiItem.SubItems(4) = "01"   'dtcPartePieza.BoundText
    glsiItem.SubItems(5) = FormatoValor(txtCantidad, "", 1)
    glsiItem.SubItems(6) = FormatoValor(txtPreUni, "", gintDecimalesMoneda)
    glsiItem.SubItems(7) = FormatoValor(txtPorcRec, "", 2)
    glsiItem.SubItems(8) = FormatoValor(txtMtoRec, "", gintDecimalesMoneda)
    glsiItem.SubItems(9) = FormatoValor(txtPreFin, "", gintDecimalesMoneda)
    glsiItem.SubItems(10) = FormatoValor(txtPorcDcto, "", 2)
    glsiItem.SubItems(11) = FormatoValor(txtMtoDcto, "", gintDecimalesMoneda)
    glsiItem.SubItems(12) = IIf(txtTipoCargo.Tag = "", TraeCargoDes(gstrIdCargo), TraeCargoDes(txtTipoCargo.Tag))
    glsiItem.SubItems(13) = IIf(txtTipoCargo.Tag = "", gstrIdCargo, txtTipoCargo.Tag)
    glsiItem.SubItems(14) = txtProveedor
    glsiItem.SubItems(15) = txtProveedor.Tag
    glsiItem.SubItems(16) = FormatoValor(txtSubTot, "", gintDecimalesMoneda)
    glsiItem.SubItems(17) = "N"
    glsiItem.SubItems(18) = lblCodigo
    
End With
IncrementaCorrelativoTrabajosTerceros gstrIdEmpresa, gstrIdSucursal
End Sub

Private Sub Form_Load()
lblCodigo = TraeIndiceTrabajosTerceros(gstrIdEmpresa, gstrIdSucursal)
Me.Label1(2).Caption = gstrMonedaLocal & " Unitario"
Me.Label(55).Caption = gstrMonedaLocal & " Recargo"
Me.Label(0).Caption = gstrMonedaLocal & " Dscto."
'FillConceptosVsCiaSeguro dtcConceptos, datConceptos, frmRecepcion.lblCompañia.Tag
'FillPartePieza dtcPartePieza, datPartesPiezas
End Sub

Private Sub Text1_Change()

End Sub

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Nuevo"
        lblCodigo = TraeIndiceTrabajosTerceros(gstrIdEmpresa, gstrIdSucursal)
        LimpiaCampos
Case "Agregar"
    Set glsiItem = frmRecepcion.lvwServiciosCarroceria.FindItem(lblCodigo)
    If glsiItem Is Nothing Then
        DownLoadDataTer
        LimpiaCampos
        lblCodigo = TraeIndiceTrabajosTerceros(gstrIdEmpresa, gstrIdSucursal)
    Else
        MsgBox "El Item Que Intenta Agregar, ya Existe en la Lista, por favor Verifique"
    End If
Case "Cerrar"
    Unload Me
End Select

End Sub

Private Sub tlbCargo_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "BuscarCargo" Then
'kjcv 24.03.20
    gstrBusca = ""
    frmTipoCargo.Show vbModal
'    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Tllr_Tipo_Cargo", "Id_Tipo_cargo", "Descripcion", "Buscar Cargo OT")
    If gstrBusca <> "" Then
        txtTipoCargo.Tag = gstrBusca
        txtTipoCargo = TraeCargoDes(gstrBusca)
        ValidaCostoCargo
    End If
End If
End Sub

Private Sub tlbOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
'kjcv 14.02.12
Dim lstrIdCliente As String
Dim lstrDescripCliente As String
lstrIdCliente = ""
lstrDescripCliente = ""

If Button.Key = "Buscar" Then
    Me.txtDescripcion.SetFocus
    Libreria.ClienteBuscar Conexion, lstrIdCliente, lstrDescripCliente, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
'    apfFormulario.BuscarRegistroClientes Conexion, gstrBusca, mstrnombre, gstrIdEmpresa
    'apfFormulario.BuscarRegistroClientes Conexion, gstrBusca, mstrnombre
    'lblProveedor.Tag = gstrBusca
    'lblProveedor.Caption = mstrNombre
    txtProveedor.Tag = lstrIdCliente
    txtProveedor = lstrDescripCliente
End If
End Sub



Private Sub txtCantidad_GotFocus()
MarcaTexto txtCantidad
End Sub


Private Sub txtCantidad_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    DownLoadDataTer
    LimpiaCampos
    lblCodigo = TraeIndiceTrabajosTerceros(gstrIdEmpresa, gstrIdSucursal)
End If
If KeyAscii = 27 Then
    Unload Me
End If
KeyAscii = CheckNumber(KeyAscii, txtCantidad, strDot)
End Sub

Private Sub txtCantidad_LostFocus()
'If Trim(txtCantidad) <> "" Then
'    curSubTotal = CCur(IIf(txtPreFin <> "", txtPreFin, "0")) * CDbl(IIf(txtCantidad <> "", txtCantidad, "0"))
'    txtSubTot = FormatoValor(curSubTotal, "", 0)
'Else
'    txtSubTot = "0"
'End If
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
If KeyAscii = 13 Then
    DownLoadDataTer
    LimpiaCampos
    lblCodigo = TraeIndiceTrabajosTerceros(gstrIdEmpresa, gstrIdSucursal)
End If
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub txtMtoDcto_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    DownLoadDataTer
    LimpiaCampos
    lblCodigo = TraeIndiceTrabajosTerceros(gstrIdEmpresa, gstrIdSucursal)
End If
If KeyAscii = 27 Then
    Unload Me
End If
KeyAscii = CheckNumber(KeyAscii, txtMtoDcto, strDot)
End Sub

Private Sub txtMtoDcto_LostFocus()
Dim dblMtoinicial As Double
With Me
    dblMtoinicial = CDbl(IIf(.txtCantidad <> "", txtCantidad, "0")) * CCur(IIf(.txtPreFin <> "", txtPreFin, "0"))
    .txtPorcDcto = PorcentajeMonto(CDbl(IIf(dblMtoinicial <> 0, dblMtoinicial, "0")), CSng(IIf(.txtMtoDcto <> "", txtMtoDcto, "0")))
    .txtSubTot = (CCur(IIf(.txtPreFin <> "", txtPreFin, "0")) * CDbl(IIf(.txtCantidad <> "", txtCantidad, "0"))) - CCur(IIf(.txtMtoDcto <> "", txtMtoDcto, "0"))
End With
End Sub

Private Sub txtMtoRec_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    DownLoadDataTer
    LimpiaCampos
    lblCodigo = TraeIndiceTrabajosTerceros(gstrIdEmpresa, gstrIdSucursal)
End If
If KeyAscii = 27 Then
    Unload Me
End If
KeyAscii = CheckNumber(KeyAscii, txtMtoRec, strDot)
End Sub

Private Sub txtMtoRec_LostFocus()
Dim dblMtoinicial As Double
With Me
    dblMtoinicial = CDbl(IIf(.txtCantidad <> "", txtCantidad, "0")) * CCur(IIf(.txtPreUni <> "", txtPreUni, "0"))
    .txtPorcRec = PorcentajeMonto(CDbl(IIf(.txtPreUni <> "", txtPreUni, "0")), CSng(IIf(.txtMtoRec <> "", txtMtoRec, "0")))
    .txtPreFin = CDbl(IIf(.txtPreUni <> "", txtPreUni, "0")) + CCur(IIf(.txtMtoRec <> "", txtMtoRec, "0"))
    .txtSubTot = CCur(IIf(.txtPreFin <> "", txtPreFin, "0")) * CDbl(IIf(.txtCantidad <> "", txtCantidad, "0"))
End With
End Sub

Private Sub txtPorcDcto_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    DownLoadDataTer
    LimpiaCampos
    lblCodigo = TraeIndiceTrabajosTerceros(gstrIdEmpresa, gstrIdSucursal)
End If
If KeyAscii = 27 Then
    Unload Me
End If
KeyAscii = CheckNumber(KeyAscii, txtPorcDcto, strDot)
End Sub

Private Sub txtPorcDcto_LostFocus()
Dim dblMtoinicial As Double

dblMtoinicial = 0
If Trim(txtPorcDcto) <> "" Then
    With Me
        dblMtoinicial = CDbl(IIf(.txtCantidad <> "", txtCantidad, "0")) * CCur(IIf(.txtPreUni <> "", txtPreUni, "0"))
        '.txtMtoRec = ValorPorcentaje(CDbl(IIf(.txtPreUni <> "", txtPreUni, "0")), CSng(IIf(.txtPorcRec <> "", txtPorcRec, "0")))
        .txtPreFin = CDbl(IIf(.txtPreUni <> "", txtPreUni, "0")) + CCur(IIf(.txtMtoRec <> "", txtMtoRec, "0"))
'        .txtMtoDcto = Format(ValorPorcentaje(CCur(IIf(.txtPreFin <> "", .txtPreFin, 0)) * CCur(IIf(.txtCantidad <> "", .txtCantidad, 0)), CSng(IIf(.txtPorcDcto <> "", txtPorcDcto, "0"))))
        .txtMtoDcto = Format(ValorPorcentaje(CCur(IIf(.txtPreFin <> "", .txtPreFin, 0)) * CCur(IIf(.txtCantidad <> "", .txtCantidad, 0)), CSng(IIf(.txtPorcDcto <> "", txtPorcDcto, "0"))))
        '.txtSubTot = CCur(IIf(.txtPreFin <> "", txtPreFin, "0")) * CDbl(IIf(.txtCantidad <> "", txtCantidad, "0"))
'        .txtSubTot = Format((CCur(IIf(.txtPreFin <> "", txtPreFin, "0")) * CDbl(IIf(.txtCantidad <> "", txtCantidad, "0"))) - CCur(IIf(.txtMtoDcto <> "", .txtMtoDcto, 0)), "", 0)
        .txtSubTot = Format((CCur(IIf(.txtPreFin <> "", txtPreFin, "0")) * CDbl(IIf(.txtCantidad <> "", txtCantidad, "0"))) - CCur(IIf(.txtMtoDcto <> "", .txtMtoDcto, 0)), "", 0)
    End With
Else
    txtPorcDcto = "0"
End If

End Sub

Private Sub txtPorcRec_GotFocus()
MarcaTexto txtPorcRec
End Sub


Private Sub txtPorcRec_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    DownLoadDataTer
    LimpiaCampos
    lblCodigo = TraeIndiceTrabajosTerceros(gstrIdEmpresa, gstrIdSucursal)
End If
If KeyAscii = 27 Then
    Unload Me
End If
KeyAscii = CheckNumber(KeyAscii, txtPorcRec, strDot)
End Sub

Private Sub txtPorcRec_LostFocus()
Dim dblMtoinicial As Double

dblMtoinicial = 0
If Trim(txtPorcRec) <> "" Then
    With Me
        dblMtoinicial = CDbl(IIf(.txtCantidad <> "", txtCantidad, "0")) * CCur(IIf(.txtPreUni <> "", txtPreUni, "0"))
        .txtMtoRec = ValorPorcentaje(CDbl(IIf(.txtPreUni <> "", txtPreUni, "0")), CSng(IIf(.txtPorcRec <> "", txtPorcRec, "0")))
        .txtPreFin = CDbl(IIf(.txtPreUni <> "", txtPreUni, "0")) + CCur(IIf(.txtMtoRec <> "", txtMtoRec, "0"))
        .txtSubTot = CCur(IIf(.txtPreFin <> "", txtPreFin, "0")) * CDbl(IIf(.txtCantidad <> "", txtCantidad, "0"))
    End With
Else
    txtPorcRec = "0"
End If
End Sub

Private Sub txtPreUni_GotFocus()
MarcaTexto txtPreUni
End Sub


Private Sub txtPreUni_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    DownLoadDataTer
    LimpiaCampos
    lblCodigo = TraeIndiceTrabajosTerceros(gstrIdEmpresa, gstrIdSucursal)
End If
If KeyAscii = 27 Then
    Unload Me
End If
KeyAscii = CheckNumber(KeyAscii, txtPreUni, strDot)
End Sub

Private Sub txtPreUni_LostFocus()
'Dim dblMtoInicial As Double
'With Me
'    dblMtoInicial = CDbl(IIf(.txtCantidad <> "", txtCantidad, "0")) * CCur(IIf(.txtPreUni <> "", txtPreUni, "0"))
'    '.txtMtoRec = Format(ValorPorcentaje(dblMtoInicial, CSng(IIf(.txtPorcRec <> "", txtPorcRec, "0"))))
'    .txtMtoRec = ValorPorcentaje(CDbl(IIf(.txtPreUni <> "", txtPreUni, "0")), CSng(IIf(.txtPorcRec <> "", txtPorcRec, "0")))
'    .txtPreFin = Format(CCur(IIf(.txtPreUni <> "", txtPreUni, "0")) + CCur(IIf(.txtMtoRec <> "", txtMtoRec, "0")))
'    .txtSubTot = Format(CCur(IIf(.txtPreFin <> "", txtPreFin, "0")) * CDbl(IIf(.txtCantidad <> "", txtCantidad, "0")))
'End With
End Sub

Private Sub txtProveedor_GotFocus()
MarcaTexto txtProveedor
End Sub

Private Sub txtProveedor_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    DownLoadDataTer
'    LimpiaCampos
'    lblCodigo = TraeIndiceTrabajosTerceros(gstrIdEmpresa, gstrIdSucursal)
'End If
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub txtProveedor_LostFocus()
txtProveedor.Tag = txtProveedor
txtProveedor = ProveedorS(txtProveedor)
If txtProveedor = "" Then
    txtProveedor.Tag = ""
    txtProveedor = ""
End If
End Sub

Private Sub txtSubTot_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    DownLoadDataTer
    LimpiaCampos
    lblCodigo = TraeIndiceTrabajosTerceros(gstrIdEmpresa, gstrIdSucursal)
End If
If KeyAscii = 27 Then
    Unload Me
End If
KeyAscii = CheckNumber(KeyAscii, txtSubTot, strDot)
End Sub

Private Sub txtSubTot_LostFocus()
Dim dblMtoRecargo As Double

    If Me.txtCantidad <> "" And Me.txtCantidad <> "0" And Me.txtPreUni <> "" And Me.txtPreUni <> "0" Then
        dblMtoRecargo = (CCur(IIf(Me.txtSubTot <> "", Me.txtSubTot, "0")) / CDbl(Me.txtCantidad)) - CCur(Me.txtPreUni)
'        Me.txtMtoRec = Format(dblMtoRecargo, "0")
        Me.txtMtoRec = Round(dblMtoRecargo, 2)
'        Me.txtPreFin = Format(CCur(Me.txtPreUni) + CCur(Me.txtMtoRec), "0")
        Me.txtPreFin = Round(CCur(Me.txtPreUni) + CCur(Me.txtMtoRec), 2)
        Me.txtPorcRec = PorcentajeMonto(CDbl(IIf(Me.txtPreUni <> "", Me.txtPreUni, "0")), CSng(IIf(Me.txtMtoRec <> "", Me.txtMtoRec, "0")))
    End If

End Sub

Private Sub txtTipoCargo_GotFocus()
MarcaTexto txtTipoCargo
End Sub

Private Sub txtTipoCargo_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub txtTipoCargo_LostFocus()
txtTipoCargo.Tag = txtTipoCargo
txtTipoCargo = TraeCargoDes(txtTipoCargo)
ValidaCostoCargo
If txtTipoCargo = "" Then
    txtTipoCargo.Tag = ""
    txtTipoCargo = ""
End If
End Sub

Private Sub txtTipoTrabajo_GotFocus()
MarcaTexto txtTipoTrabajo
End Sub

Private Sub txtTipoTrabajo_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
KeyAscii = VerificaLetraCarroceria(KeyAscii)
End Sub
Private Sub ValidaCostoCargo()
Dim lstrCostea As String
Dim dblMtoinicial As Double

If Me.txtTipoCargo <> "" Then
    'trae costo cargo
    lstrCostea = Retorna_Valor_General("Select Costea from Tllr_Tipo_Cargo where Id_Empresa='" & gstrIdEmpresa & "' and id_tipo_Cargo='" & Me.txtTipoCargo.Tag & "'", gcdynamic)
    If lstrCostea = "S" Then
        Me.txtSubTot = CDbl(Me.txtCantidad) * CDbl(Me.txtPreUni)
        Me.txtPorcRec = 0
        Me.txtMtoRec = 0
        Me.txtPorcDcto = 0
        Me.txtMtoDcto = 0
        Me.txtPreFin = Me.txtSubTot
    Else
        Me.txtSubTot = CDbl(Me.txtCantidad) * CDbl(Me.txtPreUni)
        
        'recargo
        dblMtoinicial = 0
        If Trim(txtPorcRec) <> "" Then
            With Me
                dblMtoinicial = CDbl(IIf(.txtCantidad <> "", txtCantidad, "0")) * CCur(IIf(.txtPreUni <> "", txtPreUni, "0"))
                .txtMtoRec = ValorPorcentaje(CDbl(IIf(.txtPreUni <> "", txtPreUni, "0")), CSng(IIf(.txtPorcRec <> "", txtPorcRec, "0")))
                .txtPreFin = CDbl(IIf(.txtPreUni <> "", txtPreUni, "0")) + CCur(IIf(.txtMtoRec <> "", txtMtoRec, "0"))
                .txtSubTot = CCur(IIf(.txtPreFin <> "", txtPreFin, "0")) * CDbl(IIf(.txtCantidad <> "", txtCantidad, "0"))
            End With
        Else
            txtPorcRec = "0"
        End If
        
        'descuento
        dblMtoinicial = 0
        If Trim(txtPorcDcto) <> "" Then
            With Me
                dblMtoinicial = CDbl(IIf(.txtCantidad <> "", txtCantidad, "0")) * CCur(IIf(.txtPreUni <> "", txtPreUni, "0"))
                '.txtMtoRec = ValorPorcentaje(CDbl(IIf(.txtPreUni <> "", txtPreUni, "0")), CSng(IIf(.txtPorcRec <> "", txtPorcRec, "0")))
                .txtPreFin = CDbl(IIf(.txtPreUni <> "", txtPreUni, "0")) + CCur(IIf(.txtMtoRec <> "", txtMtoRec, "0"))
                .txtMtoDcto = Format(ValorPorcentaje(CCur(IIf(.txtPreFin <> "", .txtPreFin, 0)) * CCur(IIf(.txtCantidad <> "", .txtCantidad, 0)), CSng(IIf(.txtPorcDcto <> "", txtPorcDcto, "0"))))
                '.txtSubTot = CCur(IIf(.txtPreFin <> "", txtPreFin, "0")) * CDbl(IIf(.txtCantidad <> "", txtCantidad, "0"))
                .txtSubTot = Format((CCur(IIf(.txtPreFin <> "", txtPreFin, "0")) * CDbl(IIf(.txtCantidad <> "", txtCantidad, "0"))) - CCur(IIf(.txtMtoDcto <> "", .txtMtoDcto, 0)), "", 0)
            End With
        Else
            txtPorcDcto = "0"
        End If
    
    End If
End If
End Sub

