VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmEditaServicioTercero 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edición Servicio de Tercero"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "frmEditaServicioTercero.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptarRep 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   315
      Left            =   4785
      TabIndex        =   11
      Top             =   3375
      Width           =   800
   End
   Begin VB.CommandButton cmdCancelarRep 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   5580
      TabIndex        =   12
      Top             =   3375
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   45
      TabIndex        =   13
      Top             =   0
      Width           =   6390
      Begin VB.CommandButton cmdCC 
         Caption         =   "..."
         Height          =   255
         Left            =   5880
         TabIndex        =   34
         Top             =   2900
         Width           =   375
      End
      Begin VB.TextBox txtCC 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   32
         Top             =   2880
         Width           =   585
      End
      Begin VB.TextBox txtProveedor 
         Height          =   330
         Left            =   1035
         TabIndex        =   31
         ToolTipText     =   "Puede Ingresar el Rut del Proveedor"
         Top             =   135
         Width           =   4155
      End
      Begin VB.TextBox txtTipoCargo 
         Height          =   330
         Left            =   1020
         TabIndex        =   10
         Top             =   2880
         Width           =   2850
      End
      Begin VB.TextBox txtMtoDcto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3750
         MaxLength       =   8
         TabIndex        =   8
         Top             =   1860
         Width           =   1320
      End
      Begin VB.TextBox txtPorcDcto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1020
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1845
         Width           =   795
      End
      Begin VB.TextBox txtFactura 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1020
         MaxLength       =   10
         TabIndex        =   9
         Top             =   2520
         Width           =   1170
      End
      Begin VB.TextBox txtSubTot 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3750
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2205
         Width           =   1320
      End
      Begin VB.TextBox txtPreFin 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2190
         Width           =   1170
      End
      Begin VB.TextBox txtPorcRec 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1020
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1500
         Width           =   795
      End
      Begin VB.TextBox txtMtoRec 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3750
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1515
         Width           =   1320
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   1020
         MaxLength       =   70
         TabIndex        =   0
         Top             =   840
         Width           =   4875
      End
      Begin VB.TextBox txtPreUni 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3750
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1185
         Width           =   1320
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1170
         Width           =   795
      End
      Begin MSComctlLib.Toolbar tlbCargo 
         Height          =   330
         Left            =   3960
         TabIndex        =   27
         Top             =   2865
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlAux"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Busca Cargo"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbProv 
         Height          =   330
         Left            =   5295
         TabIndex        =   28
         Top             =   165
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlAux"
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
      Begin VB.Label Label2 
         Caption         =   "CC"
         Height          =   255
         Left            =   4920
         TabIndex        =   33
         Top             =   2910
         Width           =   255
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Dscto    :"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   30
         Top             =   1860
         Width           =   810
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$ Dscto.    :"
         Height          =   195
         Index           =   0
         Left            =   2535
         TabIndex        =   29
         Top             =   1875
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Factura Nº :"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   26
         Top             =   2565
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SubTotal   :"
         Height          =   195
         Index           =   6
         Left            =   2535
         TabIndex        =   25
         Top             =   2220
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Precio Final  :"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   2925
         Width           =   870
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$ Recargo :"
         Height          =   195
         Index           =   55
         Left            =   2535
         TabIndex        =   22
         Top             =   1530
         Width           =   1080
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Recargo :"
         Height          =   195
         Index           =   56
         Left            =   90
         TabIndex        =   21
         Top             =   1515
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor  :"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   20
         Top             =   210
         Width           =   870
      End
      Begin VB.Label lblProveedor 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2745
         TabIndex        =   19
         Top             =   495
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo       :"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   18
         Top             =   540
         Width           =   855
      End
      Begin VB.Label lblCodigo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1020
         TabIndex        =   17
         Top             =   510
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   16
         Top             =   885
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "$ Unitario  :"
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   15
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad    :"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   14
         Top             =   1230
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   5955
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":000C
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":011E
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":0230
            Key             =   "Cerrar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlAux 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":0342
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":0454
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":0566
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":0678
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":078A
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":089C
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":09AE
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":0AC0
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":0BD2
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":0CE4
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":0DF6
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":0F08
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":101A
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":112C
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":123E
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":1350
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":1462
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":18B4
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioTercero.frx":1D06
            Key             =   "Copiar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEditaServicioTercero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnSW As Boolean
Dim dblTotalInicial As Double
Dim mstrnombre As String

Sub ReCalculoTer()
dblTotalInicial = Val(Format(txtCantidad, "#####0.0")) * Val(Format(txtPreFin, "#####0.0"))
txtSubTot = Format(dblTotalInicial, "###,##0.0")
If Val(txtPorcRec) > 0 Then
    txtMtoRec = Format(ValorPorcentaje(dblTotalInicial, CSng(Format(Me.txtPorcRec, "#0.0"))), "###,##0.0")
    txtSubTot = Format(dblTotalInicial - CDbl(Format(txtMtoRec, "###,##0.0")), "###,##0.0")
    Exit Sub
End If
If Val(txtMtoRec) > 0 Then
    txtPorcRec = Format(PorcentajeMonto(dblTotalInicial, CSng(Format(txtMtoRec, "#####0.0"))), "###,##0.0")
    txtSubTot = Format(dblTotalInicial - CDbl(Format(txtMtoRec, "###,##0.0")), "###,##0.0")
    Exit Sub
End If
End Sub

Sub UpLoadDataTer()
With frmRecepcion.lvwServiciosTerceros
    lblCodigo = .SelectedItem
    txtDescripcion = .SelectedItem.SubItems(3)
    'lblProveedor.Caption = .SelectedItem.SubItems(1)
    'lblProveedor.Tag = .SelectedItem.SubItems(2)
    txtProveedor = .SelectedItem.SubItems(1)
    txtProveedor.Tag = .SelectedItem.SubItems(2)
    txtFactura = .SelectedItem.SubItems(4)
    txtPreUni = SacarFormatoValor(.SelectedItem.SubItems(5), "")
    txtCantidad = SacarFormatoValor(.SelectedItem.SubItems(6), "")
    txtPorcRec = SacarFormatoValor(.SelectedItem.SubItems(7), "")
    txtMtoRec = SacarFormatoValor(.SelectedItem.SubItems(8), "")
    txtPreFin = SacarFormatoValor(.SelectedItem.SubItems(9), "")
    txtPorcDcto = SacarFormatoValor(.SelectedItem.SubItems(10), "")
    txtMtoDcto = SacarFormatoValor(.SelectedItem.SubItems(11), "")
    txtSubTot = SacarFormatoValor(.SelectedItem.SubItems(12), "")
    'lblTipoCargo.Caption = .SelectedItem.SubItems(13)
    'lblTipoCargo.Tag = .SelectedItem.SubItems(14)
    txtTipoCargo = .SelectedItem.SubItems(13)
    txtTipoCargo.Tag = .SelectedItem.SubItems(14)
    'kjcv 14.09.18
    txtCC.Tag = .SelectedItem.SubItems(16)
    txtCC = .SelectedItem.SubItems(16)
End With
End Sub

Sub DownLoadDataTer()
Dim sql As String
Dim AdoCargo As New ADODB.Recordset
Dim Cargos(1 To 9) As String
Dim ldblCont As Integer
Dim j As Integer
'kjcv 07.04.15
sql = "SELECT Id_Cargo FROM Tllr_Mecanicos_Cargo WHERE Id_Empresa='" & gstrIdEmpresa & "' and Id_Sucursal='" & gstrIdSucursal & "' and Id_Mecanico='" & gstrIdUsuario & "'"
If Conexion.SendHost(sql, AdoCargo, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If AdoCargo.EOF = False And AdoCargo.BOF = False Then
        ldblCont = 1
        AdoCargo.MoveFirst
        While AdoCargo.EOF = False
            Cargos(ldblCont) = ValorNulo(AdoCargo.Fields("Id_Cargo"))
            ldblCont = ldblCont + 1
            AdoCargo.MoveNext
        Wend
    End If
End If
Conexion.CloseHost AdoCargo

With frmRecepcion.lvwServiciosTerceros
    Set glsiItem = .SelectedItem
    'glsiItem.SubItems(1) = IIf(lblProveedor.Caption <> "", lblProveedor.Caption, "S/PROVEEDOR")
    'glsiItem.SubItems(2) = IIf(lblProveedor.Tag <> "", lblProveedor.Tag, "00")
    glsiItem.SubItems(1) = IIf(txtProveedor <> "", txtProveedor, "S/PROVEEDOR")
    glsiItem.SubItems(2) = IIf(txtProveedor.Tag <> "", txtProveedor.Tag, "19")
    glsiItem.SubItems(3) = IIf(txtDescripcion <> "", UCase(txtDescripcion), ".")
    glsiItem.SubItems(4) = IIf(txtFactura <> "", txtFactura, "S/factura")
    glsiItem.SubItems(5) = FormatoValor(txtPreUni, "", gintDecimalesMoneda)
    glsiItem.SubItems(6) = FormatoValor(txtCantidad, "", 1)
    glsiItem.SubItems(7) = FormatoValor(txtPorcRec, "", 2)
    glsiItem.SubItems(8) = FormatoValor(txtMtoRec, "", gintDecimalesMoneda)
    glsiItem.SubItems(9) = FormatoValor(txtPreFin, "", gintDecimalesMoneda)
    glsiItem.SubItems(10) = FormatoValor(txtPorcDcto, "", 2)
    glsiItem.SubItems(11) = FormatoValor(txtMtoDcto, "", gintDecimalesMoneda)
    glsiItem.SubItems(12) = FormatoValor(txtSubTot, "", gintDecimalesMoneda)
     ' inicio kjcv 07.04.15
        For j = 1 To 6
            If txtTipoCargo.Tag = Cargos(j) Then
                glsiItem.SubItems(13) = txtTipoCargo
                glsiItem.SubItems(14) = txtTipoCargo.Tag
                glsiItem.SubItems(15) = "N"
            End If
        Next j
    'kjcv 17.09.18
    If txtTipoCargo.Tag = "03" Then
        glsiItem.SubItems(16) = txtCC
    Else
        txtCC = ""
    End If
    
    If txtTipoCargo.Tag = "03" And txtCC = "" Then
        frmCentroCosto.Show vbModal
         frmRecepcion.lvwServiciosTerceros.SelectedItem.SubItems(16) = gCentroCosto
'    Else
'        frmRecepcion.lvwServiciosTerceros.SelectedItem.SubItems(16) = ""
    End If
    
    
End With
End Sub

Private Sub cmdAceptarRep_Click()
    If Me.txtProveedor = "S/PROVEEDOR" Or Me.txtProveedor = "" Then
        If frmRecepcion.dtcGarantia.BoundText <> "PRE" Then
            MsgBox "Debe Ingresar el Proveedor", vbExclamation, "Terceros"
            txtProveedor.SetFocus
            Exit Sub
        End If
    End If
DownLoadDataTer
Unload Me
End Sub

Private Sub cmdCancelarRep_Click()
Unload Me
End Sub

Private Sub cmdCC_Click()
frmCentroCosto.Show vbModal
Me.txtCC = gCentroCosto
End Sub

Private Sub Form_Activate()
If mblnSW Then
    UpLoadDataTer
    mblnSW = False
End If
End Sub

Private Sub Form_Load()
mblnSW = True

Me.Label1(2).Caption = gstrMonedaLocal & " Unitario"
Me.Label(55).Caption = gstrMonedaLocal & " Recargo"
Me.Label(0).Caption = gstrMonedaLocal & " Dscto."
End Sub

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Confirmar"
'    ReCalculoTer
    DownLoadDataTer
Case "Cancelar"
    UpLoadDataTer
Case "Cerrar"
    Unload Me
End Select
End Sub




Private Sub tlbCargo_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
'kjcv 24.03.20
    gstrBusca = ""
    frmTipoCargo.Show vbModal
'    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Tllr_Tipo_Cargo", "Id_Tipo_cargo", "Descripcion", "Buscar Cargo OT")
    If gstrBusca <> "" Then
        'lblTipoCargo.Tag = gstrBusca
        'lblTipoCargo.Caption = TraeCargoDes(gstrBusca)
        txtTipoCargo.Tag = gstrBusca
        txtTipoCargo = TraeCargoDes(gstrBusca)
        ValidaCostoCargo
    End If
End If
End Sub

Private Sub tlbProv_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    Me.txtDescripcion.SetFocus
    apfFormulario.BuscarRegistroClientes Conexion, gstrBusca, mstrnombre, gstrIdEmpresa
    'apfFormulario.BuscarRegistroClientes Conexion, gstrBusca, mstrnombre
    'lblProveedor.Tag = gstrBusca
    'lblProveedor.Caption = mstrNombre
    
    txtProveedor.Tag = gstrBusca
    txtProveedor = mstrnombre
   
   ' gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Tllr_Proveedor_Servicio", "Id_Proveedor", "Nombre", "Buscar Proveedor de Servicio")
   ' lblProveedor.Tag = gstrBusca
   ' lblProveedor.Caption = ProveedorS(gstrBusca)
End If
End Sub

Private Sub txtCantidad_GotFocus()
MarcaTexto txtCantidad
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtCantidad, strDot)
If KeyAscii = 13 Then
    DownLoadDataTer
    Unload Me
End If
End Sub


Private Sub txtCantidad_LostFocus()
'If Trim(txtCantidad) <> "" Then
'    txtSubTot = CCur(IIf(txtPreFin <> "", txtPreFin, "0")) * CDbl(IIf(txtCantidad <> "", txtCantidad, "0"))
'Else
'    txtSubTot = "0"
'End If

End Sub


Private Sub txtdescripcion_GotFocus()
MarcaTexto txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
If KeyAscii = 13 Then
    DownLoadDataTer
    Unload Me
End If
End Sub

Private Sub txtFactura_GotFocus()
MarcaTexto txtFactura

End Sub

Private Sub txtFactura_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'kjcv 19.01.16
txtFactura = Format(Trim$(txtFactura), "0000000")
    DownLoadDataTer
    Unload Me
End If
End Sub

Private Sub txtFactura_LostFocus()
'kjcv 19.01.16
txtFactura = Format(Trim$(txtFactura), "0000000")
End Sub

Private Sub txtMtoDcto_KeyPress(KeyAscii As Integer)
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

Private Sub txtMtoRec_GotFocus()
MarcaTexto txtMtoRec

End Sub

Private Sub txtMtoRec_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtMtoRec, strDot)
If KeyAscii = 13 Then
    DownLoadDataTer
    Unload Me
End If
End Sub


Private Sub txtMtoRec_LostFocus()
'Dim dblMtoinicial As Double
'
'dblMtoinicial = 0
'If Trim(txtMtoRec) <> "" Then
'    With Me
'        dblMtoinicial = CDbl(IIf(.txtCantidad <> "", txtCantidad, "0")) * CCur(IIf(.txtPreUni <> "", txtPreUni, "0"))
'        .txtPorcRec = PorcentajeMonto(CDbl(IIf(.txtPreUni <> "", txtPreUni, "0")), CSng(IIf(.txtMtoRec <> "", txtMtoRec, "0")))
'        .txtPreFin = CDbl(IIf(.txtPreUni <> "", txtPreUni, "0")) + CCur(IIf(.txtMtoRec <> "", txtMtoRec, "0"))
'        .txtSubTot = CCur(IIf(.txtPreFin <> "", txtPreFin, "0")) * CDbl(IIf(.txtCantidad <> "", txtCantidad, "0"))
'    End With
'End If

End Sub

Private Sub txtPorcDcto_KeyPress(KeyAscii As Integer)
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
        .txtMtoDcto = Format(ValorPorcentaje(CCur(IIf(.txtPreFin <> "", .txtPreFin, 0)) * CCur(IIf(.txtCantidad <> "", .txtCantidad, 0)), CSng(IIf(.txtPorcDcto <> "", txtPorcDcto, "0"))))
        '.txtSubTot = CCur(IIf(.txtPreFin <> "", txtPreFin, "0")) * CDbl(IIf(.txtCantidad <> "", txtCantidad, "0"))
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
KeyAscii = CheckNumber(KeyAscii, txtPorcRec, strDot)
If KeyAscii = 13 Then
    DownLoadDataTer
    Unload Me
End If
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

Private Sub txtPreFin_GotFocus()
MarcaTexto txtPreFin
End Sub

Private Sub txtPreFin_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtPreFin, strDot)
End Sub


Private Sub txtPreUni_Change()
'Dim dblMtoInicial As Double
'With Me
'    dblMtoInicial = CDbl(IIf(.txtCantidad <> "", txtCantidad, "0")) * CCur(IIf(.txtPreUni <> "", txtPreUni, "0"))
'    .txtMtoRec = Format(ValorPorcentaje(dblMtoInicial, CSng(IIf(.txtPorcRec <> "", txtPorcRec, "0"))))
'    .txtPreFin = Format(CCur(IIf(.txtPreUni <> "", txtPreUni, "0")) + CCur(IIf(.txtMtoRec <> "", txtMtoRec, "0")))
'    .txtSubTot = Format(CCur(IIf(.txtPreFin <> "", txtPreFin, "0")) * CDbl(IIf(.txtCantidad <> "", txtCantidad, "0")))
'End With
End Sub

Private Sub txtPreUni_GotFocus()
MarcaTexto txtPreUni

End Sub

Private Sub txtPreUni_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtPreUni, strDot)
If KeyAscii = 13 Then
    DownLoadDataTer
    Unload Me
End If
End Sub


Private Sub txtProveedor_GotFocus()
MarcaTexto txtProveedor
End Sub

Private Sub txtProveedor_LostFocus()
txtProveedor.Tag = txtProveedor
txtProveedor = ProveedorS(txtProveedor)
If txtProveedor = "" Then
    txtProveedor.Tag = ""
    txtProveedor = ""
End If
End Sub

Private Sub txtSubTot_GotFocus()
MarcaTexto txtSubTot
End Sub

Private Sub txtSubTot_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtSubTot, strDot)
If KeyAscii = 13 Then
    DownLoadDataTer
    Unload Me
End If
End Sub


Private Sub txtSubTot_LostFocus()
Dim dblMtoRecargo As Double

    If Me.txtCantidad <> "" And Me.txtCantidad <> "0" And Me.txtPreUni <> "" And Me.txtPreUni <> "0" Then
        dblMtoRecargo = (CCur(IIf(Me.txtSubTot <> "", Me.txtSubTot, "0")) / CDbl(Me.txtCantidad)) - CCur(Me.txtPreUni)
        Me.txtMtoRec = Format(dblMtoRecargo, "0")
        Me.txtPreFin = Format(CCur(Me.txtPreUni) + CCur(Me.txtMtoRec), "0")
        Me.txtPorcRec = PorcentajeMonto(CDbl(IIf(Me.txtPreUni <> "", Me.txtPreUni, "0")), CSng(IIf(Me.txtMtoRec <> "", Me.txtMtoRec, "0")))
    End If

End Sub

Private Sub txtTipoCargo_GotFocus()
MarcaTexto txtTipoCargo
End Sub

Private Sub txtTipoCargo_LostFocus()
txtTipoCargo.Tag = txtTipoCargo
txtTipoCargo = TraeCargoDes(txtTipoCargo)
ValidaCostoCargo
If txtTipoCargo = "" Then
    txtTipoCargo.Tag = frmRecepcion.lvwServiciosTerceros.SelectedItem.SubItems(14)
    txtTipoCargo = frmRecepcion.lvwServiciosTerceros.SelectedItem.SubItems(13)
End If
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

