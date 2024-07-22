VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmEditaServicioMecanica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edición Servicio Mecánica"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "frmEditaServicioMecanica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCC 
      Caption         =   "..."
      Height          =   255
      Left            =   6240
      TabIndex        =   26
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox txtCC 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5640
      MaxLength       =   4
      TabIndex        =   24
      Top             =   1395
      Width           =   585
   End
   Begin VB.TextBox txtHorasReales 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1215
      TabIndex        =   9
      Top             =   2520
      Width           =   720
   End
   Begin VB.TextBox txtMecanico 
      Height          =   330
      Left            =   1215
      TabIndex        =   6
      Top             =   1755
      Width           =   3795
   End
   Begin VB.TextBox txtTipoCargo 
      Height          =   330
      Left            =   1215
      TabIndex        =   5
      Top             =   1395
      Width           =   2895
   End
   Begin VB.CommandButton cmdAceptarRep 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   3780
      TabIndex        =   10
      Top             =   2760
      Width           =   800
   End
   Begin VB.CommandButton cmdCancelarRep 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   4575
      TabIndex        =   11
      Top             =   2760
      Width           =   800
   End
   Begin VB.TextBox txtSubTotalMec 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2100
      Width           =   2040
   End
   Begin VB.TextBox txtPrecioUnitarioMec 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2775
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   2
      Top             =   720
      Width           =   1320
   End
   Begin VB.TextBox txtPorcDescMec 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1215
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1065
      Width           =   700
   End
   Begin VB.TextBox txtMtoDescMec 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2790
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1080
      Width           =   1320
   End
   Begin VB.TextBox txtHorasMec 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1215
      MaxLength       =   4
      TabIndex        =   1
      Top             =   720
      Width           =   700
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   4920
      Top             =   -405
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
            Picture         =   "frmEditaServicioMecanica.frx":000C
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":011E
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":0230
            Key             =   "Cerrar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMecanico 
      Height          =   330
      Left            =   5115
      TabIndex        =   21
      Top             =   1740
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imlAux"
      DisabledImageList=   "imlAux"
      HotImageList    =   "imlAux"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BuscarMecanico"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlAux 
      Left            =   3330
      Top             =   -330
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
            Picture         =   "frmEditaServicioMecanica.frx":0342
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":0454
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":0566
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":0678
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":078A
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":089C
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":09AE
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":0AC0
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":0BD2
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":0CE4
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":0DF6
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":0F08
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":101A
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":112C
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":123E
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":1350
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":1462
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":18B4
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioMecanica.frx":1D06
            Key             =   "Copiar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCargo 
      Height          =   330
      Left            =   4455
      TabIndex        =   22
      Top             =   1410
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imlAux"
      DisabledImageList=   "imlAux"
      HotImageList    =   "imlAux"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BuscarCargo"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "CC"
      Height          =   255
      Left            =   5280
      TabIndex        =   25
      Top             =   1425
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Horas Reales"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   315
      Left            =   1215
      TabIndex        =   20
      Top             =   375
      Width           =   4335
   End
   Begin VB.Label lblIDServicioMec 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   315
      Left            =   1215
      TabIndex        =   19
      Top             =   15
      Width           =   1425
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mecánico:"
      Height          =   195
      Index           =   57
      Left            =   150
      TabIndex        =   18
      Top             =   1770
      Width           =   750
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Cargo"
      Height          =   195
      Index           =   58
      Left            =   135
      TabIndex        =   17
      Top             =   1425
      Width           =   780
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total"
      Height          =   195
      Index           =   54
      Left            =   150
      TabIndex        =   16
      Top             =   2130
      Width           =   690
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$ Unitario"
      Height          =   195
      Index           =   53
      Left            =   2025
      TabIndex        =   15
      Top             =   780
      Width           =   675
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$ Desc."
      Height          =   195
      Index           =   55
      Left            =   2040
      TabIndex        =   14
      Top             =   1095
      Width           =   555
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% Desc."
      Height          =   195
      Index           =   56
      Left            =   150
      TabIndex        =   13
      Top             =   1050
      Width           =   585
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      Height          =   195
      Index           =   50
      Left            =   120
      TabIndex        =   12
      Top             =   45
      Width           =   495
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      Height          =   195
      Index           =   51
      Left            =   135
      TabIndex        =   7
      Top             =   390
      Width           =   840
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Horas"
      Height          =   195
      Index           =   52
      Left            =   135
      TabIndex        =   0
      Top             =   735
      Width           =   420
   End
End
Attribute VB_Name = "frmEditaServicioMecanica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnSW As Boolean
Dim dblTotalInicial As Double



Sub UpLoadDataMec()
With frmRecepcion.lvwServiciosMecanica
    lblIDServicioMec.Caption = .SelectedItem
    lblDescripcion.Caption = .SelectedItem.SubItems(1)
    txtHorasMec = SacarFormatoValor(.SelectedItem.SubItems(2), "")
    txtPrecioUnitarioMec = SacarFormatoValor(.SelectedItem.SubItems(3), "")
    txtPorcDescMec = SacarFormatoValor(.SelectedItem.SubItems(4), "")
    txtMtoDescMec = SacarFormatoValor(.SelectedItem.SubItems(5), "")
    'lblTipoCargo.Tag = .SelectedItem.SubItems(6)
    'lblTipoCargo.Caption = .SelectedItem.SubItems(7)
    txtTipoCargo.Tag = .SelectedItem.SubItems(6)
    txtTipoCargo = .SelectedItem.SubItems(7)
    'lblMecanico.Tag = .SelectedItem.SubItems(8)
    'lblMecanico.Caption = .SelectedItem.SubItems(9)
    txtMecanico.Tag = .SelectedItem.SubItems(8)
    txtMecanico = .SelectedItem.SubItems(9)
    txtSubTotalMec = SacarFormatoValor(.SelectedItem.SubItems(10), "")
    txtHorasReales = SacarFormatoValor(.SelectedItem.SubItems(13), "")
End With
End Sub


Sub DownLoadDataMec()
Dim Sql As String
Dim AdoCargo As New ADODB.Recordset
Dim Cargos(1 To 9) As String
Dim ldblCont As Integer
Dim j As Integer
'kjcv 07.04.15
Sql = "SELECT Id_Cargo FROM Tllr_Mecanicos_Cargo WHERE Id_Empresa='" & gstrIdEmpresa & "' and Id_Sucursal='" & gstrIdSucursal & "' and Id_Mecanico='" & gstrIdUsuario & "'"
If Conexion.SendHost(Sql, AdoCargo, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
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

With frmRecepcion.lvwServiciosMecanica
    .SelectedItem.SubItems(2) = FormatoValor(txtHorasMec, "", 1)
    .SelectedItem.SubItems(4) = FormatoValor(txtPorcDescMec, "", 2)
    .SelectedItem.SubItems(3) = FormatoValor(Me.txtPrecioUnitarioMec, "", gintDecimalesMoneda)
    .SelectedItem.SubItems(5) = FormatoValor(txtMtoDescMec, "", gintDecimalesMoneda)
    ' inicio kjcv 07.04.15
        For j = 1 To 6
            If txtTipoCargo.Tag = Cargos(j) Then
                .SelectedItem.SubItems(6) = Trim(txtTipoCargo.Tag)
                .SelectedItem.SubItems(7) = Trim(txtTipoCargo)
            End If
        Next j
    'fin kjcv 07.04.15
    
''    .SelectedItem.SubItems(6) = Trim(txtTipoCargo.Tag)
''    .SelectedItem.SubItems(7) = Trim(txtTipoCargo)
    .SelectedItem.SubItems(8) = Trim(txtMecanico.Tag)
    .SelectedItem.SubItems(9) = Trim(txtMecanico)
    .SelectedItem.SubItems(10) = FormatoValor(txtSubTotalMec, "", gintDecimalesMoneda)
    .SelectedItem.SubItems(11) = "N"
    .SelectedItem.SubItems(13) = FormatoValor(txtHorasReales, "", 2)
    .SelectedItem.SubItems(16) = txtCC
     
     If txtTipoCargo.Tag = "03" And txtCC = "" Then
        frmCentroCosto.Show vbModal
        .SelectedItem.SubItems(16) = gCentroCosto
'    Else
'        .SelectedItem.SubItems(16) = ""
    End If
    
    Unload Me
End With
End Sub

Private Sub cmdAceptarRep_Click()
DownLoadDataMec
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
    UpLoadDataMec
    mblnSW = False
End If
End Sub

Private Sub Form_Load()
mblnSW = True
Me.Label(53).Caption = gstrMonedaLocal & " Unitario"
Me.Label(55).Caption = gstrMonedaLocal & " Desc."
If gstrAsignaRecursos = "S" Then
    Me.txtMecanico.Locked = True
End If
End Sub


Private Sub tlbCargo_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "BuscarCargo" Then
    Me.txtSubTotalMec.SetFocus
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

Private Sub tlbMecanico_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "BuscarMecanico" Then
    If gstrAsignaRecursos = "N" Then
        gstrBusca = apfFormulario.BuscarRegistros(Conexion, "(select * from Tllr_Mecanicos where id_empresa='" & gstrIdEmpresa & "' and id_sucursal='" & gstrIdSucursal & "' And Vigencia='S' AND Es_Recepcionista='N' AND Es_Supervisor='N' ) as Tllr_Mecanicos", "Id_Mecanico", "Nombre", "Buscar Mecánico")
        If gstrBusca <> "" Then
            'lblMecanico.Tag = gstrBusca
            'lblMecanico.Caption = TraeNombreMecanico(gstrBusca)
            txtMecanico.Tag = gstrBusca
            txtMecanico = TraeNombreMecanico(gstrBusca)
        End If
    End If
End If
End Sub

Private Sub txtHorasMec_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtHorasMec, strDot)
End Sub

Private Sub txtHorasMec_LostFocus()
If txtHorasMec <> "" Then
    txtSubTotalMec = Val(txtHorasMec) * Val(IIf(txtPrecioUnitarioMec <> "", txtPrecioUnitarioMec, 0))
End If
End Sub

Private Sub txtHorasReales_GotFocus()
MarcaTexto txtHorasReales
End Sub

Private Sub txtHorasReales_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtHorasReales, strDot)
End Sub

Private Sub txtMecanico_GotFocus()
MarcaTexto txtMecanico
End Sub

Private Sub txtMecanico_LostFocus()
txtMecanico.Tag = txtMecanico
txtMecanico = TraeNombreMecanico(txtMecanico)
If txtMecanico = "" Then
    txtMecanico.Tag = frmRecepcion.lvwServiciosMecanica.SelectedItem.SubItems(8)
    txtMecanico = frmRecepcion.lvwServiciosMecanica.SelectedItem.SubItems(9)
End If
End Sub

Private Sub txtMtoDescMec_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtMtoDescMec, strDot)
End Sub

Private Sub txtMtoDescMec_LostFocus()
If txtHorasMec <> "" And txtMtoDescMec <> "" And txtPrecioUnitarioMec <> "" Then
    dblTotalInicial = Val(txtHorasMec) * Val(txtPrecioUnitarioMec)
    txtPorcDescMec = PorcentajeMonto(dblTotalInicial, CSng(txtMtoDescMec))
    txtSubTotalMec = Val(dblTotalInicial) - Val(txtMtoDescMec)
End If
End Sub

Private Sub txtPorcDescMec_KeyDown(KeyCode As Integer, Shift As Integer)
'If txtPorcDescMec <> "" And txtHorasMec <> "" Then
'    dblTotalInicial = CDbl(txtHorasMec) * CDbl(txtPrecioUnitarioMec)
'    txtMtoDescMec = ValorPorcentaje(dblTotalInicial, CSng(txtPorcDescMec))
'    txtSubTotalMec = dblTotalInicial - CDbl(txtMtoDescMec)
'End If
End Sub

Private Sub txtPorcDescMec_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtPorcDescMec, strDot)
End Sub

Private Sub txtPorcDescMec_LostFocus()
If txtPorcDescMec <> "" And txtHorasMec <> "" Then
    dblTotalInicial = CDbl(txtHorasMec) * CDbl(txtPrecioUnitarioMec)
    txtMtoDescMec = ValorPorcentaje(dblTotalInicial, CSng(txtPorcDescMec))
    txtSubTotalMec = dblTotalInicial - CDbl(txtMtoDescMec)
End If
End Sub

Private Sub txtPrecioUnitarioMec_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtPrecioUnitarioMec, strDot)
End Sub

Private Sub txtTipoCargo_GotFocus()
MarcaTexto txtTipoCargo
End Sub

Private Sub txtTipoCargo_LostFocus()
txtTipoCargo.Tag = txtTipoCargo
txtTipoCargo = TraeCargoDes(txtTipoCargo)
ValidaCostoCargo
If txtTipoCargo = "" Then
    txtTipoCargo.Tag = frmRecepcion.lvwServiciosMecanica.SelectedItem.SubItems(6)
    txtTipoCargo = frmRecepcion.lvwServiciosMecanica.SelectedItem.SubItems(7)
End If
End Sub
Sub ValidaCostoCargo()
Dim lstrCostea As String
Dim lstrSQL As String
Dim recAux As New ADODB.Recordset

If Me.txtTipoCargo <> "" Then
    'trae costo cargo
    lstrCostea = Retorna_Valor_General("Select Costea from Tllr_Tipo_Cargo where Id_Empresa='" & gstrIdEmpresa & "' and id_tipo_Cargo='" & Me.txtTipoCargo.Tag & "'", gcdynamic)
    If lstrCostea = "S" Then
        
        If gblnPreciosMarca = True Then
            'trae costo de hora por marca
            lstrSQL = "SELECT CostoManoObra, CostoMOGarantia From Tllr_Marca_Precios_MO WHERE (Id_Marca = '" & frmRecepcion.lblIdMarca & "')"
            If Conexion.SendHost(lstrSQL, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                If Not recAux.BOF And Not recAux.EOF Then
                    Me.txtPrecioUnitarioMec = IIf(txtTipoCargo.Tag = gstrCargoGtiaFabrica, recAux!CostoMOGarantia, recAux!CostoManoObra)
                End If
            End If
        Else
            Me.txtPrecioUnitarioMec = gcurCostoManoObra
        End If
        
        Me.txtSubTotalMec = CDbl(Me.txtHorasMec) * CDbl(Me.txtPrecioUnitarioMec)
        If txtPorcDescMec <> "" And txtHorasMec <> "" Then
            dblTotalInicial = CDbl(txtHorasMec) * CDbl(txtPrecioUnitarioMec)
            txtMtoDescMec = ValorPorcentaje(dblTotalInicial, CSng(txtPorcDescMec))
            txtSubTotalMec = dblTotalInicial - CDbl(txtMtoDescMec)
        Else
            Me.txtSubTotalMec = CDbl(txtPrecioUnitarioMec) * CDbl(Me.txtHorasMec)
            Me.txtPorcDescMec = 0
            Me.txtMtoDescMec = 0
        End If
    Else
        
        If gblnPreciosMarca = True Then
            'trae costo de hora por marca
            lstrSQL = "SELECT VentaManoObra, VentaMOGarantia From Tllr_Marca_Precios_MO WHERE (Id_Marca = '" & frmRecepcion.lblIdMarca & "')"
            If Conexion.SendHost(lstrSQL, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                If Not recAux.BOF And Not recAux.EOF Then
                    Me.txtPrecioUnitarioMec = IIf(txtTipoCargo.Tag = gstrCargoGtiaFabrica, recAux!VentaMOGarantia, recAux!VentaManoObra)
                End If
            End If
        Else
'            Me.txtPrecioUnitarioMec = IIf(txtTipoCargo.Tag = gstrCargoGtiaFabrica, Retorna_Valor_General("Select PrecioManoOBraGarantia from Tllr_Parametro Where id_empresa='" & gstrIdEmpresa & "' And id_sucursal='" & gstrIdSucursal & "'", gcdynamic), gcurPrecioManoObra)
            'kjcv 09.05.20
            Me.txtPrecioUnitarioMec = IIf(txtTipoCargo.Tag = gstrCargoGtiaFabrica, Retorna_Valor_General("Select ValorMOGarantia from Tllr_mo where id_empresa='" & gstrIdEmpresa & "' and Id_Marca = '" & frmRecepcion.lblIdMarca & "'", gcdynamic), gcurPrecioManoObra)
        End If
    
        Me.txtSubTotalMec = CDbl(Me.txtHorasMec) * CDbl(Me.txtPrecioUnitarioMec)
        If txtPorcDescMec <> "" And txtHorasMec <> "" Then
            dblTotalInicial = CDbl(txtHorasMec) * CDbl(txtPrecioUnitarioMec)
            txtMtoDescMec = ValorPorcentaje(dblTotalInicial, CSng(txtPorcDescMec))
            txtSubTotalMec = dblTotalInicial - CDbl(txtMtoDescMec)
        End If
    End If
End If
End Sub
