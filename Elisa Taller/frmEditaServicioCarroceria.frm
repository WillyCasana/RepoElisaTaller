VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmEditaServicioCarroceria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edición Servicio Carrocería"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   Icon            =   "frmEditaServicioCarroceria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCC 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   5520
      MaxLength       =   4
      TabIndex        =   27
      Top             =   1800
      Width           =   585
   End
   Begin VB.CommandButton cmdCC 
      Caption         =   "..."
      Height          =   255
      Left            =   6120
      TabIndex        =   26
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmdAceptarRep 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   4890
      TabIndex        =   23
      Top             =   2880
      Width           =   800
   End
   Begin VB.CommandButton cmdCancelarRep 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   5685
      TabIndex        =   22
      Top             =   2880
      Width           =   800
   End
   Begin VB.TextBox txtTipo 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5115
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   16
      Top             =   45
      Width           =   375
   End
   Begin VB.TextBox txtSubTotalCar 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2520
      Width           =   2040
   End
   Begin VB.TextBox txtPrecioUnitarioCar 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2730
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   13
      Top             =   765
      Width           =   1320
   End
   Begin VB.TextBox txtPorcDescCar 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1155
      MaxLength       =   4
      TabIndex        =   12
      Top             =   1110
      Width           =   700
   End
   Begin VB.TextBox txtMtoDescCar 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2730
      MaxLength       =   8
      TabIndex        =   11
      Top             =   1110
      Width           =   1320
   End
   Begin VB.TextBox txtHorasCar 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1155
      MaxLength       =   4
      TabIndex        =   8
      Top             =   765
      Width           =   700
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
            Picture         =   "frmEditaServicioCarroceria.frx":000C
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":011E
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":0230
            Key             =   "Cerrar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4815
      Top             =   810
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
            Picture         =   "frmEditaServicioCarroceria.frx":0342
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":0454
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":0566
            Key             =   "Cerrar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMecanico 
      Height          =   330
      Left            =   5040
      TabIndex        =   24
      Top             =   2160
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
      Left            =   4095
      Top             =   810
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
            Picture         =   "frmEditaServicioCarroceria.frx":0678
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":078A
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":089C
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":09AE
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":0AC0
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":0BD2
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":0CE4
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":0DF6
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":0F08
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":101A
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":112C
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":123E
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":1350
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":1462
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":1574
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":1686
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":1798
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":1BEA
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioCarroceria.frx":203C
            Key             =   "Copiar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCargo 
      Height          =   330
      Left            =   4140
      TabIndex        =   25
      Top             =   1770
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
      Left            =   5160
      TabIndex        =   28
      Top             =   1830
      Width           =   255
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$ Def"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   21
      Top             =   1440
      Width           =   630
   End
   Begin VB.Label lblDef 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1155
      TabIndex        =   20
      Top             =   1455
      Width           =   1350
   End
   Begin VB.Label lblTipoCargo 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1155
      TabIndex        =   19
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label lblMecanico 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1155
      TabIndex        =   18
      Top             =   2145
      Width           =   3855
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   195
      Index           =   0
      Left            =   4740
      TabIndex        =   17
      Top             =   75
      Width           =   315
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto"
      Height          =   195
      Index           =   24
      Left            =   75
      TabIndex        =   15
      Top             =   90
      Width           =   690
   End
   Begin VB.Label lblPartePieza 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1155
      TabIndex        =   10
      Top             =   420
      Width           =   4335
   End
   Begin VB.Label lblConcepto 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1155
      TabIndex        =   9
      Top             =   60
      Width           =   3480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mecánico:"
      Height          =   195
      Index           =   57
      Left            =   105
      TabIndex        =   7
      Top             =   2190
      Width           =   750
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Cargo"
      Height          =   195
      Index           =   58
      Left            =   90
      TabIndex        =   6
      Top             =   1815
      Width           =   780
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total"
      Height          =   195
      Index           =   54
      Left            =   105
      TabIndex        =   5
      Top             =   2505
      Width           =   690
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$ Unitario"
      Height          =   195
      Index           =   53
      Left            =   1860
      TabIndex        =   4
      Top             =   825
      Width           =   795
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$ Desc."
      Height          =   195
      Index           =   55
      Left            =   1875
      TabIndex        =   3
      Top             =   1140
      Width           =   675
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% Desc."
      Height          =   195
      Index           =   56
      Left            =   105
      TabIndex        =   2
      Top             =   1095
      Width           =   585
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parte / Pieza"
      Height          =   195
      Index           =   51
      Left            =   90
      TabIndex        =   1
      Top             =   435
      Width           =   930
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Horas"
      Height          =   195
      Index           =   52
      Left            =   90
      TabIndex        =   0
      Top             =   780
      Width           =   420
   End
End
Attribute VB_Name = "frmEditaServicioCarroceria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnSW As Boolean
Dim dblTotalInicial As Double



Sub UpLoadDataCar()
With frmRecepcion.lvwServiciosCarroceria
    lblConcepto.Caption = .SelectedItem
    lblConcepto.Tag = .SelectedItem.SubItems(1)
    txtTipo = .SelectedItem.SubItems(2)
    lblPartePieza.Caption = .SelectedItem.SubItems(3)
    lblPartePieza.Tag = .SelectedItem.SubItems(4)
    txtHorasCar = SacarFormatoValor(.SelectedItem.SubItems(5), "")
    txtPrecioUnitarioCar = SacarFormatoValor(.SelectedItem.SubItems(6), "")
    txtPorcDescCar = SacarFormatoValor(.SelectedItem.SubItems(7), "")
    txtMtoDescCar = SacarFormatoValor(.SelectedItem.SubItems(8), "")
    lblDef.Caption = SacarFormatoValor(.SelectedItem.SubItems(9), "")
    lblTipoCargo.Tag = .SelectedItem.SubItems(11) 'SacarFormatoValor(.SelectedItem.SubItems(10), "")
    lblTipoCargo.Caption = .SelectedItem.SubItems(10) '.SelectedItem.SubItems(11)
    lblMecanico.Tag = .SelectedItem.SubItems(13)
    lblMecanico.Caption = .SelectedItem.SubItems(12)
    txtSubTotalCar = SacarFormatoValor(.SelectedItem.SubItems(14), "")
End With
End Sub

Sub DownLoadDataCar()
With frmRecepcion.lvwServiciosCarroceria
    .SelectedItem.SubItems(5) = FormatoValor(txtHorasCar, "", 1)
    .SelectedItem.SubItems(7) = FormatoValor(txtPorcDescCar, "", 2)
    .SelectedItem.SubItems(8) = FormatoValor(txtMtoDescCar, "", gintDecimalesMoneda)
    .SelectedItem.SubItems(9) = FormatoValor(txtSubTotalCar, "", gintDecimalesMoneda)
    .SelectedItem.SubItems(10) = lblTipoCargo.Caption
    .SelectedItem.SubItems(11) = lblTipoCargo.Tag
    .SelectedItem.SubItems(12) = lblMecanico.Caption
    .SelectedItem.SubItems(13) = lblMecanico.Tag
    .SelectedItem.SubItems(14) = FormatoValor(txtSubTotalCar, "", gintDecimalesMoneda)
    .SelectedItem.SubItems(15) = "N"
    .SelectedItem.SubItems(19) = txtCC
 
    
    If lblTipoCargo.Tag = "03" Then
        frmCentroCosto.Show vbModal
        .SelectedItem.SubItems(19) = gCentroCosto
'    Else
'        .SelectedItem.SubItems(19) = ""
    End If
    
    Unload Me
End With
End Sub

Private Sub cmdAceptarRep_Click()
DownLoadDataCar
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
    UpLoadDataCar
    mblnSW = False
End If
End Sub

Private Sub Form_Load()
mblnSW = True
Me.Label(53).Caption = gstrMonedaLocal & " Unitario"
Me.Label(55).Caption = gstrMonedaLocal & " Desc."
Me.Label(1).Caption = gstrMonedaLocal & " Def."
End Sub

Private Sub tlbCargo_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "BuscarCargo" Then
'kjcv 24.03.20
    gstrBusca = ""
    frmTipoCargo.Show vbModal
'    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Tllr_Tipo_Cargo", "Id_Tipo_cargo", "Descripcion", "Buscar Cargo OT")
    If gstrBusca <> "" Then
        lblTipoCargo.Tag = gstrBusca
        lblTipoCargo.Caption = TraeCargoDes(gstrBusca)
    End If
End If
End Sub

Private Sub tlbMecanico_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "BuscarMecanico" Then
    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Tllr_Mecanicos", "Id_Mecanico", "Nombre", "Buscar Mecánico")
    If gstrBusca <> "" Then
        lblMecanico.Tag = gstrBusca
        lblMecanico.Caption = TraeNombreMecanico(gstrBusca)
    End If
End If

End Sub

Private Sub txtHorasCar_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtHorasCar, strDot)
End Sub


Private Sub txtHorasCar_LostFocus()
If txtHorasCar <> "" Then
    txtSubTotalCar = Val(txtHorasCar) * Val(IIf(txtPrecioUnitarioCar <> "", txtPrecioUnitarioCar, 0))
End If
End Sub

Private Sub txtMtoDescCar_LostFocus()
If txtMtoDescCar <> "" Then
    dblTotalInicial = Val(txtHorasCar) * Val(txtPrecioUnitarioCar)
    txtPorcDescCar = PorcentajeMonto(dblTotalInicial, CSng(txtMtoDescCar))
    txtSubTotalCar = Val(dblTotalInicial) - Val(txtMtoDescCar)
End If
End Sub


Private Sub txtPorcDescCar_LostFocus()
If txtPorcDescCar <> "" Then
    dblTotalInicial = Val(txtHorasCar) * Val(txtPrecioUnitarioCar)
    txtMtoDescCar = ValorPorcentaje(dblTotalInicial, CSng(txtPorcDescCar))
    txtSubTotalCar = dblTotalInicial - CDbl(txtMtoDescCar)
End If
End Sub


