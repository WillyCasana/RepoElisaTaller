VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEditaOtroServicio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edición Otros Servicios"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmEditaOtroServicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHorasAsignadas 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3870
      TabIndex        =   30
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdCC 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   6240
      TabIndex        =   29
      Top             =   1470
      Width           =   375
   End
   Begin VB.TextBox txtCC 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   5640
      MaxLength       =   4
      TabIndex        =   27
      Top             =   1440
      Width           =   585
   End
   Begin VB.TextBox txtHorasReales 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1470
      TabIndex        =   8
      Top             =   3135
      Width           =   735
   End
   Begin VB.TextBox txtTipoCargoOtr 
      Height          =   330
      Left            =   1440
      TabIndex        =   5
      Top             =   1485
      Width           =   2850
   End
   Begin VB.TextBox txtMecanico 
      Height          =   330
      Left            =   1470
      TabIndex        =   6
      Top             =   1845
      Width           =   3885
   End
   Begin VB.TextBox lblDescripcion 
      Height          =   315
      Left            =   1470
      TabIndex        =   0
      Top             =   405
      Width           =   4785
   End
   Begin VB.CommandButton cmdAceptarRep 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   315
      Left            =   5010
      TabIndex        =   9
      Top             =   3360
      Width           =   800
   End
   Begin VB.CommandButton cmdCancelarRep 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   5805
      TabIndex        =   11
      Top             =   3360
      Width           =   800
   End
   Begin VB.TextBox txtSubTotalOtr 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1470
      TabIndex        =   7
      Top             =   2745
      Width           =   2040
   End
   Begin VB.TextBox txtPrecioUnitarioOtr 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3045
      MaxLength       =   8
      TabIndex        =   2
      Top             =   750
      Width           =   1320
   End
   Begin VB.TextBox txtPorcDescOtr 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1470
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1140
      Width           =   700
   End
   Begin VB.TextBox txtMtoDescOtr 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3030
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1140
      Width           =   1320
   End
   Begin VB.TextBox txtHorasOtr 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1470
      MaxLength       =   4
      TabIndex        =   1
      Top             =   750
      Width           =   700
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   6360
      Top             =   0
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
            Picture         =   "frmEditaOtroServicio.frx":000C
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":011E
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":0230
            Key             =   "Cerrar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCargo 
      Height          =   330
      Left            =   4425
      TabIndex        =   19
      Top             =   1515
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
            Key             =   "BuscarCargo"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlAux 
      Left            =   4785
      Top             =   -555
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
            Picture         =   "frmEditaOtroServicio.frx":0342
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":0454
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":0566
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":0678
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":078A
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":089C
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":09AE
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":0AC0
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":0BD2
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":0CE4
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":0DF6
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":0F08
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":101A
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":112C
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":123E
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":1350
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":1462
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":18B4
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaOtroServicio.frx":1D06
            Key             =   "Copiar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMecanico 
      Height          =   330
      Left            =   5460
      TabIndex        =   20
      Top             =   1860
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
   Begin MSComCtl2.DTPicker dtpAsigna 
      Height          =   315
      HelpContextID   =   285
      Left            =   1560
      TabIndex        =   25
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   144834561
      CurrentDate     =   42991
   End
   Begin VB.Label Label3 
      Caption         =   "Hrs. Asignadas"
      Height          =   255
      Left            =   2520
      TabIndex        =   31
      Top             =   3180
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "CC"
      Height          =   255
      Left            =   5280
      TabIndex        =   28
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Asignación"
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Hrs. Reales"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3195
      Width           =   855
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$ Desc."
      Height          =   195
      Index           =   2
      Left            =   2295
      TabIndex        =   23
      Top             =   1155
      Width           =   555
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$ Unitario"
      Height          =   195
      Index           =   1
      Left            =   2295
      TabIndex        =   22
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Cargo"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   1515
      Width           =   780
   End
   Begin VB.Label lblIDServicioOtr 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   315
      Left            =   1470
      TabIndex        =   18
      ToolTipText     =   "Digite el codigo del mecanico"
      Top             =   45
      Width           =   1425
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mecanico"
      Height          =   195
      Index           =   58
      Left            =   120
      TabIndex        =   17
      Top             =   1890
      Width           =   705
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total"
      Height          =   195
      Index           =   54
      Left            =   120
      TabIndex        =   16
      Top             =   2775
      Width           =   690
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$ Desc."
      Height          =   195
      Index           =   55
      Left            =   1455
      TabIndex        =   15
      Top             =   1155
      Width           =   555
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% Desc."
      Height          =   195
      Index           =   56
      Left            =   75
      TabIndex        =   14
      Top             =   1155
      Width           =   585
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      Height          =   195
      Index           =   50
      Left            =   120
      TabIndex        =   13
      Top             =   45
      Width           =   495
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      Height          =   195
      Index           =   51
      Left            =   120
      TabIndex        =   12
      Top             =   420
      Width           =   840
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Horas"
      Height          =   195
      Index           =   52
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   420
   End
End
Attribute VB_Name = "frmEditaOtroServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnSW As Boolean
Dim dblTotalInicial As Double
Sub ReCalculoOS()

dblTotalInicial = CDbl(SacarFormatoValor(txtHorasOtr, "")) * CDbl(SacarFormatoValor(txtPrecioUnitarioOtr, ""))
txtSubTotalOtr = FormatoValor(dblTotalInicial, "", gintDecimalesMoneda)

If Val(txtPorcDescOtr) > 0 Then
    txtMtoDescOtr = FormatoValor(ValorPorcentaje(dblTotalInicial, CSng(SacarFormatoValor(txtPorcDescOtr, ""))), "", gintDecimalesMoneda)
    txtSubTotalOtr = FormatoValor(dblTotalInicial - CDbl(SacarFormatoValor(txtMtoDescOtr, "")), "", gintDecimalesMoneda)
    Exit Sub
End If

If Val(txtMtoDescOtr) > 0 Then
    txtPorcDescOtr = FormatoValor(PorcentajeMonto(dblTotalInicial, CSng(SacarFormatoValor(txtMtoDescOtr, ""))), "", gintDecimalesMoneda)
    txtSubTotalOtr = FormatoValor(dblTotalInicial - CDbl(SacarFormatoValor(txtMtoDescOtr, "")), "", gintDecimalesMoneda)
    Exit Sub
End If

End Sub

Sub UpLoadDataOS()
With frmRecepcion.lvwOtrosServicios
    lblIDServicioOtr = .SelectedItem
    lblDescripcion = .SelectedItem.SubItems(1)
    txtHorasOtr = SacarFormatoValor(.SelectedItem.SubItems(2), "")
    txtPrecioUnitarioOtr = SacarFormatoValor(.SelectedItem.SubItems(3), "")
    'kjcv 17.04.18
    
    If frmRecepcion.cmbCuponera.ListIndex = 1 Then
    End If
    txtPorcDescOtr = SacarFormatoValor(.SelectedItem.SubItems(4), "")
    txtMtoDescOtr = SacarFormatoValor(.SelectedItem.SubItems(5), "")
    txtTipoCargoOtr.Tag = .SelectedItem.SubItems(6)
    txtTipoCargoOtr = .SelectedItem.SubItems(7)
    txtMecanico = .SelectedItem.SubItems(9)
    txtMecanico.Tag = .SelectedItem.SubItems(8)
    txtSubTotalOtr = SacarFormatoValor(.SelectedItem.SubItems(10), "")
    txtHorasReales = SacarFormatoValor(.SelectedItem.SubItems(12), "")
    txtHorasAsignadas = SacarFormatoValor(.SelectedItem.SubItems(17), "")
    
    
    
    'kjcv 14.09.18
    If txtTipoCargoOtr.Tag = "03" Then
        txtCC = .SelectedItem.SubItems(15)
    Else
        txtCC = ""
    End If
    'kjcv 15.09.17
    dtpAsigna.Value = IIf(IsNull(.SelectedItem.SubItems(16)) Or (.SelectedItem.SubItems(16) = "") Or .SelectedItem.SubItems(16) = "01/01/1900", Now(), .SelectedItem.SubItems(16))
    
End With
End Sub

Sub DownLoadDataOS()
With frmRecepcion.lvwOtrosServicios
    .SelectedItem.SubItems(1) = UCase(lblDescripcion)
    .SelectedItem.SubItems(2) = FormatoValor(txtHorasOtr, "", 2)
    .SelectedItem.SubItems(3) = FormatoValor(txtPrecioUnitarioOtr, "", gintDecimalesMoneda)
    .SelectedItem.SubItems(4) = FormatoValor(txtPorcDescOtr, "", 2)
    .SelectedItem.SubItems(5) = FormatoValor(txtMtoDescOtr, "", gintDecimalesMoneda)
    '.SelectedItem.SubItems(6) = lblTipoCargoOtr.Tag
    '.SelectedItem.SubItems(7) = lblTipoCargoOtr.Caption
    .SelectedItem.SubItems(6) = txtTipoCargoOtr.Tag
    .SelectedItem.SubItems(7) = txtTipoCargoOtr
    .SelectedItem.SubItems(8) = txtMecanico.Tag
    .SelectedItem.SubItems(9) = txtMecanico
    .SelectedItem.SubItems(10) = FormatoValor(txtSubTotalOtr, "", gintDecimalesMoneda)
    .SelectedItem.SubItems(11) = "N"
    .SelectedItem.SubItems(12) = FormatoValor(txtHorasReales, "", 2)
    .SelectedItem.SubItems(17) = FormatoValor(txtHorasAsignadas, "", 2)
    'kjcv 17.09.18
    .SelectedItem.SubItems(15) = txtCC
    If txtTipoCargoOtr.Tag = "03" And txtCC = "" Then
        frmCentroCosto.Show vbModal
        .SelectedItem.SubItems(15) = gCentroCosto
    End If
    'kjcv 15.09.17
    .SelectedItem.SubItems(16) = dtpAsigna.Value
    

'    'kjcv 29.09.17
'    If dtpAsigna.Value < frmRecepcion.pckFechaAtencion.Value Then
'        MsgBox "Debe seleccionar una fecha mayor o igual a " & Format(frmRecepcion.pckFechaAtencion.Value, "dd/mm/yyyy"), vbExclamation
'        dtpAsigna.SetFocus
'    Else
'        .SelectedItem.SubItems(16) = dtpAsigna.Value
'    End If
    
    
End With
End Sub

Private Sub cmdAceptarRep_Click()
'kjcv 29.09.17
If Not validacion() Then
        Exit Sub
End If
DownLoadDataOS

Unload Me
End Sub

Private Sub cmdCancelarRep_Click()
Unload Me
End Sub
'kjcv 29.09.17
Private Function validacion() As Boolean
Dim F1 As Date
Dim F2 As Date
Dim DifDias As Single

validacion = True
    
    F1 = CDate(frmEditaOtroServicio.dtpAsigna)
    F2 = CDate(frmRecepcion.pckFechaAtencion)
With Me
'Format(.dtpAsigna.Value, "dd/mm/yyyy") < Format(frmRecepcion.pckFechaAtencion.Value, "dd/mm/yyyy")
DifDias = DateDiff("d", F2, F1)

    If DifDias < 0 Then
        MsgBox "Debe seleccionar una fecha mayor o igual a " & Format(frmRecepcion.pckFechaAtencion.Value, "dd/mm/yyyy"), vbInformation, "Advertencia"
        dtpAsigna.SetFocus
        validacion = False
        Exit Function
    
    End If
End With
End Function




Private Sub cmdCC_Click()
frmCentroCosto.Show vbModal
Me.txtCC = gCentroCosto
End Sub

Private Sub Form_Activate()

If mblnSW Then
    UpLoadDataOS
    mblnSW = False
End If

'txtHorasOtr.Enabled = False

Me.txtHorasAsignadas.Enabled = Atributos("Glbl", "Tllr_30_0086", False, False, False, False)
'If Not Atributos("Glbl", "Tllr_30_0086", False, False, False, False) Then
'        MsgBox "Ud. No cuenta con Acceso para realizar esta operación...", vbInformation, "Advertencia"
''        Unload Me
'        Exit Sub
'    End If


End Sub

Private Sub Form_Load()
mblnSW = True
Me.Label(1).Caption = gstrMonedaLocal & " Unitario"
Me.Label(2).Caption = gstrMonedaLocal & " Desc."
Me.dtpAsigna.Value = Now()


If gstrAsignaRecursos = "S" Then
    Me.txtMecanico.Locked = True
End If
End Sub

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Confirmar"
    ReCalculoOS
    DownLoadDataOS
Case "Cancelar"
    UpLoadDataOS
Case "Cerrar"
    Unload Me
End Select
End Sub
Private Sub lblDescripcion_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub tlbCargo_ButtonClick(ByVal Button As MSComctlLib.Button)
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

If Button.Key = "BuscarCargo" Then
    Me.cmdAceptarRep.SetFocus
    'kjcv 24.03.20
    gstrBusca = ""
    frmTipoCargo.Show vbModal
'    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Tllr_Tipo_Cargo", "Id_Tipo_cargo", "Descripcion", "Buscar Cargo OT")
    If gstrBusca <> "" Then
        If gstrBusca = "03" Then
            Me.cmdCC.Enabled = True
        End If
        'kjcv 08.04.15
            For j = 1 To 9
                If gstrBusca = Cargos(j) Then
                    txtTipoCargoOtr.Tag = gstrBusca
                    txtTipoCargoOtr = TraeCargoDes(gstrBusca)
                    ValidaCostoCargo
                End If
            Next j
            

    End If
End If
End Sub

Private Sub tlbMecanico_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "BuscarMecanico" Then
    If gstrAsignaRecursos = "N" Then
        Me.cmdAceptarRep.SetFocus
        gstrBusca = apfFormulario.BuscarRegistros(Conexion, "(select * from Tllr_Mecanicos where id_empresa='" & gstrIdEmpresa & "' and id_sucursal='" & gstrIdSucursal & "' And Vigencia='S' and Es_Recepcionista='N' and Es_Supervisor='N') as Tllr_Mecanicos", "Id_Mecanico", "Nombre", "Buscar Mecánico")
        'gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Tllr_Mecanicos", "Id_Mecanico", "Nombre", "Mecanicos")
        If gstrBusca <> "" Then
            txtMecanico.Tag = gstrBusca
            txtMecanico = TraeNombreMecanico(gstrBusca)
        End If
    End If
End If
End Sub



Private Sub txtHorasOtr_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtHorasOtr, strDot)
End Sub

Private Sub txtHorasOtr_LostFocus()
If txtHorasOtr <> "" Then
    txtSubTotalOtr = CDbl(txtHorasOtr) * CDbl(IIf(txtPrecioUnitarioOtr <> "", txtPrecioUnitarioOtr, "0"))
End If
If txtHorasAsignadas.Enabled = True Then
txtHorasAsignadas = txtHorasOtr
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
    txtMecanico.Tag = frmRecepcion.lvwOtrosServicios.SelectedItem.SubItems(8)
    txtMecanico = frmRecepcion.lvwOtrosServicios.SelectedItem.SubItems(9)
End If
End Sub

Private Sub txtMtoDescOtr_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtMtoDescOtr, strDot)
End Sub

Private Sub txtMtoDescOtr_LostFocus()
If txtMtoDescOtr <> "" Then
    dblTotalInicial = CDbl(txtHorasOtr) * CDbl(txtPrecioUnitarioOtr)
    txtPorcDescOtr = PorcentajeMonto(dblTotalInicial, CSng(txtMtoDescOtr))
    txtSubTotalOtr = CDbl(dblTotalInicial) - CDbl(txtMtoDescOtr)
End If
End Sub

Private Sub txtPorcDescOtr_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtPorcDescOtr, strDot)
End Sub

Private Sub txtPorcDescOtr_LostFocus()
If txtPorcDescOtr <> "" Then
    dblTotalInicial = CDbl(txtHorasOtr) * CDbl(txtPrecioUnitarioOtr)
    txtMtoDescOtr = ValorPorcentaje(dblTotalInicial, CSng(txtPorcDescOtr))
    txtSubTotalOtr = dblTotalInicial - CDbl(txtMtoDescOtr)
End If

End Sub

Private Sub txtPrecioUnitarioOtr_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtPrecioUnitarioOtr, strDot)
End Sub

Private Sub txtSubTotalOtr_GotFocus()
MarcaTexto txtSubTotalOtr
End Sub

Private Sub txtSubTotalOtr_LostFocus()
Me.txtHorasOtr = CCur(CDbl(txtSubTotalOtr) / CDbl(txtPrecioUnitarioOtr))  'gcurPrecioManoObra)
End Sub

Private Sub txtTipoCargoOtr_GotFocus()
MarcaTexto txtTipoCargoOtr
End Sub

Private Sub txtTipoCargoOtr_LostFocus()
txtTipoCargoOtr.Tag = txtTipoCargoOtr
txtTipoCargoOtr = TraeCargoDes(txtTipoCargoOtr)
ValidaCostoCargo
If txtTipoCargoOtr = "" Then
    txtTipoCargoOtr.Tag = frmRecepcion.lvwOtrosServicios.SelectedItem.SubItems(6)
    txtTipoCargoOtr = frmRecepcion.lvwOtrosServicios.SelectedItem.SubItems(7)
End If

End Sub
Private Sub ValidaCostoCargo()
Dim lstrCostea As String
Dim lstrSQL As String
Dim recAux As New ADODB.Recordset

If Me.txtTipoCargoOtr <> "" Then
    'trae costo cargo
    lstrCostea = Retorna_Valor_General("Select Costea from Tllr_Tipo_Cargo where Id_Empresa='" & gstrIdEmpresa & "' and id_tipo_Cargo='" & Me.txtTipoCargoOtr.Tag & "'", gcdynamic)
    
    If lstrCostea = "S" Then
        If gblnPreciosMarca = True Then
            'trae costo de hora por marca
            lstrSQL = "SELECT CostoManoObra, CostoMOGarantia From Tllr_Marca_Precios_MO WHERE (Id_Marca = '" & frmRecepcion.lblIdMarca & "')"
            If Conexion.SendHost(lstrSQL, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                If Not recAux.BOF And Not recAux.EOF Then
                    txtPrecioUnitarioOtr = IIf(txtTipoCargoOtr.Tag = gstrCargoGtiaFabrica, recAux!CostoMOGarantia, recAux!CostoManoObra)
                    'kjcv 09.11.15
                    If (txtTipoCargoOtr.Tag = gstrCargoGtiaFabrica) Or (txtTipoCargoOtr.Tag = "06") Or (txtTipoCargoOtr.Tag = "07") Or (txtTipoCargoOtr.Tag = "08") Then
                        txtPrecioUnitarioOtr = recAux!CostoMOGarantia
                    Else
                        txtPrecioUnitarioOtr = recAux!CostoManoObra
                    End If
                End If
            End If
        Else
            txtPrecioUnitarioOtr = gcurCostoManoObra
        End If
        
        If txtPorcDescOtr <> "" And txtPorcDescOtr <> "0" Then
            dblTotalInicial = CDbl(txtHorasOtr) * CDbl(txtPrecioUnitarioOtr)
            txtMtoDescOtr = ValorPorcentaje(dblTotalInicial, CSng(txtPorcDescOtr))
            txtSubTotalOtr = dblTotalInicial - CDbl(txtMtoDescOtr)
        Else
            Me.txtSubTotalOtr = CDbl(txtPrecioUnitarioOtr) * CDbl(Me.txtHorasOtr)
            Me.txtPorcDescOtr = 0
            Me.txtMtoDescOtr = 0
        End If
    Else
        If gblnPreciosMarca = True Then
            'trae costo de hora por marca
            lstrSQL = "SELECT VentaManoObra, VentaMOGarantia From Tllr_Marca_Precios_MO WHERE (Id_Marca = '" & frmRecepcion.lblIdMarca & "')"
            If Conexion.SendHost(lstrSQL, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                If Not recAux.BOF And Not recAux.EOF Then
'                    Me.txtPrecioUnitarioOtr = IIf(txtTipoCargoOtr.Tag = gstrCargoGtiaFabrica, recAux!VentaMOGarantia, recAux!VentaManoObra)
                    'kjcv 09.11.15
                    If (txtTipoCargoOtr.Tag = gstrCargoGtiaFabrica) Or (txtTipoCargoOtr.Tag = "06") Or (txtTipoCargoOtr.Tag = "07") Or (txtTipoCargoOtr.Tag = "08") Then
                        Me.txtPrecioUnitarioOtr = recAux!VentaMOGarantia
                    'kjcv 09.09.16
                    ElseIf txtTipoCargoOtr.Tag = "02" Then
                        txtPrecioUnitarioOtr = Round(traeValorHoraCS(gstrIdCompañiaSeg, gstrIdEmpresa) * IIf(traeParidadMonedaMesCS("02", frmRecepcion.pckFechaAtencion, gstrIdCompañiaSeg, gstrIdEmpresa) = 0, traeParidadMoneda("02"), traeParidadMonedaMesCS("02", frmRecepcion.pckFechaAtencion, gstrIdCompañiaSeg, gstrIdEmpresa)), 2)
                    Else
                        Me.txtPrecioUnitarioOtr = recAux!VentaManoObra
                    End If
                    
                End If
            End If
        Else
            'txtPrecioUnitarioOtr = IIf(txtTipoCargoOtr.Tag = gstrCargoGtiaFabrica, Retorna_Valor_General("Select PrecioManoOBraGarantia from Tllr_Parametro Where id_empresa='" & gstrIdEmpresa & "' And id_sucursal='" & gstrIdSucursal & "'", gcdynamic), gcurPrecioManoObra)
            'kjcv 09.11.15
            If (txtTipoCargoOtr.Tag = gstrCargoGtiaFabrica) Or (txtTipoCargoOtr.Tag = "06") Or (txtTipoCargoOtr.Tag = "07") Or (txtTipoCargoOtr.Tag = "08") Then
'                txtPrecioUnitarioOtr = Retorna_Valor_General("Select PrecioManoOBraGarantia from Tllr_Parametro Where id_empresa='" & gstrIdEmpresa & "' And id_sucursal='" & gstrIdSucursal & "'", gcdynamic)
                'kjcv 09.05.20 toma valores de Tllr_MO
                txtPrecioUnitarioOtr = (Retorna_Valor_General("Select ValorMOGarantia from Tllr_mo where id_empresa='" & gstrIdEmpresa & "' and Id_Marca = '" & frmRecepcion.lblIdMarca & "'", gcdynamic))
            'kjcv 09.09.16
            ElseIf txtTipoCargoOtr.Tag = "02" Then
                txtPrecioUnitarioOtr = Round(traeValorHoraCS(gstrIdCompañiaSeg, gstrIdEmpresa) * IIf(traeParidadMonedaMesCS("02", frmRecepcion.pckFechaAtencion, gstrIdCompañiaSeg, gstrIdEmpresa) = 0, traeParidadMoneda("02"), traeParidadMonedaMesCS("02", frmRecepcion.pckFechaAtencion, gstrIdCompañiaSeg, gstrIdEmpresa)), 2)
            Else
                txtPrecioUnitarioOtr = gcurPrecioManoObra
            End If
        End If
        
        If txtPorcDescOtr <> "" Then
            dblTotalInicial = CDbl(txtHorasOtr) * CDbl(txtPrecioUnitarioOtr)
            txtMtoDescOtr = ValorPorcentaje(dblTotalInicial, CSng(txtPorcDescOtr))
            txtSubTotalOtr = dblTotalInicial - CDbl(txtMtoDescOtr)
        End If
    End If
End If
End Sub

