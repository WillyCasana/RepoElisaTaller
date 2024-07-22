VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmEditaServicioRepuesto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edición Servicio Repuesto"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "frmEditaServicioRepuesto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCC 
      Caption         =   "..."
      Height          =   255
      Left            =   5640
      TabIndex        =   21
      Top             =   1470
      Width           =   375
   End
   Begin VB.TextBox txtCC 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   5040
      MaxLength       =   4
      TabIndex        =   19
      Top             =   1440
      Width           =   585
   End
   Begin VB.TextBox txtTipoCargo 
      Height          =   330
      Left            =   1155
      TabIndex        =   5
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancelarRep 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   5190
      TabIndex        =   8
      Top             =   2205
      Width           =   800
   End
   Begin VB.CommandButton cmdAceptarRep 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   315
      Left            =   4395
      TabIndex        =   7
      Top             =   2205
      Width           =   800
   End
   Begin VB.TextBox txtSubTotalRep 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1155
      TabIndex        =   6
      Top             =   1815
      Width           =   2040
   End
   Begin VB.TextBox txtPreUniRep 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2700
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   2
      Top             =   765
      Width           =   1320
   End
   Begin VB.TextBox txtPorcDescRep 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1155
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1110
      Width           =   700
   End
   Begin VB.TextBox txtMtoDescRep 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2700
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1110
      Width           =   1320
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1155
      MaxLength       =   4
      TabIndex        =   1
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
            Picture         =   "frmEditaServicioRepuesto.frx":000C
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":011E
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":0230
            Key             =   "Cerrar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCargo 
      Height          =   330
      Left            =   4110
      TabIndex        =   18
      Top             =   1455
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
      Left            =   4860
      Top             =   2460
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
            Picture         =   "frmEditaServicioRepuesto.frx":0342
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":0454
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":0566
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":0678
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":078A
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":089C
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":09AE
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":0AC0
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":0BD2
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":0CE4
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":0DF6
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":0F08
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":101A
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":112C
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":123E
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":1350
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":1462
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":18B4
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditaServicioRepuesto.frx":1D06
            Key             =   "Copiar"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "CC"
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   1470
      Width           =   255
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   315
      Left            =   1155
      TabIndex        =   0
      Top             =   420
      Width           =   4575
   End
   Begin VB.Label lblIdItem 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   315
      Left            =   1155
      TabIndex        =   17
      Top             =   60
      Width           =   2730
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Cargo"
      Height          =   195
      Index           =   58
      Left            =   60
      TabIndex        =   16
      Top             =   1470
      Width           =   780
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total"
      Height          =   195
      Index           =   54
      Left            =   75
      TabIndex        =   15
      Top             =   1845
      Width           =   690
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$ Unitario"
      Height          =   195
      Index           =   53
      Left            =   1950
      TabIndex        =   14
      Top             =   825
      Width           =   675
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$ Desc."
      Height          =   195
      Index           =   55
      Left            =   1965
      TabIndex        =   13
      Top             =   1140
      Width           =   555
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% Desc."
      Height          =   195
      Index           =   56
      Left            =   75
      TabIndex        =   12
      Top             =   1095
      Width           =   585
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo Item"
      Height          =   195
      Index           =   50
      Left            =   45
      TabIndex        =   11
      Top             =   90
      Width           =   840
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      Height          =   195
      Index           =   51
      Left            =   60
      TabIndex        =   10
      Top             =   435
      Width           =   840
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      Height          =   195
      Index           =   52
      Left            =   105
      TabIndex        =   9
      Top             =   765
      Width           =   675
   End
End
Attribute VB_Name = "frmEditaServicioRepuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnSW As Boolean
Dim dblTotalInicial As Double


Sub UpLoadDataRep()
With Me
    .lblIdItem = frmRecepcion.lvwRepuestos.SelectedItem
    .lblDescripcion = frmRecepcion.lvwRepuestos.SelectedItem.SubItems(1)
    .txtCantidad = SacarFormatoValor(frmRecepcion.lvwRepuestos.SelectedItem.SubItems(2), "")
    .txtPreUniRep = SacarFormatoValor(frmRecepcion.lvwRepuestos.SelectedItem.SubItems(3), "")
    .txtPorcDescRep = SacarFormatoValor(frmRecepcion.lvwRepuestos.SelectedItem.SubItems(4), "")
    .txtMtoDescRep = SacarFormatoValor(frmRecepcion.lvwRepuestos.SelectedItem.SubItems(5), "")
    '.lblTipoCargo.Tag = frmRecepcion.lvwRepuestos.SelectedItem.SubItems(7)
    '.lblTipoCargo.Caption = frmRecepcion.lvwRepuestos.SelectedItem.SubItems(6)
    .txtTipoCargo.Tag = frmRecepcion.lvwRepuestos.SelectedItem.SubItems(7)
    .txtTipoCargo = frmRecepcion.lvwRepuestos.SelectedItem.SubItems(6)
    .txtSubTotalRep = SacarFormatoValor(frmRecepcion.lvwRepuestos.SelectedItem.SubItems(8), "")
    'kjcv 14.09.18
    txtCC = frmRecepcion.lvwRepuestos.SelectedItem.SubItems(14)
'    If txtTipoCargo.Tag = "03" Then
'        frmCentroCosto.Show vbModal
'        frmRecepcion.lvwRepuestos.SelectedItem.SubItems(14) = gCentroCosto
'    Else
'        frmRecepcion.lvwRepuestos.SelectedItem.SubItems(14) = ""
'    End If
End With
End Sub

Sub DownLoadDataRep()
With frmRecepcion.lvwRepuestos.SelectedItem
    .SubItems(2) = FormatoValor(txtCantidad, "", 2)
    .SubItems(3) = FormatoValor(txtPreUniRep, "", gintDecimalesMoneda)
    .SubItems(4) = FormatoValor(txtPorcDescRep, "", 2)
    .SubItems(5) = FormatoValor(txtMtoDescRep, "", gintDecimalesMoneda)
    .SubItems(7) = txtTipoCargo.Tag
    .SubItems(6) = txtTipoCargo
    .SubItems(8) = FormatoValor(txtSubTotalRep, "", gintDecimalesMoneda)
    .SubItems(10) = "N"
    'kjcv 17.09.18
    .SubItems(14) = txtCC
    'If txtTipoCargo.Tag = "03" Then
    If txtTipoCargo.Tag = "03" And txtCC = "" Then
        frmCentroCosto.Show vbModal
        frmRecepcion.lvwRepuestos.SelectedItem.SubItems(14) = gCentroCosto
'    Else
''        frmRecepcion.lvwRepuestos.SelectedItem.SubItems(14) = ""
    End If
    
End With
End Sub


Function VerificaLinea() As Boolean
With Me
    If .txtCantidad <> "" And _
    .txtPreUniRep <> "" And _
    .txtPorcDescRep <> "" And _
    .txtMtoDescRep <> "" And _
    .txtSubTotalRep <> "" Then
        VerificaLinea = True
    Else
        VerificaLinea = False
    
    End If
End With

End Function

Private Sub cmdAceptarRep_Click()

If Me.Tag <> "descto_ok" Then txtPorcDescRep_LostFocus

'valida descuento
If gblnBloqueaSubtotalRep = False Then
    If Me.txtPorcDescRep = 0 Then
        If CDbl(Me.txtCantidad) * CDbl(Me.txtPreUniRep) > CDbl(Me.txtSubTotalRep) Then
            Me.txtPorcDescRep = 100 - (Round((Me.txtSubTotalRep * 100) / (CDbl(Me.txtCantidad) * CDbl(Me.txtPreUniRep)), 2))
            Me.txtMtoDescRep = (CDbl(Me.txtCantidad) * CDbl(Me.txtPreUniRep)) - CDbl(Me.txtSubTotalRep)
        End If
    End If
End If

'If Val(Me.txtPorcDescRep) <= gintDescuentoMaximo Then  'valida si el porcentaje ingresado es mayor que el parametro
    If gblnValidaCostoRepuestos = True Then
        If (CostoRepuesto(lblIdItem, txtCantidad) <= CDbl(txtSubTotalRep)) Then 'valida si la venta es menor que el costo
            DownLoadDataRep
        Else
            If MsgBox("El Valor Venta del Repuesto " & Me.lblDescripcion & " " & Chr(13) & _
                      "Es menor que el Precio de Costo " & Chr(13) & "Desea Continuar...", vbQuestion + vbYesNo, "Confirma Precio Venta Repuesto") = vbYes Then
                
                Screen.MousePointer = 1
                gblnDescuentoRepuesto = True
                frmPermisoDiasHabiles.Show 1
                
                If NoEsLaPassword(gstrVerificacion, gstrMecanicoDiasHabiles) Then
                    DownLoadDataRep
                    Unload Me
                Else
                    MsgBox "Lo Siento, La passWord ingresada no es Correcta", vbExclamation, "Password"
                End If
                gblnDescuentoRepuesto = False
                
            End If
            Exit Sub
        End If
    Else
        DownLoadDataRep
    End If
    If frmRecepcion.dtcGarantia.BoundText = "PRE" Then
        Me.txtPreUniRep.Locked = True
    End If
    Unload Me
'Else
'    MsgBox "El descuento ingresado es mayor que el permitido", vbExclamation, "Advertencia"
'End If
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
    UpLoadDataRep
    mblnSW = False
    If frmRecepcion.dtcGarantia.BoundText = "PRE" Then
        Me.txtPreUniRep.Locked = False
    End If
    If gblnBloqueaSubtotalRep = True Then
        Me.txtSubTotalRep.Locked = True
    Else
        Me.txtSubTotalRep.Locked = False
    End If
End If

End Sub

Private Sub Form_Load()
mblnSW = True
Me.Label(53).Caption = gstrMonedaLocal & " Unitario"
Me.Label(55).Caption = gstrMonedaLocal & " Desc."
frmEditaServicioRepuesto.txtCC = gCentroCosto
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

'Const CargoGtiaFab As String = "04"
'Dim gTipoCargoActual As String
'
If Button.Key = "BuscarCargo" Then
    txtCantidad.SetFocus
    'kjcv 24.03.20
    gstrBusca = ""
    frmTipoCargo.Show vbModal
'    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Tllr_Tipo_Cargo", "Id_Tipo_cargo", "Descripcion", "Buscar Cargo OT")
   
            
    If gstrBusca <> "" Then
        
        
    ' inicio kjcv 07.04.15
        For j = 1 To 6
            If gstrBusca = Cargos(j) Then
                txtTipoCargo.Tag = gstrBusca
                txtTipoCargo = TraeCargoDes(gstrBusca)
                ValidaCostoCargo
            End If
        Next j
    'fin kjcv 07.04.15
        
''        txtTipoCargo.Tag = gstrBusca
''        txtTipoCargo = TraeCargoDes(gstrBusca)
''        ValidaCostoCargo
    End If
End If
End Sub

Private Sub txtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtSubTotalRep = Val(txtCantidad) * Val(txtPreUniRep)
End If
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtCantidad, strDot)
End Sub

Private Sub txtCantidad_LostFocus()
    txtPorcDescRep_LostFocus
End Sub

Private Sub txtMtoDescRep_GotFocus()
MarcaTexto txtMtoDescRep
End Sub


Private Sub txtMtoDescRep_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtMtoDescRep, strDot)
End Sub

Private Sub txtMtoDescRep_LostFocus()
Dim dblTotalInicial As Double

If Val(txtMtoDescRep) > 0 Then
    dblTotalInicial = Val(txtCantidad) * Val(txtPreUniRep)
    txtPorcDescRep = PorcentajeMonto(dblTotalInicial, IIf(txtMtoDescRep <> "", CSng(Val(txtMtoDescRep)), 0))
    'txtSubTotalRep = dblTotalInicial - PorcentajeMonto(dblTotalInicial, IIf(txtMtoDescRep <> "", CSng(Val(txtMtoDescRep)), 0))
    txtSubTotalRep = dblTotalInicial - txtMtoDescRep
Else
    txtSubTotalRep = Val(txtCantidad) * Val(txtPreUniRep)
End If

End Sub

Private Sub txtPorcDescRep_GotFocus()
MarcaTexto txtPorcDescRep
End Sub

Private Sub txtPorcDescRep_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtPorcDescRep, strDot)
End Sub

Private Sub txtPorcDescRep_LostFocus()
Dim lstrPswrdDescto As String

'If Val(Me.txtPorcDescRep) <= gintDescuentoMaximo Then
'    If Val(txtPorcDescRep) > 0 Then
'        dblTotalInicial = Round(Val(txtCantidad) * Val(txtPreUniRep), 2)
'        txtMtoDescRep = ValorPorcentaje(dblTotalInicial, IIf(txtPorcDescRep <> "", CSng(Val(txtPorcDescRep)), 0))
'        'txtSubTotalRep = dblTotalInicial - ValorPorcentaje(dblTotalInicial, IIf(txtPorcDescRep <> "", CSng(Val(txtPorcDescRep)), 0))
'        '//// MODIFICADO POR FDO DIAZ EL 07/12/2000 CALCULA MAL EL DESCUENTO
'        txtSubTotalRep = dblTotalInicial - txtMtoDescRep
'    Else
'        txtMtoDescRep = "0"
'        txtSubTotalRep = Val(txtCantidad) * Val(txtPreUniRep)
'    End If
'Else
'    MsgBox "El descuento ingresado es mayor que el permitido", vbExclamation, "Advertencia"
'    txtPorcDescRep.SetFocus
'End If

If Val(Me.txtPorcDescRep) <= gintDescuentoMaximo Then
    
    If Val(txtPorcDescRep) > 0 Then
        dblTotalInicial = Round(Val(txtCantidad) * Val(txtPreUniRep), 2)
        txtMtoDescRep = ValorPorcentaje(dblTotalInicial, IIf(txtPorcDescRep <> "", CSng(Val(txtPorcDescRep)), 0))
        'txtSubTotalRep = dblTotalInicial - ValorPorcentaje(dblTotalInicial, IIf(txtPorcDescRep <> "", CSng(Val(txtPorcDescRep)), 0))
        '//// MODIFICADO POR FDO DIAZ EL 07/12/2000 CALCULA MAL EL DESCUENTO
        txtSubTotalRep = dblTotalInicial - txtMtoDescRep
    Else
        txtMtoDescRep = "0"
        txtSubTotalRep = Val(txtCantidad) * Val(txtPreUniRep)
    End If
'''kjcv 13.03.17
''ElseIf txtTipoCargo.Text = "02" And (Val(Me.txtPorcDescRep) <= gintDescuentoMaximoCIA) Then
''
''        If Val(txtPorcDescRep) > 0 Then
''            dblTotalInicial = Round(Val(txtCantidad) * Val(txtPreUniRep), 2)
''            txtMtoDescRep = ValorPorcentaje(dblTotalInicial, IIf(txtPorcDescRep <> "", CSng(Val(txtPorcDescRep)), 0))
''            txtSubTotalRep = dblTotalInicial - txtMtoDescRep
''        Else
''            txtMtoDescRep = "0"
''            txtSubTotalRep = Val(txtCantidad) * Val(txtPreUniRep)
''        End If
    
Else
    Dim strSql As String
    Dim adoTemp As New ADODB.Recordset
    strSql = "Select Pswrd_Descuento from Tllr_Parametro where Id_Empresa='" & gstrIdEmpresa & "' AND Id_Sucursal='" & gstrIdSucursal & "'"
    If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            lstrPswrdDescto = adoTemp!Pswrd_Descuento
        End If
    End If

    If InputBox("Descuento mayor al permitido" & Chr(13) & "Ingresesu contraseña:", "Autorización de Descuento", "") = lstrPswrdDescto Then
        If Val(txtPorcDescRep) > 0 Then
            dblTotalInicial = Round(Val(txtCantidad) * Val(txtPreUniRep), 2)
            txtMtoDescRep = ValorPorcentaje(dblTotalInicial, IIf(txtPorcDescRep <> "", CSng(Val(txtPorcDescRep)), 0))
            'txtSubTotalRep = dblTotalInicial - ValorPorcentaje(dblTotalInicial, IIf(txtPorcDescRep <> "", CSng(Val(txtPorcDescRep)), 0))
            '//// MODIFICADO POR FDO DIAZ EL 07/12/2000 CALCULA MAL EL DESCUENTO
            txtSubTotalRep = dblTotalInicial - txtMtoDescRep
        Else
            txtMtoDescRep = "0"
            txtSubTotalRep = Val(txtCantidad) * Val(txtPreUniRep)
        End If
        Me.Tag = "descto_ok"
    Else
        MsgBox "Contraseña incorrecta", vbExclamation, "Advertencia"
        txtPorcDescRep.Text = "0"
        txtMtoDescRep = "0"
        txtSubTotalRep = Val(txtCantidad) * Val(txtPreUniRep)
        Me.Tag = ""
    End If
End If

End Sub


Private Sub txtPreUniRep_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtPreUniRep, strDot)
End Sub
Private Sub txtPreUniRep_LostFocus()
    txtPorcDescRep_LostFocus
End Sub

Private Sub txtSubTotalRep_LostFocus()
    If gblnBloqueaSubtotalRep = False Then
        If Me.txtPorcDescRep = 0 Then
            If CDbl(Me.txtCantidad) * CDbl(Me.txtPreUniRep) > CDbl(Me.txtSubTotalRep) Then
                Me.txtPorcDescRep = 100 - (Round((Me.txtSubTotalRep * 100) / (CDbl(Me.txtCantidad) * CDbl(Me.txtPreUniRep)), 2))
                Me.txtMtoDescRep = (CDbl(Me.txtCantidad) * CDbl(Me.txtPreUniRep)) - CDbl(Me.txtSubTotalRep)
            End If
        End If
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
    txtTipoCargo.Tag = frmRecepcion.lvwRepuestos.SelectedItem.SubItems(7)
    txtTipoCargo = frmRecepcion.lvwRepuestos.SelectedItem.SubItems(6)
End If

End Sub
Private Sub ValidaCostoCargo()
Dim lstrCostea As String
Dim strSql As String
Dim adoTemp As New ADODB.Recordset
Dim Sql As String
Dim dFecha As Date
Dim adoFecha As New ADODB.Recordset

    Sql = "SELECT * FROM TLLR_OT WHERE Id_OT='" & frmRecepcion.lblNroRecepcion & "' "
    Sql = Sql & " AND Id_Empresa='" & gstrIdEmpresa & "'"
    Sql = Sql & " AND Id_Sucursal='" & gstrIdSucursal & "'"
'    sql = sql & " AND Seccion_OT='" & Mid(Me.cmbSeccion, 1, 1) & "'"
    If Conexion.SendHost(Sql, adoFecha, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoFecha.BOF And Not adoFecha.EOF Then
            dFecha = adoFecha!Fecha_Emision
        End If
    End If
    Conexion.CloseHost adoFecha

If Me.txtTipoCargo <> "" Then
    'trae costo cargo
    lstrCostea = Retorna_Valor_General("Select Costea from Tllr_Tipo_Cargo where Id_Empresa='" & gstrIdEmpresa & "' and id_tipo_Cargo='" & Me.txtTipoCargo.Tag & "'", gcdynamic)
    If lstrCostea = "S" Then
        '//LREYES...
        Me.txtPorcDescRep = 0
        Me.txtMtoDescRep = 0
        'multiplicaba como 4 veces el precio unitario
        txtPreUniRep = CostoRepuesto(lblIdItem, txtCantidad) * traeParidadMonedaMes("02", dFecha)
''        'kjcv 01.03.13 cambio de taller a dolares
''        txtPreUniRep = CostoRepuesto(lblIdItem, txtCantidad)
        Me.txtSubTotalRep = CDbl(txtPreUniRep) * CDbl(Me.txtCantidad)
        
 Else
        '//LREYES...
        'kjcv 25.03.20 para cia toma precio_CIa
        strSql = "select isnull(precio_venta,0) as precio_venta,isnull(Precio_Venta_CIA,0) as Precio_CIA from stck_item where id_item='" & lblIdItem & "'"
'kjcv 26.02.13
''        strSql = "select isnull(precioventaD,0) as precio_venta from Tllr_Repuestos_OT where id_item='" & lblIdItem & "'"
''        strSql = strSql & " and Id_OT='" & frmRecepcion.lblNroRecepcion & "' "
        If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoTemp.BOF And Not adoTemp.EOF Then
                Me.txtPorcDescRep = 0
                Me.txtMtoDescRep = 0
                'kjcv 06.11.12
                If Me.txtTipoCargo.Tag = "02" Then
'
                    'kjcv 17.07.13 se agrego tipo de cambio por mes por Compañia de Seguro
                    'kjcv 25.03.20
                    txtPreUniRep = adoTemp!Precio_CIA * IIf(traeParidadMonedaMesCS("02", frmRecepcion.pckFechaAtencion, gstrIdCompañiaSeg, gstrIdEmpresa) = 0, traeParidadMoneda("02"), traeParidadMonedaMesCS("02", frmRecepcion.pckFechaAtencion, gstrIdCompañiaSeg, gstrIdEmpresa))
                    Me.txtSubTotalRep = CDbl(txtPreUniRep) * CDbl(Me.txtCantidad)
''                    Me.txtSubTotalRep = adoTemp!Precio_CIA * IIf(traeParidadMonedaMesCS("02", frmRecepcion.pckFechaAtencion, gstrIdCompañiaSeg) = 0, traeParidadMoneda("02"), traeParidadMonedaMesCS("02", frmRecepcion.pckFechaAtencion, gstrIdCompañiaSeg))
                    
                Else
                'kjcv 10.02.16
                    If Me.txtTipoCargo.Tag = "04" Or Me.txtTipoCargo.Tag = "06" Or Me.txtTipoCargo.Tag = "07" Then
                        'kjcv 10.02.16 tipo cambio gtia fabrica
                        txtPreUniRep = (1 + (gstrPorPrecioGtia / 100)) * CostoRepuesto(lblIdItem, txtCantidad) * traeParidadMonedaMesGarantia("02", frmRecepcion.pckFechaAtencion)
                        Me.txtSubTotalRep = CDbl(txtPreUniRep) * CDbl(Me.txtCantidad)
                        
                    ''kjcv 10.02.16cargo Cortesi Comercial
                    ElseIf Me.txtTipoCargo.Tag = "08" Then
                        txtPreUniRep = CostoRepuesto(lblIdItem, txtCantidad) * traeParidadMonedaMesGarantia("02", frmRecepcion.pckFechaAtencion)
                        Me.txtSubTotalRep = CDbl(txtPreUniRep) * CDbl(Me.txtCantidad)
                    Else
                    'kjcv 25..03..20 para cliente y cliente2
                        txtPreUniRep = adoTemp!Precio_Venta * traeParidadMonedaMes("02", frmRecepcion.pckFechaAtencion)
                        Me.txtSubTotalRep = CDbl(txtPreUniRep) * CDbl(Me.txtCantidad)
'                        Me.txtSubTotalRep = adoTemp!Precio_venta * traeParidadMonedaMes("02", frmRecepcion.pckFechaAtencion)
                    End If

                End If
'               'kjcv 01.03.13 Cambio de taller a dolares
''                Me.txtSubTotalRep = adoTemp!Precio_Venta
'                txtPreUniRep = Me.txtSubTotalRep
            End If
        End If
        Conexion.CloseHost adoTemp
    End If
End If
End Sub

