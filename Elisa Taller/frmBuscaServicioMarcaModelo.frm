VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBuscaServicioMarcaModelo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar Servicios"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   Icon            =   "frmBuscaServicioMarcaModelo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7515
   Begin VB.TextBox txtNroRecord 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "10"
      Top             =   1650
      Width           =   510
   End
   Begin VB.TextBox txtCodigo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   735
      Width           =   2235
   End
   Begin VB.TextBox txtDes 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Top             =   1185
      Width           =   5130
   End
   Begin VB.ComboBox cboCoincidir 
      Height          =   315
      ItemData        =   "frmBuscaServicioMarcaModelo.frx":0442
      Left            =   1320
      List            =   "frmBuscaServicioMarcaModelo.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1635
      Width           =   2220
   End
   Begin VB.CheckBox optCriterios 
      Caption         =   "Código"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   6
      Top             =   765
      Width           =   990
   End
   Begin VB.Frame fmeServicios 
      Caption         =   "Servicios del Modelo"
      Height          =   3735
      Left            =   45
      TabIndex        =   0
      Top             =   1980
      Width           =   7440
      Begin MSComctlLib.ListView lvwServicios 
         Height          =   3465
         Left            =   30
         TabIndex        =   3
         Top             =   195
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   6112
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Codigo"
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Des"
            Text            =   "Descripción"
            Object.Width           =   10019
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   "NroHoras"
            Text            =   "Nº Horas"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Valor"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ImageList ImgBarraHerramienta 
         Left            =   450
         Top             =   -45
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
               Picture         =   "frmBuscaServicioMarcaModelo.frx":04A5
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":05B7
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":0A0F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":0E67
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":12BF
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":13D1
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":14E3
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":15F5
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":1707
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":1819
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":192B
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":1A3D
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":1B4F
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":1C61
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":1D73
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":1E85
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":1F97
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":20A9
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":21BB
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":22CD
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":271F
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaServicioMarcaModelo.frx":2B71
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar tlbOpciones 
      Height          =   330
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   5775
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   582
      ButtonWidth     =   1508
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Todos"
            Key             =   "SelectAll"
            Object.ToolTipText     =   "Agrega Servicio Nuevo"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Todos"
            Key             =   "UnSelectAll"
            Object.ToolTipText     =   "Quitar Servicio"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbOpciones 
      Height          =   330
      Index           =   1
      Left            =   3645
      TabIndex        =   5
      Top             =   5805
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   582
      ButtonWidth     =   2196
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar Servicios"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleccionar"
            Key             =   "Seleccionar"
            Object.ToolTipText     =   "Seleccionar Servicio"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar"
            Key             =   "Cerrar"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox optCriterios 
      Caption         =   "Descripción"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   10
      Top             =   1230
      Width           =   1305
   End
   Begin MSComctlLib.Toolbar tlbMarca 
      Height          =   330
      Left            =   2850
      TabIndex        =   13
      Top             =   75
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
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
            Object.ToolTipText     =   "Agrega Servicio Nuevo"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbModelo 
      Height          =   330
      Left            =   7110
      TabIndex        =   15
      Top             =   60
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
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
            Object.ToolTipText     =   "Agrega Servicio Nuevo"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.UpDown updNroRecord 
      Height          =   315
      Left            =   6150
      TabIndex        =   16
      Top             =   1650
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Value           =   5
      BuddyControl    =   "txtNroRecord"
      BuddyDispid     =   196609
      OrigLeft        =   8445
      OrigTop         =   300
      OrigRight       =   8685
      OrigBottom      =   615
      Max             =   100
      Min             =   5
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nro. de Registros :"
      Height          =   195
      Index           =   1
      Left            =   4350
      TabIndex        =   18
      Top             =   1695
      Width           =   1320
   End
   Begin VB.Label lblModelo 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3885
      TabIndex        =   14
      Top             =   75
      Width           =   3195
   End
   Begin VB.Label lblMarca 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   690
      TabIndex        =   12
      Top             =   90
      Width           =   2100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   90
      X2              =   7440
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Index           =   0
      X1              =   90
      X2              =   7440
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Coincidir en :"
      Height          =   195
      Left            =   285
      TabIndex        =   11
      Top             =   1695
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Modelo :"
      Height          =   195
      Left            =   3270
      TabIndex        =   2
      Top             =   135
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Marca :"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "frmBuscaServicioMarcaModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnSW As Boolean
Dim adoPrincipal As New ADODB.Recordset
Dim mstrSql As String
Dim mstrWhere As String

Dim lsiItemSelected As Boolean
Dim lsiItem As ListItem, itmFound As ListItem
Dim intContador As Integer
Const mcintHeight As Integer = 7900
Const mcintWidth As Integer = 11900
Const mcstrMensaje As String = "Confirma Eliminar El Item Seleccionado desde "



Sub EliminarItem(intTipo As Integer, strMarca As String, strModelo As String, Optional strServicio As String, Optional strActividad As String, Optional strRepuesto As String)
Dim strSql As String

Select Case intTipo
    
    Case 0 '////////elimina servicio
        If MsgBox(mcstrMensaje & "Servicios por Modelos", 4 + 32) = vbYes Then
            strSql = "SELECT COUNT(*) AS CUANTOS FROM Tllr_Actividad_Servicio_Modelo WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' "
            If Conexion.SendHost(strSql, adoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
                With adoPrincipal
                    .MoveFirst
                    If !CUANTOS > 0 Then
                        '////////////////// TIENE ACTIVIDADES RELACIONADAS
                        MsgBox "TIENE ACTIVIDADES RELACIONADAS"
                    Else
                        '////////////////// NO TIENE ACTIVIDADES RELACIONADAS
                        MsgBox "NO TIENE ACTIVIDADES RELACIONADAS"
                        strSql = "DELETE FROM Tllr_Servicio_Modelo WHERE Id_Marca = '" & strMarca & "' AND Id_Modelo = '" & strModelo & "' AND Id_Servicio = '" & strServicio & "' "
                        Conexion.SendHost strSql, , , , gcTiempoEspera
                        lvwServicios.ListItems.Remove lvwServicios.SelectedItem.Index
                    End If
                End With
            End If
            
        End If
        
    
End Select

End Sub

Sub ServiciosdelModelo(strCondicion As String, strOrden As String)
Dim Valor As Double
Dim recAux As New ADODB.Recordset
Dim ValorHoraMarca As Double


lvwServicios.ListItems.Clear

mstrSql = "SELECT VentaManoObra, VentaMOGarantia From Tllr_Marca_Precios_MO WHERE (Id_Marca = '" & Me.lblMarca.Tag & "')"
If Conexion.SendHost(mstrSql, recAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not recAux.BOF And Not recAux.EOF Then
        ValorHoraMarca = recAux!VentaManoObra
    End If
End If

If gstrServiciosMarca = "S" Then
    'Valor = ValorHora
    mstrSql = "SELECT  TOP " & CStr(updNroRecord.Value) & " Tllr_Servicio_Modelo.Id_Servicio AS ID,"
    mstrSql = mstrSql & " Tllr_Servicio.Descripcion AS DES,"
    mstrSql = mstrSql & " Tllr_Servicio_Modelo.Horas AS TIEMPO,"
    mstrSql = mstrSql & " " & ValorHoraMarca & " AS VALOR"
    mstrSql = mstrSql & " FROM Tllr_Servicio_Modelo LEFT OUTER JOIN Tllr_Servicio "
    mstrSql = mstrSql & " On Tllr_Servicio_Modelo.Id_Marca = Tllr_Servicio.Id_Marca "
    mstrSql = mstrSql & " And Tllr_Servicio_Modelo.Id_Servicio = Tllr_Servicio.Id_Servicio"
    mstrSql = mstrSql & strCondicion & " " & strOrden
Else
    Valor = 10000
    mstrSql = "SELECT  TOP " & CStr(updNroRecord.Value) & " Tllr_Servicio_Modelo.Id_Servicio AS ID,"
    mstrSql = mstrSql & " Tllr_Servicio.Descripcion AS DES,"
    mstrSql = mstrSql & " Tllr_Servicio_Modelo.Horas AS TIEMPO,"
    mstrSql = mstrSql & " " & ValorHora(gstrIdEmpresa, gstrIdSucursal) & " AS VALOR"
    mstrSql = mstrSql & " FROM Tllr_Servicio_Modelo LEFT OUTER JOIN Tllr_Servicio ON Tllr_Servicio_Modelo.Id_Servicio = Tllr_Servicio.Id_Servicio"
    mstrSql = mstrSql & strCondicion & " " & strOrden
End If
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            While Not .EOF
                Set lsiItem = lvwServicios.ListItems.Add(, , !ID)
                lsiItem.SubItems(1) = ValorNulo(!Des)
                lsiItem.SubItems(2) = !TIEMPO
                lsiItem.SubItems(3) = Format(!Valor, "###,##0.0")
                .MoveNext
            Wend
        End If
    End With
End If

End Sub



Private Sub Form_Activate()
If mblnSW Then
'    If gstrProcedencia = "Movimientos" Then
        lblMarca.Caption = frmPresupuestoMantenciones.dtcMarca.Text
        lblMarca.Tag = frmPresupuestoMantenciones.dtcMarca.BoundText
        lblModelo.Caption = frmPresupuestoMantenciones.dtcModelo.Text
        lblModelo.Tag = frmPresupuestoMantenciones.dtcModelo.BoundText
'    ElseIf gstrProcedencia = "Temparios" Then
'        lblMarca.Caption = frmRecepcion.lblMarca.Caption
'        lblMarca.Tag = frmRecepcion.lblIdMarca.Caption
'        lblModelo.Caption = frmRecepcion.lblModelo.Caption
'        lblModelo.Tag = frmRecepcion.lblIdModelo.Caption
'    End If
    cboCoincidir.ListIndex = 0
    mblnSW = False
End If

End Sub

Private Sub Form_Load()

mblnSW = True
updNroRecord.Value = gintNroRecDefectoQry
End Sub

Private Sub lvwServicios_DblClick()
        frmPresupuestoMantenciones.txtCodigoServicio.Tag = lvwServicios.SelectedItem
        frmPresupuestoMantenciones.txtCodigoServicio.Text = lvwServicios.SelectedItem.SubItems(1)
        frmPresupuestoMantenciones.lblHorasServicio = lvwServicios.SelectedItem.SubItems(2)
        frmPresupuestoMantenciones.lblValorServicio = FormatoValor(CDbl(lvwServicios.SelectedItem.SubItems(3)) * CDbl(lvwServicios.SelectedItem.SubItems(2)), "", gintDecimalesMoneda)
        Unload Me
End Sub

Private Sub optCriterios_Click(Index As Integer)
With Me
    Select Case Index
    Case 0
        If .optCriterios(0).Value = 1 Then ' codigo
            .optCriterios(1).Value = 0
            .txtDes.Enabled = False
            .txtDes.Text = ""
            .txtCodigo.Enabled = True
            .txtCodigo.SetFocus
        Else
            .txtCodigo.Enabled = False
            .txtCodigo.Text = ""
        End If
    Case 1 '////////////////---------------descripcion
        If .optCriterios(1).Value = 1 Then
            .optCriterios(0).Value = 0
            .txtDes.Enabled = True
            .txtCodigo.Enabled = False
            .txtCodigo.Text = ""
            .txtDes.SetFocus
        Else
            .txtDes.Enabled = False
            .txtDes.Text = ""
        
        End If
    End Select
End With
End Sub

Private Sub tlbMarca_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Glbl_Marca", "Id_Marca", "Descripcion", "Busca Marca")
    lblMarca.Tag = gstrBusca
    lblMarca.Caption = MarcaD(gstrBusca)
    lblModelo.Caption = ""
End If

End Sub

Private Sub tlbModelo_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    If lblMarca.Tag <> "" Then
        gstrBusca = apfFormulario.BuscarRegistrosModelo(Conexion, "Glbl_Modelo", "Id_Modelo", "Id_Marca", "Descripcion", "Busca Modelo", lblMarca.Tag)
        lblModelo.Tag = gstrBusca
        lblModelo.Caption = ModeloD(lblMarca.Tag, gstrBusca)
    Else
        MsgBox "Seleccione la Marca"
    End If
End If
End Sub

Private Sub tlbOpciones_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Select Case Index
Case 0 '//////////////////////////////// seleccionar todos o no
    Select Case Button.Key
    Case "SelectAll" '////Todos
        SelectingItem lvwServicios, gcSelectAll
    Case "UnSelectAll" '////Ninguno
        SelectingItem lvwServicios, gcUnSelectAll
    End Select
Case 1 '//////////////////////////////// buscar, Agregar y cerrar
    'gstrIdCargo = gstrIdCargoDefecto
    Select Case Button.Key
    Case "Seleccionar" '////Seleccionar
        If Me.lvwServicios.ListItems.Count > 0 Then
            frmPresupuestoMantenciones.txtCodigoServicio.Tag = lvwServicios.SelectedItem
            frmPresupuestoMantenciones.txtCodigoServicio.Text = lvwServicios.SelectedItem.SubItems(1)
            frmPresupuestoMantenciones.lblHorasServicio = lvwServicios.SelectedItem.SubItems(2)
            frmPresupuestoMantenciones.lblValorServicio = FormatoValor(CDbl(lvwServicios.SelectedItem.SubItems(3)) * CDbl(lvwServicios.SelectedItem.SubItems(2)), "", gintDecimalesMoneda)
            Unload Me
        End If
    Case "Cerrar" '////cerrar
        Unload Me
    Case "Buscar" '/////buscar
        If optCriterios(0).Value = 1 Then '/////////////// codigo
            mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  And  Tllr_Servicio_Modelo.id_Servicio LIKE '" & MatchMode(txtCodigo, cboCoincidir.Text, apSqlServer) & "' "
            ServiciosdelModelo mstrWhere, "Order By Tllr_Servicio_Modelo.Id_Servicio"
        ElseIf optCriterios(1).Value = 1 Then '////////////////////des cripcion
            mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  And Tllr_Servicio.Descripcion LIKE '" & MatchMode(Me.txtDes, cboCoincidir.Text, apSqlServer) & "' "
            ServiciosdelModelo mstrWhere, " Order by Descripcion"
        Else
            mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  "
            ServiciosdelModelo mstrWhere, ""
        End If
    End Select
End Select

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If optCriterios(0).Value = 1 Then '/////////////// codigo
         mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  And  Tllr_Servicio_Modelo.id_Servicio LIKE '" & MatchMode(txtCodigo, cboCoincidir.Text, apSqlServer) & "' "
         ServiciosdelModelo mstrWhere, "Order By Tllr_Servicio_Modelo.Id_Servicio"
     ElseIf optCriterios(1).Value = 1 Then '////////////////////des cripcion
         mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  And Tllr_Servicio.Descripcion LIKE '" & MatchMode(txtDes, cboCoincidir.Text, apSqlServer) & "' "
         ServiciosdelModelo mstrWhere, " Order by Descripcion"
     Else
         mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  "
         ServiciosdelModelo mstrWhere, ""
     End If
End If
End Sub

Private Sub txtDes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If optCriterios(0).Value = 1 Then '/////////////// codigo
         mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  And  Tllr_Servicio_Modelo.id_Servicio LIKE '" & MatchMode(txtCodigo, cboCoincidir.Text, apSqlServer) & "' "
         ServiciosdelModelo mstrWhere, "Order By Tllr_Servicio_Modelo.Id_Servicio"
     ElseIf optCriterios(1).Value = 1 Then '////////////////////des cripcion
         mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  And Tllr_Servicio.Descripcion LIKE '" & MatchMode(txtDes, cboCoincidir.Text, apSqlServer) & "' "
         ServiciosdelModelo mstrWhere, " Order by Descripcion"
     Else
         mstrWhere = " Where Tllr_Servicio_Modelo.Id_Marca = '" & lblMarca.Tag & "' AND Tllr_Servicio_Modelo.Id_Modelo = '" & lblModelo.Tag & "'  "
         ServiciosdelModelo mstrWhere, ""
     End If
End If
End Sub
