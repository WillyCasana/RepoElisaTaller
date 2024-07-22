VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddServicioTer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar Servicio de Tercero"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   Icon            =   "frmAddServicioTer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNroRecord 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "10"
      Top             =   1650
      Width           =   540
   End
   Begin VB.TextBox txtCodigo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Top             =   735
      Width           =   2235
   End
   Begin VB.TextBox txtDes 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Top             =   1185
      Width           =   5130
   End
   Begin VB.ComboBox cboCoincidir 
      Height          =   315
      ItemData        =   "frmAddServicioTer.frx":0442
      Left            =   1305
      List            =   "frmAddServicioTer.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1635
      Width           =   2220
   End
   Begin VB.CheckBox optCriterios 
      Caption         =   "Código"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   5
      Top             =   765
      Width           =   990
   End
   Begin VB.Frame fmeServicios 
      Caption         =   "Listado Servicios"
      Height          =   3735
      Left            =   45
      TabIndex        =   0
      Top             =   1980
      Width           =   7440
      Begin MSComctlLib.ListView lvwServicios 
         Height          =   3465
         Left            =   30
         TabIndex        =   2
         Top             =   195
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   6112
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
         Left            =   6570
         Top             =   1140
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
               Picture         =   "frmAddServicioTer.frx":04A5
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":05B7
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":0A0F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":0E67
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":12BF
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":13D1
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":14E3
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":15F5
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":1707
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":1819
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":192B
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":1A3D
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":1B4F
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":1C61
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":1D73
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":1E85
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":1F97
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":20A9
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":21BB
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":22CD
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":271F
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddServicioTer.frx":2B71
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar tlbOpciones 
      Height          =   330
      Index           =   0
      Left            =   90
      TabIndex        =   3
      Top             =   5775
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
      Left            =   4530
      TabIndex        =   4
      Top             =   5790
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   582
      ButtonWidth     =   1693
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
            Object.ToolTipText     =   "Agrega Servicio Nuevo"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agregar"
            Key             =   "Agregar"
            Object.ToolTipText     =   "Quitar Servicio"
            ImageIndex      =   20
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
      TabIndex        =   9
      Top             =   1230
      Width           =   1305
   End
   Begin MSComctlLib.Toolbar tlbOpciones 
      Height          =   330
      Index           =   2
      Left            =   5580
      TabIndex        =   12
      Top             =   60
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
            Object.ToolTipText     =   "Agrega Servicio Nuevo"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.UpDown updNroRecord 
      Height          =   315
      Left            =   6210
      TabIndex        =   13
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
      Left            =   4320
      TabIndex        =   15
      Top             =   1695
      Width           =   1320
   End
   Begin VB.Label lblProveedor 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1320
      TabIndex        =   11
      Top             =   60
      Width           =   4215
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
      Left            =   315
      TabIndex        =   10
      Top             =   1695
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Proveedor :"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "frmAddServicioTer"
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


Sub ServiciosTercero(strCondicion As String, strOrden As String)
    
lvwServicios.ListItems.Clear
mstrSql = "SELECT  TOP " & CStr(updNroRecord.Value) & " Id_Servicio_Tercero, "
mstrSql = mstrSql & " Descripcion,"
mstrSql = mstrSql & " Tiempo, "
mstrSql = mstrSql & " Valor "
mstrSql = mstrSql & " From Tllr_Servicio_Tercero "
mstrSql = mstrSql & strCondicion & " " & strOrden
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            While Not .EOF
                Set lsiItem = lvwServicios.ListItems.Add(, , !Id_Servicio_Tercero)
                lsiItem.SubItems(1) = !Descripcion
                lsiItem.SubItems(2) = !TIEMPO
                lsiItem.SubItems(3) = Format(!Valor, "###,##0")
                .MoveNext
            Wend
        End If
    End With
End If

End Sub
Private Sub Form_Activate()

If mblnSW Then
    cboCoincidir.ListIndex = 0
    mblnSW = False
End If


End Sub

Private Sub Form_Load()

mblnSW = True
updNroRecord.Value = gintNroRecDefectoQry

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
    Select Case Button.Key
    Case "Agregar" '////Agregar
        For intContador = 1 To lvwServicios.ListItems.Count
            Set lvwServicios.SelectedItem = lvwServicios.ListItems(intContador)
            If lvwServicios.ListItems(intContador).Checked = True Then
                Set itmFound = frmRecepcion.lvwServiciosTerceros.FindItem(lvwServicios.SelectedItem, lvwText, , 0)
                If itmFound Is Nothing Then   ' Si no hay coincidencia                                    ' usuario y sale.
                    Set itmFound = frmRecepcion.lvwServiciosTerceros.ListItems.Add(, , lvwServicios.ListItems(intContador))
                    Set frmRecepcion.lvwServiciosTerceros.SelectedItem = itmFound
                    itmFound.SubItems(1) = lvwServicios.ListItems(intContador).SubItems(1)
                    itmFound.SubItems(2) = lblProveedor.Caption
                    itmFound.SubItems(3) = lblProveedor.Tag
                    itmFound.SubItems(4) = lvwServicios.ListItems(intContador).SubItems(2)
                    itmFound.SubItems(5) = lvwServicios.ListItems(intContador).SubItems(3)
                    itmFound.SubItems(6) = Format(0, "#0.0")
                    itmFound.SubItems(7) = Format(0, "###,##0.0")
                    itmFound.SubItems(8) = TraeCargoDes(gstrIdCargo)
                    itmFound.SubItems(9) = gstrIdCargo
                    itmFound.SubItems(10) = Format(frmRecepcion.CalculoSubTotal(mcFichaTerceros), "###,##0.0")
                End If
            End If
        Next
        Unload Me
    Case "Cerrar" '////cerrar
        Unload Me
    Case "Buscar" '/////buscar
        If lblProveedor.Tag <> "" Then
            If optCriterios(0).Value = 1 Then '/////////////// codigo
                mstrWhere = " Where Id_Proveedor = '" & lblProveedor.Tag & "' and Id_Servicio_Tercero Like '" & MatchMode(txtCodigo, cboCoincidir, apSqlServer) & "'"
                ServiciosTercero mstrWhere, "Order By Tllr_Servicio_Modelo.Id_Servicio"
            ElseIf optCriterios(1).Value = 1 Then '////////////////////des cripcion
                mstrWhere = " Where Id_Proveedor = '" & lblProveedor.Tag & "' and Descripcion Like '" & MatchMode(txtDes, cboCoincidir, apSqlServer) & "'"
                ServiciosTercero mstrWhere, " Order by Descripcion"
            Else
                mstrWhere = " Where Id_Proveedor = '" & lblProveedor.Tag & "' "
                ServiciosTercero mstrWhere, ""
            End If
        Else
            MsgBox "Seleccione un Proveedor de Servicios Externos"
        End If
    End Select
Case 2
    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Tllr_Proveedor_Servicio", "Id_Proveedor", "Nombre", "Buscar Proveedor de Servicio")
    lblProveedor.Tag = gstrBusca
    lblProveedor.Caption = ProveedorS(gstrBusca)
Case Else
    DoEvents
End Select

End Sub
