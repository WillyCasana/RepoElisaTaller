VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBuscaVehiculo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Vehículo"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   Icon            =   "frmBuscaVehiculo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlbCliente 
      Height          =   330
      Left            =   7320
      TabIndex        =   12
      Top             =   1180
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
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   15
      TabIndex        =   0
      Top             =   -45
      Width           =   7950
      Begin VB.TextBox txtVin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3960
         MaxLength       =   30
         TabIndex        =   22
         Top             =   240
         Width           =   2775
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "VIN"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   1280
         Width           =   975
      End
      Begin VB.TextBox txtColExt 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txtModelo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   17
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   16
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1560
         Width           =   4455
      End
      Begin VB.TextBox txtNroRecord 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "10"
         Top             =   840
         Width           =   555
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   7
         Top             =   1280
         Width           =   1095
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Modelo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Marca "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Patente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtPatente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin MSComctlLib.ImageList ImgBarraHerramienta 
         Left            =   7305
         Top             =   1875
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   23
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":179A
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":18AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":1D04
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":215C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":25B4
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":26C6
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":27D8
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":28EA
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":29FC
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":2B0E
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":2C20
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":2D32
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":2E44
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":2F56
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":3068
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":317A
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":328C
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":339E
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":34B0
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":35C2
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":3A14
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":3E66
               Key             =   "Copiar"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaVehiculo.frx":3F78
               Key             =   "Salir"
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.UpDown updNroRecord 
         Height          =   315
         Left            =   7080
         TabIndex        =   10
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   5
         BuddyControl    =   "txtNroRecord"
         BuddyDispid     =   196616
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
      Begin MSComctlLib.Toolbar tlbMarca 
         Height          =   330
         Left            =   2580
         TabIndex        =   13
         Top             =   555
         Width           =   450
         _ExtentX        =   794
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
      End
      Begin MSComctlLib.Toolbar tlbModelo 
         Height          =   330
         Left            =   5760
         TabIndex        =   14
         Top             =   570
         Width           =   465
         _ExtentX        =   820
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
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   2595
         TabIndex        =   20
         Top             =   1275
         Width           =   465
         _ExtentX        =   820
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
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro. de Registros :"
         Height          =   195
         Index           =   1
         Left            =   6480
         TabIndex        =   11
         Top             =   600
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " :"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   2
         Top             =   1725
         Width           =   90
      End
   End
   Begin MSComctlLib.ListView lvwResultado 
      Height          =   3060
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   5398
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
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Patente"
         Text            =   "Patente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Marca"
         Text            =   "Marca"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Modelo"
         Text            =   "Modelo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Exterior"
         Text            =   "Color "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Cliente"
         Text            =   "Cliente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "IDCLI"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotones 
      Height          =   330
      Left            =   4965
      TabIndex        =   8
      Top             =   5145
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      ButtonWidth     =   1720
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aceptar"
            Key             =   "Seleccionar"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            Key             =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar"
            Key             =   "Cerrar"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBuscaVehiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoPrincipal As New ADODB.Recordset
Dim blnSw As Boolean
Dim mstrSql As String
Dim mstrWhere As String
Dim mstrOrden As String
Dim lsiItem As ListItem


Sub FillVehiculos(strCondicion As String, strOrdenamiento As String)

lvwResultado.ListItems.Clear

mstrSql = "SELECT TOP " & updNroRecord.Value & " Tllr_Vehiculo_Cliente.Patente,"
mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor AS IDCLI,Tllr_vehiculo_Cliente.Vin,"
mstrSql = mstrSql & " Glbl_Marca.Descripcion AS Marca,"
mstrSql = mstrSql & " Glbl_Modelo.Descripcion AS Modelo,"
mstrSql = mstrSql & " Glbl_Color_Exterior.Descripcion AS ColorE,"
'mstrSql = mstrSql & " Glbl_Color_Interior.Descripcion AS ColorI,"
mstrSql = mstrSql & " Glbl_Cliente_Proveedor.Razon_Social AS Cliente"
mstrSql = mstrSql & " FROM Glbl_Marca RIGHT OUTER JOIN Glbl_Modelo ON Glbl_Marca.Id_Marca = Glbl_Modelo.Id_Marca RIGHT OUTER JOIN Tllr_Vehiculo_Cliente LEFT OUTER JOIN Glbl_Color_Exterior ON Tllr_Vehiculo_Cliente.Id_Color_Exterior = Glbl_Color_Exterior.Id_Color_Exterior LEFT OUTER JOIN Glbl_Cliente_Proveedor ON Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor = Glbl_Cliente_Proveedor.Id_Cliente_Proveedor ON Glbl_Modelo.Id_Marca = Tllr_Vehiculo_Cliente.Id_Marca AND Glbl_Modelo.Id_Modelo = Tllr_Vehiculo_Cliente.Id_Modelo "
mstrSql = mstrSql & strCondicion & " " & strOrdenamiento
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            While Not .EOF
                Set lsiItem = lvwResultado.ListItems.Add(, , !Patente)
                lsiItem.SubItems(1) = ValorNulo(!Marca)
                lsiItem.SubItems(2) = ValorNulo(!Modelo)
                lsiItem.SubItems(3) = ValorNulo(!ColorE)
'                lsiItem.SubItems(4) = ValorNulo(!ColorI)
                lsiItem.SubItems(4) = ValorNulo(!Cliente)
                lsiItem.SubItems(5) = ValorNulo(!IDCLI)
                .MoveNext
            Wend
        End If
    End With
End If
End Sub







Private Sub cckCriterios_Click(Index As Integer)
If cckCriterios(0).Value = 1 Then '//////PATENTE
    txtPatente.Enabled = True
    txtPatente.SetFocus
Else
    txtPatente.Enabled = False
    txtPatente.Text = ""
End If

If cckCriterios(1).Value = 1 Then '//////MARCA
    txtMarca.Enabled = True
    txtMarca.SetFocus
Else
    txtMarca.Enabled = False
    txtMarca.Text = ""
End If
If cckCriterios(2).Value = 1 Then '//////MODELO
    txtModelo.Enabled = True
    txtModelo.SetFocus
Else
    txtModelo.Enabled = False
    txtModelo.Text = ""
End If

If cckCriterios(3).Value = 1 Then '//////CLIENTE
    txtCliente.Enabled = True
    txtCliente.SetFocus
Else
    txtCliente.Enabled = False
    txtCliente.Text = ""
End If

If cckCriterios(4).Value = 1 Then '//////VIN
    txtVin.Enabled = True
    txtVin.SetFocus
Else
    txtVin.Enabled = False
    txtVin.Text = ""
End If

If cckCriterios(5).Value = 1 Then '//////COLOR EXTERIOR
    txtColExt.Enabled = True
    txtColExt.SetFocus
Else
    txtColExt.Enabled = False
    txtColExt.Text = ""
End If


End Sub




Private Sub Form_Activate()
If blnSw Then
    Me.Tag = gstrProcedencia
    blnSw = False
End If
End Sub

Private Sub Form_Load()
blnSw = True
Me.cckCriterios(0).Caption = gstrNombrePatente
Me.lvwResultado.ColumnHeaders(1).Text = gstrNombrePatente
updNroRecord.Value = gintNroRecDefectoQry
Screen.MousePointer = 0
End Sub
Private Sub lvwResultado_DblClick()
If gstrProcedencia = "Movimientos" Then
    If lvwResultado.ListItems.Count > 0 Then
        With frmRecepcion
            .txtPatente = lvwResultado.SelectedItem
            .DatosVehiculo .txtPatente
            Unload Me
        End With
    End If
ElseIf gstrProcedencia = "Mantenedor" Then
    gstrBusca = lvwResultado.SelectedItem
    Unload Me
End If
End Sub
Private Sub tlbBotones_ButtonClick(ByVal Button As MSComctlLib.Button)
mstrWhere = ""
mstrOrden = ""
gstrProcedencia = Me.Tag
Select Case Button.Key
Case "Seleccionar"
    If gstrProcedencia = "Movimientos" Then
        If lvwResultado.ListItems.Count > 0 Then
            With frmRecepcion
                .txtPatente = lvwResultado.SelectedItem
                
                If ValidaCliente(lvwResultado.SelectedItem.SubItems(5)) Then
                'kjcv 09.01.14
                    If ConsultaPatente(txtPatente) = True Then
                        MsgBox "No hay Cupo en el Taller...", vbCritical, "Elisa"
                        .DatosVehiculo .txtPatente
                    Else
                        .DatosVehiculo .txtPatente
                    End If
                End If
'                .DatosVehiculo .txtPatente
                Unload Me
            End With
        End If
    ElseIf gstrProcedencia = "Mantenedor" Then
        If lvwResultado.ListItems.Count > 0 Then
            gstrBusca = lvwResultado.SelectedItem
            Unload Me
        End If
    ElseIf gstrProcedencia = "ReservaHora" Then
        If Me.lvwResultado.ListItems.Count > 0 Then
            gstrBusca = Me.lvwResultado.SelectedItem
            Unload Me
        End If
    ElseIf gstrProcedencia = "Campañas" Then
        If Me.lvwResultado.ListItems.Count > 0 Then
            gstrBusca = Me.lvwResultado.SelectedItem
            Unload Me
        End If
    End If
    
Case "Buscar"
    With Me
        If cckCriterios(0).Value = 1 Then ' //////////PATENTE
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " And Tllr_Vehiculo_Cliente.Patente LIKE '" & IIf(txtPatente <> "", txtPatente & "%", "%") & "'"
                mstrOrden = mstrOrden & ",Tllr_Vehiculo_Cliente.Patente"
            Else
                mstrWhere = "WHERE Tllr_Vehiculo_Cliente.Patente LIKE '" & IIf(txtPatente <> "", txtPatente & "%", "%") & "'"
                mstrOrden = "Order by Tllr_Vehiculo_Cliente.Patente"
            End If
        End If
'////////////////////////////////////////////////////////////////////////////////////////////////////
        If cckCriterios(1).Value = 1 Then ' //////////MARCA
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " And Glbl_Marca.Descripcion LIKE '" & IIf(txtMarca <> "", txtMarca & "%", "%") & "'"
                mstrOrden = mstrOrden & ",Glbl_Marca.Descripcion"
            Else
                mstrWhere = "WHERE Glbl_Marca.Descripcion LIKE '" & IIf(txtMarca <> "", txtMarca & "%", "%") & "'"
                mstrOrden = "Order by Glbl_Marca.Descripcion"
            End If
        End If
'////////////////////////////////////////////////////////////////////////////////////////////////////
        If cckCriterios(2).Value = 1 Then ' //////////MODELO
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " And Glbl_Modelo.Descripcion LIKE '" & IIf(txtModelo <> "", txtModelo & "%", "%") & "'"
                mstrOrden = mstrOrden & ",Glbl_Modelo.Descripcion"
            Else
                mstrWhere = "WHERE Glbl_Modelo.Descripcion LIKE  '" & IIf(txtModelo <> "", txtModelo & "%", "%") & "'"
                mstrOrden = "Order by Glbl_Modelo.Descripcion"
            End If
        End If
'////////////////////////////////////////////////////////////////////////////////////////////////////
        If cckCriterios(3).Value = 1 Then ' //////////CLIENTE
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " And Glbl_Cliente_Proveedor.Razon_Social LIKE '" & IIf(txtCliente <> "", txtCliente & "%", "%") & "'"
                mstrOrden = mstrOrden & ",Glbl_Cliente_Proveedor.Razon_Social"
            Else
                mstrWhere = "Where Glbl_Cliente_Proveedor.Razon_Social LIKE '" & IIf(txtCliente <> "", txtCliente & "%", "%") & "'"
                mstrOrden = "Order by Glbl_Cliente_Proveedor.Razon_Social"
            End If
        End If
'////////////////////////////////////////////////////////////////////////////////////////////////////
'        If cckCriterios(4).Value = 1 Then ' //////////COLOR INTERIOR
'            If mstrWhere <> "" Then
'
'
''            Glbl_Color_Exterior.Descripcion AS ColorE,"
''mstrSql = mstrSql & " Glbl_Color_Interior.Descripcion"
'                mstrWhere = mstrWhere & " And Glbl_Color_Interior.Descripcion LIKE '" & IIf(txtColInt <> "", txtColInt & "%", "%") & "'"
'                mstrOrden = mstrOrden & ",Glbl_Color_Interior.Descripcion"
'            Else
'                mstrWhere = "Where Glbl_Color_Interior.Descripcion LIKE '" & IIf(txtColInt <> "", txtColInt & "%", "%") & "'"
'                mstrOrden = "Order by Glbl_Color_Interior.Descripcion"
'            End If
'        End If
'////////////////////////////////////////////////////////////////////////////////////////////////////
        
        If cckCriterios(4).Value = 1 Then ' //////////VIN
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " And Tllr_Vehiculo_Cliente.Vin LIKE '%" & IIf(Me.txtVin <> "", txtVin & "%", "%") & "'"
                mstrOrden = mstrOrden & ",Tllr_Vehiculo_Cliente.Vin"
            Else
                mstrWhere = "Where Tllr_Vehiculo_Cliente.Vin LIKE '%" & IIf(txtVin <> "", txtVin & "%", "%") & "'"
                mstrOrden = "Order by Tllr_Vehiculo_Cliente.Vin"
            End If
        End If
        
        If cckCriterios(5).Value = 1 Then ' //////////COLOR EXTERIOR
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " And Glbl_Color_Exterior.Descripcion LIKE '" & IIf(txtColExt <> "", txtColExt & "%", "%") & "'"
                mstrOrden = mstrOrden & ",Glbl_Color_Exterior.Descripcion"
            Else
                mstrWhere = "Where Glbl_Color_Exterior.Descripcion LIKE '" & IIf(txtColExt <> "", txtColExt & "%", "%") & "'"
                mstrOrden = "Order by Glbl_Color_Exterior.Descripcion"
            End If
        End If
    End With
    If mstrWhere <> "" Then
        FillVehiculos mstrWhere, mstrOrden
    Else
        FillVehiculos "", ""
    End If
Case "Cerrar"
    Unload Me
End Select
End Sub

Private Sub tlbCliente_ButtonClick(ByVal Button As MSComctlLib.Button)

If cckCriterios(3).Value = 1 Then
    If Button.Key = "Buscar" Then
        gstrProcedencia = "Buscar"
        gstrBusca = apfFormulario.BuscarRegistros(Conexion, "GLBL_CLIENTE_PROVEEDOR", "Id_Cliente_Proveedor", "Razon_Social", "Buscar Cliente")
        If gstrBusca <> "" Then
            If gstrProcedencia = "Buscar" Then
                With frmBuscaVehiculo
                    .txtCliente.Tag = gstrBusca
                    .txtCliente = ClienteD(gstrBusca)
                End With
            ElseIf gstrProcedencia = "Mantenedor" Then
                With frmMantenedorVehiculoCliente
                    .lblCliente.Tag = gstrBusca
                    .lblCliente.Caption = ClienteD(gstrBusca)
                End With
            End If
            
        End If
    End If
End If
End Sub
Private Sub tlbMarca_ButtonClick(ByVal Button As MSComctlLib.Button)
If cckCriterios(1).Value = 1 Then
    If Button.Key = "Buscar" Then
        gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Glbl_Marca", "Id_Marca", "Descripcion", "Buscar Marca")
        If gstrBusca <> "" Then
            txtMarca.Tag = gstrBusca
            txtMarca = MarcaD(gstrBusca)
        End If
    End If
End If
End Sub

Private Sub tlbModelo_ButtonClick(ByVal Button As MSComctlLib.Button)
If cckCriterios(2).Value = 1 Then
    If Button.Key = "Buscar" Then
        gstrBusca = apfFormulario.BuscarRegistrosModelo(Conexion, "Glbl_Modelo", "Id_Modelo", "Id_Marca", "Descripcion", "Buscar Modelo", txtMarca)
        If gstrBusca <> "" Then
            txtModelo.Tag = gstrBusca
            txtModelo = ModeloD(txtMarca.Tag, gstrBusca)
        End If
    End If
End If
End Sub


Private Sub txtCliente_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub


Private Sub txtPatente_KeyPress(KeyAscii As Integer)
'If gstrValidaPatente = "S" Then
'    KeyAscii = CheckIdCar(txtPatente.SelStart, mdLLNNNN, UpCaseLetter(KeyAscii))
'End If
'KeyAscii = UpCaseLetter(KeyAscii)
'kjcv 24-01-12 Valida Letras y numeros
If (KeyAscii <> 8) And Not (KeyAscii >= 48 And KeyAscii <= 57) And Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
    KeyAscii = 0: Beep
Else
    KeyAscii = UpCaseLetter(KeyAscii)
End If

End Sub
