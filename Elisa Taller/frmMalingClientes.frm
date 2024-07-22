VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmMailingClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mailing Clientes"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "frmMalingClientes.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport rptOT 
      Left            =   3915
      Top             =   6615
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir Listado"
      Height          =   360
      Left            =   8025
      TabIndex        =   24
      Top             =   7155
      Width           =   1680
   End
   Begin VB.Frame Frame2 
      Height          =   3450
      Left            =   60
      TabIndex        =   5
      Top             =   -15
      Width           =   11370
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   10350
         TabIndex        =   39
         Top             =   630
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   9360
         TabIndex        =   38
         Top             =   630
         Width           =   915
      End
      Begin VB.CheckBox Check1 
         Caption         =   "No ha Venido al Taller"
         Height          =   195
         Left            =   135
         TabIndex        =   37
         Top             =   1035
         Width           =   2085
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2850
         Left            =   7290
         TabIndex        =   36
         Top             =   540
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   5027
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame3 
         Caption         =   "Estado"
         Height          =   525
         Left            =   135
         TabIndex        =   30
         Top             =   2160
         Width           =   3870
         Begin VB.OptionButton optVigente 
            Caption         =   "Vigente"
            Height          =   195
            Left            =   888
            TabIndex        =   34
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optTodas 
            Caption         =   "Todas"
            Height          =   195
            Left            =   75
            TabIndex        =   33
            Top             =   240
            Value           =   -1  'True
            Width           =   810
         End
         Begin VB.OptionButton optCerrada 
            Caption         =   "Emitidas"
            Height          =   195
            Left            =   2739
            TabIndex        =   32
            Top             =   240
            Width           =   945
         End
         Begin VB.OptionButton optLiquidada 
            Caption         =   "Liquidada"
            Height          =   195
            Left            =   1746
            TabIndex        =   31
            Top             =   240
            Width           =   990
         End
      End
      Begin MSComctlLib.ListView lsvtipoOt 
         Height          =   1785
         Left            =   4140
         TabIndex        =   29
         Top             =   1620
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   3149
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo OT"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Frame Frame1 
         Height          =   510
         Left            =   135
         TabIndex        =   25
         Top             =   2790
         Width           =   3900
         Begin VB.OptionButton OptAmbas 
            Caption         =   "Ambas"
            Height          =   240
            Left            =   2775
            TabIndex        =   28
            Top             =   180
            Value           =   -1  'True
            Width           =   840
         End
         Begin VB.OptionButton optCarroceria 
            Caption         =   "Carrocería"
            Height          =   240
            Left            =   90
            TabIndex        =   27
            Top             =   180
            Width           =   1215
         End
         Begin VB.OptionButton optMecanica 
            Caption         =   "Mecánica"
            Height          =   240
            Left            =   1485
            TabIndex        =   26
            Top             =   180
            Width           =   1065
         End
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Fecha Emisión (Final)"
         Height          =   195
         Index           =   7
         Left            =   2055
         TabIndex        =   23
         Top             =   1545
         Width           =   1920
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Recepcionista"
         Height          =   195
         Index           =   5
         Left            =   4140
         TabIndex        =   19
         Top             =   945
         Width           =   1395
      End
      Begin VB.TextBox txtRecepcionista 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4140
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1215
         Width           =   2820
      End
      Begin VB.TextBox txtPatente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   105
         MaxLength       =   6
         TabIndex        =   13
         Top             =   525
         Width           =   1065
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Patente"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   330
         Width           =   900
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Marca "
         Height          =   195
         Index           =   2
         Left            =   1185
         TabIndex        =   11
         Top             =   315
         Width           =   915
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Modelo"
         Height          =   195
         Index           =   3
         Left            =   4110
         TabIndex        =   10
         Top             =   345
         Width           =   885
      End
      Begin VB.TextBox txtMarca 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1185
         MaxLength       =   50
         TabIndex        =   8
         Top             =   540
         Width           =   2880
      End
      Begin VB.TextBox txtModelo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4110
         MaxLength       =   50
         TabIndex        =   7
         Top             =   555
         Width           =   2880
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Fecha Emisión (Inicial)"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   1545
         Width           =   1920
      End
      Begin MSComctlLib.ImageList ImgBarraHerramienta 
         Left            =   10485
         Top             =   2730
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
               Picture         =   "frmMalingClientes.frx":000C
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":011E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":0576
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":09CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":0E26
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":0F38
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":104A
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":115C
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":126E
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":1380
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":1492
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":15A4
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":16B6
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":17C8
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":18DA
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":19EC
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":1AFE
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":1C10
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":1D22
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":1E34
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":2286
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMalingClientes.frx":26D8
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbMarca 
         Height          =   330
         Left            =   3525
         TabIndex        =   15
         Top             =   210
         Width           =   495
         _ExtentX        =   873
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
         Enabled         =   0   'False
      End
      Begin MSComctlLib.Toolbar tlbModelo 
         Height          =   330
         Left            =   6465
         TabIndex        =   16
         Top             =   225
         Width           =   510
         _ExtentX        =   900
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
         Enabled         =   0   'False
      End
      Begin MSComctlLib.Toolbar tlbRecep 
         Height          =   330
         Left            =   6390
         TabIndex        =   20
         Top             =   900
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
         Enabled         =   0   'False
      End
      Begin MSComCtl2.DTPicker pckFechaDesde 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24576001
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckFechaHasta 
         Height          =   315
         Left            =   2055
         TabIndex        =   22
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24576001
         CurrentDate     =   36776
      End
      Begin MSComCtl2.UpDown updNroRecord 
         Height          =   315
         Left            =   10770
         TabIndex        =   14
         Top             =   -375
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   393216
         Value           =   5
         BuddyControl    =   "txtNroRecord"
         BuddyDispid     =   196628
         OrigLeft        =   10950
         OrigTop         =   525
         OrigRight       =   11190
         OrigBottom      =   840
         Max             =   999
         Min             =   5
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtNroRecord 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10230
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "10"
         Top             =   -375
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label4 
         Caption         =   "Km. Hasta"
         Height          =   240
         Left            =   10395
         TabIndex        =   41
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label3 
         Caption         =   "Km. Desde"
         Height          =   285
         Left            =   9360
         TabIndex        =   40
         Top             =   270
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "Asistencia a Revisiones"
         Height          =   510
         Left            =   7650
         TabIndex        =   35
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Registros"
         Height          =   195
         Index           =   8
         Left            =   10260
         TabIndex        =   17
         Top             =   -570
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdBuscarOT 
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   360
      Left            =   6225
      TabIndex        =   0
      Top             =   7140
      Width           =   1680
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   9750
      TabIndex        =   1
      Top             =   7170
      Width           =   1680
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   3600
      Left            =   90
      TabIndex        =   4
      Top             =   3465
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   6350
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   21
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N° OT"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Estado"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Patente"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cliente"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Marca"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Modelo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Fecha Ingreso"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Recepcionista"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Seccion"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Tipo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Id_Seccion"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "TMEC"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "TCAR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "TOTR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "TTER"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "TREP"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "TMAT"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "TINS"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Text            =   "TNETO"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Text            =   "TIVA"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Text            =   "TOTAL"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   7
      Left            =   1980
      TabIndex        =   3
      Top             =   7200
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registros Encontrados :"
      Height          =   195
      Index           =   6
      Left            =   165
      TabIndex        =   2
      Top             =   7185
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmMailingClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SW As Boolean
'Sub ResumenOT(pstrIdEmpresa As String, _
'            pstrIdSucursal As String, _
'            pstrIdOT As String, _
'            pstrSeccion As String)
'
'gstrSql = "SELECT Top 1 Tllr_OT.Patente, "
'gstrSql = gstrSql & " Glbl_Modelo.Descripcion AS MODELO,"
'gstrSql = gstrSql & " Glbl_Marca.Descripcion AS MARCA,"
'gstrSql = gstrSql & " Glbl_Cliente_Proveedor.Razon_Social AS CLIENTE,"
'gstrSql = gstrSql & " Tllr_OT.Total_Mecanica AS TMEC,"
'gstrSql = gstrSql & " Tllr_OT.Total_Carroceria AS TCAR,"
'gstrSql = gstrSql & " Tllr_OT.Total_Otros AS TOTR,"
'gstrSql = gstrSql & " Tllr_OT.Total_Terceros AS TTER,"
'gstrSql = gstrSql & " Tllr_OT.Total_Repuestos AS TREP,"
'gstrSql = gstrSql & " Tllr_OT.Total_OT AS TNETO,"
'gstrSql = gstrSql & " Tllr_OT.Total_Materiales AS TMAT,"
'gstrSql = gstrSql & " Tllr_OT.Total_Insumos AS TINS,"
'gstrSql = gstrSql & " Tllr_OT.Total_OT_Iva AS TOTAL, "
'gstrSql = gstrSql & " Tllr_OT.Total_IVA AS TIVA,"
'gstrSql = gstrSql & " Tllr_OT.Estado"
'gstrSql = gstrSql & " FROM Tllr_Vehiculo_Cliente RIGHT OUTER JOIN Tllr_OT ON Tllr_Vehiculo_Cliente.Patente = Tllr_OT.Patente LEFT OUTER JOIN Glbl_Cliente_Proveedor ON Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor = Glbl_Cliente_Proveedor.Id_Cliente_Proveedor LEFT OUTER JOIN Glbl_Modelo LEFT OUTER JOIN Glbl_Marca ON Glbl_Modelo.Id_Marca = Glbl_Marca.Id_Marca ON Tllr_Vehiculo_Cliente.Id_Modelo = Glbl_Modelo.Id_Modelo AND Tllr_Vehiculo_Cliente.Id_Marca = Glbl_Modelo.Id_Marca"
'gstrSql = gstrSql & " WHERE (Tllr_OT.Id_Empresa = '" & pstrIdEmpresa & "') AND"
'gstrSql = gstrSql & " (Tllr_OT.Id_Sucursal = '" & pstrIdSucursal & "') AND "
'gstrSql = gstrSql & " (Tllr_OT.Id_OT = '" & pstrIdOT & "') AND"
'gstrSql = gstrSql & " (Tllr_OT.Seccion_OT = '" & pstrSeccion & "')"
'If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
'    If Not gadoPrincipal.BOF And Not gadoPrincipal.EOF Then
'        .MoveFirst
'        With frmResumenOT
'            .lblIdOT = pstrIdOT
'            .lblSeccion = pstrSeccion
'            .lblEstado
'            .lblPatente
'            .lblCliente
'            .lblMarca
'            .lblModelo
'            .lblTotalMec
'            .lblTotalCar
'            .lblTotalOtr
'            .lblTotalTer
'            .lblTotalRep
'            .lblTotalMat
'            .lblTotalIns
'            .lblSubTotal
'            .lblIva
'            .lblTotalOT
'        End With
'    End If
'End If
'
'End Sub

Sub ImprimirConsulta()
Dim Dbsnueva As Database
Dim tabla As DAO.Recordset
Dim i As Integer
Dim GcamBaseTem As String

    'Devuelve la ruta del directorio Windows
    Dim rc As Long
    Dim WinPath As String
    WinPath = Space$(300)
    rc = GetWindowsDirectory(WinPath, 300)
    GcamBaseTem = Trim$(WinPath)
    GcamBaseTem = Mid(GcamBaseTem, 1, Len(GcamBaseTem) - 1) & "\Temp"
    '---------------------------------------
    
    If lvDetalle.ListItems.Count = 0 Then
      MsgBox "No existen elementos en la lista", vbExclamation, "Imprimir"
      Exit Sub
    End If

'    Screen.MousePointer = 11
'    Dim wrkPredeterminado As Workspace
'    Dim prpBucle As Property
'    Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
'    If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
'    Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
'    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (NroOT text,Estado text,Patente text,Cliente text,Marca text,Modelo text,FechaIngreso date,Recepcionista text,Seccion text,Tipo text, TIVA TEXT, TNETO TEXT, TOTAL TEXT)"
'    Set tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
'    For i = 1 To lvDetalle.ListItems.Count
'        Set lvDetalle.SelectedItem = lvDetalle.ListItems(i)
'        tabla.AddNew
'        tabla!NroOT = IIf(lvDetalle.SelectedItem = "", " ", lvDetalle.SelectedItem)
'        tabla!Estado = IIf(lvDetalle.SelectedItem.SubItems(1) = "", " ", lvDetalle.SelectedItem.SubItems(1))
'        tabla!Patente = IIf(lvDetalle.SelectedItem.SubItems(2) = "", " ", lvDetalle.SelectedItem.SubItems(2))
'        tabla!CLIENTE = IIf(lvDetalle.SelectedItem.SubItems(3) = "", " ", lvDetalle.SelectedItem.SubItems(3))
'        tabla!Marca = IIf(lvDetalle.SelectedItem.SubItems(4) = "", " ", lvDetalle.SelectedItem.SubItems(4))
'        tabla!Modelo = IIf(lvDetalle.SelectedItem.SubItems(5) = "", " ", lvDetalle.SelectedItem.SubItems(5))
'        tabla!FechaIngreso = DateValue(IIf(lvDetalle.SelectedItem.SubItems(6) = "", " ", lvDetalle.SelectedItem.SubItems(6)))
'        tabla!Recepcionista = IIf(lvDetalle.SelectedItem.SubItems(7) = "", " ", lvDetalle.SelectedItem.SubItems(7))
'        tabla!Seccion = IIf(lvDetalle.SelectedItem.SubItems(8) = "", " ", lvDetalle.SelectedItem.SubItems(8))
'        tabla!tipo = IIf(lvDetalle.SelectedItem.SubItems(9) = "", " ", lvDetalle.SelectedItem.SubItems(9))
'        tabla!Tiva = IIf(lvDetalle.SelectedItem.SubItems(19) = "", " ", lvDetalle.SelectedItem.SubItems(19))
'        tabla!Tneto = IIf(lvDetalle.SelectedItem.SubItems(18) = "", " ", lvDetalle.SelectedItem.SubItems(18))
'        tabla!TOTAL = IIf(lvDetalle.SelectedItem.SubItems(20) = "", " ", lvDetalle.SelectedItem.SubItems(20))
'        tabla.Update
'    Next i
'   tabla.Close
'
'   With rptOT
'        '//MODIFICADO POR FDO DIAZ EL 29/11/2000
'        '.ReportFileName = "\\POMPEYO_NT\SERINFO\REPORTES\TLLR" & "\ResumenOt.rpt"
'        .ReportFileName = gstrPathReporte & "\ResumenOt.rpt"
'        .WindowTitle = "Informe de Ordenes de Trabajo"
'        .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
'        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
'        .Formulas(1) = "TITULO='RESUMEN VALORIZADO DE OT'"
'        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
'        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
'        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
'        .Formulas(5) = "SUMNETO='" & Me.stbTotales.Panels(2).Text & "'"
'        .Formulas(6) = "SUMIVA='" & Me.stbTotales.Panels(4).Text & "'"
'        .Formulas(7) = "SUMTOTAL='" & Me.stbTotales.Panels(6).Text & "'"
'        .Destination = crptToWindow
'        .Action = True
'   End With
'
'   Dbsnueva.Close
'   Screen.MousePointer = 1

End Sub


Private Sub cckCriterios_Click(Index As Integer)
Select Case Index
Case 0
'    If cckCriterios(Index).Value = 0 Then
'        txtNroOt.Enabled = False
'        txtNroOt = ""
'    Else
'        txtNroOt.Enabled = True
'        txtNroOt.SetFocus
'    End If
Case 1
    If cckCriterios(Index).Value = 0 Then
        txtPatente.Enabled = False
        txtPatente = ""
        
    Else
        txtPatente.Enabled = True
        txtPatente.SetFocus
    End If
Case 2
    If cckCriterios(Index).Value = 0 Then
        tlbMarca.Enabled = False
        txtMarca.Enabled = False
        txtMarca = ""
    Else
        tlbMarca.Enabled = True
        txtMarca = ""
        txtMarca.Enabled = True
        txtMarca.SetFocus
    End If
Case 3
    If cckCriterios(Index).Value = 0 Then
        txtModelo.Enabled = False
        txtModelo = ""
    Else
        txtModelo.Enabled = True
        txtModelo.SetFocus
    End If
Case 4
'    If cckCriterios(Index).Value = 0 Then
'        tlbCliente.Enabled = False
'        txtCliente.Enabled = False
'        txtCliente = ""
'    Else
'        tlbCliente.Enabled = True
'        txtCliente.Enabled = True
'        txtCliente.SetFocus
'    End If
Case 5
    If cckCriterios(Index).Value = 0 Then
        tlbRecep.Enabled = False
        txtRecepcionista.Enabled = False
        txtRecepcionista = ""
    Else
        tlbRecep.Enabled = True
        txtRecepcionista.Enabled = True
        txtRecepcionista.SetFocus
    End If
Case 6
    If cckCriterios(Index).Value = 0 Then
        pckFechaDesde.Enabled = False
    Else
        pckFechaDesde.Enabled = True
        pckFechaDesde.SetFocus
    End If
Case 7
    If cckCriterios(Index).Value = 0 Then
        pckFechaHasta.Enabled = False
    Else
        pckFechaHasta.Enabled = True
        pckFechaHasta.SetFocus
    End If
Case 8
'    If cckCriterios(Index).Value = 0 Then
'        pckLiquida.Enabled = False
'    Else
'        pckLiquida.Enabled = True
'        pckLiquida.SetFocus
'    End If
End Select
End Sub


Private Sub cckTipoOt_Click(Index As Integer)
End Sub

Private Sub cmdBuscarOT_Click()
Dim i As Integer
Dim mstrsql As String
Dim mstrWhere As String
Dim adoTemp As ADODB.Recordset
Dim AdoAux As ADODB.Recordset
Dim itmItem As ListItem
Dim Item As ListItem
Dim mstrEstado As String

    lvDetalle.ListItems.Clear
mstrWhere = ""
With Me
'    If .cckCriterios(0).Value = 1 Then  '////////// nro ot
'        If mstrWhere <> "" Then
'            mstrWhere = mstrWhere & " and Tllr_Ot.Id_Ot LIKE '" & MatchMode(txtNroOt, "Cualquier Parte del Campo", apSqlServer) & "'"
'        Else
'            mstrWhere = " Where Tllr_Ot.Id_Ot LIKE '" & MatchMode(txtNroOt, "Cualquier Parte del Campo", apSqlServer) & "'"
'        End If
'    End If
    
    If .cckCriterios(1).Value = 1 Then  '////////// patente
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Tllr_Ot.PATENTE LIKE '" & MatchMode(.txtPatente, "Comienzo del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Tllr_Ot.PATENTE LIKE '" & MatchMode(.txtPatente, "Comienzo del Campo", apSqlServer) & "'"
        End If
    End If
    
    If .cckCriterios(2).Value = 1 Then  '////////// marca
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Glbl_Marca.Descripcion LIKE '" & MatchMode(.txtMarca, "Comienzo del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Glbl_Marca.Descripcion LIKE '" & MatchMode(.txtMarca, "Comienzo del Campo", apSqlServer) & "'"
        End If
    End If
    
    If .cckCriterios(3).Value = 1 Then  '////////// modelo
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Glbl_Modelo.Descripcion LIKE '" & MatchMode(.txtModelo, "Comienzo del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Glbl_Modelo.Descripcion LIKE '" & MatchMode(.txtModelo, "Comienzo del Campo", apSqlServer) & "'"
        End If
    End If
    
'    If .cckCriterios(4).Value = 1 Then  '////////// cliente
'        If mstrWhere <> "" Then
'            mstrWhere = mstrWhere & " and Glbl_Cliente_Proveedor.Razon_Social LIKE '" & MatchMode(.txtCliente, "Comienzo del Campo", apSqlServer) & "'"
'        Else
'            mstrWhere = " Where Glbl_Cliente_Proveedor.Razon_Social LIKE '" & MatchMode(.txtCliente, "Comienzo del Campo", apSqlServer) & "'"
'        End If
'    End If
    
    If .cckCriterios(5).Value = 1 Then  '////////// recepcionista
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Tllr_Mecanicos.Nombre LIKE '" & MatchMode(.txtRecepcionista, "Comienzo del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Tllr_Mecanicos.Nombre LIKE '" & MatchMode(.txtRecepcionista, "Comienzo del Campo", apSqlServer) & "'"
        End If
    End If
    
    If .cckCriterios(6).Value = 1 Then  '////////// fecha inicio
        If .cckCriterios(7).Value = 1 Then  '////////// fecha termino
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " AND fecha_emision between '" & pckFechaDesde.Value & "' and '" & pckFechaHasta.Value & "'"
            Else
                mstrWhere = " WHERE fecha_emision between '" & pckFechaDesde.Value & "' and '" & pckFechaHasta.Value & "'"
            End If
        Else
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " AND fecha_emision = '" & pckFechaDesde.Value & "' "
            Else
                mstrWhere = " WHERE fecha_emision = '" & pckFechaDesde.Value & "' "
            End If
        End If
    Else
        If .cckCriterios(7).Value = 1 Then  '////////// fecha termino
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " AND fecha_emision = '" & pckFechaHasta.Value & "' "
            Else
                mstrWhere = " WHERE fecha_emision = '" & pckFechaHasta.Value & "' "
            End If
        End If
    End If
    
    
    If mstrWhere <> "" Then
        mstrWhere = mstrWhere & " AND TLLR_OT.ID_EMPRESA= '" & gstrIdEmpresa & "' AND TLLR_OT.ID_SUCURSAL='" & gstrIdSucursal & "' "
    Else
        mstrWhere = " WHERE TLLR_OT.ID_EMPRESA= '" & gstrIdEmpresa & "' AND TLLR_OT.ID_SUCURSAL='" & gstrIdSucursal & "' "
    End If
    
    
    If Me.optCarroceria.Value = True Then ' POR CARROCERIA
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Tllr_Ot.Seccion_Ot = 'C'"
        Else
            mstrWhere = " Where Tllr_Ot.Seccion_Ot = 'C'"
        End If
    End If
    
    If Me.optMecanica.Value = True Then ' POR MECANICA
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Tllr_Ot.Seccion_Ot = 'M'"
        Else
            mstrWhere = " Where Tllr_Ot.Seccion_Ot = 'M'"
        End If
    End If
    
    '//////////////////estado
    If optTodas.Value = True Then
        mstrEstado = "IN ('V','A','L','C','F','B')"
    ElseIf optVigente.Value = True Then
        mstrEstado = "IN ('V','A')"
    ElseIf optLiquidada.Value = True Then
        mstrEstado = "IN ('L','C')"
    ElseIf optCerrada.Value = True Then
        mstrEstado = "IN ('F','B')"
'    ElseIf optNula.Value = True Then
'        mstrEstado = "IN ('N')"
    End If
    
    If mstrEstado <> "" Then
        mstrWhere = mstrWhere & " And Tllr_OT.Estado  " & mstrEstado
    End If
    
    Dim lsw As Double
    lsw = False
    For i = 1 To Me.lsvtipoOt.ListItems.Count 'R
        If Me.lsvtipoOt.ListItems(i).Checked Then 'Si esta checheada agrega al where
            
                If lsw = False Then 'Si es el primero usa AND
                    If mstrWhere <> "" Then
                        mstrWhere = mstrWhere & " and (Tllr_Ot.Id_Garantia = '" & Me.lsvtipoOt.ListItems(i).ListSubItems(1) & "'"
                    Else
                        mstrWhere = " Where (Tllr_Ot.Id_Garantia = '" & Me.lsvtipoOt.ListItems(i).ListSubItems(1) & "'"
                    End If
                    lsw = True
                Else
                    If mstrWhere <> "" Then
                        mstrWhere = mstrWhere & " OR Tllr_Ot.Id_Garantia = '" & Me.lsvtipoOt.ListItems(i).ListSubItems(1) & "'"
                    Else
                        mstrWhere = " Where Tllr_Ot.Id_Garantia = '" & Me.lsvtipoOt.ListItems(i).ListSubItems(1) & "'"
                    End If
                End If
        End If
    Next
    
    'Si alguna vez paso cierra el parentesis
     If lsw = True Then 'Si es el ultimo entonces cierra parentesis
        mstrWhere = mstrWhere & ")"
    End If
End With
'/////////////////////////////////////////////////////////////////////////////////
    mstrsql = "SELECT  Tllr_OT.Id_OT, "
    mstrsql = mstrsql & " Tllr_OT.Seccion_OT  AS SEC, "
    mstrsql = mstrsql & " Tllr_OT.Patente AS PAT,"
    mstrsql = mstrsql & " Tllr_Vehiculo_Cliente.Id_Marca AS IDMAR,"
    mstrsql = mstrsql & " Glbl_Marca.Descripcion AS MARCA,"
    mstrsql = mstrsql & " Tllr_Vehiculo_Cliente.Id_Modelo AS IDMOD,"
    mstrsql = mstrsql & " Glbl_Modelo.Descripcion AS MODELO,"
    mstrsql = mstrsql & " Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor AS IDCLI,"
    mstrsql = mstrsql & " Glbl_Cliente_Proveedor.Razon_Social AS CLIENTE,"
    mstrsql = mstrsql & " Tllr_OT.Fecha_Emision AS FEC, "
    mstrsql = mstrsql & " Tllr_OT.Estado AS EST, "
    mstrsql = mstrsql & " Tllr_OT.RealizadoPor AS IDREC,"
    mstrsql = mstrsql & " Tllr_Mecanicos.Nombre AS RECEP, "
    mstrsql = mstrsql & " Tllr_OT.Id_Garantia AS IDGAR,"
    mstrsql = mstrsql & " Tllr_Garantias.Descripcion AS GAR,"
    
    mstrsql = mstrsql & " Tllr_OT.Total_Mecanica AS TMEC,"
    mstrsql = mstrsql & " Tllr_OT.Total_Carroceria AS TCAR,"
    mstrsql = mstrsql & " Tllr_OT.Total_Otros AS TOTR,"
    mstrsql = mstrsql & " Tllr_OT.Total_Terceros AS TTER,"
    mstrsql = mstrsql & " Tllr_OT.Total_Repuestos AS TREP,"
    mstrsql = mstrsql & " Tllr_OT.Total_Materiales AS TMAT,"
    mstrsql = mstrsql & " Tllr_OT.Total_Insumos AS TINS,"
    mstrsql = mstrsql & " Tllr_OT.Total_OT AS TNETO,"
    mstrsql = mstrsql & " Tllr_OT.Total_IVA AS TIVA, "
    mstrsql = mstrsql & " Tllr_OT.Total_OT_Iva AS TOTAL "
    
    mstrsql = mstrsql & " FROM Tllr_Garantias RIGHT OUTER JOIN Tllr_OT ON Tllr_Garantias.Id_Garantia = Tllr_OT.Id_Garantia LEFT OUTER JOIN Tllr_Mecanicos ON Tllr_OT.RealizadoPor = Tllr_Mecanicos.Id_Mecanico LEFT OUTER Join Glbl_Modelo LEFT OUTER JOIN Glbl_Marca ON Glbl_Modelo.Id_Marca = Glbl_Marca.Id_Marca RIGHT OUTER JOIN Tllr_Vehiculo_Cliente ON Glbl_Modelo.Id_Modelo = Tllr_Vehiculo_Cliente.Id_Modelo AND Glbl_Modelo.Id_Marca = Tllr_Vehiculo_Cliente.Id_Marca LEFT OUTER Join Glbl_Cliente_Proveedor ON Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor = Glbl_Cliente_Proveedor.Id_Cliente_Proveedor ON Tllr_OT.Patente = Tllr_Vehiculo_Cliente.Patente   "
    
    mstrsql = mstrsql & mstrWhere
    mstrsql = mstrsql & "  ORDER BY ID_OT"
    
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrsql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
    With adoTemp
       If Not .BOF And Not .EOF Then
          While Not .EOF
              Set itmItem = lvDetalle.ListItems.Add(, , !Id_OT)
              itmItem.SubItems(1) = ValorNulo(IIf(!Est = "L", "LIQUIDADA", IIf(!Est = "V", "VIGENTE", IIf(!Est = "N", "NULA", IIf(!Est = "B", "BOLETEADA", "FACTURADA")))))
              itmItem.SubItems(2) = ValorNulo(!Pat)
              itmItem.SubItems(3) = ValorNulo(!CLIENTE)
              itmItem.SubItems(4) = ValorNulo(!Marca)
              itmItem.SubItems(5) = ValorNulo(!Modelo)
              itmItem.SubItems(6) = ValorNulo(!FEC)
              itmItem.SubItems(7) = ValorNulo(!RECEP)
              itmItem.SubItems(8) = ValorNulo(IIf(!Sec = "M", "MECANICA", "CARROCERIA"))
              itmItem.SubItems(9) = ValorNulo(!GAR)
              itmItem.SubItems(10) = ValorNulo(!Sec)
              
              itmItem.SubItems(11) = ValorNulo(!TMEC)
              itmItem.SubItems(12) = ValorNulo(!TCAR)
              itmItem.SubItems(13) = ValorNulo(!TOTR)
              itmItem.SubItems(14) = ValorNulo(!TTER)
              itmItem.SubItems(15) = ValorNulo(!TREP)
              itmItem.SubItems(16) = ValorNulo(!TMAT)
              itmItem.SubItems(17) = ValorNulo(!TINS)
              itmItem.SubItems(18) = FormatoValor(ValorNulo(!Tneto), "$", 0)
              itmItem.SubItems(19) = FormatoValor(ValorNulo(!Tiva), "$", 0)
              itmItem.SubItems(20) = FormatoValor(ValorNulo(!TOTAL), "$", 0)
              adoTemp.MoveNext
          Wend
       End If
    End With
    
    'Ahora crea la linea de Totales
'              Set itmItem = lvDetalle.ListItems.Add(, , "TOTALES :")
'              itmItem.SubItems(18) = TotalSeccion(Me.lvDetalle, 18)
'              itmItem.SubItems(19) = TotalSeccion(Me.lvDetalle, 19)
'              itmItem.SubItems(20) = TotalSeccion(Me.lvDetalle, 20)
'    With Me.stbTotales
'        .Panels(2).Text = FormatoValor(TotalSeccionFormato(lvDetalle, 18), "$", 0)
'        .Panels(4).Text = FormatoValor(TotalSeccionFormato(lvDetalle, 19), "$", 0)
'        .Panels(6).Text = FormatoValor(TotalSeccionFormato(lvDetalle, 20), "$", 0)
'    End With
    End If
    Screen.MousePointer = 1
    lblTotal(7).Caption = lvDetalle.ListItems.Count
    
    
End Sub






Private Sub cmdImprimir_Click()
If lvDetalle.ListItems.Count > 0 Then
    ImprimirConsulta
Else
    MsgBox "no"
End If
End Sub

Private Sub cmdResumenOT_Click()
With frmResumenOT
    .lblIdOT = lvDetalle.SelectedItem
    .lblSeccion = lvDetalle.SelectedItem.SubItems(8)
    .lblEstado = lvDetalle.SelectedItem.SubItems(1)
    .lblPatente = lvDetalle.SelectedItem.SubItems(2)
    .lblCliente = lvDetalle.SelectedItem.SubItems(3)
    .lblMarca = lvDetalle.SelectedItem.SubItems(4)
    .lblModelo = lvDetalle.SelectedItem.SubItems(5)
    .lblTotalMec = lvDetalle.SelectedItem.SubItems(11)
    .lblTotalCar = lvDetalle.SelectedItem.SubItems(12)
    .lblTotalOtr = lvDetalle.SelectedItem.SubItems(13)
    .lblTotalTer = lvDetalle.SelectedItem.SubItems(14)
    .lblTotalRep = lvDetalle.SelectedItem.SubItems(15)
    .lblTotalMat = lvDetalle.SelectedItem.SubItems(16)
    .lblTotalIns = lvDetalle.SelectedItem.SubItems(17)
    .lblsubtotal = lvDetalle.SelectedItem.SubItems(18)
    .lblIva = lvDetalle.SelectedItem.SubItems(19)
    .lblTotalOT = lvDetalle.SelectedItem.SubItems(20)
    .Show vbModal
End With
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSeleccionar_Click()
If Not lvDetalle.SelectedItem Is Nothing Then
    gstrBusca = lvDetalle.SelectedItem
    gstrSeccion = lvDetalle.SelectedItem.SubItems(10)
End If
Unload Me
End Sub




Private Sub Form_Activate()

If SW Then
    pckFechaDesde = BOM(Date)
    pckFechaHasta = EOM(Date)
    cmdImprimir.Enabled = Atributos("Glbl", "Tllr_30_0010", True, True, True, True)
    SW = False
End If

End Sub

Private Sub Form_Load()
Dim adopaso As ADODB.Recordset
Dim Item As ListItem
SW = True

    If Not Conexion.SendHost("Select Descripcion, Id_Garantia From Tllr_Garantias", adopaso, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        MsgBox "Error en Conexion con el Host...", vbCritical, "Stock Pro"
        End
    End If

    If Not (adopaso.EOF = True And adopaso.BOF = True) Then
        Do Until adopaso.EOF
                Set Item = Me.lsvtipoOt.ListItems.Add(, , ValorNulo(adopaso.Fields(0)))
                Item.SubItems(1) = ValorNulo(adopaso.Fields(1))
            adopaso.MoveNext
        Loop
    End If
    adopaso.Close
End Sub

Private Sub lvDetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ReOrdenaLista lvDetalle, ColumnHeader
End Sub

Private Sub lvDetalle_DblClick()
If cmdSeleccionar.Enabled = True Then cmdSeleccionar.Value = True
End Sub

Private Sub tlbCliente_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    gstrBusca = apfFormulario.BuscarRegistroClientes(Conexion, "Id_Cliente_Proveedor", "Razon_Social")
    txtCliente.Tag = gstrBusca
    txtCliente = ClienteD(gstrBusca)
End If
End Sub

Private Sub tlbMarca_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Glbl_Marca", "Id_Marca", "Descripcion", "Busca Marca")
    txtMarca.Tag = gstrBusca
    txtMarca = MarcaD(gstrBusca)
End If
End Sub

Private Sub tlbModelo_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    gstrBusca = apfFormulario.BuscarRegistrosModelo(Conexion, "Glbl_Modelo", "Id_Modelo", "Id_Marca", "Descripcion", "Busca Modelo", IIf(Me.txtMarca.Tag <> "", txtMarca.Tag, "01"))
    txtModelo.Tag = gstrBusca
    txtModelo = ModeloD(txtMarca.Tag, gstrBusca)
End If
End Sub

Private Sub tlbRecep_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "Tllr_Mecanicos", "Id_Mecanico", "Nombre", "Busca Mecanico")
    txtRecepcionista.Tag = gstrBusca
    txtRecepcionista = MecanicoD(gstrBusca)
End If
End Sub


Private Sub txtCliente_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtMarca_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtModelo_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtNroRecord_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtNroRecord, strComa)
End Sub

Private Sub txtPatente_KeyPress(KeyAscii As Integer)
'KeyAscii = CheckIdCar(txtPatente.SelStart, mdLLNNNN, UpCaseLetter(KeyAscii))
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtRecepcionista_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub
Function TotalSeccion(lvwObjeto As ListView, IndiceSubItem As Integer) As Double
Dim intS As Integer
Dim dblPreSuma As Double
dblPreSuma = 0
With lvwObjeto
    For intS = 1 To .ListItems.Count
        Set .SelectedItem = .ListItems(intS)
        dblPreSuma = dblPreSuma + CDbl(SacarFormatoValor(IIf(.SelectedItem.SubItems(IndiceSubItem) <> "", .SelectedItem.SubItems(IndiceSubItem), 0), ""))
    Next
End With
TotalSeccion = dblPreSuma
End Function


Function TotalSeccionFormato(lvwObjeto As ListView, IndiceSubItem As Integer) As Double
Dim intS As Integer
Dim dblPreSuma As Double
dblPreSuma = 0
With lvwObjeto
    For intS = 1 To .ListItems.Count
        Set .SelectedItem = .ListItems(intS)
        dblPreSuma = dblPreSuma + Val(SacarFormatoValor(.SelectedItem.SubItems(IndiceSubItem), "$"))
    Next
End With
TotalSeccionFormato = dblPreSuma
End Function

