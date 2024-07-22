VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmConsultaMecanicoRepuestos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Mecanico Repuestos"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "frmConsultaMecanicoRepuesto.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport rptPatente 
      Left            =   3960
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir Informe"
      Height          =   360
      Left            =   7995
      TabIndex        =   21
      Top             =   7200
      Width           =   1680
   End
   Begin VB.Frame Frame2 
      Height          =   2145
      Left            =   60
      TabIndex        =   5
      Top             =   -15
      Width           =   11370
      Begin VB.Frame Frame1 
         Caption         =   "Estado"
         Height          =   525
         Left            =   5640
         TabIndex        =   22
         Top             =   975
         Width           =   4680
         Begin VB.OptionButton optLiquidada 
            Caption         =   "Liquidada"
            Height          =   195
            Left            =   1746
            TabIndex        =   27
            Top             =   240
            Width           =   990
         End
         Begin VB.OptionButton optNula 
            Caption         =   "Nula"
            Height          =   195
            Left            =   3840
            TabIndex        =   26
            Top             =   270
            Width           =   675
         End
         Begin VB.OptionButton optCerrada 
            Caption         =   "Facturadas"
            Height          =   195
            Left            =   2739
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   1110
         End
         Begin VB.OptionButton optTodas 
            Caption         =   "Todas"
            Height          =   195
            Left            =   75
            TabIndex        =   24
            Top             =   240
            Width           =   810
         End
         Begin VB.OptionButton optVigente 
            Caption         =   "Vigente"
            Height          =   195
            Left            =   888
            TabIndex        =   23
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "F. Emisión (Fin)"
         Height          =   195
         Index           =   7
         Left            =   1800
         TabIndex        =   20
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.TextBox txtPatente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   6
         TabIndex        =   14
         Top             =   600
         Width           =   1020
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Patente"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Marca "
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   12
         Top             =   315
         Width           =   870
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Modelo"
         Height          =   195
         Index           =   3
         Left            =   5640
         TabIndex        =   11
         Top             =   360
         Width           =   840
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Cliente"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   795
      End
      Begin VB.TextBox txtCliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1185
         Width           =   5175
      End
      Begin VB.TextBox txtMarca 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   8
         Top             =   600
         Width           =   3435
      End
      Begin VB.TextBox txtModelo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5640
         MaxLength       =   50
         TabIndex        =   7
         Top             =   555
         Width           =   4635
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "F. Emisión (Ini)"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   1545
         Value           =   1  'Checked
         Width           =   1320
      End
      Begin MSComctlLib.ImageList ImgBarraHerramienta 
         Left            =   10680
         Top             =   120
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
               Picture         =   "frmConsultaMecanicoRepuesto.frx":000C
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":011E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":0576
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":09CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":0E26
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":0F38
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":104A
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":115C
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":126E
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":1380
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":1492
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":15A4
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":16B6
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":17C8
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":18DA
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":19EC
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":1AFE
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":1C10
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":1D22
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":1E34
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":2286
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmConsultaMecanicoRepuesto.frx":26D8
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbMarca 
         Height          =   330
         Left            =   4920
         TabIndex        =   15
         Top             =   240
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
         Enabled         =   0   'False
      End
      Begin MSComctlLib.Toolbar tlbModelo 
         Height          =   330
         Left            =   9840
         TabIndex        =   16
         Top             =   240
         Width           =   345
         _ExtentX        =   609
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
      Begin MSComctlLib.Toolbar tlbCliente 
         Height          =   330
         Left            =   4920
         TabIndex        =   17
         Top             =   930
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
         TabIndex        =   18
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   83427329
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckFechaHasta 
         Height          =   315
         Left            =   1800
         TabIndex        =   19
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   83427329
         CurrentDate     =   36776
      End
      Begin MSDataListLib.DataCombo dtcSupervisor 
         Bindings        =   "frmConsultaMecanicoRepuesto.frx":27EA
         Height          =   315
         Left            =   3720
         TabIndex        =   28
         Top             =   1755
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc datSupervisor 
         Height          =   330
         Left            =   5895
         Top             =   1800
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label2 
         Caption         =   "Mecanico"
         Height          =   255
         Left            =   3720
         TabIndex        =   29
         Top             =   1560
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdBuscarOT 
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   360
      Left            =   6240
      TabIndex        =   0
      Top             =   7200
      Width           =   1680
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   9750
      TabIndex        =   1
      Top             =   7200
      Width           =   1680
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   2010
      Left            =   75
      TabIndex        =   4
      Top             =   2175
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   3545
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N° OT"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Estado"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha Emisión"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Patente"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Marca"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Modelo"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Horas Realizadas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Seccion"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvDetalleRepuestosOT 
      Height          =   2010
      Left            =   75
      TabIndex        =   30
      Top             =   4680
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   3545
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   6967
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Familia"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Cantidad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Costo  Unitario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Subtotal"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbTotales 
      Height          =   315
      Left            =   8400
      TabIndex        =   32
      Top             =   4200
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Suma - Horas"
            TextSave        =   "Suma - Horas"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2469
            MinWidth        =   2469
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbTotalCosto 
      Height          =   315
      Left            =   8400
      TabIndex        =   33
      Top             =   6720
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Suma - Costos"
            TextSave        =   "Suma - Costos"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2469
            MinWidth        =   2469
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Repuestos Usados"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   7
      Left            =   1920
      TabIndex        =   3
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registros Encontrados :"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   2
      Top             =   7320
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmConsultaMecanicoRepuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SW As Boolean

Sub ImprimirConsulta()
Dim Dbsnueva As Database
Dim Tabla As DAO.Recordset
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

    Screen.MousePointer = 11
    Dim wrkPredeterminado As Workspace
    Dim prpBucle As Property
    Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
    If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (NroOT text,Estado text,FechaIngreso Text,Recepcionista text,Seccion text,Tipo text,Kilometros text,Trabajo text,Valor Double,Patente Text,Cliente Text,Marca Text,Modelo Text)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To lvDetalle.ListItems.Count
        Set lvDetalle.SelectedItem = lvDetalle.ListItems(i)
        Tabla.AddNew
        Tabla!NroOT = IIf(lvDetalle.SelectedItem = "", " ", lvDetalle.SelectedItem)
        Tabla!estado = IIf(lvDetalle.SelectedItem.SubItems(1) = "", " ", lvDetalle.SelectedItem.SubItems(1))
        Tabla!FechaIngreso = IIf(lvDetalle.SelectedItem.SubItems(2) = "", "", lvDetalle.SelectedItem.SubItems(2))
        Tabla!Recepcionista = IIf(lvDetalle.SelectedItem.SubItems(3) = "", " ", lvDetalle.SelectedItem.SubItems(3))
        Tabla!Seccion = IIf(lvDetalle.SelectedItem.SubItems(4) = "", " ", lvDetalle.SelectedItem.SubItems(4))
        Tabla!Tipo = IIf(lvDetalle.SelectedItem.SubItems(5) = "", " ", lvDetalle.SelectedItem.SubItems(5))
        Tabla!Kilometros = IIf(lvDetalle.SelectedItem.SubItems(6) = "", " ", lvDetalle.SelectedItem.SubItems(6))
        Tabla!Trabajo = IIf(lvDetalle.SelectedItem.SubItems(7) = "", " ", lvDetalle.SelectedItem.SubItems(7))
        Tabla!Valor = IIf(lvDetalle.SelectedItem.SubItems(8) = "", 0, SacarFormatoValor(lvDetalle.SelectedItem.SubItems(8), gstrMonedaLocal))
        Tabla!Patente = IIf(txtPatente = "", " ", txtPatente)   ' IIf(lvDetalle.SelectedItem.SubItems(8) = "", " ", lvDetalle.SelectedItem.SubItems(8))
        Tabla!Cliente = txtCliente
        Tabla!Marca = txtMarca
        Tabla!Modelo = txtModelo
        
        Tabla.Update
    Next i
   Tabla.Close
   
   With rptPatente
        .ReportFileName = gstrPathReporte & "\HistoricoPatente.Rpt"
        .WindowTitle = "Historico Por Placa"
        .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='Historico Por Placa'"
        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
        .Formulas(5) = "desde='" & pckFechaDesde & "'"
        .Formulas(6) = "hasta='" & pckFechaHasta & "'"
        .Destination = crptToWindow
        .Action = True
   End With
   
   Dbsnueva.Close
   Screen.MousePointer = 1

End Sub


Private Sub cckCriterios_Click(Index As Integer)
Select Case Index
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
    If cckCriterios(Index).Value = 0 Then
        tlbCliente.Enabled = False
        txtCliente.Enabled = False
        txtCliente = ""
    Else
        tlbCliente.Enabled = True
        txtCliente.Enabled = True
        txtCliente.SetFocus
    End If
Case 5
    If cckCriterios(Index).Value = 0 Then
        'tlbRecep.Enabled = False
        'txtRecepcionista.Enabled = False
        'txtRecepcionista = ""
    Else
        'tlbRecep.Enabled = True
        'txtRecepcionista.Enabled = True
        'txtRecepcionista.SetFocus
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
End Select
End Sub
Private Sub cmdBuscarOT_Click()
Dim mstrSql As String
Dim lstrSql As String
Dim mstrWhere As String
Dim mstrWhere2 As String
Dim AdoTemp As New ADODB.Recordset
Dim AdoAux As New ADODB.Recordset
Dim itmItem As ListItem
Dim mstrEstado As String
Dim ContLinea As Integer
Dim mdblSumaHoras As Double
Dim mstrNumeroDocumento As String

lvDetalle.ListItems.Clear
lvDetalleRepuestosOT.ListItems.Clear
mstrWhere = ""
mstrWhere2 = ""
With Me
    
    If .cckCriterios(1).Value = 1 Then  '////////// patente
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Tllr_Ot.PATENTE LIKE '" & MatchMode(.txtPatente, "Comienzo del Campo", apSqlServer) & "'"
            mstrWhere2 = mstrWhere2 & " and Tllr_Ot.PATENTE LIKE '" & MatchMode(.txtPatente, "Comienzo del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Tllr_Ot.PATENTE LIKE '" & MatchMode(.txtPatente, "Comienzo del Campo", apSqlServer) & "'"
            mstrWhere2 = " Where Tllr_Ot.PATENTE LIKE '" & MatchMode(.txtPatente, "Comienzo del Campo", apSqlServer) & "'"
        End If
    End If
    
    If .cckCriterios(2).Value = 1 Then  '////////// marca
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Glbl_Marca.Descripcion LIKE '" & MatchMode(.txtMarca, "Comienzo del Campo", apSqlServer) & "'"
            mstrWhere2 = mstrWhere2 & " and Glbl_Marca.Descripcion LIKE '" & MatchMode(.txtMarca, "Comienzo del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Glbl_Marca.Descripcion LIKE '" & MatchMode(.txtMarca, "Comienzo del Campo", apSqlServer) & "'"
            mstrWhere2 = " Where Glbl_Marca.Descripcion LIKE '" & MatchMode(.txtMarca, "Comienzo del Campo", apSqlServer) & "'"
        End If
    End If
    
    If .cckCriterios(3).Value = 1 Then  '////////// modelo
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Glbl_Modelo.Descripcion LIKE '" & MatchMode(.txtModelo, "Comienzo del Campo", apSqlServer) & "'"
            mstrWhere2 = mstrWhere2 & " and Glbl_Modelo.Descripcion LIKE '" & MatchMode(.txtModelo, "Comienzo del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Glbl_Modelo.Descripcion LIKE '" & MatchMode(.txtModelo, "Comienzo del Campo", apSqlServer) & "'"
            mstrWhere2 = " Where Glbl_Modelo.Descripcion LIKE '" & MatchMode(.txtModelo, "Comienzo del Campo", apSqlServer) & "'"
        End If
    End If
    
    If .cckCriterios(4).Value = 1 Then  '////////// cliente
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Glbl_Cliente_Proveedor.Razon_Social LIKE '" & MatchMode(.txtCliente, "Comienzo del Campo", apSqlServer) & "'"
            mstrWhere2 = mstrWhere2 & " and Glbl_Cliente_Proveedor.Razon_Social LIKE '" & MatchMode(.txtCliente, "Comienzo del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Glbl_Cliente_Proveedor.Razon_Social LIKE '" & MatchMode(.txtCliente, "Comienzo del Campo", apSqlServer) & "'"
            mstrWhere2 = " Where Glbl_Cliente_Proveedor.Razon_Social LIKE '" & MatchMode(.txtCliente, "Comienzo del Campo", apSqlServer) & "'"
        End If
    End If
    
    If .dtcSupervisor.Text <> "" Then
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Tllr_Otro_Ot.Mecanico_Asignado='" & .dtcSupervisor.BoundText & "'"
            mstrWhere2 = mstrWhere2 & " and Tllr_Mecanica_Ot.Mecanico_Designado='" & .dtcSupervisor.BoundText & "'"
        Else
            mstrWhere = " Where Tllr_Otro_Ot.Mecanico_Asignado='" & .dtcSupervisor.BoundText & "'"
            mstrWhere2 = " Where Tllr_Mecanica_Ot.Mecanico_Designado='" & .dtcSupervisor.BoundText & "'"
        End If
    Else
        MsgBox "Debe Ingresar un valor en Mecanico", vbExclamation, "Advertencia"
        dtcSupervisor.SetFocus
        Exit Sub
    End If
    
    If .cckCriterios(6).Value = 1 Then  '////////// fecha inicio
        If .cckCriterios(7).Value = 1 Then  '////////// fecha termino
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " AND fecha_emision between '" & pckFechaDesde.Value & "' and '" & pckFechaHasta.Value & " 23:59:59" & "'"
                mstrWhere2 = mstrWhere2 & " AND fecha_emision between '" & pckFechaDesde.Value & "' and '" & pckFechaHasta.Value & " 23:59:59" & "'"
            Else
                mstrWhere = " WHERE fecha_emision between '" & pckFechaDesde.Value & "' and '" & pckFechaHasta.Value & " 23:59:59" & "'"
                mstrWhere2 = " WHERE fecha_emision between '" & pckFechaDesde.Value & "' and '" & pckFechaHasta.Value & " 23:59:59" & "'"
            End If
        Else
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " AND fecha_emision = '" & pckFechaDesde.Value & "' "
                mstrWhere2 = mstrWhere2 & " AND fecha_emision = '" & pckFechaDesde.Value & "' "
            Else
                mstrWhere = " WHERE fecha_emision = '" & pckFechaDesde.Value & "' "
                mstrWhere2 = " WHERE fecha_emision = '" & pckFechaDesde.Value & "' "
            End If
        End If
    Else
        If .cckCriterios(7).Value = 1 Then  '////////// fecha termino
            If mstrWhere <> "" Then
                mstrWhere = " AND fecha_emision = '" & pckFechaHasta.Value & "' "
                mstrWhere2 = " AND fecha_emision = '" & pckFechaHasta.Value & "' "
            Else
                mstrWhere = " WHERE fecha_emision = '" & pckFechaHasta.Value & "' "
                mstrWhere2 = " WHERE fecha_emision = '" & pckFechaHasta.Value & "' "
            End If
        End If
    End If
    
     '////////// empresa y sucursal
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " AND Tllr_Ot.ID_EMPRESA= '" & gstrIdEmpresa & "' AND Tllr_Ot.ID_SUCURSAL='" & gstrIdSucursal & "' "
            mstrWhere2 = mstrWhere2 & " AND Tllr_Ot.ID_EMPRESA= '" & gstrIdEmpresa & "' AND Tllr_Ot.ID_SUCURSAL='" & gstrIdSucursal & "' "
        Else
            mstrWhere = " WHERE Tllr_Ot.ID_EMPRESA= '" & gstrIdEmpresa & "' AND Tllr_Ot.ID_SUCURSAL='" & gstrIdSucursal & "' "
            mstrWhere2 = " WHERE Tllr_Ot.ID_EMPRESA= '" & gstrIdEmpresa & "' AND Tllr_Ot.ID_SUCURSAL='" & gstrIdSucursal & "' "
        End If
    '//////////////////estado
            If optTodas.Value = True Then
                mstrEstado = "IN ('L','F','B','C','V','A','N')"
            ElseIf optVigente.Value = True Then
                mstrEstado = "IN ('V','A')"
            ElseIf optLiquidada.Value = True Then
                mstrEstado = "IN ('L','C')"
            ElseIf optCerrada.Value = True Then
                mstrEstado = "IN ('F','B')"
            ElseIf optNula.Value = True Then
                mstrEstado = "IN ('N')"
            End If
        If mstrEstado <> "" Then
            mstrWhere = mstrWhere & " And Tllr_OT.Estado  " & mstrEstado
            mstrWhere2 = mstrWhere2 & " And Tllr_OT.Estado  " & mstrEstado
        End If
End With
'/////////////////////////////////////////////////////////////////////////////////
    
    
    'horas de otros servicios
'    mstrSql = "SELECT SUM(Tllr_Otro_OT.Horas) AS Horas, Tllr_Otro_OT.Id_OT, "
'    mstrSql = mstrSql & "Tllr_OT.Estado, Tllr_OT.Fecha_Emision, Tllr_OT.Patente, "
'    mstrSql = mstrSql & "Tllr_OT.Seccion_OT, "
'    mstrSql = mstrSql & "Glbl_Marca.Descripcion AS Marca, "
'    mstrSql = mstrSql & "Glbl_Modelo.Descripcion AS Modelo "
'    mstrSql = mstrSql & "FROM Glbl_Cliente_Proveedor INNER JOIN "
'    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente ON "
'    mstrSql = mstrSql & "Glbl_Cliente_Proveedor.Id_Cliente_Proveedor = Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor "
'    mstrSql = mstrSql & "Inner Join Tllr_Otro_OT INNER JOIN Tllr_OT ON "
'    mstrSql = mstrSql & "Tllr_Otro_OT.Id_Empresa = Tllr_OT.Id_Empresa AND "
'    mstrSql = mstrSql & "Tllr_Otro_OT.Id_Sucursal = Tllr_OT.Id_Sucursal AND "
'    mstrSql = mstrSql & "Tllr_Otro_OT.Id_OT = Tllr_OT.Id_OT AND "
'    mstrSql = mstrSql & "Tllr_Otro_OT.Seccion_OT = Tllr_OT.Seccion_OT ON "
'    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente.Patente = Tllr_OT.Patente INNER JOIN "
'    mstrSql = mstrSql & "Glbl_Modelo INNER JOIN Glbl_Marca ON "
'    mstrSql = mstrSql & "Glbl_Modelo.Id_Marca = Glbl_Marca.Id_Marca ON "
'    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente.Id_Modelo = Glbl_Modelo.Id_Modelo AND "
'    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente.Id_Marca = Glbl_Modelo.Id_Marca "
'    'where
'    mstrSql = mstrSql & mstrWhere & " "
'    'group by
'    mstrSql = mstrSql & "GROUP BY Tllr_Otro_OT.Id_OT, Tllr_OT.Estado, "
'    mstrSql = mstrSql & "Tllr_OT.Fecha_Emision, Tllr_OT.Patente, Tllr_OT.Seccion_OT, "
'    mstrSql = mstrSql & "Glbl_Marca.Descripcion , Glbl_Modelo.Descripcion "
'
'    'mstrSql = mstrSql & "  ORDER BY Tllr_OT.Id_OT"
    
    
    mstrSql = "SELECT SUM(Tllr_Otro_OT.Horas) AS SUMAOTRO, "
    mstrSql = mstrSql & "(SELECT SUM(Tllr_Mecanica_OT.Horas) AS Horas "
    mstrSql = mstrSql & "FROM Glbl_Cliente_Proveedor INNER JOIN "
    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente ON "
    mstrSql = mstrSql & "Glbl_Cliente_Proveedor.Id_Cliente_Proveedor = Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor "
    mstrSql = mstrSql & "Inner Join Tllr_Mecanica_OT INNER JOIN Tllr_OT ON "
    mstrSql = mstrSql & "Tllr_Mecanica_OT.Id_Empresa = Tllr_OT.Id_Empresa AND "
    mstrSql = mstrSql & "Tllr_Mecanica_OT.Id_Sucursal = Tllr_OT.Id_Sucursal AND "
    mstrSql = mstrSql & "Tllr_Mecanica_OT.Id_OT = Tllr_OT.Id_OT AND "
    mstrSql = mstrSql & "Tllr_Mecanica_OT.Seccion_OT = Tllr_OT.Seccion_OT ON "
    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente.Patente = Tllr_OT.Patente INNER JOIN "
    mstrSql = mstrSql & "Glbl_Modelo INNER JOIN "
    mstrSql = mstrSql & "Glbl_Marca ON Glbl_Modelo.Id_Marca = Glbl_Marca.Id_Marca ON "
    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente.Id_Modelo = Glbl_Modelo.Id_Modelo "
    mstrSql = mstrSql & "AND Tllr_Vehiculo_Cliente.Id_Marca = Glbl_Modelo.Id_Marca "
    mstrSql = mstrSql & mstrWhere2 & " "
    mstrSql = mstrSql & "GROUP BY Tllr_Mecanica_OT.Id_OT, Tllr_OT.Estado, "
    mstrSql = mstrSql & "Tllr_OT.Fecha_Emision, Tllr_OT.Patente, "
    mstrSql = mstrSql & "Tllr_OT.Seccion_OT, Glbl_Marca.Descripcion, "
    mstrSql = mstrSql & "Glbl_Modelo.Descripcion) AS SUMAMECANICA, "
    mstrSql = mstrSql & "Tllr_Otro_OT.Id_OT, Tllr_OT.Estado, Tllr_OT.Fecha_Emision, "
    mstrSql = mstrSql & "Tllr_OT.Patente, Tllr_OT.Seccion_OT, "
    mstrSql = mstrSql & "Glbl_Marca.Descripcion AS Marca, "
    mstrSql = mstrSql & "Glbl_Modelo.Descripcion AS Modelo "
    mstrSql = mstrSql & "FROM Glbl_Cliente_Proveedor INNER JOIN "
    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente ON "
    mstrSql = mstrSql & "Glbl_Cliente_Proveedor.Id_Cliente_Proveedor = Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor "
    mstrSql = mstrSql & "Inner Join Tllr_Otro_OT INNER JOIN Tllr_OT ON "
    mstrSql = mstrSql & "Tllr_Otro_OT.Id_Empresa = Tllr_OT.Id_Empresa AND "
    mstrSql = mstrSql & "Tllr_Otro_OT.Id_Sucursal = Tllr_OT.Id_Sucursal AND "
    mstrSql = mstrSql & "Tllr_Otro_OT.Id_OT = Tllr_OT.Id_OT AND "
    mstrSql = mstrSql & "Tllr_Otro_OT.Seccion_OT = Tllr_OT.Seccion_OT ON "
    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente.Patente = Tllr_OT.Patente INNER JOIN "
    mstrSql = mstrSql & "Glbl_Modelo INNER JOIN "
    mstrSql = mstrSql & "Glbl_Marca ON "
    mstrSql = mstrSql & "Glbl_Modelo.Id_Marca = Glbl_Marca.Id_Marca ON "
    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente.Id_Modelo = Glbl_Modelo.Id_Modelo AND "
    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente.Id_Marca = Glbl_Modelo.Id_Marca "
    mstrSql = mstrSql & mstrWhere & " "
    mstrSql = mstrSql & "GROUP BY Tllr_Otro_OT.Id_OT, Tllr_OT.Estado, "
    mstrSql = mstrSql & "Tllr_OT.Fecha_Emision, Tllr_OT.Patente, Tllr_OT.Seccion_OT, "
    mstrSql = mstrSql & "Glbl_Marca.Descripcion , Glbl_Modelo.Descripcion "
    
    Screen.MousePointer = 11
    mdblSumaHoras = 0
    If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
    With AdoTemp
       If Not .BOF And Not .EOF Then
          While Not .EOF
              Set itmItem = lvDetalle.ListItems.Add(, , !Id_OT)
              If !estado = "F" Or !estado = "B" Then
                 mstrNumeroDocumento = TraeNumeroDocumento(!Seccion_OT, !Id_OT, "")
              Else
                mstrNumeroDocumento = ""
              End If
              itmItem.SubItems(1) = ValorNulo(IIf(!estado = "L", "LIQUIDADA", IIf(!estado = "V", "VIGENTE", IIf(!estado = "N", "NULA", IIf(!estado = "C", "CERRADA", IIf(!estado = "F", "FACTURADA", IIf(!estado = "B", "BOLETEADA", "OTRO"))))))) & "(" & mstrNumeroDocumento & ")"
              itmItem.SubItems(2) = Format(ValorNulo(!Fecha_Emision), "dd/mm/yyyy")
              itmItem.SubItems(3) = ValorNulo(!Patente)
              itmItem.SubItems(4) = ValorNulo(!Marca)
              itmItem.SubItems(5) = ValorNulo(!Modelo)
              itmItem.SubItems(6) = FormatoValor(IIf(IsNull(!SumaOtro), 0, !SumaOtro) + IIf(IsNull(!SumaMecanica), 0, !SumaMecanica), "", 2)
              itmItem.SubItems(7) = ValorNulo(!Seccion_OT)
              mdblSumaHoras = mdblSumaHoras + (IIf(IsNull(!SumaOtro), 0, !SumaOtro) + IIf(IsNull(!SumaMecanica), 0, !SumaMecanica))
              AdoTemp.MoveNext
          Wend
       End If
    End With
    End If
    stbTotales.Panels(2) = FormatoValor(mdblSumaHoras, "", 2)
    Screen.MousePointer = 1
    lblTotal(7).Caption = lvDetalle.ListItems.Count
    mstrEstado = ""
End Sub
Private Sub cmdImprimir_Click()
If lvDetalle.ListItems.Count > 0 Then
    'ImprimirConsulta
Else
    MsgBox "no"
End If
End Sub

Private Sub cmdResumenOT_Click()
If Not lvDetalle.SelectedItem Is Nothing Then
With frmResumenOT
    .lblIdOT = lvDetalle.SelectedItem
    .lblSeccion = lvDetalle.SelectedItem.SubItems(9)
    .lblestado = lvDetalle.SelectedItem.SubItems(1)
    .lblPatente = lvDetalle.SelectedItem.SubItems(2)
    .lblCliente = lvDetalle.SelectedItem.SubItems(3)
    .lblMarca = lvDetalle.SelectedItem.SubItems(4)
    .lblModelo = lvDetalle.SelectedItem.SubItems(5)
    .lblTotalMec = FormatoValor(lvDetalle.SelectedItem.SubItems(12), "", gintDecimalesMoneda)
    .lblTotalCar = FormatoValor(lvDetalle.SelectedItem.SubItems(13), "", gintDecimalesMoneda)
    .lblTotalOtr = FormatoValor(lvDetalle.SelectedItem.SubItems(14), "", gintDecimalesMoneda)
    .lblTotalTer = FormatoValor(lvDetalle.SelectedItem.SubItems(15), "", gintDecimalesMoneda)
    .lblTotalRep = FormatoValor(lvDetalle.SelectedItem.SubItems(16), "", gintDecimalesMoneda)
    .lblTotalMat = FormatoValor(lvDetalle.SelectedItem.SubItems(17), "", gintDecimalesMoneda)
    .lblTotalIns = FormatoValor(lvDetalle.SelectedItem.SubItems(18), "", gintDecimalesMoneda)
    .lblsubtotal = FormatoValor(lvDetalle.SelectedItem.SubItems(19), "", gintDecimalesMoneda)
    .lblIva = FormatoValor(lvDetalle.SelectedItem.SubItems(20), "", gintDecimalesMoneda)
    .lblTotalOT = FormatoValor(lvDetalle.SelectedItem.SubItems(21), "", gintDecimalesMoneda)
    .ReCalculo
    .Show vbModal
End With
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSeleccionar_Click()
If Not lvDetalle.SelectedItem Is Nothing Then
    gstrBusca = lvDetalle.SelectedItem
    gstrSeccion = lvDetalle.SelectedItem.SubItems(11)
End If
Unload Me
End Sub




Private Sub Form_Activate()

If SW Then
    pckFechaDesde = BOM(Date)
    pckFechaHasta = EOM(Date)
    FillMecanicos dtcSupervisor, datSupervisor
    'cmdImprimir.Enabled = Atributos("Glbl", "Tllr_30_0010", True, True, True, True)
    SW = False
End If

End Sub

Private Sub Form_Load()
SW = True
End Sub

Private Sub lvDetalle_Click()
TraeRepuestosAsociados
End Sub

Private Sub lvDetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'ReOrdenaLista lvDetalle, ColumnHeader
End Sub

Private Sub tlbCliente_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
    gstrBusca = apfFormulario.BuscarRegistroClientes(Conexion, "Id_Cliente_Proveedor", "Razon_Social", gstrIdEmpresa)
    'gstrBusca = apfFormulario.BuscarRegistroClientes(Conexion, "Id_Cliente_Proveedor", "Razon_Social")
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

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtMarca_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtModelo_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtPatente_KeyPress(KeyAscii As Integer)
'KeyAscii = CheckIdCar(txtPatente.SelStart, mdLLNNNN, UpCaseLetter(KeyAscii))
'KeyAscii = UpCaseLetter(KeyAscii)
'kjcv 24-01-12 Valida Letras y numeros
If (KeyAscii <> 8) And Not (KeyAscii >= 48 And KeyAscii <= 57) And Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
    KeyAscii = 0: Beep
Else
    KeyAscii = UpCaseLetter(KeyAscii)
End If

End Sub


Sub TraeRepuestosAsociados()
Dim mstrSql As String
Dim AdoTemp As New ADODB.Recordset
Dim AdoAux As New ADODB.Recordset
Dim itmItemAux As ListItem
Dim pdblCostoPromedio As Double
Dim pdblSumaCostos As Double

Me.lvDetalleRepuestosOT.ListItems.Clear

'With Me

    mstrSql = "SELECT Stck_Item.Id_Familia + '°' + Stck_Item.Prefijo + Stck_Item.Basico + Stck_Item.Sufijo AS Item," _
     & "Glbl_Familia.Descripcion as Familia, Tllr_Repuestos_OT.Cantidad,Tllr_Repuestos_OT.Valor, " _
     & "Tllr_Repuestos_OT.Porcentaje_Descuento,Tllr_Repuestos_OT.Monto_Descuento, " _
     & "Tllr_Repuestos_OT.SubTotal , Stck_Item.Precio_Costo, Stck_Item.Descripcion as Repuesto FROM Tllr_Repuestos_OT INNER JOIN Tllr_OT ON " _
     & "Tllr_Repuestos_OT.Id_Empresa = Tllr_OT.Id_Empresa AND " _
     & "Tllr_Repuestos_OT.Id_Sucursal = Tllr_OT.Id_Sucursal AND " _
     & "Tllr_Repuestos_OT.Id_OT = Tllr_OT.Id_OT AND " _
     & "Tllr_Repuestos_OT.Seccion_OT = Tllr_OT.Seccion_OT INNER JOIN " _
     & "Stck_Item ON " _
     & "Tllr_Repuestos_OT.Id_Item = Stck_Item.Id_Item INNER JOIN " _
     & "Glbl_Familia ON Stck_Item.Id_Familia = Glbl_Familia.Id_Familia " _
     & "Where Tllr_Repuestos_Ot.Id_Ot='" & Me.lvDetalle.SelectedItem & "' and " _
     & "Tllr_Repuestos_Ot.Seccion_Ot='" & Me.lvDetalle.SelectedItem.SubItems(7) & "'"
    
    pdblSumaCostos = 0
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
    With AdoTemp
       If Not .BOF And Not .EOF Then
          While Not .EOF
              Set itmItemAux = lvDetalleRepuestosOT.ListItems.Add(, , !Item)
              itmItemAux.SubItems(1) = ValorNulo(!Repuesto)
              itmItemAux.SubItems(2) = ValorNulo(!Familia)
              itmItemAux.SubItems(3) = FormatoValor(!cantidad, "", 2)
              pdblCostoPromedio = Round(Costo_Promedio_Repuesto(lvDetalle.SelectedItem, !Item), gintDecimalesMoneda)
              If !cantidad <> 0 Then
                itmItemAux.SubItems(4) = FormatoValor(pdblCostoPromedio / !cantidad, "", gintDecimalesMoneda) 'Precio costo
              Else
                itmItemAux.SubItems(4) = "0"
              End If
              itmItemAux.SubItems(5) = FormatoValor(pdblCostoPromedio, "", gintDecimalesMoneda)
              pdblSumaCostos = pdblSumaCostos + pdblCostoPromedio
              AdoTemp.MoveNext
          Wend
          stbTotalCosto.Panels(2) = FormatoValor(pdblSumaCostos, "", gintDecimalesMoneda)
       End If
    End With
    End If
    Screen.MousePointer = 1
    
End Sub
