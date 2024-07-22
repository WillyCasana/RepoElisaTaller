VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmBuscaOTSrvt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de O/T Serviteca"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "frmBuscaOTSrvt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11475
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rptOT 
      Left            =   3945
      Top             =   6210
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir Informe"
      Height          =   360
      Left            =   6195
      TabIndex        =   32
      Top             =   6240
      Width           =   1680
   End
   Begin VB.Frame Frame2 
      Height          =   2145
      Left            =   60
      TabIndex        =   6
      Top             =   -15
      Width           =   11370
      Begin VB.CheckBox cckCriterios 
         Caption         =   "F. Liquidación (Fin)"
         Enabled         =   0   'False
         Height          =   195
         Index           =   9
         Left            =   4725
         TabIndex        =   40
         Top             =   1560
         Width           =   1680
      End
      Begin VB.Frame Frame1 
         Caption         =   "Estado"
         Height          =   525
         Left            =   6960
         TabIndex        =   34
         Top             =   1560
         Width           =   4320
         Begin VB.OptionButton optLiquidada 
            Caption         =   "Liquidada"
            Height          =   195
            Left            =   1746
            TabIndex        =   39
            Top             =   240
            Width           =   990
         End
         Begin VB.OptionButton optNula 
            Caption         =   "Nula"
            Height          =   195
            Left            =   3600
            TabIndex        =   38
            Top             =   240
            Width           =   675
         End
         Begin VB.OptionButton optCerrada 
            Caption         =   "Cerrada"
            Height          =   195
            Left            =   2739
            TabIndex        =   37
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optTodas 
            Caption         =   "Todas"
            Height          =   195
            Left            =   75
            TabIndex        =   36
            Top             =   240
            Value           =   -1  'True
            Width           =   810
         End
         Begin VB.OptionButton optVigente 
            Caption         =   "Vigente"
            Height          =   195
            Left            =   888
            TabIndex        =   35
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdResumenOT 
         Caption         =   "Ver Resumen"
         Height          =   360
         Left            =   10005
         TabIndex        =   33
         Top             =   1155
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "F. Emisión (Fin)"
         Height          =   195
         Index           =   7
         Left            =   1530
         TabIndex        =   30
         Top             =   1545
         Width           =   1365
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "F. Liquidación (Ini)"
         Enabled         =   0   'False
         Height          =   195
         Index           =   8
         Left            =   3045
         TabIndex        =   29
         Top             =   1560
         Width           =   1680
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Recepcionista"
         Height          =   195
         Index           =   5
         Left            =   4620
         TabIndex        =   25
         Top             =   945
         Width           =   1395
      End
      Begin VB.TextBox txtRecepcionista 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4620
         MaxLength       =   50
         TabIndex        =   24
         Top             =   1185
         Width           =   3675
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Nro OT"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   23
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox txtNroOt 
         Height          =   300
         Left            =   105
         MaxLength       =   15
         TabIndex        =   22
         Top             =   525
         Width           =   2670
      End
      Begin VB.TextBox txtPatente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2850
         MaxLength       =   6
         TabIndex        =   16
         Top             =   525
         Width           =   1020
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Patente"
         Height          =   195
         Index           =   1
         Left            =   2865
         TabIndex        =   15
         Top             =   330
         Width           =   855
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Marca "
         Height          =   195
         Index           =   2
         Left            =   3930
         TabIndex        =   14
         Top             =   315
         Width           =   870
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Modelo"
         Height          =   195
         Index           =   3
         Left            =   6855
         TabIndex        =   13
         Top             =   345
         Width           =   840
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Cliente"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   945
         Width           =   795
      End
      Begin VB.TextBox txtNroRecord 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10380
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "0"
         Top             =   525
         Width           =   555
      End
      Begin VB.TextBox txtCliente 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1185
         Width           =   4455
      End
      Begin VB.TextBox txtMarca 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3930
         MaxLength       =   50
         TabIndex        =   9
         Top             =   540
         Width           =   2835
      End
      Begin VB.TextBox txtModelo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6855
         MaxLength       =   50
         TabIndex        =   8
         Top             =   555
         Width           =   2835
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "F. Emisión (Ini)"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   1545
         Width           =   1320
      End
      Begin MSComctlLib.ImageList ImgBarraHerramienta 
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
            NumListImages   =   22
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":0442
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":0554
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":09AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":0E04
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":125C
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":136E
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":1480
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":1592
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":16A4
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":17B6
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":18C8
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":19DA
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":1AEC
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":1BFE
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":1D10
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":1E22
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":1F34
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":2046
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":2158
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":226A
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":26BC
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBuscaOTSrvt.frx":2B0E
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.UpDown updNroRecord 
         Height          =   315
         Left            =   10920
         TabIndex        =   17
         Top             =   525
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         BuddyControl    =   "txtNroRecord"
         BuddyDispid     =   196622
         OrigLeft        =   10950
         OrigTop         =   525
         OrigRight       =   11190
         OrigBottom      =   840
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComctlLib.Toolbar tlbMarca 
         Height          =   330
         Left            =   6270
         TabIndex        =   18
         Top             =   210
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
         Left            =   9210
         TabIndex        =   19
         Top             =   225
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
      Begin MSComctlLib.Toolbar tlbCliente 
         Height          =   330
         Left            =   4140
         TabIndex        =   20
         Top             =   855
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
      Begin MSComctlLib.Toolbar tlbRecep 
         Height          =   330
         Left            =   7860
         TabIndex        =   26
         Top             =   885
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
         TabIndex        =   27
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   82968577
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckFechaHasta 
         Height          =   315
         Left            =   1530
         TabIndex        =   28
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   82968577
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckLiquidaIni 
         Height          =   315
         Left            =   3045
         TabIndex        =   31
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   82968577
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckLiquidaFin 
         Height          =   315
         Left            =   4725
         TabIndex        =   41
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   82968577
         CurrentDate     =   36776
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Registros"
         Height          =   195
         Index           =   8
         Left            =   10410
         TabIndex        =   21
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdBuscarOT 
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   360
      Left            =   4440
      TabIndex        =   0
      Top             =   6255
      Width           =   1680
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "Seleccionar"
      Height          =   360
      Left            =   7980
      TabIndex        =   1
      Top             =   6255
      Width           =   1680
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   9750
      TabIndex        =   2
      Top             =   6255
      Width           =   1680
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   3930
      Left            =   45
      TabIndex        =   5
      Top             =   2175
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   6932
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
      NumItems        =   22
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N° OT"
         Object.Width           =   1411
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
         Text            =   "Fono"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Marca"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Modelo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Fecha Apertura"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Recepcionista"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Dias Llamados"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Tipo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Id_Seccion"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "TMEC"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "TCAR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "TOTR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "TTER"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "TREP"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "TMAT"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "TINS"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "TNETO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "TIVA"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "TOTAL"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   7
      Left            =   1935
      TabIndex        =   4
      Top             =   6390
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registros Encontrados :"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   3
      Top             =   6390
      Width           =   1695
   End
End
Attribute VB_Name = "frmBuscaOTSrvt"
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
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (NroOT text,Estado text,Patente text,Cliente text,Marca text,Modelo text,FechaIngreso date,Recepcionista text,Dias text)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To lvDetalle.ListItems.Count
        Set lvDetalle.SelectedItem = lvDetalle.ListItems(i)
        Tabla.AddNew
        Tabla!NroOT = IIf(lvDetalle.SelectedItem = "", " ", lvDetalle.SelectedItem)
        Tabla!estado = IIf(lvDetalle.SelectedItem.SubItems(1) = "", " ", lvDetalle.SelectedItem.SubItems(1))
        Tabla!Patente = IIf(lvDetalle.SelectedItem.SubItems(2) = "", " ", lvDetalle.SelectedItem.SubItems(2))
        Tabla!Cliente = IIf(lvDetalle.SelectedItem.SubItems(3) = "", " ", lvDetalle.SelectedItem.SubItems(3))
        Tabla!Marca = IIf(lvDetalle.SelectedItem.SubItems(5) = "", " ", lvDetalle.SelectedItem.SubItems(5))
        Tabla!Modelo = IIf(lvDetalle.SelectedItem.SubItems(6) = "", " ", lvDetalle.SelectedItem.SubItems(6))
        Tabla!FechaIngreso = DateValue(IIf(lvDetalle.SelectedItem.SubItems(7) = "", " ", lvDetalle.SelectedItem.SubItems(7)))
        'tabla!FechaLiq = IIf(lvDetalle.SelectedItem.SubItems(7) = "", " ", lvDetalle.SelectedItem.SubItems(7))
        Tabla!Recepcionista = IIf(lvDetalle.SelectedItem.SubItems(8) = "", " ", lvDetalle.SelectedItem.SubItems(8))
        Tabla!dias = IIf(lvDetalle.SelectedItem.SubItems(9) = "", " ", lvDetalle.SelectedItem.SubItems(9))
        'tabla!tipo = IIf(lvDetalle.SelectedItem.SubItems(10) = "", " ", lvDetalle.SelectedItem.SubItems(10))
        Tabla.Update
    Next i
   Tabla.Close
   
   With rptOT
        .ReportFileName = gstrPathReporte & "\ListadoOtSrvt.rpt"
        .WindowTitle = "Reporte de Ordenes de Trabajo"
        .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='LISTADO DE ORDENES DE TRABAJO DE SERVITECA'"
        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
        .Destination = crptToWindow
        .Action = True
   End With
   
   Dbsnueva.Close
   Screen.MousePointer = 1

End Sub


Private Sub cckCriterios_Click(Index As Integer)
Select Case Index
Case 0
    If cckCriterios(Index).Value = 0 Then
        txtNroOt.Enabled = False
        txtNroOt = ""
    Else
        txtNroOt.Enabled = True
        txtNroOt.SetFocus
    End If
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
    If cckCriterios(Index).Value = 0 Then
        pckLiquidaIni.Enabled = False
    Else
        pckLiquidaIni.Enabled = True
        pckLiquidaIni.SetFocus
    End If
Case 9
    If cckCriterios(Index).Value = 0 Then
        pckLiquidaFin.Enabled = False
    Else
        pckLiquidaFin.Enabled = True
        pckLiquidaFin.SetFocus
    End If
End Select
End Sub

Private Sub cmdBuscarOT_Click()
Dim mstrSql As String
Dim mstrWhere As String
Dim AdoTemp As New ADODB.Recordset
Dim AdoAux As New ADODB.Recordset
Dim itmItem As ListItem
Dim mstrEstado As String

lvDetalle.ListItems.Clear
mstrWhere = ""
With Me
    If .cckCriterios(0).Value = 1 Then  '////////// nro ot
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Srvt_Ot.Id_Ot LIKE '" & MatchMode(txtNroOt, "Cualquier Parte del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Srvt_Ot.Id_Ot LIKE '" & MatchMode(txtNroOt, "Cualquier Parte del Campo", apSqlServer) & "'"
        End If
    End If
    
    If .cckCriterios(1).Value = 1 Then  '////////// patente
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Srvt_Ot.PATENTE LIKE '" & MatchMode(.txtPatente, "Comienzo del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Srvt_Ot.PATENTE LIKE '" & MatchMode(.txtPatente, "Comienzo del Campo", apSqlServer) & "'"
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
    
    If .cckCriterios(4).Value = 1 Then  '////////// cliente
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Glbl_Cliente_Proveedor.Razon_Social LIKE '" & MatchMode(.txtCliente, "Comienzo del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Glbl_Cliente_Proveedor.Razon_Social LIKE '" & MatchMode(.txtCliente, "Comienzo del Campo", apSqlServer) & "'"
        End If
    End If
    
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
                mstrWhere = mstrWhere & " AND fecha_apertura between '" & pckFechaDesde.Value & "' and '" & pckFechaHasta.Value & " 23:59:00'"
            Else
                mstrWhere = " WHERE fecha_apertura between '" & pckFechaDesde.Value & "' and '" & pckFechaHasta.Value & " 23:59:00'"
            End If
        Else
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " AND fecha_apertura = '" & pckFechaDesde.Value & "' "
            Else
                mstrWhere = " WHERE fecha_apertura = '" & pckFechaDesde.Value & "' "
            End If
        End If
    Else
        If .cckCriterios(7).Value = 1 Then  '////////// fecha termino
            If mstrWhere <> "" Then
                mstrWhere = " AND fecha_apertura = '" & pckFechaHasta.Value & "' "
            Else
                mstrWhere = " WHERE fecha_apertura = '" & pckFechaHasta.Value & "' "
            End If
        End If
    End If

        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " AND Srvt_Ot.ID_EMPRESA= '" & gstrIdEmpresa & "' AND Srvt_Ot.ID_SUCURSAL='" & gstrIdSucursal & "' "
        Else
            mstrWhere = " WHERE Srvt_Ot.ID_EMPRESA= '" & gstrIdEmpresa & "' AND Srvt_Ot.ID_SUCURSAL='" & gstrIdSucursal & "' "
        End If
    '//////////////////estado
        If optTodas.Value = True Then
            mstrEstado = "('N','V','L','F','B')"
        ElseIf optVigente.Value = True Then
            mstrEstado = "('V')"
        ElseIf optLiquidada.Value = True Then
            mstrEstado = "('L')"
        ElseIf optCerrada.Value = True Then
            mstrEstado = "('F','B')"
        ElseIf optNula.Value = True Then
            mstrEstado = "('N')"
        End If
        
        If mstrEstado <> "" Then
            mstrWhere = mstrWhere & " And Srvt_OT.Estado In " & mstrEstado
        End If
End With
'/////////////////////////////////////////////////////////////////////////////////
'    mstrSql = "SELECT " & IIf(Val(txtNroRecord) > 0, "TOP " & CInt(Val(txtNroRecord)) & "", "") & " Srvt_OT.Id_OT, "
'    mstrSql = mstrSql & " Srvt_Ot.PATENTE,"
'    mstrSql = mstrSql & " Srvt_ot.Estado,"
'    mstrSql = mstrSql & " Srvt_Ot.Fecha_Apertura,"
'    mstrSql = mstrSql & " Srvt_ot.Id_Cliente_Proveedor,"
'    mstrSql = mstrSql & " Srvt_ot.Id_mecanico,"
'    mstrSql = mstrSql & " Srvt_ot.Dias_llamado,"
'    mstrSql = mstrSql & " Glbl_Cliente_proveedor.Razon_Social,"
'    mstrSql = mstrSql & " Glbl_Cliente_proveedor.Telefono,"
'    mstrSql = mstrSql & " Tllr_mecanicos.nombre as Recep, "
'    mstrSql = mstrSql & " TLLR_VEHICULO_CLIENTE.ID_MARCA,"
'    mstrSql = mstrSql & " TLLR_VEHICULO_CLIENTE.ID_MODELO,"
'    mstrSql = mstrSql & " Glbl_Marca.Descripcion as Marca,"
'    mstrSql = mstrSql & " Glbl_Modelo.Descripcion as Modelo"
'    mstrSql = mstrSql & " FROM Srvt_ot"
'    mstrSql = mstrSql & " INNER JOIN GLBL_CLIENTE_PROVEEDOR ON GLBL_CLIENTE_PROVEEDOR.ID_CLIENTE_PROVEEDOR = SRVT_OT.ID_CLIENTE_PROVEEDOR"
'    mstrSql = mstrSql & " INNER JOIN TLLR_MECANICOS ON TLLR_MECANICOS.ID_MECANICO = SRVT_OT.ID_MECANICO"
'    mstrSql = mstrSql & " INNER JOIN TLLR_VEHICULO_CLIENTE ON TLLR_VEHICULO_CLIENTE.PATENTE = SRVT_OT.PATENTE"
'    mstrSql = mstrSql & " INNER JOIN GLBL_MARCA ON GLBL_MARCA.ID_MARCA = TLLR_VEHICULO_CLIENTE.ID_MARCA"
'    mstrSql = mstrSql & " INNER JOIN GLBL_MODELO ON GLBL_MODELO.ID_MODELO = TLLR_VEHICULO_CLIENTE.ID_MODELO"
    
    mstrSql = "SELECT " & IIf(Val(txtNroRecord) > 0, "TOP " & CInt(Val(txtNroRecord)) & "", "") & " Srvt_OT.Id_OT, "
    mstrSql = mstrSql & "Srvt_OT.Patente, Srvt_OT.Estado, Srvt_OT.Fecha_Apertura, Srvt_OT.Id_Cliente_Proveedor, "
    mstrSql = mstrSql & "Srvt_OT.Id_Mecanico, Srvt_OT.Dias_LLamado, Glbl_Cliente_Proveedor.Razon_Social, Glbl_Cliente_Proveedor.Telefono, "
    mstrSql = mstrSql & "Tllr_Mecanicos.Nombre AS Recep, Tllr_Vehiculo_Cliente.Id_Marca, Tllr_Vehiculo_Cliente.Id_Modelo, "
    mstrSql = mstrSql & "Glbl_Marca.Descripcion AS Marca, Glbl_Modelo.Descripcion AS Modelo "
    mstrSql = mstrSql & "FROM Srvt_OT INNER JOIN "
    mstrSql = mstrSql & "Glbl_Cliente_Proveedor ON Glbl_Cliente_Proveedor.Id_Cliente_Proveedor = Srvt_OT.Id_Cliente_Proveedor INNER JOIN "
    mstrSql = mstrSql & "Tllr_Mecanicos ON Tllr_Mecanicos.Id_Mecanico = Srvt_OT.Id_Mecanico INNER JOIN "
    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente ON Tllr_Vehiculo_Cliente.Patente = Srvt_OT.Patente INNER JOIN "
    mstrSql = mstrSql & "Glbl_Modelo ON Tllr_Vehiculo_Cliente.Id_Marca = Glbl_Modelo.Id_Marca AND "
    mstrSql = mstrSql & "Tllr_Vehiculo_Cliente.Id_Modelo = Glbl_Modelo.Id_Modelo INNER JOIN "
    mstrSql = mstrSql & "Glbl_Marca ON Glbl_Modelo.Id_Marca = Glbl_Marca.Id_Marca "
    mstrSql = mstrSql & mstrWhere
    mstrSql = mstrSql & "  ORDER BY Srvt_Ot.ID_OT"
    
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
    With AdoTemp
       If Not .BOF And Not .EOF Then
          While Not .EOF
              Set itmItem = lvDetalle.ListItems.Add(, , !Id_OT)
              itmItem.SubItems(1) = ValorNulo(IIf(!estado = "L", "LIQUIDADA", IIf(!estado = "V", "VIGENTE", IIf(!estado = "F", "FACTURADA", IIf(!estado = "B", "BOLETEADA", IIf(!estado = "N", "NULA", IIf(!estado = "C", "CERRADA", "OTRO")))))))
              itmItem.SubItems(2) = ValorNulo(!Patente)
              itmItem.SubItems(3) = ValorNulo(!Razon_Social)
              itmItem.SubItems(4) = ValorNulo(!Telefono)
              itmItem.SubItems(5) = ValorNulo(!Marca)
              itmItem.SubItems(6) = ValorNulo(!Modelo)
              itmItem.SubItems(7) = Format(ValorNulo(!Fecha_Apertura), "dd/mm/yyyy")
              itmItem.SubItems(8) = ValorNulo(!RECEP)
              itmItem.SubItems(9) = ValorNulo(!Dias_LLamado)
              AdoTemp.MoveNext
          Wend
       End If
    End With
    End If
    Screen.MousePointer = 1
    lblTotal(7).Caption = lvDetalle.ListItems.Count
    mstrEstado = ""
End Sub
Private Sub cmdImprimir_Click()
If lvDetalle.ListItems.Count > 0 Then
    ImprimirConsulta
Else
    MsgBox "no"
End If
End Sub

Private Sub cmdResumenOT_Click()
If Not lvDetalle.SelectedItem Is Nothing Then
With frmResumenOT
    .lblIdOT = lvDetalle.SelectedItem
    .lblSeccion = lvDetalle.SelectedItem.SubItems(8)
    .lblestado = lvDetalle.SelectedItem.SubItems(1)
    .lblPatente = lvDetalle.SelectedItem.SubItems(2)
    .lblCliente = lvDetalle.SelectedItem.SubItems(3)
    .lblMarca = lvDetalle.SelectedItem.SubItems(4)
    .lblModelo = lvDetalle.SelectedItem.SubItems(5)
    .lblTotalMec = FormatoValor(lvDetalle.SelectedItem.SubItems(11), "", gintDecimalesMoneda)
    .lblTotalCar = FormatoValor(lvDetalle.SelectedItem.SubItems(12), "", gintDecimalesMoneda)
    .lblTotalOtr = FormatoValor(lvDetalle.SelectedItem.SubItems(13), "", gintDecimalesMoneda)
    .lblTotalTer = FormatoValor(lvDetalle.SelectedItem.SubItems(14), "", gintDecimalesMoneda)
    .lblTotalRep = FormatoValor(lvDetalle.SelectedItem.SubItems(15), "", gintDecimalesMoneda)
    .lblTotalMat = FormatoValor(lvDetalle.SelectedItem.SubItems(16), "", gintDecimalesMoneda)
    .lblTotalIns = FormatoValor(lvDetalle.SelectedItem.SubItems(17), "", gintDecimalesMoneda)
    .lblsubtotal = FormatoValor(lvDetalle.SelectedItem.SubItems(18), "", gintDecimalesMoneda)
    .lblIva = FormatoValor(lvDetalle.SelectedItem.SubItems(19), "", gintDecimalesMoneda)
    .lblTotalOT = FormatoValor(lvDetalle.SelectedItem.SubItems(20), "", gintDecimalesMoneda)
    .ReCalculo
    .Show vbModal
End With
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSeleccionar_Click()
If Trim$(gstrBusca) <> "" Then
    Load frmOtServiteca
    frmOtServiteca.TraeDesdeFuera
    Unload Me
    frmOtServiteca.Show
End If
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
SW = True
Me.cckCriterios(1).Caption = gstrNombrePatente
End Sub

Private Sub lvDetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ReOrdenaLista lvDetalle, ColumnHeader
End Sub

Private Sub lvDetalle_DblClick()
If cmdSeleccionar.Enabled = True Then cmdSeleccionar.Value = True
End Sub

Private Sub lvDetalle_ItemClick(ByVal Item As MSComctlLib.ListItem)
gstrBusca = CStr(CDbl(Me.lvDetalle.SelectedItem))
Me.cmdSeleccionar.Enabled = True
End Sub

Private Sub optCerrada_Click()
If optCerrada.Value = True Then
    pckLiquidaIni.Enabled = False
    pckLiquidaFin.Enabled = False
    cckCriterios(8).Enabled = False
    cckCriterios(9).Enabled = False
End If
End Sub

Private Sub optLiquidada_Click()

If optLiquidada.Value = True Then
    pckLiquidaIni.Enabled = True
    pckLiquidaFin.Enabled = True
    cckCriterios(8).Enabled = True
    cckCriterios(9).Enabled = True
End If

End Sub

Private Sub optNula_Click()
If optNula.Value = True Then
    pckLiquidaIni.Enabled = False
    pckLiquidaFin.Enabled = False
    cckCriterios(8).Enabled = False
    cckCriterios(9).Enabled = False
End If
End Sub

Private Sub optVigente_Click()
If optVigente.Value = True Then
    pckLiquidaIni.Enabled = False
    pckLiquidaFin.Enabled = False
    cckCriterios(8).Enabled = False
    cckCriterios(9).Enabled = False
End If
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
'KeyAscii = UpCaseLetter(KeyAscii)
'kjcv 24-01-12 Valida Letras y numeros
If (KeyAscii <> 8) And Not (KeyAscii >= 48 And KeyAscii <= 57) And Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
    KeyAscii = 0: Beep
Else
    KeyAscii = UpCaseLetter(KeyAscii)
End If

End Sub

Private Sub txtRecepcionista_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub
