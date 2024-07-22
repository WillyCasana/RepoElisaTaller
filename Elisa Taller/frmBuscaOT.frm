VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBuscaOT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de O/T"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "frmBuscaOT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   11475
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdExportar 
      Left            =   4680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport rptOT 
      Left            =   3945
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      Caption         =   "Imprimir Informe"
      Height          =   360
      Left            =   6120
      TabIndex        =   32
      Top             =   6480
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Frame Frame2 
      Height          =   2145
      Left            =   60
      TabIndex        =   6
      Top             =   345
      Width           =   11370
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "VIN"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   2880
         TabIndex        =   45
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox txtVin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2880
         TabIndex        =   44
         Top             =   525
         Width           =   1335
      End
      Begin VB.CheckBox cckPresupuesto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Presupuestos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9720
         TabIndex        =   43
         Top             =   1170
         Width           =   1575
      End
      Begin VB.CheckBox cckReservas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Reservas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8400
         TabIndex        =   42
         Top             =   1170
         Width           =   1125
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "F. Liquidación (Fin)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   4725
         TabIndex        =   40
         Top             =   1545
         Width           =   1680
      End
      Begin VB.Frame Frame1 
         Caption         =   "Estado"
         Height          =   525
         Left            =   6615
         TabIndex        =   34
         Top             =   1575
         Width           =   4680
         Begin VB.OptionButton optLiquidada 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Liquidada"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1746
            TabIndex        =   39
            Top             =   240
            Width           =   990
         End
         Begin VB.OptionButton optNula 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Nula"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3960
            TabIndex        =   38
            Top             =   270
            Width           =   675
         End
         Begin VB.OptionButton optCerrada 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Facturadas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2739
            TabIndex        =   37
            Top             =   240
            Width           =   1230
         End
         Begin VB.OptionButton optTodas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Todas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   36
            Top             =   240
            Value           =   -1  'True
            Width           =   810
         End
         Begin VB.OptionButton optVigente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Vigente"
            ForeColor       =   &H80000008&
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
         Left            =   9960
         TabIndex        =   33
         Top             =   480
         Width           =   1275
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "F. Emisión (Fin)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   1530
         TabIndex        =   30
         Top             =   1545
         Width           =   1365
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "F. Liquidación (Ini)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   3045
         TabIndex        =   29
         Top             =   1545
         Width           =   1680
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Recepcionista"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   4500
         TabIndex        =   25
         Top             =   945
         Width           =   1395
      End
      Begin VB.TextBox txtRecepcionista 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4500
         MaxLength       =   50
         TabIndex        =   24
         Top             =   1185
         Width           =   3675
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Nro OT"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   23
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox txtNroOt 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   105
         MaxLength       =   15
         TabIndex        =   22
         Top             =   525
         Width           =   1590
      End
      Begin VB.TextBox txtPatente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   16
         Top             =   525
         Width           =   1020
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Placa"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   15
         Top             =   300
         Width           =   855
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Marca "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   4440
         TabIndex        =   14
         Top             =   300
         Width           =   870
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Modelo"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   7320
         TabIndex        =   13
         Top             =   300
         Width           =   840
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Cliente"
         ForeColor       =   &H80000008&
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
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1185
         Width           =   4335
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   9
         Top             =   525
         Width           =   2715
      End
      Begin VB.TextBox txtModelo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7320
         MaxLength       =   50
         TabIndex        =   8
         Top             =   525
         Width           =   2475
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "F. Emisión (Ini)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   1545
         Width           =   1320
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
         BuddyDispid     =   196625
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
         Left            =   6720
         TabIndex        =   18
         Top             =   240
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
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
         Left            =   9360
         TabIndex        =   19
         Top             =   225
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
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
         Left            =   4020
         TabIndex        =   20
         Top             =   855
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
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
         Left            =   7740
         TabIndex        =   26
         Top             =   885
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
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
         Format          =   101974017
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
         Format          =   101974017
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
         Format          =   101974017
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
         Format          =   101974017
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
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdBuscarOT 
      Appearance      =   0  'Flat
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   360
      Left            =   4440
      TabIndex        =   0
      Top             =   6480
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.CommandButton cmdSeleccionar 
      Appearance      =   0  'Flat
      Caption         =   "Seleccionar"
      Height          =   360
      Left            =   9720
      TabIndex        =   1
      Top             =   6480
      Width           =   1680
   End
   Begin VB.CommandButton cmdSalir 
      Appearance      =   0  'Flat
      Caption         =   "Salir"
      Height          =   360
      Left            =   7800
      TabIndex        =   2
      Top             =   6480
      Visible         =   0   'False
      Width           =   1680
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   3930
      Left            =   75
      TabIndex        =   5
      Top             =   2520
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   6932
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   27
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
         Text            =   "VIN"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Cliente"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cod. Cliente"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Fono"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Modelo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Fecha Emisión"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Fecha Liquidación"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Recepcionista"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Seccion"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Tipo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Trabajo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Object.Tag             =   "N"
         Text            =   "Id_Seccion"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Object.Tag             =   "N"
         Text            =   "TMEC"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Object.Tag             =   "N"
         Text            =   "TCAR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Object.Tag             =   "N"
         Text            =   "TOTR"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Object.Tag             =   "N"
         Text            =   "TTER"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Object.Tag             =   "N"
         Text            =   "TREP"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Object.Tag             =   "N"
         Text            =   "TMAT"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Object.Tag             =   "N"
         Text            =   "TINS"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Object.Tag             =   "N"
         Text            =   "TNETO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Object.Tag             =   "N"
         Text            =   "TIVA"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Object.Tag             =   "N"
         Text            =   "TOTAL"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Comentario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "CIA Seguros"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Height          =   330
      Left            =   120
      TabIndex        =   46
      Top             =   0
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar "
            ImageKey        =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Seleccion1"
            Object.ToolTipText     =   "Configurar Lista"
            ImageKey        =   "Seleccion1"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir "
            ImageKey        =   "Imprimir"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Cristal"
                  Text            =   "Informe Crystal Report"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Medida"
                  Text            =   "Informe a Medida"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Configurar"
                  Text            =   "Configurar Reporte"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exportar"
            Object.ToolTipText     =   "Exportar a Excel"
            ImageKey        =   "Excel"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "CotizaPerdida"
            Object.ToolTipText     =   "Cotización Perdida"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Vender"
            Object.ToolTipText     =   "Pasar a Venta"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Agenda"
            Object.ToolTipText     =   "Agenda Diaria"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ultimo Registro (Ctrl+U)"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Renovar"
            Object.ToolTipText     =   "Renovar Registros (Ctrl+R)"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (Ctrl + Q)"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   5640
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   46
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":179A
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":18AC
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":19BE
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":1AD0
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":1BE2
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":1CF4
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":1E06
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":1F18
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":202A
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":213C
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":224E
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":2360
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":2472
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":2584
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":2696
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":27A8
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":28BA
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":2D0C
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":315E
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":3270
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":33CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":3528
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":3684
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":37E0
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":42AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":4700
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":4864
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":4CC0
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":4E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":6128
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":66C4
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":6820
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":697C
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":6CD0
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":7024
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":7378
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":76CC
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":7A20
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":7F64
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":8076
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":83CA
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":871E
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":8832
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":8B86
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":8C98
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscaOT.frx":8FEC
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   7
      Left            =   1935
      TabIndex        =   4
      Top             =   6600
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registros Encontrados :"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   3
      Top             =   6600
      Width           =   1695
   End
End
Attribute VB_Name = "frmBuscaOT"
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
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (NroOT text,Estado text,Patente text,VIN text,Cliente text,IDCLI text,Marca text,Modelo text,FechaIngreso date,Recepcionista text,Seccion text,Tipo text)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To lvDetalle.ListItems.Count
        Set lvDetalle.SelectedItem = lvDetalle.ListItems(i)
        Tabla.AddNew
        Tabla!NroOT = IIf(lvDetalle.SelectedItem = "", " ", Val(Mid(lvDetalle.SelectedItem, 6, 10)))
        Tabla!estado = IIf(lvDetalle.SelectedItem.SubItems(1) = "", " ", lvDetalle.SelectedItem.SubItems(1))
        Tabla!Patente = IIf(lvDetalle.SelectedItem.SubItems(2) = "", " ", lvDetalle.SelectedItem.SubItems(2))
        Tabla!VIN = IIf(lvDetalle.SelectedItem.SubItems(3) = "", " ", lvDetalle.SelectedItem.SubItems(3))
        Tabla!Cliente = IIf(lvDetalle.SelectedItem.SubItems(4) = "", " ", lvDetalle.SelectedItem.SubItems(4))
        Tabla!idCLI = IIf(lvDetalle.SelectedItem.SubItems(5) = "", " ", lvDetalle.SelectedItem.SubItems(5))
        Tabla!Marca = IIf(lvDetalle.SelectedItem.SubItems(6) = "", " ", lvDetalle.SelectedItem.SubItems(6))
        Tabla!Modelo = IIf(lvDetalle.SelectedItem.SubItems(7) = "", " ", lvDetalle.SelectedItem.SubItems(7))
        Tabla!FechaIngreso = DateValue(IIf(lvDetalle.SelectedItem.SubItems(8) = "", " ", lvDetalle.SelectedItem.SubItems(8)))
        Tabla!Recepcionista = IIf(lvDetalle.SelectedItem.SubItems(10) = "", " ", Iniciales(lvDetalle.SelectedItem.SubItems(10)))
        Tabla!Seccion = IIf(lvDetalle.SelectedItem.SubItems(11) = "", " ", lvDetalle.SelectedItem.SubItems(11))
        Tabla!Tipo = IIf(lvDetalle.SelectedItem.SubItems(12) = "", " ", lvDetalle.SelectedItem.SubItems(12))
        Tabla.Update

    Next i
   Tabla.Close
   Dbsnueva.Close
   With rptOT
        .ReportFileName = gstrPathReporte & "\OTS.rpt"
        .WindowTitle = "Reporte de Ordenes de Trabajo"
        .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='LISTADO DE ORDENES DE TRABAJO'"
        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
        .Formulas(5) = "NombreRut='" & gstrNombreRut & "'"
        .Formulas(6) = "NombrePatente='" & gstrNombrePatente & "'"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = True
   End With

'   Dbsnueva.Close
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
        Me.tlbModelo.Enabled = False
        txtModelo.Enabled = False
        txtModelo = ""
    Else
        Me.tlbModelo.Enabled = True
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
Case 10
    If cckCriterios(Index).Value = 0 Then
        txtVin.Enabled = False
        txtVin = ""
    Else
        txtVin.Enabled = True
        txtVin.SetFocus
    End If
End Select
End Sub




Private Sub cmdBuscarOT_Click()
Dim mstrSQL As String
Dim mstrWhere As String
Dim adoTemp As New ADODB.Recordset
Dim AdoAux As New ADODB.Recordset
Dim itmItem As ListItem
Dim mstrEstado As String
Dim mstrNumeroDocumento As String

lvDetalle.ListItems.Clear
mstrWhere = "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "'"
With Me
    If .cckCriterios(0).Value = 1 Then  '////////// nro ot
        mstrWhere = mstrWhere & ",'" & MatchMode(txtNroOt, "Cualquier Parte del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(1).Value = 1 Then  '////////// patente
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtPatente, "Comienzo del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(10).Value = 1 Then  '////////// vin
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtVin, "Cualquier Parte del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(2).Value = 1 Then  '////////// marca
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtMarca, "Comienzo del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(3).Value = 1 Then  '////////// modelo
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtModelo, "Comienzo del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(4).Value = 1 Then  '////////// cliente
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtCliente, "Comienzo del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(5).Value = 1 Then  '////////// recepcionista
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtRecepcionista, "Comienzo del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(6).Value = 1 And .cckCriterios(7).Value = 1 Then   '////////// fecha iniciosi y terminosi
        mstrWhere = mstrWhere & ",'" & pckFechaDesde.Value & "','" & pckFechaHasta.Value & " 23:59:00" & "'"
    ElseIf .cckCriterios(6).Value = 0 And .cckCriterios(7).Value = 0 Then  '////////// fecha iniciono y terminono
        mstrWhere = mstrWhere & ",'',''"
    ElseIf .cckCriterios(6).Value = 1 And .cckCriterios(7).Value = 0 Then   '////////// fecha iniciosi y terminono
        mstrWhere = mstrWhere & ",'" & pckFechaDesde.Value & "',''"
    ElseIf .cckCriterios(6).Value = 0 And .cckCriterios(7).Value = 1 Then  '////////// fecha iniciono y terminosi
        mstrWhere = ",'','" & pckFechaHasta.Value & "'"
    End If
    
    '//////////////////////////////////////////////////////
    If .cckCriterios(8).Value = 1 And .cckCriterios(9).Value = 1 Then   '////////// fecha iniciosi y terminosi
        mstrWhere = mstrWhere & ",'" & pckLiquidaIni.Value & "','" & pckLiquidaFin.Value & " 23:59:00" & "'"
    ElseIf .cckCriterios(8).Value = 0 And .cckCriterios(9).Value = 0 Then  '////////// fecha iniciono y terminono
        mstrWhere = mstrWhere & ",'',''"
    ElseIf .cckCriterios(8).Value = 1 And .cckCriterios(9).Value = 0 Then   '////////// fecha iniciosi y terminono
        mstrWhere = mstrWhere & ",'" & pckLiquidaIni.Value & "',''"
    ElseIf .cckCriterios(8).Value = 0 And .cckCriterios(9).Value = 1 Then  '////////// fecha iniciono y terminosi
        mstrWhere = mstrWhere & ",'','" & pckLiquidaFin.Value & "'"
    End If

    '//////////////////estado
    
   If Me.cckReservas.Value = 1 And Me.cckPresupuesto.Value = 1 Then
       mstrWhere = mstrWhere & ",'RP'"
   ElseIf Me.cckReservas.Value = 1 And Me.cckPresupuesto.Value = 0 Then
       mstrWhere = mstrWhere & ",'R'"
   ElseIf Me.cckReservas.Value = 0 And Me.cckPresupuesto.Value = 1 Then
       mstrWhere = mstrWhere & ",'P'"
   ElseIf optTodas.Value = True Then
       mstrWhere = mstrWhere & ",'T'"
   ElseIf optVigente.Value = True Then
       mstrWhere = mstrWhere & ",'V'"
   ElseIf optLiquidada.Value = True Then
       mstrWhere = mstrWhere & ",'L'"
   ElseIf optCerrada.Value = True Then
       mstrWhere = mstrWhere & ",'F'"
   ElseIf optNula.Value = True Then
       mstrWhere = mstrWhere & ",'N'"
   
   End If
        
End With
'/////////////////////////////////////////////////////////////////////////////////
    
    mstrSQL = "Exec Tllr_listado_OT " & mstrWhere
    
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With adoTemp
       If Not .BOF And Not .EOF Then
          While Not .EOF
              Set itmItem = lvDetalle.ListItems.Add(, , !Id_OT)
              If !est = "B" Or !est = "F" Then
                'mstrNumeroDocumento = IIf(!NUMDOCUMENTO <> "S/N", !NUMDOCUMENTO, TraeNumeroDocumento(!Sec, !Id_OT, ""))
                mstrNumeroDocumento = !NUMDOCUMENTO
              End If
              itmItem.SubItems(1) = ValorNulo(IIf(!est = "L", "LIQUIDADA", IIf(!est = "V", "VIGENTE", IIf(!est = "N", "NULA", IIf(!est = "C", "CERRADA", IIf(!est = "F", "FACTURADA (" & mstrNumeroDocumento & ")", IIf(!est = "B", "BOLETEADA (" & mstrNumeroDocumento & ")", "OTRO")))))))
              itmItem.SubItems(2) = ValorNulo(!Pat)
              itmItem.SubItems(3) = ValorNulo(!VIN)
              itmItem.SubItems(4) = ValorNulo(!Cliente)
              itmItem.SubItems(5) = ValorNulo(!idCLI)
              itmItem.SubItems(6) = ValorNulo(!FONO)
              itmItem.SubItems(7) = ValorNulo(!Modelo)
              itmItem.SubItems(8) = Format(ValorNulo(!FEC), "dd/mm/yyyy")
              itmItem.SubItems(9) = Format(ValorNulo(!FECLIQ), "dd/mm/yyyy")
              itmItem.SubItems(10) = ValorNulo(!RECEP)
              itmItem.SubItems(11) = ValorNulo(IIf(!Sec = "M", "MECANICA", "CARROCERIA"))
              itmItem.SubItems(12) = ValorNulo(!GAR)
              itmItem.SubItems(13) = ValorNulo(!Sec)
              
              itmItem.SubItems(14) = ValorNulo(!TMEC)
              itmItem.SubItems(15) = ValorNulo(!TCAR)
              itmItem.SubItems(16) = ValorNulo(!TOTR)
              itmItem.SubItems(17) = ValorNulo(!TTER)
              itmItem.SubItems(18) = ValorNulo(!TREP)
              itmItem.SubItems(19) = ValorNulo(!TMAT)
              itmItem.SubItems(20) = ValorNulo(!TINS)
              itmItem.SubItems(21) = ValorNulo(!Tneto)
              itmItem.SubItems(22) = ValorNulo(!Tiva)
              itmItem.SubItems(23) = ValorNulo(!Total)
              adoTemp.MoveNext
          Wend
       End If
    End With
    End If
    Screen.MousePointer = 1
    lblTotal(7).Caption = lvDetalle.ListItems.Count
    mstrEstado = ""
End Sub
Private Sub ImprimirInforme()
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
    .lblSeccion = lvDetalle.SelectedItem.SubItems(11)
    .lblestado = lvDetalle.SelectedItem.SubItems(1)
    .lblPatente = lvDetalle.SelectedItem.SubItems(2)
    .lblCliente = lvDetalle.SelectedItem.SubItems(4)
    .lblMarca = lvDetalle.SelectedItem.SubItems(6)
    .lblModelo = lvDetalle.SelectedItem.SubItems(7)
    .lblTotalMec = FormatoValor(lvDetalle.SelectedItem.SubItems(14), "", gintDecimalesMoneda)
    .lblTotalCar = FormatoValor(lvDetalle.SelectedItem.SubItems(15), "", gintDecimalesMoneda)
    .lblTotalOtr = FormatoValor(lvDetalle.SelectedItem.SubItems(16), "", gintDecimalesMoneda)
    .lblTotalTer = FormatoValor(lvDetalle.SelectedItem.SubItems(17), "", gintDecimalesMoneda)
    .lblTotalRep = FormatoValor(lvDetalle.SelectedItem.SubItems(18), "", gintDecimalesMoneda)
    .lblTotalMat = FormatoValor(lvDetalle.SelectedItem.SubItems(19), "", gintDecimalesMoneda)
    .lblTotalIns = FormatoValor(lvDetalle.SelectedItem.SubItems(20), "", gintDecimalesMoneda)
    .lblsubtotal = FormatoValor(lvDetalle.SelectedItem.SubItems(21), "", gintDecimalesMoneda)
    .lblIva = FormatoValor(lvDetalle.SelectedItem.SubItems(22), "", gintDecimalesMoneda)
    .lblTotalOT = FormatoValor(lvDetalle.SelectedItem.SubItems(23), "", gintDecimalesMoneda)
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
    gstrEstadoOT = Mid(Me.lvDetalle.SelectedItem.SubItems(1), 1, 1)
End If
Unload Me
End Sub
Private Sub Form_Activate()

If SW Then
    If Not Atributos("Glbl", "Tllr_30_0010", True, True, True, True) Then
        MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
        Unload Me
        Exit Sub
    End If

    pckFechaDesde = BOM(Date)
    pckFechaHasta = EOM(Date)
    pckLiquidaIni = BOM(Date)
    pckLiquidaFin = EOM(Date)
    cmdImprimir.Enabled = Atributos("Glbl", "Tllr_30_0010", True, True, True, True)
    If gstrProcedencia = "Presupuestos" Then
        cckPresupuesto.Value = 1
    End If
    SW = False
End If

End Sub

Private Sub Form_Load()
SW = True
Me.cckCriterios(1).Caption = gstrNombrePatente
Me.lvDetalle.ColumnHeaders(3).Text = gstrNombrePatente
End Sub

Private Sub lvDetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ReOrdenaLista lvDetalle, ColumnHeader
End Sub

Private Sub lvDetalle_DblClick()
If cmdSeleccionar.Enabled = True Then cmdSeleccionar.Value = True
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

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Screen.MousePointer = vbHourglass
    Select Case Button.Key
        Case "Buscar"
            BuscarOt
        Case "Exportar"
            ExportarDatos Me.lvDetalle, Me.cdExportar, Me.hwnd
        Case "Seleccion1"
            ObjHideColumnHeader Me.lvDetalle
        Case "Cerrar"
            Unload Me
    End Select
    Screen.MousePointer = vbDefault

End Sub

Private Sub tlbBarraHerramientas_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Screen.MousePointer = vbHourglass
    Select Case ButtonMenu.Index
        Case 1
            ImprimirInforme
        Case 2
            ImprimeMiReporte Me.lvDetalle, Me.cdExportar, "LISTADO DE OT"
        Case 3
            Me.cdExportar.CancelError = False
            Me.cdExportar.Flags = cdlCFPrinterFonts
            Me.cdExportar.ShowFont

    End Select
    Screen.MousePointer = vbDefault

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

Private Sub BuscarOt()
Dim mstrSQL As String
Dim mstrWhere As String
Dim adoTemp As New ADODB.Recordset
Dim AdoAux As New ADODB.Recordset
Dim itmItem As ListItem
Dim mstrEstado As String
Dim mstrNumeroDocumento As String

lvDetalle.ListItems.Clear
mstrWhere = "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "'"
With Me
    If .cckCriterios(0).Value = 1 Then  '////////// nro ot
        mstrWhere = mstrWhere & ",'" & MatchMode(txtNroOt, "Cualquier Parte del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(1).Value = 1 Then  '////////// patente
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtPatente, "Comienzo del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(10).Value = 1 Then  '////////// vin
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtVin, "Cualquier Parte del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(2).Value = 1 Then  '////////// marca
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtMarca, "Comienzo del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(3).Value = 1 Then  '////////// modelo
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtModelo, "Comienzo del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(4).Value = 1 Then  '////////// cliente
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtCliente, "Comienzo del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(5).Value = 1 Then  '////////// recepcionista
        mstrWhere = mstrWhere & ",'" & MatchMode(.txtRecepcionista, "Comienzo del Campo", apSqlServer) & "'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    If .cckCriterios(6).Value = 1 And .cckCriterios(7).Value = 1 Then   '////////// fecha iniciosi y terminosi
        mstrWhere = mstrWhere & ",'" & pckFechaDesde.Value & "','" & pckFechaHasta.Value & " 23:59:00" & "'"
    ElseIf .cckCriterios(6).Value = 0 And .cckCriterios(7).Value = 0 Then  '////////// fecha iniciono y terminono
        mstrWhere = mstrWhere & ",'',''"
    ElseIf .cckCriterios(6).Value = 1 And .cckCriterios(7).Value = 0 Then   '////////// fecha iniciosi y terminono
        mstrWhere = mstrWhere & ",'" & pckFechaDesde.Value & "',''"
    ElseIf .cckCriterios(6).Value = 0 And .cckCriterios(7).Value = 1 Then  '////////// fecha iniciono y terminosi
        mstrWhere = ",'','" & pckFechaHasta.Value & "'"
    End If
    
    '//////////////////////////////////////////////////////
    If .cckCriterios(8).Value = 1 And .cckCriterios(9).Value = 1 Then   '////////// fecha iniciosi y terminosi
        mstrWhere = mstrWhere & ",'" & pckLiquidaIni.Value & "','" & pckLiquidaFin.Value & " 23:59:00" & "'"
    ElseIf .cckCriterios(8).Value = 0 And .cckCriterios(9).Value = 0 Then  '////////// fecha iniciono y terminono
        mstrWhere = mstrWhere & ",'',''"
    ElseIf .cckCriterios(8).Value = 1 And .cckCriterios(9).Value = 0 Then   '////////// fecha iniciosi y terminono
        mstrWhere = mstrWhere & ",'" & pckLiquidaIni.Value & "',''"
    ElseIf .cckCriterios(8).Value = 0 And .cckCriterios(9).Value = 1 Then  '////////// fecha iniciono y terminosi
        mstrWhere = mstrWhere & ",'','" & pckLiquidaFin.Value & "'"
    End If

    '//////////////////estado
    
   If Me.cckReservas.Value = 1 And Me.cckPresupuesto.Value = 1 Then
       mstrWhere = mstrWhere & ",'RP'"
   ElseIf Me.cckReservas.Value = 1 And Me.cckPresupuesto.Value = 0 Then
       mstrWhere = mstrWhere & ",'R'"
   ElseIf Me.cckReservas.Value = 0 And Me.cckPresupuesto.Value = 1 Then
       mstrWhere = mstrWhere & ",'P'"
   ElseIf optTodas.Value = True Then
       mstrWhere = mstrWhere & ",'T'"
   ElseIf optVigente.Value = True Then
       mstrWhere = mstrWhere & ",'V'"
   ElseIf optLiquidada.Value = True Then
       mstrWhere = mstrWhere & ",'L'"
   ElseIf optCerrada.Value = True Then
       mstrWhere = mstrWhere & ",'F'"
   ElseIf optNula.Value = True Then
       mstrWhere = mstrWhere & ",'N'"
   
   End If
        
End With
'/////////////////////////////////////////////////////////////////////////////////
    
    mstrSQL = "Exec Tllr_listado_OT " & mstrWhere
    
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With adoTemp
       If Not .BOF And Not .EOF Then
          While Not .EOF
              Set itmItem = lvDetalle.ListItems.Add(, , !Id_OT)
              If !est = "B" Or !est = "F" Then
                'mstrNumeroDocumento = IIf(!NUMDOCUMENTO <> "S/N", !NUMDOCUMENTO, TraeNumeroDocumento(!Sec, !Id_OT, ""))
                mstrNumeroDocumento = !NUMDOCUMENTO
              End If
              itmItem.SubItems(1) = ValorNulo(IIf(!est = "L", "LIQUIDADA", IIf(!est = "V", "VIGENTE", IIf(!est = "N", "NULA", IIf(!est = "C", "CERRADA", IIf(!est = "F", "FACTURADA (" & mstrNumeroDocumento & ")", IIf(!est = "B", "BOLETEADA (" & mstrNumeroDocumento & ")", "OTRO")))))))
              itmItem.SubItems(2) = ValorNulo(!Pat)
              itmItem.SubItems(3) = ValorNulo(!VIN)
              itmItem.SubItems(4) = ValorNulo(!Cliente)
              itmItem.SubItems(5) = ValorNulo(!idCLI)
              itmItem.SubItems(6) = ValorNulo(!FONO)
              itmItem.SubItems(7) = ValorNulo(!Modelo)
              itmItem.SubItems(8) = Format(ValorNulo(!FEC), "dd/mm/yyyy")
              itmItem.SubItems(9) = Format(ValorNulo(!FECLIQ), "dd/mm/yyyy")
              itmItem.SubItems(10) = ValorNulo(!RECEP)
              itmItem.SubItems(11) = ValorNulo(IIf(!Sec = "M", "MECANICA", "CARROCERIA"))
              itmItem.SubItems(12) = ValorNulo(!GAR)
              itmItem.SubItems(13) = ValorNulo(!TIP)
              
              itmItem.SubItems(14) = ValorNulo(!Sec)
              
              itmItem.SubItems(15) = ValorNulo(!TMEC)
              itmItem.SubItems(16) = ValorNulo(!TCAR)
              itmItem.SubItems(17) = ValorNulo(!TOTR)
              itmItem.SubItems(18) = ValorNulo(!TTER)
              itmItem.SubItems(19) = ValorNulo(!TREP)
              itmItem.SubItems(20) = ValorNulo(!TMAT)
              itmItem.SubItems(20) = ValorNulo(!TINS)
              itmItem.SubItems(21) = ValorNulo(!Tneto)
              itmItem.SubItems(22) = ValorNulo(!Tiva)
              itmItem.SubItems(23) = ValorNulo(!Total)
              itmItem.SubItems(25) = ValorNulo(!Comentario)
              itmItem.SubItems(26) = ValorNulo(!CIASEG)
              
              adoTemp.MoveNext
          Wend
       End If
    End With
    End If
    Screen.MousePointer = 1
    lblTotal(7).Caption = lvDetalle.ListItems.Count
    mstrEstado = ""

End Sub
