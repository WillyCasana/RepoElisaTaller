VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRentabilidadOT 
   Caption         =   "Rentabilidad OT"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   Icon            =   "frmRentabilidadOT.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExcel 
      Appearance      =   0  'Flat
      Caption         =   "Excel"
      Height          =   360
      Left            =   7095
      TabIndex        =   30
      Top             =   7920
      Width           =   840
   End
   Begin Crystal.CrystalReport rptOT 
      Left            =   3960
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      Caption         =   "Imprimir Informe"
      Height          =   360
      Left            =   8010
      TabIndex        =   29
      Top             =   7920
      Width           =   1680
   End
   Begin VB.Frame Frame2 
      Height          =   3825
      Left            =   120
      TabIndex        =   6
      Top             =   -15
      Width           =   11370
      Begin VB.Frame Frame3 
         Caption         =   "Estado"
         Height          =   510
         Left            =   120
         TabIndex        =   40
         Top             =   2880
         Width           =   3525
         Begin VB.OptionButton optFacturadas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Facturadas"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2280
            TabIndex        =   43
            Top             =   180
            Value           =   -1  'True
            Width           =   1200
         End
         Begin VB.OptionButton optTodas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Todas"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   45
            TabIndex        =   42
            Top             =   180
            Width           =   855
         End
         Begin VB.OptionButton optLiquidadas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Liquidadas"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1080
            TabIndex        =   41
            Top             =   180
            Width           =   1065
         End
      End
      Begin VB.TextBox txtPorcentaje 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3960
         TabIndex        =   38
         Text            =   "0"
         Top             =   2280
         Width           =   1185
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Renta <= a"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   3960
         TabIndex        =   37
         Top             =   2025
         Width           =   1185
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sección"
         Height          =   510
         Left            =   10800
         TabIndex        =   31
         Top             =   1080
         Visible         =   0   'False
         Width           =   3525
         Begin VB.OptionButton optMecanica 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Mecánica"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1320
            TabIndex        =   34
            Top             =   180
            Width           =   1065
         End
         Begin VB.OptionButton optCarroceria 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Carrocería"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   45
            TabIndex        =   33
            Top             =   180
            Width           =   1215
         End
         Begin VB.OptionButton OptAmbas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Ambas"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2520
            TabIndex        =   32
            Top             =   180
            Value           =   -1  'True
            Width           =   840
         End
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Fec. Final"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   1800
         TabIndex        =   27
         Top             =   2025
         Value           =   1  'Checked
         Width           =   1920
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Costo Mano Obra"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   6000
         TabIndex        =   26
         Top             =   2025
         Width           =   1545
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Recepcionista"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   4200
         TabIndex        =   22
         Top             =   1185
         Width           =   1395
      End
      Begin VB.TextBox txtRecepcionista 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   21
         Top             =   1425
         Width           =   3975
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Nro OT"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   20
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox txtNroOt 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   105
         MaxLength       =   15
         TabIndex        =   19
         Top             =   525
         Width           =   1470
      End
      Begin VB.TextBox txtPatente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   15
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
         Left            =   1680
         TabIndex        =   14
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
         Left            =   2760
         TabIndex        =   13
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
         Left            =   5400
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   1185
         Width           =   795
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1425
         Width           =   3795
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   9
         Top             =   525
         Width           =   2565
      End
      Begin VB.TextBox txtModelo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5400
         MaxLength       =   50
         TabIndex        =   8
         Top             =   525
         Width           =   2835
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Fec. Inicial"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   2025
         Value           =   1  'Checked
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
               Picture         =   "frmRentabilidadOT.frx":179A
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":18AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":1D04
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":215C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":25B4
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":26C6
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":27D8
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":28EA
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":29FC
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":2B0E
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":2C20
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":2D32
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":2E44
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":2F56
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":3068
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":317A
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":328C
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":339E
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":34B0
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":35C2
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":3A14
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadOT.frx":3E66
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbMarca 
         Height          =   330
         Left            =   4920
         TabIndex        =   16
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
         Left            =   7800
         TabIndex        =   17
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
         Left            =   3480
         TabIndex        =   18
         Top             =   1095
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
         Left            =   7800
         TabIndex        =   23
         Top             =   1140
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
         TabIndex        =   24
         Top             =   2235
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   113377281
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckFechaHasta 
         Height          =   315
         Left            =   1800
         TabIndex        =   25
         Top             =   2235
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   113377281
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckLiquida 
         Height          =   315
         Left            =   5280
         TabIndex        =   28
         Top             =   3240
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   113377281
         CurrentDate     =   36776
      End
      Begin MSComctlLib.ListView lsvtipoOt 
         Height          =   1110
         Left            =   8400
         TabIndex        =   35
         Top             =   120
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   1958
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
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
      Begin MSAdodcLib.Adodc datCostoManoObra 
         Height          =   330
         Left            =   6795
         Top             =   135
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
      Begin MSDataListLib.DataCombo dtcCostoManoObra 
         Bindings        =   "frmRentabilidadOT.frx":3F78
         Height          =   315
         Left            =   6000
         TabIndex        =   36
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSComctlLib.ListView lsvCargos 
         Height          =   1110
         Left            =   8400
         TabIndex        =   39
         Top             =   1440
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   1958
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo de Cargo"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   2
         EndProperty
      End
      Begin MSComctlLib.ListView lsvTrabajo 
         Height          =   1110
         Left            =   8400
         TabIndex        =   44
         Top             =   2640
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   1958
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo de Cargo"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.CommandButton cmdBuscarOT 
      Appearance      =   0  'Flat
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   360
      Left            =   5280
      TabIndex        =   0
      Top             =   7920
      Width           =   1680
   End
   Begin VB.CommandButton cmdSalir 
      Appearance      =   0  'Flat
      Caption         =   "Salir"
      Height          =   360
      Left            =   9750
      TabIndex        =   2
      Top             =   7920
      Width           =   1680
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   3930
      Left            =   240
      TabIndex        =   5
      Top             =   3840
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
      Appearance      =   0
      NumItems        =   36
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N° OT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CARGO"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "N° DCTO."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Terc. Costo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Terc. Venta"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Terc. Diferencia"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Terc. Margen"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Rep. Costo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Rep. Venta"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Rep. Diferencia"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Rep. Margen"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Insumos Costo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Insumos Venta"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Insumos Dif."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "Insumos Margen"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "M.Obra Costo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "M.Obra Venta"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Text            =   "M.Obra Diferencia"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Text            =   "M.Obra Margen"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Text            =   "Carr. Costo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Text            =   "Carr. Venta"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Text            =   "Carr. Diferencia"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   22
         Text            =   "Carr. Margen"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Text            =   "Materiales"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Text            =   "Seguro Taller"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Text            =   "Descuentos"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   26
         Text            =   "Costo OT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   27
         Text            =   "Venta"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   28
         Text            =   "Deducibles"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   29
         Text            =   "Venta "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   30
         Text            =   "(S/.) Diferencia"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   31
         Text            =   "Margen Final"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   32
         Text            =   "DOCUMENTO"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "Fecha Facturación"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "Recepcionista"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "CIA Seguros"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "Seleccionar"
      Height          =   360
      Left            =   7980
      TabIndex        =   1
      Top             =   5730
      Width           =   1680
   End
   Begin MSComDlg.CommonDialog cdExportar 
      Left            =   3000
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   7
      Left            =   1935
      TabIndex        =   4
      Top             =   7080
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registros Encontrados :"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   3
      Top             =   8040
      Width           =   1695
   End
End
Attribute VB_Name = "frmRentabilidadOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SW As Boolean
Dim mstrSQL As String
Dim AdoPrincipal As New ADODB.Recordset

Sub ImprimirConsulta(strSalida As String)
Dim Dbsnueva As Database
Dim Tabla As DAO.Recordset
Dim i As Integer
Dim GcamBaseTem As String
Dim OTSeleccionada As String
Dim CargoSeleccionado As String

    On Error GoTo Solucion
    
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
''    If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    If Dir(gstrPathReporte & "\BDNueva.mdb") <> "" Then Kill gstrPathReporte & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
''    Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (NroOT text,NroDoc text, Cargo text, COSTOTER DOUBLE, TOTALTER DOUBLE, DIFTER DOUBLE, MARGENTER DOUBLE,COSTOREP DOUBLE, TOTALREP DOUBLE, DIFREP DOUBLE, MARGENREP DOUBLE, TOTALINS DOUBLE, COSTOMANO DOUBLE, TOTALMANO DOUBLE, DIFMANO DOUBLE, MARGEMANO DOUBLE, COSTOCARRO DOUBLE, TOTALCARRO DOUBLE, DIFCARRO DOUBLE, MARGENCARRO DOUBLE, DESCUENTOS DOUBLE, COSTOOT DOUBLE, VENTA DOUBLE, DEDUCIBLES DOUBLE, TOTALOT DOUBLE, DIFTOTAL DOUBLE, MARGENFINAL DOUBLE, SEGUROTALLER DOUBLE )"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To lvDetalle.ListItems.Count
        Set lvDetalle.SelectedItem = lvDetalle.ListItems(i)
        Tabla.AddNew
        Tabla!NroOT = IIf(lvDetalle.SelectedItem = "", " ", Mid(lvDetalle.SelectedItem, 6, 10))
        Tabla!CARGO = IIf(lvDetalle.SelectedItem.SubItems(1) = "", " ", lvDetalle.SelectedItem.SubItems(1))
        Tabla!NroDoc = IIf(lvDetalle.SelectedItem.SubItems(2) = "", " ", lvDetalle.SelectedItem.SubItems(2))
        Tabla!COSTOTER = CDbl(IIf(lvDetalle.SelectedItem.SubItems(3) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(3), "S/.")))
        Tabla!TOTALTER = CDbl(IIf(lvDetalle.SelectedItem.SubItems(4) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(4), "S/.")))
        Tabla!DIFTER = CDbl(IIf(lvDetalle.SelectedItem.SubItems(5) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(5), "S/.")))
        Tabla!MARGENTER = CDbl(IIf(lvDetalle.SelectedItem.SubItems(6) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(6), "%")))
        Tabla!COSTOREP = CDbl(IIf(lvDetalle.SelectedItem.SubItems(7) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(7), "S/.")))
        Tabla!TOTALREP = CDbl(IIf(lvDetalle.SelectedItem.SubItems(8) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(8), "S/.")))
        Tabla!DIFREP = CDbl(IIf(lvDetalle.SelectedItem.SubItems(9) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(9), "S/.")))
        Tabla!MARGENREP = CDbl(IIf(lvDetalle.SelectedItem.SubItems(10) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(10), "%")))
        Tabla!TOTALINS = CDbl(IIf(lvDetalle.SelectedItem.SubItems(12) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(12), "S/.")))
        Tabla!COSTOMANO = CDbl(IIf(lvDetalle.SelectedItem.SubItems(15) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(15), "S/.")))
        Tabla!TOTALMANO = CDbl(IIf(lvDetalle.SelectedItem.SubItems(16) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(16), "S/.")))
        Tabla!DIFMANO = CDbl(IIf(lvDetalle.SelectedItem.SubItems(17) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(17), "S/.")))
        Tabla!MARGEMANO = CDbl(IIf(lvDetalle.SelectedItem.SubItems(18) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(18), "%")))
        Tabla!COSTOCARRO = CDbl(IIf(lvDetalle.SelectedItem.SubItems(19) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(19), "S/.")))
        Tabla!TOTALCARRO = CDbl(IIf(lvDetalle.SelectedItem.SubItems(20) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(20), "S/.")))
        Tabla!DIFCARRO = CDbl(IIf(lvDetalle.SelectedItem.SubItems(21) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(21), "S/.")))
        Tabla!MARGENCARRO = CDbl(IIf(lvDetalle.SelectedItem.SubItems(22) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(22), "%")))
        Tabla!Descuentos = CDbl(IIf(lvDetalle.SelectedItem.SubItems(25) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(25), "S/.")))
        Tabla!COSTOOT = CDbl(IIf(lvDetalle.SelectedItem.SubItems(26) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(26), "S/.")))
        Tabla!VENTA = CDbl(IIf(lvDetalle.SelectedItem.SubItems(27) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(27), "S/.")))
        Tabla!DEDUCIBLES = CDbl(IIf(lvDetalle.SelectedItem.SubItems(28) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(28), "S/.")))
        Tabla!TotalOT = CDbl(IIf(lvDetalle.SelectedItem.SubItems(29) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(29), "S/.")))
        Tabla!DIFTOTAL = CDbl(IIf(lvDetalle.SelectedItem.SubItems(30) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(30), "S/.")))
        Tabla!MARGENFINAL = CDbl(IIf(lvDetalle.SelectedItem.SubItems(31) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(31), "%")))
        Tabla!SeguroTaller = CDbl(IIf(lvDetalle.SelectedItem.SubItems(24) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(24), "S/.")))
        Tabla.Update
    Next i
    Tabla.Close
    Dbsnueva.Close
    
    If strSalida = "Excel" Then
'        With rptOT
'            .Destination = crptToFile
'            .PrintFileType = crptExcel50Tab
'             '//MODIFICADO POR FDO DIAZ EL 29/11/2000
'             .ReportFileName = gstrPathReporte & "\RENOTEX.rpt"
'             .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
'             .Destination = crptToFile
'             .Action = True
'        End With
        ExportarDatos Me.lvDetalle, Me.cdExportar, Me.hwnd
    Else
        With rptOT
             '//MODIFICADO POR FDO DIAZ EL 29/11/2000
             .ReportFileName = gstrPathReporte & "\RENOT.rpt"
             .WindowTitle = "Rentabilidad por OT"
'             .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
             .DataFiles(0) = gstrPathReporte & "\BDNueva.mdb"
             .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
             .Formulas(1) = "TITULO='RENTABILIDAD POR OT'"
             .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
             .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
             .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
             
             
             If Me.cckCriterios(6).Value = 1 Then
                .Formulas(5) = "desde='" & Me.pckFechaDesde & "'"
                .Formulas(6) = "hasta='" & Me.pckFechaHasta & "'"
             End If
             
             If Me.optCarroceria.Value = True Then        ' POR CARROCERIA
                .Formulas(7) = "SECCION='CARROCERIA'"
             ElseIf Me.optMecanica.Value = True Then    ' POR MECANICA
                .Formulas(7) = "SECCION='MECANICA'"
             Else
                .Formulas(7) = "SECCION='MECANICA Y CARROCERIA'"
             End If
             
             OTSeleccionada = ""
             For i = 1 To Me.lsvtipoOt.ListItems.Count 'R
                If Me.lsvtipoOt.ListItems(i).Checked Then 'Si esta checkeada agrega a la formula
                    OTSeleccionada = IIf(OTSeleccionada = "", Me.lsvtipoOt.ListItems(i), OTSeleccionada & " - " & Me.lsvtipoOt.ListItems(i))
                End If
             Next
             .Formulas(8) = "TIPOOT='" & IIf(OTSeleccionada <> "", OTSeleccionada, "TODAS") & "'"
             .Formulas(9) = "MARCA='" & IIf(Me.txtMarca = "", "TODAS", Me.txtMarca) & "'"
             .Formulas(10) = "MODELO='" & IIf(Me.txtModelo = "", "TODOS", Me.txtModelo) & "'"
             .Formulas(11) = "CLIENTE='" & IIf(Me.txtCliente = "", "TODOS", Me.txtCliente) & "'"
             .Formulas(12) = "RECEPCIONISTA='" & IIf(Me.txtRecepcionista = "", "TODOS", Me.txtRecepcionista) & "'"
             .Formulas(13) = "TDecimal=" & gintDecimalesMoneda
             .Formulas(14) = "TSigla='" & gstrMonedaLocal & "'"
             
             If Me.optFacturadas.Value = True Then        ' ESTADO
                .Formulas(15) = "ESTADO='FACTURADAS'"
             ElseIf Me.optLiquidadas.Value = True Then
                .Formulas(15) = "ESTADO='LIQUIDADAS'"
             ElseIf Me.optTodas.Value = True Then
                .Formulas(15) = "ESTADO='TODAS'"
             End If
             
             CargoSeleccionado = ""
             For i = 1 To Me.lsvCargos.ListItems.Count 'R
                If Me.lsvCargos.ListItems(i).Checked Then 'Si esta checkeada agrega a la formula
                    CargoSeleccionado = IIf(CargoSeleccionado = "", Me.lsvCargos.ListItems(i), CargoSeleccionado & " - " & Me.lsvCargos.ListItems(i))
                End If
             Next
             .Formulas(16) = "CARGO='" & IIf(CargoSeleccionado <> "", CargoSeleccionado, "TODOS") & "'"
             .Formulas(17) = "OT='" & IIf(Me.txtNroOt = "", "TODAS", Me.txtNroOt) & "'"
             .Formulas(18) = "PATENTE='" & IIf(Me.txtPatente = "", "TODAS", Me.txtPatente) & "'"
             .Formulas(19) = "PORCRENTA='" & IIf(Me.txtPorcentaje = "0", "TODAS", Me.txtPorcentaje & " %") & "'"
             .Formulas(20) = "COSTOMO='" & IIf(Me.dtcCostoManoObra.Text = "", "FIJO", Me.dtcCostoManoObra.Text) & "'"
             
             .Destination = crptToWindow
             .Action = True
        End With
    End If
''    Dbsnueva.Close
    Screen.MousePointer = 1
   Exit Sub
Solucion:
    If Err.Number <> 0 Then
        MsgBox "Impresión Cancelada por el usuario", vbExclamation, "Imprimir"
        Screen.MousePointer = 1
        
    End If
    
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
        dtcCostoManoObra.Enabled = False
    Else
        dtcCostoManoObra.Enabled = True
        dtcCostoManoObra.SetFocus
    End If
Case 9
    If cckCriterios(Index).Value = 0 Then
        txtPorcentaje.Enabled = False
        txtPorcentaje = "0"
    Else
        txtPorcentaje = "0"
        txtPorcentaje.Enabled = True
        txtPorcentaje.SetFocus
    End If
    
End Select
End Sub


Private Sub cmdBuscarOT_Click()
Dim mstrSQL As String
Dim mstrWhere As String
Dim adoTemp As New ADODB.Recordset
Dim AdoAux As New ADODB.Recordset
Dim itmItem As ListItem
Dim lstrSQL As String
Dim AdoTot As New ADODB.Recordset
Dim CostoTerceros As Double
Dim CostoRepuestos As Double
Dim CostoInsumos As Double
Dim CostoManoObra As Double
Dim CostoCarroceria As Double
Dim ValorHora As Double
Dim i As Integer
Dim TotalDescuentos As Double
Dim SumaCostos As Double
Dim TotalOT As Double
Dim OtAux As String
Dim swAumentaLista As Boolean
Dim ValoresRetornados As VentaRepuestos
Dim mstrNumeroDocumento As String
Dim CostoInsumosPesos As Double
Dim CostoInsumosPorc As Double
Dim lstrCostea As String

CostoTerceros = 0
CostoRepuestos = 0
CostoInsumos = 0
CostoManoObra = 0
CostoCarroceria = 0
TotalDescuentos = 0
SumaCostos = 0
swAumentaLista = True

If Me.cckCriterios(8).Value = 1 Then  '/// factor costo mano de obra
    lstrSQL = "Select total From Tllr_CostoManoObra where Id_Empresa = '" & gstrIdEmpresa & "' and Id_Sucursal = '" & gstrIdSucursal & "' and id_mes='" & Me.dtcCostoManoObra.BoundText & "'"
    If Conexion.SendHost(lstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
           ValorHora = IIf(IsNull(adoTemp!Total), 0, adoTemp!Total)
        Else
           ValorHora = 0
        End If
    End If
    Conexion.CloseHost adoTemp
Else
    lstrSQL = "Select Valor_Mano_Costo From Tllr_Parametro where Id_Empresa = '" & gstrIdEmpresa & "' and Id_Sucursal = '" & gstrIdSucursal & "'"
    If Conexion.SendHost(lstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
           ValorHora = IIf(IsNull(adoTemp!VALOR_MANO_COSTO), 0, adoTemp!VALOR_MANO_COSTO)
        Else
           ValorHora = 0
        End If
    End If
    Conexion.CloseHost adoTemp
End If

'costo insumos
lstrSQL = "Select CostoInsumosPorc,CostoInsumosPesos From Tllr_Parametro where Id_Empresa = '" & gstrIdEmpresa & "' and Id_Sucursal = '" & gstrIdSucursal & "'"
If Conexion.SendHost(lstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not adoTemp.BOF And Not adoTemp.EOF Then
       CostoInsumosPesos = IIf(IsNull(adoTemp!CostoInsumosPesos), 0, adoTemp!CostoInsumosPesos)
       CostoInsumosPorc = IIf(IsNull(adoTemp!CostoInsumosPorc), 0, adoTemp!CostoInsumosPorc)
    Else
       CostoInsumosPesos = 0
       CostoInsumosPorc = 0
    End If
End If
Conexion.CloseHost adoTemp


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
    
    If .optCarroceria.Value = True Then ' POR CARROCERIA
        mstrWhere = mstrWhere & ",'C'"
    ElseIf .optMecanica.Value = True Then ' POR MECANICA
        mstrWhere = mstrWhere & ",'M'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    Dim lsw As Double
    lsw = False
    For i = 1 To Me.lsvtipoOt.ListItems.Count 'R
        If Me.lsvtipoOt.ListItems(i).Checked Then 'Si esta checkeada agrega al where
            If lsw = False Then 'Si es el primero usa AND
                mstrWhere = mstrWhere & ",'" & Chr(34) & Me.lsvtipoOt.ListItems(i).ListSubItems(1) & Chr(34)
                lsw = True
            Else
                mstrWhere = mstrWhere & "," & Chr(34) & Me.lsvtipoOt.ListItems(i).ListSubItems(1) & Chr(34)
            End If
        End If
    Next
    
    'Si alguna vez paso cierra el parentesis
     If lsw = True Then 'Si es el ultimo entonces cierra parentesis
        mstrWhere = mstrWhere & "'"
     Else
        mstrWhere = mstrWhere & ",''"
     End If
     
     'Tipo Trabajo
     lsw = False
    For i = 1 To Me.lsvTrabajo.ListItems.Count 'R
        If Me.lsvTrabajo.ListItems(i).Checked Then 'Si esta checkeada agrega al where
            If lsw = False Then 'Si es el primero usa AND
                mstrWhere = mstrWhere & ",'" & Chr(34) & Me.lsvTrabajo.ListItems(i).ListSubItems(1) & Chr(34)
                lsw = True
            Else
                mstrWhere = mstrWhere & "," & Chr(34) & Me.lsvTrabajo.ListItems(i).ListSubItems(1) & Chr(34)
            End If
        End If
    Next
    
    'Si alguna vez paso cierra el parentesis
     If lsw = True Then 'Si es el ultimo entonces cierra parentesis
        mstrWhere = mstrWhere & "'"
     Else
        mstrWhere = mstrWhere & ",''"
     End If
     
     
     
    'tipos de cargo
    lsw = False
    For i = 1 To Me.lsvCargos.ListItems.Count 'R
        If Me.lsvCargos.ListItems(i).Checked Then 'Si esta checkeada agrega al where
            If lsw = False Then 'Si es el primero usa AND
                mstrWhere = mstrWhere & ",'" & Chr(34) & Me.lsvCargos.ListItems(i).ListSubItems(1) & Chr(34)
                lsw = True
            Else
                mstrWhere = mstrWhere & "," & Chr(34) & Me.lsvCargos.ListItems(i).ListSubItems(1) & Chr(34)
            End If
        End If
    Next
    
    'Si alguna vez paso cierra el parentesis
    If lsw = True Then 'Si es el ultimo entonces cierra comillas
       mstrWhere = mstrWhere & "'"
    Else
       mstrWhere = mstrWhere & ",''"
    End If
    
    If .optFacturadas.Value = True Then ' facturadas
'        mstrWhere = mstrWhere & ",'F'"
        'kjcv 29.02.12
        mstrWhere = mstrWhere & ",'F',''"
    ElseIf .optLiquidadas.Value = True Then ' liquidadas
'        mstrWhere = mstrWhere & ",'L'"
        'kjcv 29.02.12
        mstrWhere = mstrWhere & ",'L',''"
    Else
'        mstrWhere = mstrWhere & ",''"  'todas
        'kjcv 29.02.12
         mstrWhere = mstrWhere & ",'',''"  'todas
    End If

    
End With

    '/// llama al procedimiento almacenado
    mstrSQL = "Exec Tllr_Rentabilidad_OT " & mstrWhere
    
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With adoTemp
       If Not .BOF And Not .EOF Then
          OtAux = !Id_OT   'Para comparar con la ot siguiente
          
          
          While Not .EOF
              'rescata el tipo de costo del cargo, es decir, muestra solo Costo si es afirmativo
              lstrCostea = Retorna_Valor_General("Select Costea from Tllr_Tipo_Cargo where Id_Empresa='" & gstrIdEmpresa & "' and id_tipo_Cargo='" & !Id_Cargo & "'", gcdynamic)
              
              'verifica si los costos son por marca
              If gblnPreciosMarca = True Then
                    mstrSQL = "SELECT CostoManoObra, CostoMOGarantia From Tllr_Marca_Precios_MO WHERE (Id_Marca = '" & !Id_Marca & "')"
                    If Conexion.SendHost(mstrSQL, AdoAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                        If Not AdoAux.BOF And Not AdoAux.EOF Then
                            ValorHora = IIf(!Id_Cargo = gstrCargoGtiaFabrica, AdoAux!CostoMOGarantia, AdoAux!CostoManoObra)
                        End If
                    End If
              End If
                
              If !Insumos <> 0 Then
                  If CostoInsumosPesos <> 0 Or CostoInsumosPorc <> 0 Then
                      CostoInsumos = IIf(CostoInsumosPesos = 0, (!Insumos * CostoInsumosPorc) / 100, CostoInsumosPesos)
                  Else
                      CostoInsumos = 0
                  End If
              Else
                  CostoInsumos = 0
              End If
              
              CostoTerceros = IIf(IsNull(!CostoTerceros), 0, !CostoTerceros)
'              CostoRepuestos = (IIf(IsNull(!SumaConsumoRepuesto), 0, !SumaConsumoRepuesto) - IIf(IsNull(!SumaDevolucionRepuesto), 0, !SumaDevolucionRepuesto))
              'kjcv 29.02.12
              CostoRepuestos = (IIf(IsNull(!CostoRepuestos), 0, !CostoRepuestos))
              CostoManoObra = (IIf(IsNull(!HorasMOM), 0, !HorasMOM) + IIf(IsNull(!HorasMOO), 0, !HorasMOO)) * ValorHora
              CostoCarroceria = IIf(IsNull(!CostoCarroceria), 0, !CostoCarroceria)
              
              SumaCostos = CostoInsumos + CostoTerceros + CostoRepuestos + CostoManoObra + CostoCarroceria
              
              TotalDescuentos = TotalDescuentos + IIf(IsNull(!Mdt), 0, !Mdt) + IIf(IsNull(!Mdr), 0, !Mdr) + IIf(IsNull(!Mdm), 0, !Mdm) + IIf(IsNull(!Mdo), 0, !Mdo) + IIf(IsNull(!Mdc), 0, !Mdc)
              TotalOT = !Total_General
                
              Dim dblTotalOT As Double
              If TotalOT <> 0 Then
                dblTotalOT = CDbl(((TotalOT - SumaCostos) * 100) / TotalOT)
              Else
                dblTotalOT = 0
              End If
              
              If dblTotalOT <= IIf(Me.cckCriterios(9).Value = 1, CDbl(txtPorcentaje), 501) Then  'rentabilidad de ot < a porcentaje
                '// numero ot
                mstrNumeroDocumento = ValorNulo(!Nro_Factura_Emitida)
                
                If swAumentaLista = False Then
                    swAumentaLista = True
                Else
                    Set itmItem = lvDetalle.ListItems.Add(, , !Id_OT)
                End If
                itmItem.SubItems(1) = ValorNulo(!Descripcion)
                itmItem.SubItems(2) = !estado & "-" & mstrNumeroDocumento
                '// terceros
                itmItem.SubItems(3) = FormatoValor(CostoTerceros, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(4) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!total_terceros, gstrMonedaLocal, gintDecimalesMoneda))
                itmItem.SubItems(5) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!total_terceros - CostoTerceros, gstrMonedaLocal, gintDecimalesMoneda))
                If !total_terceros <> 0 Then
                  itmItem.SubItems(6) = IIf(lstrCostea = "S", FormatoValor(0, "%", 2), FormatoValor(((!total_terceros - CostoTerceros) * 100) / !total_terceros, "%", 2))
                Else
                  itmItem.SubItems(6) = FormatoValor(0, "%", 2)
                End If
                '// repuestos
                itmItem.SubItems(7) = FormatoValor(CostoRepuestos, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(8) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!total_repuestos, gstrMonedaLocal, gintDecimalesMoneda))
                itmItem.SubItems(9) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!total_repuestos - CostoRepuestos, gstrMonedaLocal, gintDecimalesMoneda))
                If !total_repuestos <> 0 Then
                  itmItem.SubItems(10) = IIf(lstrCostea = "S", FormatoValor(0, "%", 2), FormatoValor(((!total_repuestos - CostoRepuestos) * 100) / !total_repuestos, "%", 2))
                Else
                  itmItem.SubItems(10) = FormatoValor(0, "%", 2)
                End If
                '//insumos
                itmItem.SubItems(11) = FormatoValor(CostoInsumos, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(12) = FormatoValor(!Insumos, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(13) = FormatoValor(!Insumos - CostoInsumos, gstrMonedaLocal, gintDecimalesMoneda)
                If !Insumos <> 0 Then
                    itmItem.SubItems(14) = FormatoValor(((!Insumos - CostoInsumos) * 100) / !Insumos, "%", 2)
                Else
                    itmItem.SubItems(14) = FormatoValor(0, "%", 2)
                End If
                '// mano de obra
                itmItem.SubItems(15) = FormatoValor(CostoManoObra, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(16) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!total_mano_obra, gstrMonedaLocal, gintDecimalesMoneda))
                itmItem.SubItems(17) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!total_mano_obra - CostoManoObra, gstrMonedaLocal, gintDecimalesMoneda))
                If !total_mano_obra <> 0 Then
                  itmItem.SubItems(18) = IIf(lstrCostea = "S", FormatoValor(0, "%", 2), FormatoValor(((!total_mano_obra - CostoManoObra) * 100) / !total_mano_obra, "%", 2))
                Else
                  itmItem.SubItems(18) = FormatoValor(0, "%", 2)
                End If
                '// carroceria
                itmItem.SubItems(19) = FormatoValor(CostoCarroceria, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(20) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!total_carroceria, gstrMonedaLocal, gintDecimalesMoneda))
                itmItem.SubItems(21) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!total_carroceria - CostoCarroceria, gstrMonedaLocal, gintDecimalesMoneda))
                If !total_carroceria <> 0 Then
                  itmItem.SubItems(22) = IIf(lstrCostea = "S", FormatoValor(0, "%", 2), FormatoValor(((!total_carroceria - CostoCarroceria) * 100) / !total_carroceria, "%", 2))
                Else
                  itmItem.SubItems(22) = FormatoValor(0, "%", 2)
                End If
                
                '//Materiales
                itmItem.SubItems(23) = FormatoValor(ValorNulo(!Materiales), gstrMonedaLocal, gintDecimalesMoneda)
                '//Seguro Taller
                itmItem.SubItems(24) = FormatoValor(ValorNulo(!SeguroTaller), gstrMonedaLocal, gintDecimalesMoneda)
                
                '// Totales OT
                itmItem.SubItems(25) = FormatoValor(TotalDescuentos, gstrMonedaLocal, gintDecimalesMoneda)   '// Descuentos
                itmItem.SubItems(26) = FormatoValor(SumaCostos, gstrMonedaLocal, gintDecimalesMoneda)        '// Total Costos
                itmItem.SubItems(27) = FormatoValor(!Total_General, gstrMonedaLocal, gintDecimalesMoneda)    '// venta Ot
'                itmItem.SubItems(28) = IIf(!Id_Cargo = gstrCargoDeducibleMenos, 0, FormatoValor(ValorNulo(!deducible_pesos), gstrMonedaLocal, gintDecimalesMoneda)) '// Deducibles
                'kjcv 09.04.13 Deducibles
                itmItem.SubItems(28) = FormatoValor(ValorNulo(!deducible_pesos), gstrMonedaLocal, gintDecimalesMoneda) '// Deducibles
                itmItem.SubItems(29) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!Total_General, gstrMonedaLocal, gintDecimalesMoneda)) '// Total Ot
                
'                itmItem.SubItems(28) = FormatoValor(ValorNulo(!deducible_pesos), gstrMonedaLocal, gintDecimalesMoneda) '// Deducibles
'                If !deducible_pesos <> 0 Then
'                    If !Id_Cargo = gstrCargoDeducibleMenos Then
'                        itmItem.SubItems(29) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!Total_General + !deducible_pesos, gstrMonedaLocal, gintDecimalesMoneda)) '// Total Ot
'                    ElseIf !Id_Cargo = gstrCargoDeducibleMas Then
'                        itmItem.SubItems(29) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!Total_General - !deducible_pesos, gstrMonedaLocal, gintDecimalesMoneda)) '// Total Ot
'                    Else
'                        itmItem.SubItems(29) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!Total_General, gstrMonedaLocal, gintDecimalesMoneda)) '// Total Ot
'                    End If
'                Else
'                    itmItem.SubItems(29) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!Total_General, gstrMonedaLocal, gintDecimalesMoneda)) '// Total Ot
'                End If
                
                'FormatoValor(TotalOT, gstrmonedalocal, 0)
                
                If TotalOT <> 0 Then
                  itmItem.SubItems(30) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!Total_General - SumaCostos, gstrMonedaLocal, gintDecimalesMoneda))
                  itmItem.SubItems(31) = IIf(lstrCostea = "S", FormatoValor(0, "%", 2), FormatoValor(((!Total_General - SumaCostos) * 100) / TotalOT, "%", 2))
                Else
                  itmItem.SubItems(30) = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
                  itmItem.SubItems(31) = FormatoValor(0, "%", 2)
                End If
                'para ordenar por fecha
                itmItem.SubItems(32) = !estado & "-" & Format(mstrNumeroDocumento, "0000000000")
                itmItem.SubItems(33) = ValorNulo(!Fecha_Facturacion)
                'kjcv 28.10.13 se adiciono recepcionista
                itmItem.SubItems(34) = ValorNulo(!Recepcionista)
                itmItem.SubItems(35) = ValorNulo(!CIASEG)
                
              End If
              TotalDescuentos = 0
              SumaCostos = 0
              adoTemp.MoveNext
              
              'si la ot se repite (cia y  deducible-cliente) solo se deja la cia
'              If Not .BOF And Not .EOF Then
'                If OtAux = !Id_OT Then
'                  If !id_cargo = "02" Then
'                      swAumentaLista = False
'                  End If
'                Else
'                  OtAux = !Id_OT
'                End If
'              End If
              
          Wend
       End If
    End With
    End If
    Screen.MousePointer = 1
    lblTotal(7).Caption = lvDetalle.ListItems.Count
    
End Sub

Private Sub cmdExcel_Click()
    If lvDetalle.ListItems.Count > 0 Then
        ImprimirConsulta "Excel"
    Else
        MsgBox "No Existen elemenetos en la lista"
    End If
End Sub

Private Sub cmdImprimir_Click()
If lvDetalle.ListItems.Count > 0 Then
    ImprimirConsulta "Impresora"
    
Else
    MsgBox "No Existen elementos en la lista"
End If
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

    If Not Atributos("Glbl", "Tllr_30_0080", True, True, True, True) Then
        MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
        Unload Me
        Exit Sub
    End If

    pckFechaDesde = BOM(Date)
    pckFechaHasta = EOM(Date)
    FillCostoManoObra
    SW = False
End If

End Sub

Private Sub Form_Load()
Dim AdoPaso As New ADODB.Recordset
Dim Item As ListItem
SW = True

    If Not Conexion.SendHost("Select Descripcion, Id_Garantia From Tllr_Garantias Where Vigencia='S' and Id_Empresa='" & gstrIdEmpresa & "' ", AdoPaso, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        MsgBox "Error en Conexion con el Host...", vbCritical, "Stock Pro"
        End
    End If

    If Not (AdoPaso.EOF = True And AdoPaso.BOF = True) Then
        Do Until AdoPaso.EOF
            Set Item = Me.lsvtipoOt.ListItems.Add(, , ValorNulo(AdoPaso.Fields(0)))
            Item.SubItems(1) = ValorNulo(AdoPaso.Fields(1))
            AdoPaso.MoveNext
        Loop
    End If
    AdoPaso.Close
    
    'cargos
    If Not Conexion.SendHost("Select Descripcion, Id_Tipo_Cargo From Tllr_Tipo_Cargo Where Id_Empresa='" & gstrIdEmpresa & "' and Vigencia='S' and Id_Empresa='" & gstrIdEmpresa & "'  ", AdoPaso, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        MsgBox "Error en Conexion con el Host...", vbCritical, "Taller Pro"
        End
    End If

    If Not (AdoPaso.EOF = True And AdoPaso.BOF = True) Then
        Do Until AdoPaso.EOF
            Set Item = Me.lsvCargos.ListItems.Add(, , ValorNulo(AdoPaso.Fields(0)))
            Item.SubItems(1) = ValorNulo(AdoPaso.Fields(1))
            AdoPaso.MoveNext
        Loop
    End If
    AdoPaso.Close
    
    'Trabajos
    If Not Conexion.SendHost("Select Descripcion, Id_Trabajo From Tllr_Trabajo Where Vigencia='S' and Id_Empresa='" & gstrIdEmpresa & "' ", AdoPaso, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        MsgBox "Error en Conexion con el Host...", vbCritical, "Stock Pro"
        End
    End If

    If Not (AdoPaso.EOF = True And AdoPaso.BOF = True) Then
        Do Until AdoPaso.EOF
            Set Item = Me.lsvTrabajo.ListItems.Add(, , ValorNulo(AdoPaso.Fields(0)))
            Item.SubItems(1) = ValorNulo(AdoPaso.Fields(1))
            AdoPaso.MoveNext
        Loop
    End If
    AdoPaso.Close
    
    
     MesManoObra(1) = "ENERO"
     MesManoObra(2) = "FEBRERO"
     MesManoObra(3) = "MARZO"
     MesManoObra(4) = "ABRIL"
     MesManoObra(5) = "MAYO"
     MesManoObra(6) = "JUNIO"
     MesManoObra(7) = "JULIO"
     MesManoObra(8) = "AGOSTO"
     MesManoObra(9) = "SEPTIEMBRE"
    MesManoObra(10) = "OCTUBRE"
    MesManoObra(11) = "NOVIEMBRE"
    MesManoObra(12) = "DICIEMBRE"
    
    Me.cckCriterios(1).Caption = gstrNombrePatente
End Sub

Private Sub lvDetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Select Case ColumnHeader.Index
Case 3
    ReOrdenaListaNumero lvDetalle, 32
Case Else
    ReOrdenaLista lvDetalle, ColumnHeader
End Select
End Sub

Private Sub tlbCliente_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "Buscar" Then
'    gstrBusca = apfFormulario.BuscarRegistroClientes(Conexion, "Id_Cliente_Proveedor", "Razon_Social", gstrIdEmpresa)
'    'gstrBusca = apfFormulario.BuscarRegistroClientes(Conexion, "Id_Cliente_Proveedor", "Razon_Social")
'    txtCliente.Tag = gstrBusca
'    txtCliente = ClienteD(gstrBusca)
    'kjcv 02-02-2012
    gstrRutCliente = ""
    gstrNombreCliente = ""
    Libreria.ClienteBuscar Conexion, gstrRutCliente, gstrNombreCliente, gstrIdEmpresa, gstrIdSucursal, gstrIdUsuario
         If gstrRutCliente <> "" Then
            Me.txtCliente = gstrNombreCliente
            Me.txtCliente.Tag = gstrRutCliente
        End If
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

'Private Sub txtNroRecord_KeyPress(KeyAscii As Integer)
'KeyAscii = CheckNumber(KeyAscii, txtNroRecord, strComa)
'End Sub

Private Sub txtPatente_KeyPress(KeyAscii As Integer)
''KeyAscii = CheckIdCar(txtPatente.SelStart, mdLLNNNN, UpCaseLetter(KeyAscii))
'KeyAscii = UpCaseLetter(KeyAscii)
'kjcv 24-01-12 Valida Letras y numeros
If (KeyAscii <> 8) And Not (KeyAscii >= 48 And KeyAscii <= 57) And Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
    KeyAscii = 0: Beep
Else
    KeyAscii = UpCaseLetter(KeyAscii)
End If

End Sub

Private Sub txtPorcentaje_GotFocus()
MarcaTexto txtPorcentaje
End Sub

Private Sub txtPorcentaje_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtPorcentaje, strDot)
End Sub

Private Sub txtPorcentaje_LostFocus()
If cckCriterios(9).Value = 1 Then
    If CDbl(txtPorcentaje) < 0 Or CDbl(txtPorcentaje) > 100 Then
        MsgBox "Valor del porcentaje Mal Ingresado", vbExclamation, "% Rentabilidad de OT"
        txtPorcentaje.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub txtRecepcionista_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub
Sub FillCostoManoObra()

    dtcCostoManoObra.Enabled = True
    mstrSQL = "SELECT Id_Mes AS Codigo, Descripcion AS Nombre FROM Tllr_CostoManoObra Where id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Ano=" & Year(Date)
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With datCostoManoObra
            Set .Recordset = AdoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcCostoManoObra.ListField = "Nombre"
                dtcCostoManoObra.BoundColumn = "Codigo"
                dtcCostoManoObra.BoundText = .Recordset!Codigo
                If .Recordset.RecordCount < 2 Then dtcCostoManoObra.Enabled = False
            End If
        End With
    End If ' por el otro
    Set AdoPrincipal = New ADODB.Recordset
    Conexion.CloseHost AdoPrincipal
    dtcCostoManoObra.Text = MesManoObra(Month(Date))
End Sub

