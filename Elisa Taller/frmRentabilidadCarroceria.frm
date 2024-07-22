VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRentabilidadCarroceria 
   Caption         =   "Rentabilidad Carroceria"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13095
   Icon            =   "frmRentabilidadCarroceria.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   13095
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExcel 
      Appearance      =   0  'Flat
      Caption         =   "Excel"
      Height          =   360
      Left            =   8655
      TabIndex        =   30
      Top             =   6855
      Width           =   840
   End
   Begin Crystal.CrystalReport rptOT 
      Left            =   3945
      Top             =   6810
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
      Top             =   7440
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Frame Frame2 
      Height          =   2625
      Left            =   60
      TabIndex        =   6
      Top             =   -15
      Width           =   12915
      Begin VB.Frame Frame3 
         Caption         =   "Estado"
         Height          =   510
         Left            =   4080
         TabIndex        =   42
         Top             =   1920
         Width           =   3525
         Begin VB.OptionButton optLiquidadas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Liquidadas"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1080
            TabIndex        =   45
            Top             =   180
            Width           =   1065
         End
         Begin VB.OptionButton optTodas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Todas"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   45
            TabIndex        =   44
            Top             =   180
            Width           =   855
         End
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
      End
      Begin VB.CheckBox chkShowInternas 
         Caption         =   "Mostrar Internas"
         Height          =   255
         Left            =   4080
         TabIndex        =   40
         Top             =   1800
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox chkDescuentos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Aplica Descuentos"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   8040
         TabIndex        =   39
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtPorcentaje 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   8505
         TabIndex        =   38
         Text            =   "0"
         Top             =   525
         Width           =   1065
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Renta <= a"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   8505
         TabIndex        =   37
         Top             =   300
         Width           =   1185
      End
      Begin VB.Frame Frame1 
         Height          =   510
         Left            =   11040
         TabIndex        =   31
         Top             =   1200
         Visible         =   0   'False
         Width           =   3765
         Begin VB.OptionButton optMecanica 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Mecánica"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1395
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
            Left            =   2775
            TabIndex        =   32
            Top             =   180
            Value           =   -1  'True
            Width           =   840
         End
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Fecha Emisión (Final)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   2055
         TabIndex        =   27
         Top             =   1905
         Value           =   1  'Checked
         Width           =   1920
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Costo Mano Obra"
         Height          =   195
         Index           =   8
         Left            =   5640
         TabIndex        =   26
         Top             =   1560
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Recepcionista"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   4215
         TabIndex        =   22
         Top             =   1065
         Width           =   1395
      End
      Begin VB.TextBox txtRecepcionista 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4275
         MaxLength       =   50
         TabIndex        =   21
         Top             =   1305
         Width           =   3495
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
         Width           =   1725
      End
      Begin VB.TextBox txtPatente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1905
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
         Left            =   1920
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
         Left            =   2985
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
         Left            =   5595
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
         Top             =   1065
         Width           =   795
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1305
         Width           =   3915
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2985
         MaxLength       =   50
         TabIndex        =   9
         Top             =   525
         Width           =   2565
      End
      Begin VB.TextBox txtModelo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   5595
         MaxLength       =   50
         TabIndex        =   8
         Top             =   525
         Width           =   2835
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Fecha Emisión (Inicial)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   1905
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
               Picture         =   "frmRentabilidadCarroceria.frx":179A
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":18AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":1D04
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":215C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":25B4
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":26C6
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":27D8
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":28EA
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":29FC
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":2B0E
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":2C20
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":2D32
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":2E44
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":2F56
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":3068
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":317A
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":328C
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":339E
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":34B0
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":35C2
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":3A14
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadCarroceria.frx":3E66
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbMarca 
         Height          =   330
         Left            =   5145
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
         Left            =   7950
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
         Left            =   3690
         TabIndex        =   18
         Top             =   975
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
         Left            =   7365
         TabIndex        =   23
         Top             =   1005
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
         Top             =   2115
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   177799169
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckFechaHasta 
         Height          =   315
         Left            =   2055
         TabIndex        =   25
         Top             =   2115
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   177799169
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckLiquida 
         Height          =   315
         Left            =   5985
         TabIndex        =   28
         Top             =   1755
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   177799169
         CurrentDate     =   36776
      End
      Begin MSComctlLib.ListView lsvtipoOt 
         Height          =   1110
         Left            =   9780
         TabIndex        =   35
         Top             =   225
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
         Bindings        =   "frmRentabilidadCarroceria.frx":3F78
         Height          =   315
         Left            =   5520
         TabIndex        =   36
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSComctlLib.ListView lsvCargos 
         Height          =   1110
         Left            =   9840
         TabIndex        =   41
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
   End
   Begin VB.CommandButton cmdBuscarOT 
      Appearance      =   0  'Flat
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   360
      Left            =   6765
      TabIndex        =   0
      Top             =   6855
      Width           =   1680
   End
   Begin VB.CommandButton cmdSalir 
      Appearance      =   0  'Flat
      Caption         =   "Salir"
      Height          =   360
      Left            =   9750
      TabIndex        =   2
      Top             =   6855
      Width           =   1680
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   3930
      Left            =   45
      TabIndex        =   5
      Top             =   2775
      Width           =   12915
      _ExtentX        =   22781
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
      NumItems        =   37
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N° OT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Sección/N°Factura"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Planc. Costo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Planc. Venta"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Planc. Diferencia"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Planc. Margen"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Pint. Costo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Pint. Venta"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Pint. Diferencia"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Pint. Margen"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Arm/Des Costo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Arm/Des Vental"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Arm/Des Diferencia"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Arm/Des Margen"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "Terc. Costo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "Terc. Venta"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Terc. Diferencia"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Text            =   "Terc. Margen"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Text            =   "Rep. Costo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Text            =   "Rep. Venta"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Text            =   "Rep. Diferencia"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Text            =   "Rep. Margen"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   22
         Text            =   "Insumos"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Text            =   "M.Obra Costo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Text            =   "M.Obra Venta"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Text            =   "M.Obra Dif."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   26
         Text            =   "M.Obra Margen"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   27
         Text            =   "Total Costo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   28
         Text            =   "Total Neto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   29
         Text            =   "Descuento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   30
         Text            =   "Total Venta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   31
         Text            =   "(S/.) Diferencia"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   32
         Text            =   "Margen Final"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   33
         Text            =   "Cargo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "Fecha Facturación"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "Recepcionista"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Text            =   "Placa"
         Object.Width           =   2999
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
      Left            =   2760
      Top             =   6840
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
      Top             =   6990
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registros Encontrados :"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   3
      Top             =   6990
      Width           =   1695
   End
End
Attribute VB_Name = "frmRentabilidadCarroceria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SW As Boolean
Dim mstrSql As String
Dim adoPrincipal As New ADODB.Recordset

Sub ImprimirConsulta(strSalida As String)
Dim Dbsnueva As Database
Dim Tabla As DAO.Recordset
Dim i As Integer
Dim GcamBaseTem As String
Dim OTSeleccionada As String

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
'    If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    'kjcv 23.04.13
    If Dir(gstrPathReporte & "\BDNueva.mdb") <> "" Then Kill gstrPathReporte & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    'Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (NroOT text,Cargo text, COSTODES DOUBLE, VENTADES DOUBLE, DIFDES DOUBLE, MARGENDES DOUBLE,COSTOPIN DOUBLE, VENTAPIN DOUBLE, DIFPIN DOUBLE, MARGENPIN DOUBLE, COSTOARM DOUBLE, VENTAARM DOUBLE, DIFARM DOUBLE, MARGENARM DOUBLE, TOTALCOSTO DOUBLE, TOTALNETO DOUBLE, DESCUENTOS DOUBLE, TOTALVENTA DOUBLE, DIFTOTAL DOUBLE, MARGENFINAL DOUBLE )"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
        
    For i = 1 To lvDetalle.ListItems.Count
        Set lvDetalle.SelectedItem = lvDetalle.ListItems(i)
        Tabla.AddNew
        Tabla!NroOT = IIf(lvDetalle.SelectedItem = "", " ", lvDetalle.SelectedItem)
        Tabla!CARGO = IIf(lvDetalle.SelectedItem.SubItems(1) = "", " ", lvDetalle.SelectedItem.SubItems(1))
        Tabla!COSTODES = CDbl(IIf(lvDetalle.SelectedItem.SubItems(2) = "", " ", lvDetalle.SelectedItem.SubItems(2)))
        Tabla!VENTADES = CDbl(IIf(lvDetalle.SelectedItem.SubItems(3) = "", " ", lvDetalle.SelectedItem.SubItems(3)))
        Tabla!DIFDES = CDbl(IIf(lvDetalle.SelectedItem.SubItems(4) = "", " ", lvDetalle.SelectedItem.SubItems(4)))
        Tabla!MARGENDES = CDbl(IIf(lvDetalle.SelectedItem.SubItems(5) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(5), "%")))
        Tabla!COSTOPIN = CDbl(IIf(lvDetalle.SelectedItem.SubItems(6) = "", " ", lvDetalle.SelectedItem.SubItems(6)))
        Tabla!VENTAPIN = CDbl(IIf(lvDetalle.SelectedItem.SubItems(7) = "", " ", lvDetalle.SelectedItem.SubItems(7)))
        Tabla!DIFPIN = CDbl(IIf(lvDetalle.SelectedItem.SubItems(8) = "", " ", lvDetalle.SelectedItem.SubItems(8)))
        Tabla!MARGENPIN = CDbl(IIf(lvDetalle.SelectedItem.SubItems(9) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(9), "%")))
        Tabla!COSTOARM = CDbl(IIf(lvDetalle.SelectedItem.SubItems(10) = "", " ", lvDetalle.SelectedItem.SubItems(10)))
        Tabla!VENTAARM = CDbl(IIf(lvDetalle.SelectedItem.SubItems(11) = "", " ", lvDetalle.SelectedItem.SubItems(11)))
        Tabla!DIFARM = CDbl(IIf(lvDetalle.SelectedItem.SubItems(12) = "", " ", lvDetalle.SelectedItem.SubItems(12)))
        Tabla!MARGENARM = CDbl(IIf(lvDetalle.SelectedItem.SubItems(13) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(13), "%")))
        Tabla!TOTALCOSTO = CDbl(IIf(lvDetalle.SelectedItem.SubItems(14) = "", " ", lvDetalle.SelectedItem.SubItems(14)))
        Tabla!TOTALNETO = CDbl(IIf(lvDetalle.SelectedItem.SubItems(15) = "", " ", lvDetalle.SelectedItem.SubItems(15)))
        Tabla!Descuentos = CDbl(IIf(lvDetalle.SelectedItem.SubItems(16) = "", " ", lvDetalle.SelectedItem.SubItems(16)))
        Tabla!TOTALVENTA = CDbl(IIf(lvDetalle.SelectedItem.SubItems(17) = "", " ", lvDetalle.SelectedItem.SubItems(17)))
        Tabla!DIFTOTAL = CDbl(IIf(lvDetalle.SelectedItem.SubItems(18) = "", " ", lvDetalle.SelectedItem.SubItems(18)))
        Tabla!MARGENFINAL = CDbl(IIf(lvDetalle.SelectedItem.SubItems(19) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(19), "%")))
        Tabla.Update
    Next i
   Tabla.Close
   
   Dbsnueva.Close
   
    If strSalida = "Excel" Then
   '        With rptOT
'            .Destination = crptToFile
'            .PrintFileType = crptExcel50Tab
'             '//MODIFICADO POR FDO DIAZ EL 29/11/2000
'             .ReportFileName = gstrPathReporte & "\RENCARROCERIAEX.rpt"
'             .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
'             .Destination = crptToFile
'             .Action = True
'        End With
 'kjcv 23.04.13 Exportar en excel
        ExportarDatos Me.lvDetalle, Me.cdExportar, Me.hwnd
    Else
        With rptOT
             '//MODIFICADO POR FDO DIAZ EL 29/11/2000
             .ReportFileName = gstrPathReporte & "\RENCARROCERIA.rpt"
             .WindowTitle = "Rentabilidad por Carroceria"
'             .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
             .DataFiles(0) = gstrPathReporte & "\BDNueva.mdb"
             .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
             .Formulas(1) = "TITULO='RENTABILIDAD POR CARROCERIA'"
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
             .Destination = crptToWindow
             .Action = True
        End With
    End If
''   Dbsnueva.Close
   Screen.MousePointer = 1
   Exit Sub
Solucion:
    If Err.Number <> 0 Then
        MsgBox "Impresión Cancelada por el usuario", vbExclamation, "Imprimir"
        Screen.MousePointer = 1
        Exit Sub
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
Dim mstrSql As String
Dim mstrWhere As String
Dim adoTemp As New ADODB.Recordset
Dim AdoAux As New ADODB.Recordset
Dim itmItem As ListItem
Dim lstrSQL As String
Dim AdoTot As New ADODB.Recordset
Dim CostoDesabolladura As Double
Dim VentaDesabolladura As Double
Dim VentaTerceros As Double
Dim VentaRepuestos As Double
Dim VentaManoObra As Double
Dim CostoPintura As Double
Dim VentaPintura As Double
Dim CostoArmeDesarme As Double
Dim VentaArmeDesarme As Double
Dim ValorHora As Double
Dim i As Integer
Dim TotalDescuentos As Double
Dim SumaCostos As Double
Dim SumaVentas As Double
Dim TotalOT As Double
Dim mstrNumeroDocumento As String
Dim lstrCostea As String
Dim CostoTerceros As Double
Dim CostoRepuestos As Double
Dim CostoInsumos As Double
Dim CostoManoObra As Double

CostoDesabolladura = 0
VentaDesabolladura = 0
CostoPintura = 0
VentaPintura = 0
CostoArmeDesarme = 0
VentaArmeDesarme = 0
TotalDescuentos = 0
SumaCostos = 0
SumaVentas = 0

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
                mstrWhere = mstrWhere & ",' and (Tllr_Ot.Id_Garantia=" & Chr(34) & Me.lsvtipoOt.ListItems(i).ListSubItems(1) & Chr(34)
                lsw = True
            Else
                mstrWhere = mstrWhere & " OR Tllr_Ot.Id_Garantia=" & Chr(34) & Me.lsvtipoOt.ListItems(i).ListSubItems(1) & Chr(34)
            End If
        End If
    Next
    
   
    
    
    'Si alguna vez paso cierra el parentesis
     If lsw = True Then 'Si es el ultimo entonces cierra parentesis
        mstrWhere = mstrWhere & ")'"
     Else
        mstrWhere = mstrWhere & ",''"
     End If
     
      'kjcv 22.04.13
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
    
      'kjcv 30.10.13 se agrega Estado de facturacion
    If .optFacturadas.Value = True Then ' facturadas
        mstrWhere = mstrWhere & ",'F'"
    ElseIf .optLiquidadas.Value = True Then ' liquidadas
        mstrWhere = mstrWhere & ",'L'"
    Else
         mstrWhere = mstrWhere & ",''"  'todas
    End If
    
    
     
'    If Me.chkShowInternas.Value = vbChecked Then
'        mstrWhere = mstrWhere & ",'1'"
'    Else
'        mstrWhere = mstrWhere & ",'0'"
'    End If
    
End With

'/////////////////////////////////////////////////////////////////////////////////
    
    '/// llama al procedimiento almacenado
    If Me.chkDescuentos.Value = 0 Then
        mstrSql = "Exec Tllr_Rentabilidad_Carroceria " & mstrWhere
    Else
        mstrSql = "Exec Tllr_Rentabilidad_Carroceria_Descuentos " & mstrWhere
    End If
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With adoTemp
       If Not .BOF And Not .EOF Then
          While Not .EOF
          
            'rescata el tipo de costo del cargo, es decir, muestra solo Costo si es afirmativo
            lstrCostea = Retorna_Valor_General("Select Costea from Tllr_Tipo_Cargo where Id_empresa='" & gstrIdEmpresa & "' and id_tipo_Cargo='" & !Id_Cargo & "'", gcdynamic)
            Dim s As Integer
            s = adoTemp.RecordCount
            
            'kjcv 04.11.13 se agrega campos de MO, terceros,insumos
            'verifica si los costos son por marca
            If gblnPreciosMarca = True Then
                  mstrSql = "SELECT CostoManoObra, CostoMOGarantia From Tllr_Marca_Precios_MO WHERE (Id_Marca = '" & !Id_Marca & "')"
                  If Conexion.SendHost(mstrSql, AdoAux, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                      If Not AdoAux.BOF And Not AdoAux.EOF Then
                          ValorHora = IIf(!Id_Cargo = gstrCargoGtiaFabrica, AdoAux!CostoMOGarantia, AdoAux!CostoManoObra)
                      End If
                  End If
            End If
            
                         
              
                            
            
            If !Total_General <> !deducible_pesos Then
                CostoDesabolladura = IIf(IsNull(!CostoD), 0, !CostoD)
                VentaDesabolladura = IIf(IsNull(!VentaD), 0, !VentaD)
                
                CostoPintura = IIf(IsNull(!CostoP), 0, !CostoP)
                VentaPintura = IIf(IsNull(!VentaP), 0, !VentaP)
                
                CostoArmeDesarme = IIf(IsNull(!CostoA), 0, !CostoA)
                VentaArmeDesarme = IIf(IsNull(!VentaA), 0, !VentaA)
                
                CostoTerceros = IIf(IsNull(!CostoTerceros), 0, !CostoTerceros)
                VentaTerceros = IIf(IsNull(!total_terceros), 0, !total_terceros)
                
                CostoRepuestos = (IIf(IsNull(!CostoRepuestos), 0, !CostoRepuestos))
                VentaRepuestos = IIf(IsNull(!total_repuestos), 0, !total_repuestos)
                
                CostoManoObra = (IIf(IsNull(!HorasMOM), 0, !HorasMOM) + IIf(IsNull(!HorasMOO), 0, !HorasMOO)) * ValorHora
                VentaManoObra = IIf(IsNull(!total_mano_obra), 0, !total_mano_obra)
                
                
                
'                TotalDescuentos = IIf(Me.chkDescuentos.Value = 1, 0, IIf(IsNull(!Descuentos), 0, !Descuentos))
                TotalDescuentos = IIf(Me.chkDescuentos.Value = 1, 0, IIf(IsNull(!Descuentos), 0, !Descuentos)) + IIf(IsNull(!Mdt), 0, !Mdt) + IIf(IsNull(!Mdr), 0, !Mdr) + IIf(IsNull(!Mdm), 0, !Mdm) + IIf(IsNull(!Mdo), 0, !Mdo)
'                SumaCostos = CostoDesabolladura + CostoPintura + CostoArmeDesarme
'kjcv 04.11.13
                SumaCostos = CostoDesabolladura + CostoPintura + CostoArmeDesarme + CostoTerceros + CostoRepuestos + CostoManoObra
'                SumaVentas = VentaDesabolladura + VentaPintura + VentaArmeDesarme
                
                SumaVentas = VentaDesabolladura + VentaPintura + VentaArmeDesarme + VentaTerceros + VentaRepuestos + VentaManoObra + !Insumos

                TotalOT = !Total_General
              
'                TotalOT = SumaVentas - TotalDescuentos
                
                'inicializa variable para comparar con el tex box de renta <= a
                Dim dblTotalOT As Double
                If TotalOT <> 0 Then
                  dblTotalOT = CDbl(((TotalOT - SumaCostos) * 100) / TotalOT)
                  If dblTotalOT <= IIf(Me.cckCriterios(9).Value = 1, CDbl(txtPorcentaje), 501) Then  'rentabilidad de ot < a porcentaje
                    
                    '// numero ot
                    Set itmItem = lvDetalle.ListItems.Add(, , !Id_OT)
                    itmItem.SubItems(1) = IIf(!Seccion_OT = "M", "MECANICA", "CARROCERIA") & "(" & !Nro_Factura_Emitida & ")"
                    '// desabolladura
                    itmItem.SubItems(2) = FormatoValor(CostoDesabolladura, gstrMonedaLocal, gintDecimalesMoneda)
                    itmItem.SubItems(3) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(VentaDesabolladura, gstrMonedaLocal, gintDecimalesMoneda))
                    itmItem.SubItems(4) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(VentaDesabolladura - CostoDesabolladura, gstrMonedaLocal, gintDecimalesMoneda))
                    If VentaDesabolladura <> 0 Then
                      itmItem.SubItems(5) = IIf(lstrCostea = "S", FormatoValor(0, "%", 2), FormatoValor(((VentaDesabolladura - CostoDesabolladura) * 100) / VentaDesabolladura, "%", 2))
                    Else
                      itmItem.SubItems(5) = FormatoValor(0, "%", 2)
                    End If
                    '// pintura
                    itmItem.SubItems(6) = FormatoValor(CostoPintura, gstrMonedaLocal, gintDecimalesMoneda)
                    itmItem.SubItems(7) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(VentaPintura, gstrMonedaLocal, gintDecimalesMoneda))
                    itmItem.SubItems(8) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(VentaPintura - CostoPintura, gstrMonedaLocal, gintDecimalesMoneda))
                    If VentaPintura <> 0 Then
                      itmItem.SubItems(9) = IIf(lstrCostea = "S", FormatoValor(0, "%", 2), FormatoValor(((VentaPintura - CostoPintura) * 100) / VentaPintura, "%", 2))
                    Else
                      itmItem.SubItems(9) = FormatoValor(0, "%", 2)
                    End If
                    '// arme y desarme
                    itmItem.SubItems(10) = FormatoValor(CostoArmeDesarme, gstrMonedaLocal, gintDecimalesMoneda)
                    itmItem.SubItems(11) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(VentaArmeDesarme, gstrMonedaLocal, gintDecimalesMoneda))
                    itmItem.SubItems(12) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(VentaArmeDesarme - CostoArmeDesarme, gstrMonedaLocal, gintDecimalesMoneda))
                    If VentaArmeDesarme <> 0 Then
                      itmItem.SubItems(13) = IIf(lstrCostea = "S", FormatoValor(0, "%", 2), FormatoValor(((VentaArmeDesarme - CostoArmeDesarme) * 100) / VentaArmeDesarme, "%", 2))
                    Else
                      itmItem.SubItems(13) = FormatoValor(0, "%", 2)
                    End If
                    
'                    '// Totales OT
'                    itmItem.SubItems(14) = FormatoValor(SumaCostos, gstrMonedaLocal, gintDecimalesMoneda)        '// Total Costos
'                    itmItem.SubItems(15) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(SumaVentas, gstrMonedaLocal, gintDecimalesMoneda))       '// Total Ot sin descuentos
'                    itmItem.SubItems(16) = FormatoValor(TotalDescuentos, gstrMonedaLocal, gintDecimalesMoneda)   '// Descuentos
'                    itmItem.SubItems(17) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(TotalOT, gstrMonedaLocal, gintDecimalesMoneda))          '// Total Ot sin descuentos
'
'                    If TotalOT <> 0 Then
'                      itmItem.SubItems(18) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(TotalOT - SumaCostos, gstrMonedaLocal, gintDecimalesMoneda))
'                      itmItem.SubItems(19) = IIf(lstrCostea = "S", FormatoValor(0, "%", 2), FormatoValor(((TotalOT - SumaCostos) * 100) / TotalOT, "%", 2))
'                    Else
'                      itmItem.SubItems(18) = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
'                      itmItem.SubItems(19) = FormatoValor(0, "%", 2)
'                    End If
                'kjcv 04.11.13 se agrgea otros campos
                '// terceros
                itmItem.SubItems(14) = FormatoValor(CostoTerceros, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(15) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!total_terceros, gstrMonedaLocal, gintDecimalesMoneda))
                itmItem.SubItems(16) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!total_terceros - CostoTerceros, gstrMonedaLocal, gintDecimalesMoneda))
                If !total_terceros <> 0 Then
                  itmItem.SubItems(17) = IIf(lstrCostea = "S", FormatoValor(0, "%", 2), FormatoValor(((!total_terceros - CostoTerceros) * 100) / !total_terceros, "%", 2))
                Else
                  itmItem.SubItems(17) = FormatoValor(0, "%", 2)
                End If
                '// repuestos
                itmItem.SubItems(18) = FormatoValor(CostoRepuestos, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(19) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!total_repuestos, gstrMonedaLocal, gintDecimalesMoneda))
                itmItem.SubItems(20) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!total_repuestos - CostoRepuestos, gstrMonedaLocal, gintDecimalesMoneda))
                If !total_repuestos <> 0 Then
                  itmItem.SubItems(21) = IIf(lstrCostea = "S", FormatoValor(0, "%", 2), FormatoValor(((!total_repuestos - CostoRepuestos) * 100) / !total_repuestos, "%", 2))
                Else
                  itmItem.SubItems(21) = FormatoValor(0, "%", 2)
                End If
                '//insumos
                itmItem.SubItems(22) = FormatoValor(!Insumos, gstrMonedaLocal, gintDecimalesMoneda)
                '// mano de obra
                itmItem.SubItems(23) = FormatoValor(CostoManoObra, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(24) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!total_mano_obra, gstrMonedaLocal, gintDecimalesMoneda))
                itmItem.SubItems(25) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(!total_mano_obra - CostoManoObra, gstrMonedaLocal, gintDecimalesMoneda))
                If !total_mano_obra <> 0 Then
                  itmItem.SubItems(26) = IIf(lstrCostea = "S", FormatoValor(0, "%", 2), FormatoValor(((!total_mano_obra - CostoManoObra) * 100) / !total_mano_obra, "%", 2))
                Else
                  itmItem.SubItems(26) = FormatoValor(0, "%", 2)
                End If
                
                
                    
                ' // Totales OT
                itmItem.SubItems(27) = FormatoValor(SumaCostos, gstrMonedaLocal, gintDecimalesMoneda)        '// Total Costos
                itmItem.SubItems(28) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(SumaVentas, gstrMonedaLocal, gintDecimalesMoneda))       '// Total Ot sin descuentos
                itmItem.SubItems(29) = FormatoValor(TotalDescuentos, gstrMonedaLocal, gintDecimalesMoneda)   '// Descuentos
                itmItem.SubItems(30) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(TotalOT, gstrMonedaLocal, gintDecimalesMoneda))          '// Total Ot sin descuentos

                If TotalOT <> 0 Then
                  itmItem.SubItems(31) = IIf(lstrCostea = "S", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(TotalOT - SumaCostos, gstrMonedaLocal, gintDecimalesMoneda))
                  itmItem.SubItems(32) = IIf(lstrCostea = "S", FormatoValor(0, "%", 2), FormatoValor(((TotalOT - SumaCostos) * 100) / TotalOT, "%", 2))
                Else
                  itmItem.SubItems(31) = FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda)
                  itmItem.SubItems(32) = FormatoValor(0, "%", 2)
                End If
                 'kjcv 31.10.13 se agrega cargo ,fecha facturacion y recepcionista de Ot
                itmItem.SubItems(33) = !CARGO
                itmItem.SubItems(34) = !Fecha_Facturacion
                itmItem.SubItems(35) = !Recepcionista
                itmItem.SubItems(36) = !Patente
                    
                  End If
                End If
                TotalDescuentos = 0
                SumaCostos = 0
                SumaVentas = 0
            End If
            .MoveNext
          Wend
       End If
    End With
    End If
    Screen.MousePointer = 1
    lblTotal(7).Caption = lvDetalle.ListItems.Count
    
End Sub

Private Sub cmdExcel_Click()
    If lvDetalle.ListItems.Count > 0 Then
'        ImprimirConsulta "Excel"
        ExportarDatos Me.lvDetalle, Me.cdExportar, Me.hwnd
    Else
        MsgBox "No Existen elemenetos en la lista"
    End If
End Sub

Private Sub cmdImprimir_Click()
If lvDetalle.ListItems.Count > 0 Then
    ImprimirConsulta "Impresora"
Else
    MsgBox "No Existen elemenetos en la lista"
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

    If Not Atributos("Glbl", "Tllr_30_0110", True, True, True, True) Then
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
Dim item As ListItem
SW = True

    If Not Conexion.SendHost("Select Descripcion, Id_Garantia From Tllr_Garantias Where Vigencia='S' and Id_Empresa='" & gstrIdEmpresa & "' ", AdoPaso, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        MsgBox "Error en Conexion con el Host...", vbCritical, "Taller Pro"
        End
    End If

    If Not (AdoPaso.EOF = True And AdoPaso.BOF = True) Then
        Do Until AdoPaso.EOF
                Set item = Me.lsvtipoOt.ListItems.Add(, , ValorNulo(AdoPaso.Fields(0)))
                item.SubItems(1) = ValorNulo(AdoPaso.Fields(1))
            AdoPaso.MoveNext
        Loop
    End If
    AdoPaso.Close
    
    'kjcv 22.04.13 LLenar Cargos
    If Not Conexion.SendHost("Select Descripcion, Id_Tipo_Cargo From Tllr_Tipo_Cargo Where Vigencia='S' and Id_Empresa='" & gstrIdEmpresa & "' ", AdoPaso, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        MsgBox "Error en Conexion con el Host...", vbCritical, "Taller Pro"
        End
    End If

    If Not (AdoPaso.EOF = True And AdoPaso.BOF = True) Then
        Do Until AdoPaso.EOF
            Set item = Me.lsvCargos.ListItems.Add(, , ValorNulo(AdoPaso.Fields(0)))
            item.SubItems(1) = ValorNulo(AdoPaso.Fields(1))
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
ReOrdenaLista lvDetalle, ColumnHeader
End Sub

Private Sub lvDetalle_DblClick()
If cmdSeleccionar.Enabled = True Then cmdSeleccionar.Value = True
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

'Private Sub txtNroRecord_KeyPress(KeyAscii As Integer)
'KeyAscii = CheckNumber(KeyAscii, txtNroRecord, strComa)
'End Sub

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
    mstrSql = "SELECT Id_Mes AS Codigo, Descripcion AS Nombre FROM Tllr_CostoManoObra"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With datCostoManoObra
            Set .Recordset = adoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcCostoManoObra.ListField = "Nombre"
                dtcCostoManoObra.BoundColumn = "Codigo"
                dtcCostoManoObra.BoundText = .Recordset!Codigo
                If .Recordset.RecordCount < 2 Then dtcCostoManoObra.Enabled = False
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
    dtcCostoManoObra.Text = MesManoObra(Month(Date))
End Sub

