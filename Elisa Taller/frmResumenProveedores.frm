VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Begin VB.Form frmResumenProveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Proveedores"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "frmResumenProveedores.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbTotales 
      Height          =   315
      Left            =   900
      TabIndex        =   36
      Top             =   6300
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Suma - Valor"
            TextSave        =   "Suma - Valor"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Suma - Subtotal"
            TextSave        =   "Suma - Subtotal"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Suma - P.Final"
            TextSave        =   "Suma - P.Final"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2646
            MinWidth        =   2646
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
   Begin Crystal.CrystalReport rptOT 
      Left            =   3945
      Top             =   6210
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
      Caption         =   "Imprimir Listado"
      Height          =   360
      Left            =   8025
      TabIndex        =   30
      Top             =   6840
      Width           =   1680
   End
   Begin VB.Frame Frame2 
      Height          =   2730
      Left            =   60
      TabIndex        =   5
      Top             =   -15
      Width           =   11370
      Begin VB.Frame Frame4 
         Caption         =   "Proveedor de"
         Height          =   510
         Left            =   4860
         TabIndex        =   45
         Top             =   1035
         Width           =   3165
         Begin VB.OptionButton optCarroce 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Carroceria"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1575
            TabIndex        =   47
            Top             =   225
            Width           =   1095
         End
         Begin VB.OptionButton optTerceros 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Terceros"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   135
            TabIndex        =   46
            Top             =   225
            Value           =   -1  'True
            Width           =   1005
         End
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Recepcionista"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   4800
         TabIndex        =   44
         Top             =   2115
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Frame Frame3 
         Caption         =   "Estado"
         Height          =   525
         Left            =   135
         TabIndex        =   38
         Top             =   2115
         Width           =   4560
         Begin VB.OptionButton optVigente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Vigente"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   888
            TabIndex        =   43
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optTodas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Todas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   42
            Top             =   240
            Width           =   810
         End
         Begin VB.OptionButton optCerrada 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Facturadas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2739
            TabIndex        =   41
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optNula 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Nula"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3840
            TabIndex        =   40
            Top             =   240
            Width           =   675
         End
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
      End
      Begin MSComctlLib.ListView lsvtipoOt 
         Height          =   1515
         Left            =   8325
         TabIndex        =   37
         Top             =   975
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   2672
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
      Begin VB.Frame Frame1 
         Caption         =   "Sección"
         Height          =   555
         Left            =   8520
         TabIndex        =   32
         Top             =   2520
         Visible         =   0   'False
         Width           =   3765
         Begin VB.OptionButton OptAmbas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Ambas"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2775
            TabIndex        =   35
            Top             =   225
            Value           =   -1  'True
            Width           =   840
         End
         Begin VB.OptionButton optCarroceria 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Carrocería"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   34
            Top             =   225
            Width           =   1215
         End
         Begin VB.OptionButton optMecanica 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Mecánica"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1425
            TabIndex        =   33
            Top             =   225
            Width           =   1065
         End
      End
      Begin VB.CommandButton cmdResumenOT 
         Caption         =   "Ver Resumen Total OT"
         Height          =   360
         Left            =   9405
         TabIndex        =   31
         Top             =   2280
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Fecha Emisión (Final)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   2055
         TabIndex        =   28
         Top             =   1545
         Value           =   1  'Checked
         Width           =   1920
      End
      Begin VB.CheckBox cckCriterios 
         Caption         =   "Fecha Liquidación"
         Height          =   195
         Index           =   8
         Left            =   3990
         TabIndex        =   27
         Top             =   1590
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.TextBox txtRecepcionista 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   23
         Top             =   2355
         Visible         =   0   'False
         Width           =   3555
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Nro OT"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   22
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox txtNroOt 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   105
         MaxLength       =   15
         TabIndex        =   21
         Top             =   525
         Width           =   2670
      End
      Begin VB.TextBox txtPatente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2850
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
         Left            =   2865
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
         Left            =   3930
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
         Left            =   6855
         TabIndex        =   12
         Top             =   300
         Width           =   840
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Proveedor"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   945
         Width           =   1515
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1185
         Width           =   4455
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3930
         MaxLength       =   50
         TabIndex        =   8
         Top             =   525
         Width           =   2835
      End
      Begin VB.TextBox txtModelo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   6855
         MaxLength       =   50
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   1545
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
               Picture         =   "frmResumenProveedores.frx":038A
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":049C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":08F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":0D4C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":11A4
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":12B6
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":13C8
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":14DA
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":15EC
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":16FE
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":1810
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":1922
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":1A34
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":1B46
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":1C58
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":1D6A
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":1E7C
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":1F8E
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":20A0
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":21B2
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":2604
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResumenProveedores.frx":2A56
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbMarca 
         Height          =   330
         Left            =   6270
         TabIndex        =   17
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
         TabIndex        =   18
         Top             =   210
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
         TabIndex        =   19
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
         TabIndex        =   24
         Top             =   2025
         Visible         =   0   'False
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
         TabIndex        =   25
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   178257921
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckFechaHasta 
         Height          =   315
         Left            =   2055
         TabIndex        =   26
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   178257921
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckLiquida 
         Height          =   315
         Left            =   4005
         TabIndex        =   29
         Top             =   1800
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   178257921
         CurrentDate     =   36776
      End
      Begin MSComCtl2.UpDown updNroRecord 
         Height          =   315
         Left            =   10770
         TabIndex        =   16
         Top             =   -375
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   5
         BuddyControl    =   "txtNroRecord"
         BuddyDispid     =   196632
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
         TabIndex        =   10
         Text            =   "10"
         Top             =   -375
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Registros"
         Height          =   195
         Index           =   8
         Left            =   10260
         TabIndex        =   20
         Top             =   -570
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdBuscarOT 
      Appearance      =   0  'Flat
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   360
      Left            =   6225
      TabIndex        =   0
      Top             =   6825
      Width           =   1680
   End
   Begin VB.CommandButton cmdSalir 
      Appearance      =   0  'Flat
      Caption         =   "Salir"
      Height          =   360
      Left            =   9750
      TabIndex        =   1
      Top             =   6855
      Width           =   1680
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   3555
      Left            =   75
      TabIndex        =   4
      Top             =   2715
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   6271
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
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N° OT"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Estado/N°Factura"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sección"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Placa"
         Object.Width           =   1764
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
         Text            =   "Tipo OT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Proveedor"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Nº Factura"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "% Recargo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "(S/.) Recargo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Subtotal"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "% Dscto."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "(S/.) Dscto."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Precio Final"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   7
      Left            =   1935
      TabIndex        =   3
      Top             =   6930
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registros Encontrados :"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   2
      Top             =   6930
      Width           =   1695
   End
End
Attribute VB_Name = "frmResumenProveedores"
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
'    If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    If Dir(gstrPathReporte & "\BDNueva.mdb") <> "" Then Kill gstrPathReporte & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
   
'    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (NroOT text,Estado text,Seccion text,Patente text,Cliente text,Marca text,Modelo text,FechaIngreso date,Tipo text, Proveedor text, Factura TEXT, Valor Double, PorcRecargo Text, MontoRecargo Double, Subtotal Double, PorcDscto Text, MontoDscto Double, PrecioFinal Double)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To lvDetalle.ListItems.Count
        Set lvDetalle.SelectedItem = lvDetalle.ListItems(i)
        Tabla.AddNew
        Tabla!NroOT = IIf(lvDetalle.SelectedItem = "", " ", lvDetalle.SelectedItem)
        Tabla!estado = IIf(lvDetalle.SelectedItem.SubItems(1) = "", " ", lvDetalle.SelectedItem.SubItems(1))
        Tabla!Seccion = IIf(lvDetalle.SelectedItem.SubItems(2) = "", " ", lvDetalle.SelectedItem.SubItems(2))
        Tabla!Patente = IIf(lvDetalle.SelectedItem.SubItems(3) = "", " ", lvDetalle.SelectedItem.SubItems(3))
        Tabla!Marca = IIf(lvDetalle.SelectedItem.SubItems(4) = "", " ", lvDetalle.SelectedItem.SubItems(4))
        Tabla!Modelo = IIf(lvDetalle.SelectedItem.SubItems(5) = "", " ", lvDetalle.SelectedItem.SubItems(5))
        Tabla!FechaIngreso = DateValue(IIf(lvDetalle.SelectedItem.SubItems(6) = "", " ", lvDetalle.SelectedItem.SubItems(6)))
        Tabla!Tipo = IIf(lvDetalle.SelectedItem.SubItems(7) = "", " ", lvDetalle.SelectedItem.SubItems(7))
        Tabla!Proveedor = IIf(lvDetalle.SelectedItem.SubItems(8) = "", " ", lvDetalle.SelectedItem.SubItems(8))
        Tabla!Factura = IIf(lvDetalle.SelectedItem.SubItems(9) = "", " ", lvDetalle.SelectedItem.SubItems(9))
        Tabla!Valor = IIf(lvDetalle.SelectedItem.SubItems(10) = "", 0, SacarFormatoValor(lvDetalle.SelectedItem.SubItems(10), "S/."))
        Tabla!PorcRecargo = IIf(lvDetalle.SelectedItem.SubItems(11) = "", 0, lvDetalle.SelectedItem.SubItems(11))
        Tabla!MontoRecargo = IIf(lvDetalle.SelectedItem.SubItems(12) = "", 0, SacarFormatoValor(lvDetalle.SelectedItem.SubItems(12), "S/."))
        Tabla!SubTotal = IIf(lvDetalle.SelectedItem.SubItems(13) = "", 0, SacarFormatoValor(lvDetalle.SelectedItem.SubItems(13), "S/."))
        Tabla!PorcDscto = IIf(lvDetalle.SelectedItem.SubItems(14) = "", 0, SacarFormatoValor(lvDetalle.SelectedItem.SubItems(14), "%"))
        Tabla!MontoDscto = IIf(lvDetalle.SelectedItem.SubItems(15) = "", 0, SacarFormatoValor(lvDetalle.SelectedItem.SubItems(15), "S/."))
        Tabla!PrecioFinal = IIf(lvDetalle.SelectedItem.SubItems(16) = "", 0, SacarFormatoValor(lvDetalle.SelectedItem.SubItems(16), "S/."))
        Tabla.Update
    Next i
   Tabla.Close
   Dbsnueva.Close
   
   With rptOT
        .ReportFileName = gstrPathReporte & "\ResumenProveedor.rpt"
        .WindowTitle = "Reporte de Proveedores"
        .DataFiles(0) = gstrPathReporte & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='RESUMEN VALORIZADO DE PROVEEDORES'"
        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
        If Me.cckCriterios(6).Value = 1 Then
            .Formulas(5) = "desde='" & Me.pckFechaDesde & "'"
            .Formulas(6) = "hasta='" & Me.pckFechaHasta & "'"
        End If
        .Formulas(7) = "TDecimal=" & gintDecimalesMoneda
        .Formulas(8) = "NombrePatente='" & gstrNombrePatente & "'"
        
        .Destination = crptToWindow
        .Action = True
   End With
   
''   Dbsnueva.Close
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
        pckLiquida.Enabled = False
    Else
        pckLiquida.Enabled = True
        pckLiquida.SetFocus
    End If
End Select
End Sub


Private Sub cckTipoOt_Click(Index As Integer)
End Sub

Private Sub Check1_Click()

End Sub


Private Sub cmdBuscarOT_Click()
Dim i As Integer
Dim mstrSql As String
Dim mstrWhere As String
Dim adoTemp As New ADODB.Recordset
Dim AdoAux As New ADODB.Recordset
Dim itmItem As ListItem
Dim item As ListItem
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
    
    '// Estado
    If optTodas.Value = True Then
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
    
    If Me.optTerceros.Value = True Then
        mstrWhere = mstrWhere & ",'T'"
    Else
        mstrWhere = mstrWhere & ",'C'"
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
    
End With
'/////////////////////////////////////////////////////////////////////////////////
    
    '/// llama al procedimiento almacenado
    mstrSql = "Exec Tllr_Resumen_Proveedor " & mstrWhere
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With adoTemp
       If Not .BOF And Not .EOF Then
          While Not .EOF
              If !Total_General <> !deducible_pesos Then
                Set itmItem = lvDetalle.ListItems.Add(, , !Id_OT)
                If !estado = "B" Or !estado = "F" Then
                  mstrNumeroDocumento = ValorNulo(!NUMDOCUMENTO)
                Else
                  mstrNumeroDocumento = "S/N"
                End If
                itmItem.SubItems(1) = ValorNulo(IIf(!estado = "L", "LIQUIDADA", IIf(!estado = "V", "VIGENTE", IIf(!estado = "N", "NULA", IIf(!estado = "B", "BOLETEADA", "FACTURADA"))))) & "(" & mstrNumeroDocumento & ")"
                itmItem.SubItems(2) = ValorNulo(IIf(!Seccion_OT = "M", "MECANICA", "CARROCERIA"))
                itmItem.SubItems(3) = ValorNulo(!Patente)
                itmItem.SubItems(4) = ValorNulo(!Marca)
                itmItem.SubItems(5) = ValorNulo(!Modelo)
                itmItem.SubItems(6) = ValorNulo(!Fecha_Emision)
                itmItem.SubItems(7) = ValorNulo(TraeTipoOT(!Id_Garantia))
                itmItem.SubItems(8) = ValorNulo(!Proveedor)
                If Me.optTerceros.Value = True Then
                  itmItem.SubItems(9) = ValorNulo(!Factura)
                Else
                  itmItem.SubItems(9) = ""
                End If
                itmItem.SubItems(10) = FormatoValor(ValorNulo(!Valor), gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(11) = IIf(!Id_Tipo_Cargo = "05", FormatoValor(0, "%", 2), FormatoValor(ValorNulo(!Porcentaje_Recargo), "%", 2))
                itmItem.SubItems(12) = IIf(!Id_Tipo_Cargo = "05", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(ValorNulo(!monto_recargo), gstrMonedaLocal, gintDecimalesMoneda))
                itmItem.SubItems(13) = IIf(!Id_Tipo_Cargo = "05", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(ValorNulo(!PrecioFinal), gstrMonedaLocal, gintDecimalesMoneda))
                itmItem.SubItems(14) = IIf(!Id_Tipo_Cargo = "05", FormatoValor(0, "%", 2), FormatoValor(ValorNulo(!PorcDscto), "%", 2))
                itmItem.SubItems(15) = IIf(!Id_Tipo_Cargo = "05", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(ValorNulo(!MontoDscto), gstrMonedaLocal, gintDecimalesMoneda))
                itmItem.SubItems(16) = IIf(!Id_Tipo_Cargo = "05", FormatoValor(0, gstrMonedaLocal, gintDecimalesMoneda), FormatoValor(ValorNulo(!SubTotal), gstrMonedaLocal, gintDecimalesMoneda))
              End If
              adoTemp.MoveNext
          Wend
       End If
    End With
    
    'Ahora crea la linea de Totales
'              Set itmItem = lvDetalle.ListItems.Add(, , "TOTALES :")
'              itmItem.SubItems(18) = TotalSeccion(Me.lvDetalle, 18)
'              itmItem.SubItems(19) = TotalSeccion(Me.lvDetalle, 19)
'              itmItem.SubItems(20) = TotalSeccion(Me.lvDetalle, 20)
    With Me.stbTotales
        .Panels(2).Text = FormatoValor(TotalSeccionFormato(lvDetalle, 10), gstrMonedaLocal, gintDecimalesMoneda)
        .Panels(4).Text = FormatoValor(TotalSeccionFormato(lvDetalle, 13), gstrMonedaLocal, gintDecimalesMoneda)
        .Panels(6).Text = FormatoValor(TotalSeccionFormato(lvDetalle, 16), gstrMonedaLocal, gintDecimalesMoneda)
    End With
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
    .lblestado = lvDetalle.SelectedItem.SubItems(1)
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
    gstrSeccion = lvDetalle.SelectedItem.SubItems(2)
End If
Unload Me
End Sub




Private Sub Form_Activate()

If SW Then

    If Not Atributos("Glbl", "Tllr_30_0100", True, True, True, True) Then
        MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
        Unload Me
        Exit Sub
    End If

    pckFechaDesde = BOM(Date)
    pckFechaHasta = EOM(Date)
    SW = False
End If

End Sub

Private Sub Form_Load()
Dim AdoPaso As New ADODB.Recordset
Dim item As ListItem
SW = True

    If Not Conexion.SendHost("Select Descripcion, Id_Garantia From Tllr_Garantias Where Vigencia='S' and Id_Empresa='" & gstrIdEmpresa & "'", AdoPaso, adOpenKeyset, adLockOptimistic, 10) = apOk Then
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
    
    Me.cckCriterios(1).Caption = gstrNombrePatente
End Sub

Private Sub lvDetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ReOrdenaLista lvDetalle, ColumnHeader
End Sub

Private Sub lvDetalle_DblClick()
'If cmdSeleccionar.Enabled = True Then cmdSeleccionar.Value = True
End Sub

Private Sub tlbCliente_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim mstrnombre As String
If Button.Key = "Buscar" Then
    apfFormulario.BuscarRegistroClientes Conexion, gstrBusca, mstrnombre, gstrIdEmpresa
    'apfFormulario.BuscarRegistroClientes Conexion, gstrBusca, mstrnombre
    txtCliente.Tag = gstrBusca
    txtCliente = mstrnombre
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
        dblPreSuma = dblPreSuma + Val(SacarFormatoValor(.SelectedItem.SubItems(IndiceSubItem), gstrMonedaLocal))
    Next
End With
TotalSeccionFormato = dblPreSuma
End Function

