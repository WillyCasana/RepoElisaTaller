VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Begin VB.Form frmFacturarCargosInternos 
   Caption         =   "Facturar Cargos Internos"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFacturarCargosInternos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   7560
      Width           =   11415
      Begin VB.CommandButton cmdSalir 
         Appearance      =   0  'Flat
         Caption         =   "Salir"
         Height          =   375
         Left            =   9600
         TabIndex        =   33
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdBuscarOT 
         Appearance      =   0  'Flat
         Caption         =   "Buscar"
         Default         =   -1  'True
         Height          =   375
         Left            =   6000
         TabIndex        =   32
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdImprimir 
         Appearance      =   0  'Flat
         Caption         =   "Imprimir Informe"
         Height          =   375
         Left            =   7800
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdAplicar 
         Appearance      =   0  'Flat
         Caption         =   "Facturar Ahora"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   27
      Top             =   10080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   26
      Top             =   10080
      Visible         =   0   'False
      Width           =   495
   End
   Begin Crystal.CrystalReport rptOT 
      Left            =   5520
      Top             =   9480
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
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   11415
      Begin MSComCtl2.DTPicker pckFechaDesde 
         Height          =   325
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   47448065
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckFechaHasta 
         Height          =   325
         Left            =   1920
         TabIndex        =   19
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   47448065
         CurrentDate     =   36776
      End
      Begin MSComctlLib.ListView lsvCargos 
         Height          =   2055
         Left            =   8400
         TabIndex        =   25
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo de Cargo"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.TextBox txtModelo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1320
         Width           =   4575
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   6
         Top             =   600
         Width           =   5055
      End
      Begin VB.TextBox txtNroOt 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         MaxLength       =   15
         TabIndex        =   16
         Top             =   600
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3480
         TabIndex        =   21
         Top             =   2880
         Visible         =   0   'False
         Width           =   3525
         Begin VB.OptionButton optMecanica 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Mecánica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1320
            TabIndex        =   24
            Top             =   240
            Width           =   1065
         End
         Begin VB.OptionButton optCarroceria 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Carrocería"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton OptAmbas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Ambas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2520
            TabIndex        =   22
            Top             =   240
            Value           =   -1  'True
            Width           =   840
         End
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Fec. Final"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   1920
         TabIndex        =   20
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Nro OT"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtPatente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Placa"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Marca "
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Modelo"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Cliente"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Fec. Inicial"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1455
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
               Picture         =   "frmFacturarCargosInternos.frx":038A
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":049C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":08F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":0D4C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":11A4
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":12B6
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":13C8
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":14DA
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":15EC
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":16FE
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":1810
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":1922
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":1A34
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":1B46
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":1C58
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":1D6A
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":1E7C
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":1F8E
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":20A0
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":21B2
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":2604
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacturarCargosInternos.frx":2A56
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbMarca 
         Height          =   330
         Left            =   7920
         TabIndex        =   13
         Top             =   300
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
         Left            =   7920
         TabIndex        =   14
         Top             =   1020
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
      Begin MSComctlLib.Toolbar tlbCliente 
         Height          =   330
         Left            =   4200
         TabIndex        =   15
         Top             =   1020
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
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   3720
         TabIndex        =   34
         Top             =   2040
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         Format          =   178782209
         CurrentDate     =   42163
      End
      Begin VB.Label lblfecha 
         Caption         =   "Fecha Facturación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         TabIndex        =   35
         Top             =   1800
         Width           =   2025
      End
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6800
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N° OT"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Seccion"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cargo"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fecha Liquidación"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Placa"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Marca"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Modelo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Cliente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "IDEmpresa"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "IdSucursal"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "IdCargo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Fecha"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "TotalNeto"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "DNI/RUC"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList imgBotones 
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
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFacturarCargosInternos.frx":2B68
            Key             =   "selec"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFacturarCargosInternos.frx":2FBA
            Key             =   "noselec"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBotones 
      Height          =   660
      Index           =   0
      Left            =   120
      TabIndex        =   28
      Top             =   6720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1164
      ButtonWidth     =   2910
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgBotones"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Selecc. Todos "
            Key             =   "Selecc"
            Object.ToolTipText     =   "Seleccionar Todos los elementos de la lista."
            ImageKey        =   "selec"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Desmarcar Todos"
            Key             =   "NoSelecc"
            Object.ToolTipText     =   "Desmarcar Todos los elementos de la lista."
            ImageKey        =   "noselec"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registros Encontrados :"
      Height          =   255
      Index           =   6
      Left            =   8880
      TabIndex        =   0
      Top             =   6840
      Width           =   2055
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   195
      Index           =   7
      Left            =   10875
      TabIndex        =   1
      Top             =   6840
      Width           =   660
   End
End
Attribute VB_Name = "frmFacturarCargosInternos"
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
    If Dir(gstrPathReporte & "\BDNueva.mdb") <> "" Then Kill gstrPathReporte & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    'Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (NroOT text,Seccion text, Cargo text, fecha date, Patente text, Marca text, Modelo text, Cliente text)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To lvDetalle.ListItems.Count
        Set lvDetalle.SelectedItem = lvDetalle.ListItems(i)
        Tabla.AddNew
        Tabla!NroOT = IIf(lvDetalle.SelectedItem = "", " ", Mid(lvDetalle.SelectedItem, 6, 10))
        Tabla!Seccion = IIf(lvDetalle.SelectedItem.SubItems(1) = "", " ", lvDetalle.SelectedItem.SubItems(1))
        Tabla!CARGO = IIf(lvDetalle.SelectedItem.SubItems(2) = "", " ", lvDetalle.SelectedItem.SubItems(2))
        Tabla!Fecha = IIf(lvDetalle.SelectedItem.SubItems(3) = "", " ", Mid(lvDetalle.SelectedItem.SubItems(3), 1, 10))
        Tabla!Patente = IIf(lvDetalle.SelectedItem.SubItems(4) = "", " ", lvDetalle.SelectedItem.SubItems(4))
        Tabla!Marca = IIf(lvDetalle.SelectedItem.SubItems(5) = "", " ", lvDetalle.SelectedItem.SubItems(5))
        Tabla!Modelo = IIf(lvDetalle.SelectedItem.SubItems(6) = "", " ", lvDetalle.SelectedItem.SubItems(6))
        Tabla!Cliente = IIf(lvDetalle.SelectedItem.SubItems(7) = "", " ", lvDetalle.SelectedItem.SubItems(7))
        Tabla.Update
    Next i
    Tabla.Close
    Dbsnueva.Close
    
    With rptOT
         '//MODIFICADO POR FDO DIAZ EL 29/11/2000
         .ReportFileName = gstrPathReporte & "\CargosaFacturar.rpt"
         .WindowTitle = "Facturar Cargos Internos"
'         .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
         .DataFiles(0) = gstrPathReporte & "\BDNueva.mdb"
         .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
         .Formulas(1) = "TITULO='OT Internas'"
         .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
         .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
         .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
         If Me.cckCriterios(6).Value = 1 Then
            .Formulas(5) = "desde='" & Me.pckFechaDesde & "'"
            .Formulas(6) = "hasta='" & Me.pckFechaHasta & "'"
         End If
         .Formulas(7) = "NombrePatente='" & gstrNombrePatente & "'"
         
         .Destination = crptToWindow
         .Action = True
    End With

    Screen.MousePointer = 1
   
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
        pckFechaDesde.Enabled = False
    Else
        pckFechaDesde.Enabled = True
        pckFechaDesde.SetFocus
    End If
Case 6
    If cckCriterios(Index).Value = 0 Then
        pckFechaHasta.Enabled = False
    Else
        pckFechaHasta.Enabled = True
        pckFechaHasta.SetFocus
    End If
    
End Select
End Sub


Private Sub cmdAplicar_Click()
Dim dblNumeroFacturaInterna As Double
Dim dblNumeroVenta As Double
Dim lstrCobradorCliente As String
Dim lstrVendedorCliente As String
Dim lstrVendedorMeson As String
Dim lstrTipoRescate As String
Dim lstrIdVenta As String
Dim lstrMonedaLocal As String
Dim Total_Neto As Double
Dim Total_Iva As Double
Dim Total_Ot As Double
Dim i As Integer
Dim SwLimpiaLista As String

SwLimpiaLista = "OFF"
If Me.lvDetalle.ListItems.Count > 0 Then
    
    If SituacionLista(Me.lvDetalle).VACIA Then
        MsgBox "No ha seleccionado OTs para facturar"
        Exit Sub
    End If
    
    'confirma y pide caja y cajero
    blnContinuar = False
    frmCajaCajero.Show vbModal
    If blnContinuar = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    If Me.Text1 <> "" And Me.Text2 <> "" Then  'que haya elegido caja y cajero
        For i = 1 To Me.lvDetalle.ListItems.Count
            If Me.lvDetalle.ListItems(i).Checked = True Then  'si esta chequeado
                SwLimpiaLista = "ON"
                'rescata correlativos y variables
                mstrSql = "select max(ultimo_numero) as Numero from vpro_correlativo_Facturas_internas "
                mstrSql = mstrSql & "Where Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Caja='" & Me.Text1 & "' And Id_Cajero='" & Me.Text2 & "'"
                If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                        dblNumeroFacturaInterna = IIf(IsNull(adoPrincipal!NUMERO), 1, adoPrincipal!NUMERO + 1)
                    End If
                End If
                
                mstrSql = "select max(Id_Numero_Venta) as Numero From Vpro_Facturacion "
                mstrSql = mstrSql & "Where Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Caja='" & Me.Text1 & "' And Id_Cajero='" & Me.Text2 & "'"
                If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                        dblNumeroVenta = IIf(IsNull(adoPrincipal!NUMERO), 1, adoPrincipal!NUMERO + 1)
                    End If
                End If
                
                mstrSql = "Select Id_Cobrador_Cliente,Id_Vendedor_Cliente From Glbl_Cliente_Proveedor "
                mstrSql = mstrSql & "Where Id_Cliente_Proveedor='" & Me.lvDetalle.ListItems(i).SubItems(13) & "'"
                If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                        lstrCobradorCliente = ValorNulo(adoPrincipal!Id_Cobrador_Cliente)
                        lstrVendedorCliente = ValorNulo(adoPrincipal!Id_Vendedor_Cliente)
                    End If
                End If
                
                'vendedor meson
                mstrSql = "Select Top 1 Id_Vendedor_Meson From Vpro_Vendedor_Meson Where Vigencia='S'"
                If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                        lstrVendedorMeson = ValorNulo(adoPrincipal!Id_Vendedor_Meson)
                    End If
                End If
                
                'tipo de rescate
                mstrSql = "Select Cod_Taller, Cod_dyp from Vpro_Parametros_Globales where Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
                If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                        lstrTipoRescate = IIf(Me.lvDetalle.ListItems(i).SubItems(1) = "M", adoPrincipal!Cod_Taller, adoPrincipal!Cod_dyp)
                    End If
                End If
                
                'tipo de venta
                mstrSql = "Select top 1 id_tipo_Venta from Vpro_Tipo_Venta where Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Vigencia='S'"
                If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                        lstrIdVenta = ValorNulo(adoPrincipal!id_tipo_Venta)
                    End If
                End If
                
                'moneda local
                mstrSql = "Select Id_Moneda_Local from Tllr_Parametro where Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
                If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                        lstrMonedaLocal = ValorNulo(adoPrincipal!Id_Moneda_Local)
                    End If
                End If
                
                Total_Neto = CDbl(Me.lvDetalle.ListItems(i).SubItems(12))
                Total_Iva = Total_Neto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaCeroPto)
                Total_Ot = Total_Neto * IVA(gstrIdEmpresa, gstrIdSucursal, gcIvaUnoPto)
                
                'inserta registro en vpro_facturacion
                mstrSql = "Insert into Vpro_Facturacion (Id_Tipo_Cargo,Id_Numero_Venta,Id_Empresa,Id_Cobrador_Cliente,Id_Sucursal,Id_Caja,Id_Vendedor_Meson,Id_Tipo_Rescate,Id_Cajero, "
                mstrSql = mstrSql & "Id_Vendedor_Cliente,Id_cliente_proveedor,Fecha_Venta,Numero_Documento,Fecha_Facturacion,Fecha_Vencto_Factura,Descto_al_total,descto_al_total_Porcent, "
                mstrSql = mstrSql & "Neto,Iva,Total,Comentario,Tipo_Docto,Centralizado,Estado_Facturacion,Estado_rescate,Numero_Rescate,Seccion_Ot,id_tipo_venta,t_comision_l_factura, "
                mstrSql = mstrSql & "Tipo_Fact_Autopro,RebajaStock,Por_Lote,Id_Moneda,Paridad) Values ("
                mstrSql = mstrSql & "'" & Me.lvDetalle.ListItems(i).SubItems(10) & "',"  'tipo Cargo
                mstrSql = mstrSql & dblNumeroVenta & ","
                mstrSql = mstrSql & "'" & gstrIdEmpresa & "',"
                mstrSql = mstrSql & "'" & lstrCobradorCliente & "',"
                mstrSql = mstrSql & "'" & gstrIdSucursal & "',"
                mstrSql = mstrSql & "'" & Me.Text1 & "',"
                mstrSql = mstrSql & "'" & lstrVendedorMeson & "',"
                mstrSql = mstrSql & "'" & lstrTipoRescate & "',"
                mstrSql = mstrSql & "'" & Me.Text2 & "',"
                mstrSql = mstrSql & "'" & lstrVendedorCliente & "',"
                mstrSql = mstrSql & "'" & Me.lvDetalle.ListItems(i).SubItems(13) & "',"
                mstrSql = mstrSql & "'" & Date & "',"
                mstrSql = mstrSql & dblNumeroFacturaInterna & ","
                'kjcv 08.06.15
                mstrSql = mstrSql & "'" & Format(dtpFecha.Value, "DD/MM/YYYY") & "',"
                'mstrSQL = mstrSQL & "'" & Date & "',"
                mstrSql = mstrSql & "'" & Date & "',"
                mstrSql = mstrSql & "0,0," & Total_Neto & "," & Total_Iva & "," & Total_Ot & ","
                mstrSql = mstrSql & "'" & "Facturado Desde Taller" & "',"
                mstrSql = mstrSql & "'" & "FI" & "',"
                mstrSql = mstrSql & "'" & "N" & "',"
                mstrSql = mstrSql & "'" & "E" & "',"
                mstrSql = mstrSql & "'" & "F" & "',"
                mstrSql = mstrSql & "'" & Me.lvDetalle.ListItems(i) & "',"
                mstrSql = mstrSql & "'" & Me.lvDetalle.ListItems(i).SubItems(1) & "',"
                mstrSql = mstrSql & "'" & lstrIdVenta & "',"
                mstrSql = mstrSql & "0" & ","
                mstrSql = mstrSql & "'" & "N" & "',"
                mstrSql = mstrSql & "'" & "S" & "',"
                mstrSql = mstrSql & "'" & "N" & "',"
                mstrSql = mstrSql & "'" & lstrMonedaLocal & "',"
                mstrSql = mstrSql & "1" & ")"
                Conexion.SendHost mstrSql, , , , gcTiempoEspera
                
                'inserta registro del último número factura interna
                mstrSql = "Insert Into Vpro_Correlativo_Facturas_Internas (Id_Empresa,Id_Sucursal,Id_Caja,Id_Cajero,Ultimo_Numero) Values ("
                mstrSql = mstrSql & "'" & gstrIdEmpresa & "',"
                mstrSql = mstrSql & "'" & gstrIdSucursal & "',"
                mstrSql = mstrSql & "'" & Me.Text1 & "',"
                mstrSql = mstrSql & "'" & Me.Text2 & "',"
                mstrSql = mstrSql & dblNumeroFacturaInterna & ")"
                Conexion.SendHost mstrSql, , , , gcTiempoEspera
                
                'actualiza tablas de taller
                mstrSql = "Update Tllr_Facturacion Set Estado='F',"
                mstrSql = mstrSql & "Nro_Factura_Emitida='" & dblNumeroFacturaInterna & "',"
                'kjcv 08.06.15
                mstrSql = mstrSql & "Fecha_Facturacion='" & Format(dtpFecha.Value, "DD/MM/YYYY") & "',"
'                mstrSQL = mstrSQL & "Fecha_Facturacion='" & Format(Now, "DD/MM/YYYY") & "',"
                mstrSql = mstrSql & "Total_Neto=" & Total_Neto & ","
                mstrSql = mstrSql & "Iva=" & Total_Iva & ","
                mstrSql = mstrSql & "Valor_Afecto=" & Total_Neto & ","
                mstrSql = mstrSql & "Total=" & Total_Ot
                mstrSql = mstrSql & " Where Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Ot='" & Me.lvDetalle.ListItems(i) & "' And Seccion_Ot='" & Me.lvDetalle.ListItems(i).SubItems(1) & "' And Id_Cargo='" & Me.lvDetalle.ListItems(i).SubItems(10) & "'"
                Conexion.SendHost mstrSql, , , , gcTiempoEspera
                
                mstrSql = "Update Tllr_Ot Set Estado='F',"
                mstrSql = mstrSql & "Nro_Factura_Emitida='" & dblNumeroFacturaInterna & "'"
                mstrSql = mstrSql & " Where Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Ot='" & Me.lvDetalle.ListItems(i) & "' And Seccion_Ot='" & Me.lvDetalle.ListItems(i).SubItems(1) & "'"
                Conexion.SendHost mstrSql, , , , gcTiempoEspera
                
                
                'actualiza movimientos de la ot
                mstrSql = "Update Tllr_Mecanica_Ot Set Facturado='S'"
                mstrSql = mstrSql & " Where Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Ot='" & Me.lvDetalle.ListItems(i) & "' And Seccion_Ot='" & Me.lvDetalle.ListItems(i).SubItems(1) & "' And Id_Tipo_Cargo='" & Me.lvDetalle.ListItems(i).SubItems(10) & "'"
                Conexion.SendHost mstrSql, , , , gcTiempoEspera
                
                mstrSql = "Update Tllr_Carroceria_Ot Set Facturado='S'"
                mstrSql = mstrSql & " Where Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Ot='" & Me.lvDetalle.ListItems(i) & "' And Seccion_Ot='" & Me.lvDetalle.ListItems(i).SubItems(1) & "' And Id_Tipo_Cargo='" & Me.lvDetalle.ListItems(i).SubItems(10) & "'"
                Conexion.SendHost mstrSql, , , , gcTiempoEspera
                
                mstrSql = "Update Tllr_Otro_Ot Set Facturado='S'"
                mstrSql = mstrSql & " Where Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Ot='" & Me.lvDetalle.ListItems(i) & "' And Seccion_Ot='" & Me.lvDetalle.ListItems(i).SubItems(1) & "' And Id_Tipo_Cargo='" & Me.lvDetalle.ListItems(i).SubItems(10) & "'"
                Conexion.SendHost mstrSql, , , , gcTiempoEspera
                
                mstrSql = "Update Tllr_Terceros_Ot Set Facturado='S'"
                mstrSql = mstrSql & " Where Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Ot='" & Me.lvDetalle.ListItems(i) & "' And Seccion_Ot='" & Me.lvDetalle.ListItems(i).SubItems(1) & "' And Id_Tipo_Cargo='" & Me.lvDetalle.ListItems(i).SubItems(10) & "'"
                Conexion.SendHost mstrSql, , , , gcTiempoEspera
                
                mstrSql = "Update Tllr_Repuestos_Ot Set Facturado='S'"
                mstrSql = mstrSql & " Where Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Ot='" & Me.lvDetalle.ListItems(i) & "' And Seccion_Ot='" & Me.lvDetalle.ListItems(i).SubItems(1) & "' And Id_Tipo_Cargo='" & Me.lvDetalle.ListItems(i).SubItems(10) & "'"
                Conexion.SendHost mstrSql, , , , gcTiempoEspera
                
            End If
        Next
    End If
    'limpia lista segun el sw
    If SwLimpiaLista = "ON" Then
        Me.lvDetalle.ListItems.Clear
    End If
    Screen.MousePointer = vbDefault
End If
End Sub

Private Sub cmdBuscarOT_Click()
Dim mstrSql As String
Dim mstrWhere As String
Dim adoTemp As New ADODB.Recordset
Dim AdoAux As New ADODB.Recordset
Dim itmItem As ListItem
Dim lstrSQL As String
Dim AdoTot As New ADODB.Recordset
Dim i As Integer

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
    
    If .cckCriterios(5).Value = 1 And .cckCriterios(6).Value = 1 Then   '////////// fecha iniciosi y terminosi
        mstrWhere = mstrWhere & ",'" & pckFechaDesde.Value & "','" & pckFechaHasta.Value & " 23:59:00" & "'"
    ElseIf .cckCriterios(5).Value = 0 And .cckCriterios(6).Value = 0 Then  '////////// fecha iniciono y terminono
        mstrWhere = mstrWhere & ",'',''"
    ElseIf .cckCriterios(5).Value = 1 And .cckCriterios(6).Value = 0 Then   '////////// fecha iniciosi y terminono
        mstrWhere = mstrWhere & ",'" & pckFechaDesde.Value & "',''"
    ElseIf .cckCriterios(5).Value = 0 And .cckCriterios(6).Value = 1 Then  '////////// fecha iniciono y terminosi
        mstrWhere = mstrWhere & ",'','" & pckFechaHasta.Value & "'"
    End If
    
    If .optCarroceria.Value = True Then ' POR CARROCERIA
        mstrWhere = mstrWhere & ",'C'"
    ElseIf .optMecanica.Value = True Then ' POR MECANICA
        mstrWhere = mstrWhere & ",'M'"
    Else
        mstrWhere = mstrWhere & ",''"
    End If
    
    Dim lsw As Double
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
    
End With

    '/// llama al procedimiento almacenado
    mstrSql = "Exec Tllr_Facturar_Cargos_Internos " & mstrWhere
    
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With adoTemp
       If Not .BOF And Not .EOF Then
          While Not .EOF
            Set itmItem = lvDetalle.ListItems.Add(, , !Id_OT)
            itmItem.SubItems(1) = ValorNulo(!Seccion_OT)
            itmItem.SubItems(2) = ValorNulo(!Descripcion)
            itmItem.SubItems(3) = ValorNulo(!Fecha_Liquidacion)
            itmItem.SubItems(4) = ValorNulo(!Patente)
            itmItem.SubItems(5) = ValorNulo(!Marca)
            itmItem.SubItems(6) = ValorNulo(!Modelo)
            itmItem.SubItems(7) = ValorNulo(!Cliente)
            itmItem.SubItems(8) = ValorNulo(!Id_Empresa)
            itmItem.SubItems(9) = ValorNulo(!Id_Sucursal)
            itmItem.SubItems(10) = ValorNulo(!Id_Cargo)
            itmItem.SubItems(11) = ValorNulo(Format(!Fecha_Liquidacion, "YYYY/MM/DD"))
            itmItem.SubItems(12) = ValorNulo(!Total_Neto)
            itmItem.SubItems(13) = ValorNulo(!rut)
            adoTemp.MoveNext
          Wend
       End If
    End With
    End If
    Screen.MousePointer = 1
    lblTotal(7).Caption = lvDetalle.ListItems.Count
    
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

Private Sub Form_Activate()

If SW Then

    If Not Atributos("Glbl", "Tllr_20_0150", True, True, True, True) Then
        MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
        Unload Me
        Exit Sub
    End If

    pckFechaDesde = BOM(Date)
    pckFechaHasta = EOM(Date)
    Me.dtpFecha = Date
    SW = False
End If

End Sub

Private Sub Form_Load()
Dim AdoPaso As New ADODB.Recordset
Dim item As ListItem
    
SW = True

'cargos
mstrSql = "Select Descripcion, Id_Tipo_Cargo From Tllr_Tipo_Cargo Where Id_Empresa='" & gstrIdEmpresa & "' and Vigencia='S' And InternaSN='S'"
If Conexion.SendHost(mstrSql, AdoPaso, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With AdoPaso
        If Not .BOF And Not .EOF Then
            While Not .EOF
                Set item = Me.lsvCargos.ListItems.Add(, , ValorNulo(AdoPaso.Fields(0)))
                item.SubItems(1) = ValorNulo(AdoPaso.Fields(1))
                AdoPaso.MoveNext
            Wend
        End If
    End With
    AdoPaso.Close
End If

'    'cargos
'    If Not Conexion.SendHost("Select Descripcion, Id_Tipo_Cargo From Tllr_Tipo_Cargo Where Vigencia='S' And InternaSN='S'", AdoPaso, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
'        MsgBox "Error en Conexion con el Host...", vbCritical, "Taller Pro"
'        End
'    End If
'
'    If Not (AdoPaso.EOF = True And AdoPaso.BOF = True) Then
'        Do Until AdoPaso.EOF
'            Set Item = Me.lsvCargos.ListItems.Add(, , ValorNulo(AdoPaso.Fields(0)))
'            Item.SubItems(1) = ValorNulo(AdoPaso.Fields(1))
'            AdoPaso.MoveNext
'        Loop
'    End If
'    AdoPaso.Close
    
    Me.cckCriterios(1).Caption = gstrNombrePatente
End Sub

Private Sub lvDetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Select Case ColumnHeader.Index
Case 4
    ReOrdenaListaNumero lvDetalle, 11
Case Else
    ReOrdenaLista lvDetalle, ColumnHeader
End Select
End Sub

Private Sub lvDetalle_ItemCheck(ByVal item As MSComctlLib.ListItem)

ActivaDesactivaBotonesListas Me.lvDetalle, Me, 0

End Sub

Private Sub tlbBotones_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

Screen.MousePointer = vbHourglass

Select Case Index
    Case 0 'Lista Ots Internas
        Select Case Button.Key
            Case "Selecc"
                SeleccionarTodo Me.lvDetalle
                Me.tlbBotones(Index).Buttons(1).Enabled = False
                Me.tlbBotones(Index).Buttons(2).Enabled = True
            Case "NoSelecc"
                DesmarcarTodo Me.lvDetalle
                Me.tlbBotones(Index).Buttons(1).Enabled = True
                Me.tlbBotones(Index).Buttons(2).Enabled = False
        End Select
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
''KeyAscii = CheckIdCar(txtPatente.SelStart, mdLLNNNN, UpCaseLetter(KeyAscii))
'KeyAscii = UpCaseLetter(KeyAscii)
'kjcv 24-01-12 Valida Letras y numeros
If (KeyAscii <> 8) And Not (KeyAscii >= 48 And KeyAscii <= 57) And Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
    KeyAscii = 0: Beep
Else
    KeyAscii = UpCaseLetter(KeyAscii)
End If

End Sub



