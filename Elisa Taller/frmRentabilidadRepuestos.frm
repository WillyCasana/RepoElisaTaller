VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Begin VB.Form frmRentabilidadRepuestos 
   Caption         =   "Rentabilidad Repuestos"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   Icon            =   "frmRentabilidadRepuestos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExcel 
      Appearance      =   0  'Flat
      Caption         =   "Excel"
      Enabled         =   0   'False
      Height          =   360
      Left            =   7200
      TabIndex        =   31
      Top             =   6240
      Width           =   840
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
      Caption         =   "Imprimir Informe"
      Height          =   360
      Left            =   8040
      TabIndex        =   30
      Top             =   6240
      Width           =   1680
   End
   Begin VB.Frame Frame2 
      Height          =   2145
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   11370
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Renta <= a"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   9840
         TabIndex        =   38
         Top             =   300
         Width           =   1185
      End
      Begin VB.TextBox txtPorcentaje 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9840
         TabIndex        =   37
         Text            =   "0"
         Top             =   525
         Width           =   1185
      End
      Begin VB.Frame Frame1 
         Height          =   510
         Left            =   11160
         TabIndex        =   33
         Top             =   1440
         Visible         =   0   'False
         Width           =   3765
         Begin VB.OptionButton OptAmbas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Ambas"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2775
            TabIndex        =   36
            Top             =   180
            Value           =   -1  'True
            Width           =   840
         End
         Begin VB.OptionButton optCarroceria 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Carrocería"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   45
            TabIndex        =   35
            Top             =   180
            Width           =   1215
         End
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
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Fecha Emisión (Final)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   2055
         TabIndex        =   29
         Top             =   1545
         Value           =   1  'Checked
         Width           =   1920
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Recepcionista"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   4620
         TabIndex        =   25
         Top             =   945
         Width           =   1395
      End
      Begin VB.TextBox txtRecepcionista 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4620
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
         Width           =   2670
      End
      Begin VB.TextBox txtPatente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2850
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
         Left            =   2865
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
         Left            =   3930
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
         Left            =   6855
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
         Text            =   "10"
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
         Width           =   4455
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3930
         MaxLength       =   50
         TabIndex        =   9
         Top             =   525
         Width           =   2835
      End
      Begin VB.TextBox txtModelo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   6855
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
               Picture         =   "frmRentabilidadRepuestos.frx":179A
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":18AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":1D04
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":215C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":25B4
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":26C6
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":27D8
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":28EA
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":29FC
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":2B0E
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":2C20
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":2D32
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":2E44
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":2F56
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":3068
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":317A
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":328C
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":339E
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":34B0
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":35C2
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":3A14
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRentabilidadRepuestos.frx":3E66
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.UpDown updNroRecord 
         Height          =   315
         Left            =   10920
         TabIndex        =   17
         Top             =   525
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   5
         BuddyControl    =   "txtNroRecord"
         BuddyDispid     =   196621
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
         Format          =   177078273
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckFechaHasta 
         Height          =   315
         Left            =   2055
         TabIndex        =   28
         Top             =   1755
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   177078273
         CurrentDate     =   36776
      End
      Begin MSComctlLib.ListView lsvtipoOt 
         Height          =   1110
         Left            =   8400
         TabIndex        =   32
         Top             =   960
         Width           =   2730
         _ExtentX        =   4815
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
      Left            =   5400
      TabIndex        =   0
      Top             =   6240
      Width           =   1680
   End
   Begin VB.CommandButton cmdSalir 
      Appearance      =   0  'Flat
      Caption         =   "Salir"
      Height          =   360
      Left            =   9720
      TabIndex        =   2
      Top             =   6240
      Width           =   1680
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   3930
      Left            =   75
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
      Appearance      =   0
      NumItems        =   21
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N° OT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CARGO/N°Factura"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Rep. Costo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Rep. Venta"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rep. Diferencia"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Rep. Margen"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Lub. Costo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Lub. Venta"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Lub. Diferencia"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Lub. Margen"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Mat. Costo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Mat. Venta"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Mat. Diferencia"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Mat. Margen"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "Ins. Costo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "Ins. Venta"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Ins. Diferencia"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Text            =   "Ins. Margen"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Text            =   "Dsctos(%)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Text            =   "Monto Dcto."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "IdCargo"
         Object.Width           =   0
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
Attribute VB_Name = "frmRentabilidadRepuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SW As Boolean

Sub ImprimirConsulta(strSalida As String)
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
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (NroOT text,Cargo text, COSTOREP DOUBLE, TOTALREP DOUBLE, DIFREP DOUBLE, MARGENREP DOUBLE,COSTOLUB DOUBLE, TOTALLUB DOUBLE, DIFLUB DOUBLE, MARGENLUB DOUBLE,COSTOMAT DOUBLE, TOTALMAT DOUBLE, DIFMAT DOUBLE, MARGENMAT DOUBLE, COSTOINS DOUBLE, TOTALINS DOUBLE, DIFINS DOUBLE, MARGENINS DOUBLE, DESCUENTOS DOUBLE, DSCTOSPORC DOUBLE, IDCARGO TEXT)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To lvDetalle.ListItems.Count
        Set lvDetalle.SelectedItem = lvDetalle.ListItems(i)
        Tabla.AddNew
        Tabla!NroOT = IIf(lvDetalle.SelectedItem = "", " ", Mid(lvDetalle.SelectedItem, 6, 10))
        Tabla!CARGO = IIf(lvDetalle.SelectedItem.SubItems(1) = "", " ", lvDetalle.SelectedItem.SubItems(1))
        Tabla!COSTOREP = CDbl(IIf(lvDetalle.SelectedItem.SubItems(2) = "", " ", lvDetalle.SelectedItem.SubItems(2)))
        Tabla!TOTALREP = CDbl(IIf(lvDetalle.SelectedItem.SubItems(3) = "", " ", lvDetalle.SelectedItem.SubItems(3)))
        Tabla!DIFREP = CDbl(IIf(lvDetalle.SelectedItem.SubItems(4) = "", " ", lvDetalle.SelectedItem.SubItems(4)))
        Tabla!MARGENREP = CDbl(IIf(lvDetalle.SelectedItem.SubItems(5) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(5), "%")))
        Tabla!COSTOLUB = CDbl(IIf(lvDetalle.SelectedItem.SubItems(6) = "", " ", lvDetalle.SelectedItem.SubItems(6)))
        Tabla!TOTALLUB = CDbl(IIf(lvDetalle.SelectedItem.SubItems(7) = "", " ", lvDetalle.SelectedItem.SubItems(7)))
        Tabla!DIFLUB = CDbl(IIf(lvDetalle.SelectedItem.SubItems(8) = "", " ", lvDetalle.SelectedItem.SubItems(8)))
        Tabla!MARGENLUB = CDbl(IIf(lvDetalle.SelectedItem.SubItems(9) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(9), "%")))
        Tabla!COSTOMAT = CDbl(IIf(lvDetalle.SelectedItem.SubItems(10) = "", " ", lvDetalle.SelectedItem.SubItems(10)))
        Tabla!TOTALMAT = CDbl(IIf(lvDetalle.SelectedItem.SubItems(11) = "", " ", lvDetalle.SelectedItem.SubItems(11)))
        Tabla!DIFMAT = CDbl(IIf(lvDetalle.SelectedItem.SubItems(12) = "", " ", lvDetalle.SelectedItem.SubItems(12)))
        Tabla!MARGENMAT = CDbl(IIf(lvDetalle.SelectedItem.SubItems(13) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(13), "%")))
        Tabla!COSTOINS = CDbl(IIf(lvDetalle.SelectedItem.SubItems(14) = "", " ", lvDetalle.SelectedItem.SubItems(14)))
        Tabla!TOTALINS = CDbl(IIf(lvDetalle.SelectedItem.SubItems(15) = "", " ", lvDetalle.SelectedItem.SubItems(15)))
        Tabla!DIFINS = CDbl(IIf(lvDetalle.SelectedItem.SubItems(16) = "", " ", lvDetalle.SelectedItem.SubItems(16)))
        Tabla!MARGENINS = CDbl(IIf(lvDetalle.SelectedItem.SubItems(17) = "", " ", SacarFormatoValor(lvDetalle.SelectedItem.SubItems(17), "%")))
        Tabla!DSCTOSPORC = CDbl(SacarFormatoValor(IIf(lvDetalle.SelectedItem.SubItems(18) = "", " ", lvDetalle.SelectedItem.SubItems(18)), "%"))
        Tabla!Descuentos = CDbl(SacarFormatoValor(IIf(lvDetalle.SelectedItem.SubItems(19) = "", " ", lvDetalle.SelectedItem.SubItems(19)), gstrMonedaLocal))
        Tabla!IDCARGO = IIf(lvDetalle.SelectedItem.SubItems(20) = "", " ", lvDetalle.SelectedItem.SubItems(20))
        Tabla.Update
    Next i
    Tabla.Close
   
    If strSalida = "Excel" Then
        With rptOT
            .Destination = crptToFile
            .PrintFileType = crptExcel50Tab
             .ReportFileName = gstrPathReporte & "\RENOTEX.rpt"
             .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
             .Destination = crptToFile
             .Action = True
        End With
    Else
        With rptOT
             .ReportFileName = gstrPathReporte & "\RENREPUESTOS.rpt"
             .WindowTitle = "Rentabilidad De Repuestos por OT"
             .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
             .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
             .Formulas(1) = "TITULO='RENTABILIDAD DE REPUESTOS POR OT'"
             .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
             .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
             .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
             .Formulas(5) = "Desde='" & Me.pckFechaDesde & "'"
             .Formulas(6) = "Hasta='" & Me.pckFechaHasta & "'"
             .Formulas(7) = "TDecimal=" & gintDecimalesMoneda

             .Destination = crptToWindow
             .Action = True
        End With
    End If
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
Dim CostoRepuestos As Double
Dim CostoLubricantes As Double
Dim CostoMateriales As Double
Dim CostoInsumos As Double
Dim VentaRepuestos As Double
Dim VentaLubricantes As Double
Dim VentaMateriales As Double
Dim VentaInsumos As Double
Dim ValorHora As Double
Dim i As Integer
Dim TotalDescuentos As Double
Dim TotalOT As Double
Dim ValoresRetornados As VentaRepuestos
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
     
     'traspasa parametros de lubricantes, materiales e insumos
     mstrWhere = mstrWhere & ",'" & gstrCodigoLubricantes & "','" & gstrCodigoMateriales & "','" & gstrCodigoInsumos & "'"
End With

'/////////////////////////////////////////////////////////////////////////////////
    
    '/// llama al procedimiento almacenado
    mstrSql = "Exec Tllr_Rentabilidad_Repuestos " & mstrWhere
    Screen.MousePointer = 11
    If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With adoTemp
       If Not .BOF And Not .EOF Then
          While Not .EOF
              'inicializa variables
              CostoRepuestos = 0
              CostoLubricantes = 0
              CostoMateriales = 0
              CostoInsumos = 0
              VentaRepuestos = 0
              VentaLubricantes = 0
              VentaMateriales = 0
              VentaInsumos = 0
              TotalDescuentos = 0
    
              'Repuestos
              CostoRepuestos = IIf(IsNull(!SumaConsumoR), 0, !SumaConsumoR) - IIf(IsNull(!SumaDevolucionR), 0, !SumaDevolucionR)
              VentaRepuestos = IIf(IsNull(!VentaR), 0, !VentaR)
              
              'Lubricantes
              CostoLubricantes = IIf(IsNull(!SumaConsumoL), 0, !SumaConsumoL) - IIf(IsNull(!SumaDevolucionL), 0, !SumaDevolucionL)
              VentaLubricantes = IIf(IsNull(!VentaL), 0, !VentaL)
              
              'Materiales
              CostoMateriales = IIf(IsNull(!SumaConsumoM), 0, !SumaConsumoM) - IIf(IsNull(!SumaDevolucionM), 0, !SumaDevolucionM)
              VentaMateriales = IIf(IsNull(!VentaM), 0, !VentaM)
              
              'Insumos
              CostoInsumos = IIf(IsNull(!SumaConsumoI), 0, !SumaConsumoI) - IIf(IsNull(!SumaDevolucionI), 0, !SumaDevolucionI)
              VentaInsumos = IIf(IsNull(!VentaI), 0, !VentaI)
              
              TotalDescuentos = IIf(IsNull(!Descuentos), 0, !Descuentos)
              
              TotalOT = VentaRepuestos + VentaLubricantes + VentaMateriales
              
              Dim dblTotalOT As Double
              If TotalOT <> 0 Then
                dblTotalOT = CDbl(((TotalOT - (CostoRepuestos + CostoLubricantes + CostoMateriales)) * 100) / TotalOT)
              Else
                dblTotalOT = 0
              End If
              
              If dblTotalOT <= IIf(Me.cckCriterios(9).Value = 1, CDbl(txtPorcentaje), 501) Then  'rentabilidad de ot < a porcentaje
                Set itmItem = lvDetalle.ListItems.Add(, , !Id_OT)
                
                mstrNumeroDocumento = IIf(Not IsNull(!Nro_Factura_Emitida), !Nro_Factura_Emitida, "S/N")
                
                itmItem.SubItems(1) = ValorNulo(!Descripcion) & "(" & mstrNumeroDocumento & ")"
                
                itmItem.SubItems(2) = FormatoValor(CostoRepuestos, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(3) = FormatoValor(VentaRepuestos, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(4) = FormatoValor(VentaRepuestos - CostoRepuestos, gstrMonedaLocal, gintDecimalesMoneda)
                If VentaRepuestos <> 0 Then
                  itmItem.SubItems(5) = FormatoValor(((VentaRepuestos - CostoRepuestos) * 100) / VentaRepuestos, "%", 2)
                Else
                  itmItem.SubItems(5) = FormatoValor(0, "%", 2)
                End If
                
                itmItem.SubItems(6) = FormatoValor(CostoLubricantes, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(7) = FormatoValor(VentaLubricantes, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(8) = FormatoValor(VentaLubricantes - CostoLubricantes, gstrMonedaLocal, gintDecimalesMoneda)
                If VentaLubricantes <> 0 Then
                  itmItem.SubItems(9) = FormatoValor(((VentaLubricantes - CostoLubricantes) * 100) / VentaLubricantes, "%", 2)
                Else
                  itmItem.SubItems(9) = FormatoValor(0, "%", 2)
                End If
                
                itmItem.SubItems(10) = FormatoValor(CostoMateriales, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(11) = FormatoValor(VentaMateriales, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(12) = FormatoValor(VentaMateriales - CostoMateriales, gstrMonedaLocal, gintDecimalesMoneda)
                If VentaMateriales <> 0 Then
                  itmItem.SubItems(13) = FormatoValor(((VentaMateriales - CostoMateriales) * 100) / VentaMateriales, "%", 2)
                Else
                  itmItem.SubItems(13) = FormatoValor(0, "%", 2)
                End If
                
                itmItem.SubItems(14) = FormatoValor(CostoInsumos, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(15) = FormatoValor(VentaInsumos, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(16) = FormatoValor(VentaInsumos - CostoInsumos, gstrMonedaLocal, gintDecimalesMoneda)
                If VentaInsumos <> 0 Then
                  itmItem.SubItems(17) = FormatoValor(((VentaInsumos - CostoInsumos) * 100) / VentaInsumos, "%", 2)
                Else
                  itmItem.SubItems(17) = FormatoValor(0, "%", 2)
                End If
                
                'descuentos
                If SacarFormatoValor(itmItem.SubItems(3), gstrMonedaLocal) <> "0" Then
                  itmItem.SubItems(18) = FormatoValor(TotalDescuentos * 100 / (Val(SacarFormatoValor(itmItem.SubItems(3), gstrMonedaLocal)) + Val(SacarFormatoValor(itmItem.SubItems(7), gstrMonedaLocal)) + Val(SacarFormatoValor(itmItem.SubItems(11), gstrMonedaLocal)) + Val(SacarFormatoValor(itmItem.SubItems(15), gstrMonedaLocal)) + TotalDescuentos), "%", 2)
                Else
                  itmItem.SubItems(18) = FormatoValor(0, "%", 2)
                End If
                itmItem.SubItems(19) = FormatoValor(TotalDescuentos, gstrMonedaLocal, gintDecimalesMoneda)
                itmItem.SubItems(20) = !Id_Cargo
                
              End If
              adoTemp.MoveNext
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
    If Not Atributos("Glbl", "Tllr_30_0040", True, True, True, True) Then
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

Private Sub txtNroRecord_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtNroRecord, strComa)
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
