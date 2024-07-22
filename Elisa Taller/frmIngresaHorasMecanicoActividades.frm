VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmIngresaHorasMecanicoActividades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Horas de Mecanicos por Actividades"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11535
   Icon            =   "frmIngresaHorasMecanicoActividades.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   11535
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGrabarOtrosServ 
      Height          =   315
      Left            =   11040
      Picture         =   "frmIngresaHorasMecanicoActividades.frx":179A
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Guardar Cambios"
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton cmdGrabarActividades 
      Height          =   315
      Left            =   11040
      Picture         =   "frmIngresaHorasMecanicoActividades.frx":1CCC
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Guardar Cambios"
      Top             =   5560
      Width           =   375
   End
   Begin Crystal.CrystalReport rptPatente 
      Left            =   3960
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Height          =   1515
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   11475
      Begin VB.Frame Frame1 
         Caption         =   "Estado"
         Height          =   525
         Left            =   7440
         TabIndex        =   19
         Top             =   920
         Width           =   3960
         Begin VB.OptionButton optLiquidada 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Liquidada"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1746
            TabIndex        =   23
            Top             =   240
            Width           =   990
         End
         Begin VB.OptionButton optCerrada 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Facturadas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2739
            TabIndex        =   22
            Top             =   240
            Width           =   1110
         End
         Begin VB.OptionButton optTodas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Todas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "F. Emisión (Fin)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   1560
         TabIndex        =   18
         Top             =   240
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.TextBox txtPatente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3120
         MaxLength       =   6
         TabIndex        =   12
         Top             =   555
         Width           =   1020
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Placa"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3120
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Marca "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   4320
         TabIndex        =   10
         Top             =   240
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
         TabIndex        =   9
         Top             =   240
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
         TabIndex        =   8
         Top             =   940
         Width           =   795
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1150
         Width           =   3615
      End
      Begin VB.TextBox txtMarca 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   6
         Top             =   555
         Width           =   2955
      End
      Begin VB.TextBox txtModelo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   5
         Top             =   555
         Width           =   4155
      End
      Begin VB.CheckBox cckCriterios 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "F. Emisión (Ini)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   1  'Checked
         Width           =   1320
      End
      Begin MSComctlLib.Toolbar tlbMarca 
         Height          =   330
         Left            =   6840
         TabIndex        =   13
         Top             =   240
         Width           =   450
         _ExtentX        =   794
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
         Left            =   10800
         TabIndex        =   14
         Top             =   240
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
         Left            =   3360
         TabIndex        =   15
         Top             =   880
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
         TabIndex        =   16
         Top             =   555
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   83427329
         CurrentDate     =   36776
      End
      Begin MSComCtl2.DTPicker pckFechaHasta 
         Height          =   315
         Left            =   1560
         TabIndex        =   17
         Top             =   555
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   83427329
         CurrentDate     =   36776
      End
      Begin MSDataListLib.DataCombo dtcSupervisor 
         Bindings        =   "frmIngresaHorasMecanicoActividades.frx":21FE
         Height          =   315
         Left            =   3840
         TabIndex        =   24
         Top             =   1150
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
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
         Left            =   5880
         Top             =   1200
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
         Left            =   3840
         TabIndex        =   25
         Top             =   960
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   1725
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   3043
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
      Appearance      =   0
      NumItems        =   11
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
         Text            =   "Placa"
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
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "idmarca"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "idmodelo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "HorasReales"
         Object.Width           =   882
      EndProperty
   End
   Begin MSComctlLib.ListView lvDetalleActiv 
      Height          =   1650
      Left            =   120
      TabIndex        =   26
      Top             =   3900
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   2910
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
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Horas"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Mecanico"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cod. Mecánico"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Horas Reales"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbTotales 
      Height          =   315
      Left            =   8400
      TabIndex        =   28
      Top             =   3240
      Visible         =   0   'False
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
   Begin MSComctlLib.ListView lvDetalleOtroS 
      Height          =   1410
      Left            =   120
      TabIndex        =   29
      Top             =   5880
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   2487
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
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Horas"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Mecanico"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cod. Mecánico"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Horas Reales"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar"
            ImageKey        =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (Ctrl+C)"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   8280
      Top             =   7080
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
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":221A
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":232C
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":243E
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":2550
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":2662
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":2774
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":2886
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":2998
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":2AAA
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":2BBC
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":2CCE
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":2DE0
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":2EF2
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":3004
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":3116
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":3228
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":333A
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":378C
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":3BDE
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":3CF0
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":3E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":3FA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":4104
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":4260
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":4D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":5180
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":52E4
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":5740
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":589C
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":6BA8
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":7144
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":72A0
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":73FC
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":7750
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":7AA4
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":7DF8
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":814C
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":84A0
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":89E4
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":8AF6
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":8E4A
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":919E
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":92B2
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":9606
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":9718
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngresaHorasMecanicoActividades.frx":9A6C
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Otros Servicios"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Temparios"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   3680
      Width           =   1215
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   7
      Left            =   1920
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   7320
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmIngresaHorasMecanicoActividades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SW As Boolean
Dim itmAux As ListItem
Dim mstrSql As String
Dim AdoTemp As New ADODB.Recordset
Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean
Dim lstrCodigoServicio As String
Dim lstrCodigoMecanico As String
Dim lstrNombreMecanico As String

Sub ImprimirConsulta()
Dim Dbsnueva As Database
Dim Tabla As DAO.Recordset
Dim i As Integer
Dim GcamBaseTem As String
Dim HorasRealesActividades As Double
Dim HorasRealesOtro As Double

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
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (NroOT text,Estado text,FechaIngreso Text,Patente Text,Marca Text,Modelo Text,HorasTempario Double, HorasReales Double)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To lvDetalle.ListItems.Count
        Set lvDetalle.SelectedItem = lvDetalle.ListItems(i)
        Tabla.AddNew
        Tabla!NroOT = IIf(lvDetalle.SelectedItem = "", " ", lvDetalle.SelectedItem)
        Tabla!estado = IIf(lvDetalle.SelectedItem.SubItems(1) = "", " ", lvDetalle.SelectedItem.SubItems(1))
        Tabla!FechaIngreso = IIf(lvDetalle.SelectedItem.SubItems(2) = "", "", lvDetalle.SelectedItem.SubItems(2))
        Tabla!Patente = IIf(lvDetalle.SelectedItem.SubItems(3) = "", " ", lvDetalle.SelectedItem.SubItems(3))
        Tabla!Marca = IIf(lvDetalle.SelectedItem.SubItems(4) = "", "", lvDetalle.SelectedItem.SubItems(4))
        Tabla!Modelo = IIf(lvDetalle.SelectedItem.SubItems(5) = "", "", lvDetalle.SelectedItem.SubItems(5))
        HorasRealesActividades = Retorna_Valor_General("Select isnull(Sum(isnull(HorasReales,0)),0) as Horas from Tllr_Actividades_Mecanico where id_ot='" & Me.lvDetalle.ListItems(i) & "' And Seccion_Ot='" & Me.lvDetalle.ListItems(i).SubItems(7) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'", gcdynamic)
        HorasRealesOtro = Retorna_Valor_General("Select isnull(Sum(isnull(HorasReales,0)),0) as Horas from Tllr_Otro_Ot where id_ot='" & Me.lvDetalle.ListItems(i) & "' And Seccion_Ot='" & Me.lvDetalle.ListItems(i).SubItems(7) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'", gcdynamic)
        Tabla!HorasTempario = IIf(lvDetalle.SelectedItem.SubItems(6) = "", "", lvDetalle.SelectedItem.SubItems(6))
        Tabla!HorasReales = HorasRealesActividades + HorasRealesOtro
        
        Tabla.Update
    Next i
    Tabla.Close
   
   With rptPatente
        .ReportFileName = gstrPathReporte & "\ActividadesOt.Rpt"
        .WindowTitle = "Horas Por OT"
        .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='Horas por OT'"
        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
        .Formulas(5) = "desde='" & pckFechaDesde & "'"
        .Formulas(6) = "hasta='" & pckFechaHasta & "'"
        .Formulas(7) = "NombrePlaca='" & gstrNombrePatente & "'"
        .WindowState = crptMaximized
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
Private Sub Buscar()
Dim mstrSql As String
Dim lstrSql As String
Dim mstrWhere As String
Dim AdoTemp As New ADODB.Recordset
Dim AdoAux As New ADODB.Recordset
Dim itmItem As ListItem
Dim mstrEstado As String
Dim ContLinea As Integer
Dim mdblSumaHoras As Double
Dim mstrNumeroDocumento As String

lvDetalle.ListItems.Clear
lvDetalleActiv.ListItems.Clear
lvDetalleOtroS.ListItems.Clear
mstrWhere = ""
With Me
    
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
    
    If .cckCriterios(4).Value = 1 Then  '////////// cliente
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Glbl_Cliente_Proveedor.Razon_Social LIKE '" & MatchMode(.txtCliente, "Comienzo del Campo", apSqlServer) & "'"
        Else
            mstrWhere = " Where Glbl_Cliente_Proveedor.Razon_Social LIKE '" & MatchMode(.txtCliente, "Comienzo del Campo", apSqlServer) & "'"
        End If
    End If
    
    If .dtcSupervisor.Text <> "" Then
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " and Tllr_Otro_Ot.Mecanico_Asignado='" & .dtcSupervisor.BoundText & "'"
        Else
            mstrWhere = " Where Tllr_Otro_Ot.Mecanico_Asignado='" & .dtcSupervisor.BoundText & "'"
        End If
    End If
    
    If .cckCriterios(6).Value = 1 Then  '////////// fecha inicio
        If .cckCriterios(7).Value = 1 Then  '////////// fecha termino
            If mstrWhere <> "" Then
                mstrWhere = mstrWhere & " AND fecha_emision between '" & pckFechaDesde.Value & "' and '" & pckFechaHasta.Value & " 23:59:00" & "'"
            Else
                mstrWhere = " WHERE fecha_emision between '" & pckFechaDesde.Value & "' and '" & pckFechaHasta.Value & " 23:59:00" & "'"
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
                mstrWhere = " AND fecha_emision = '" & pckFechaHasta.Value & "' "
            Else
                mstrWhere = " WHERE fecha_emision = '" & pckFechaHasta.Value & "' "
            End If
        End If
    End If
    
     '////////// empresa y sucursal
        If mstrWhere <> "" Then
            mstrWhere = mstrWhere & " AND Tllr_Ot.ID_EMPRESA= '" & gstrIdEmpresa & "' AND Tllr_Ot.ID_SUCURSAL='" & gstrIdSucursal & "' "
        Else
            mstrWhere = " WHERE Tllr_Ot.ID_EMPRESA= '" & gstrIdEmpresa & "' AND Tllr_Ot.ID_SUCURSAL='" & gstrIdSucursal & "' "
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
            End If
        If mstrEstado <> "" Then
            mstrWhere = mstrWhere & " And Tllr_OT.Estado  " & mstrEstado
        End If
End With
'/////////////////////////////////////////////////////////////////////////////////
    
    mstrSql = "SELECT SUM(Tllr_Otro_OT.Horas) AS SUMAOTRO, "
    mstrSql = mstrSql & "(SELECT SUM(Tllr_Mecanica_OT.Horas) AS Horas "
    mstrSql = mstrSql & "FROM Tllr_Mecanica_OT INNER JOIN "
    mstrSql = mstrSql & "Tllr_OT ON Tllr_Mecanica_OT.Id_Empresa = Tllr_OT.Id_Empresa AND Tllr_Mecanica_OT.Id_Sucursal = Tllr_OT.Id_Sucursal AND "
    mstrSql = mstrSql & "Tllr_Mecanica_OT.Id_OT = Tllr_OT.Id_OT And Tllr_Mecanica_OT.Seccion_OT = Tllr_OT.Seccion_OT "
    mstrSql = mstrSql & "WHERE (Tllr_Mecanica_OT.Id_Empresa = Tllr_OT.Id_Empresa) AND (Tllr_Mecanica_OT.Id_Sucursal = Tllr_Ot.ID_Sucursal) AND (Tllr_Mecanica_OT.Seccion_OT = Tllr_Ot.Seccion_Ot) AND "
    mstrSql = mstrSql & "(Tllr_Mecanica_OT.Id_OT = Tllr_Otro_Ot.Id_Ot) "
    mstrSql = mstrSql & "GROUP BY Tllr_Mecanica_OT.Id_OT, Tllr_Mecanica_OT.Seccion_OT, Tllr_Mecanica_OT.Id_Sucursal, Tllr_Mecanica_OT.Id_Empresa, "
    mstrSql = mstrSql & "Tllr_Mecanica_OT.Id_OT) AS SUMAMECANICA, "
    mstrSql = mstrSql & "Tllr_Otro_OT.Id_OT , Tllr_OT.Estado, Tllr_OT.fecha_emision, Tllr_OT.Patente, Tllr_OT.Seccion_OT,Tllr_OT.Nro_Factura_Emitida, Glbl_Marca.Id_Marca, Glbl_Marca.Descripcion "
    mstrSql = mstrSql & "AS Marca, Glbl_Modelo.Descripcion AS Modelo, Glbl_Modelo.Id_Modelo "
    mstrSql = mstrSql & "FROM Glbl_Cliente_Proveedor INNER JOIN Tllr_Vehiculo_Cliente "
    mstrSql = mstrSql & "ON Glbl_Cliente_Proveedor.Id_Cliente_Proveedor = Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor Inner Join Tllr_Otro_OT "
    mstrSql = mstrSql & "INNER JOIN Tllr_OT ON Tllr_Otro_OT.Id_Empresa = Tllr_OT.Id_Empresa AND Tllr_Otro_OT.Id_Sucursal = Tllr_OT.Id_Sucursal "
    mstrSql = mstrSql & "AND Tllr_Otro_OT.Id_OT = Tllr_OT.Id_OT AND Tllr_Otro_OT.Seccion_OT = Tllr_OT.Seccion_OT "
    mstrSql = mstrSql & "ON Tllr_Vehiculo_Cliente.Patente = Tllr_OT.Patente INNER JOIN Glbl_Modelo "
    mstrSql = mstrSql & "INNER JOIN Glbl_Marca ON Glbl_Modelo.Id_Marca = Glbl_Marca.Id_Marca "
    mstrSql = mstrSql & "ON Tllr_Vehiculo_Cliente.Id_Modelo = Glbl_Modelo.Id_Modelo AND Tllr_Vehiculo_Cliente.Id_Marca = Glbl_Modelo.Id_Marca "
    mstrSql = mstrSql & mstrWhere & " "
    mstrSql = mstrSql & "GROUP BY Tllr_Otro_OT.Id_OT, Tllr_OT.Estado, Tllr_OT.Fecha_Emision, Tllr_OT.Patente, Tllr_OT.Seccion_OT, Glbl_Marca.Descripcion , Glbl_Modelo.Descripcion,Tllr_Ot.Nro_Factura_Emitida, Glbl_Marca.Id_Marca, Glbl_Modelo.Id_Modelo"

    Screen.MousePointer = 11
    mdblSumaHoras = 0
    If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
    With AdoTemp
       If Not .BOF And Not .EOF Then
          While Not .EOF
              Set itmItem = lvDetalle.ListItems.Add(, , !Id_OT)
              If !estado = "F" Or !estado = "B" Then
                 mstrNumeroDocumento = ValorNulo(!Nro_Factura_Emitida)
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
              itmItem.SubItems(8) = ValorNulo(!Id_Marca)
              itmItem.SubItems(9) = ValorNulo(!Id_Modelo)
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
Private Sub ImprimirReporte()
If lvDetalle.ListItems.Count > 0 Then
    ImprimirConsulta
End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdGrabarActividades_Click()
Dim i As Integer

Screen.MousePointer = vbHourglass

'elimina primero los movimientos
mstrSql = "Delete from Tllr_actividades_Mecanico WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_OT='" & Me.lvDetalle.SelectedItem & "' AND Seccion_OT ='" & Me.lvDetalle.SelectedItem.SubItems(7) & "'"
If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
    MsgBox "Problemas para Actualizar Actividades de Mecanicos", vbExclamation, "ElisaTaller"
End If

'agrega uno por uno las actividades de la lista
For i = 1 To Me.lvDetalleActiv.ListItems.Count
    If lvDetalleActiv.ListItems(i).ForeColor = 192 Then
        lstrCodigoServicio = Me.lvDetalleActiv.ListItems(i)
        i = i + 1
    End If
    mstrSql = "Insert Into Tllr_Actividades_Mecanico (Id_Empresa,Id_Sucursal,Id_Ot,Seccion_Ot,"
    mstrSql = mstrSql & "Id_Servicio,Id_Actividad,Descripcion,Id_Mecanico,HorasActividad,Valor,Subtotal,HorasReales,FechaEmision) "
    mstrSql = mstrSql & "Values ('"
    mstrSql = mstrSql & gstrIdEmpresa & "','" & gstrIdSucursal & "','" & Me.lvDetalle.SelectedItem & "','"
    mstrSql = mstrSql & Me.lvDetalle.SelectedItem.SubItems(7) & "','" & Me.lvDetalle.SelectedItem.SubItems(8) & "-" & lstrCodigoServicio & "','" & Me.lvDetalleActiv.ListItems(i) & "','"
    mstrSql = mstrSql & Me.lvDetalleActiv.ListItems(i).SubItems(1) & "','" & Me.lvDetalleActiv.ListItems(i).SubItems(5) & "'," & Me.lvDetalleActiv.ListItems(i).SubItems(2) & ","
    mstrSql = mstrSql & Me.lvDetalleActiv.ListItems(i).SubItems(3) & "," & Me.lvDetalleActiv.ListItems(i).SubItems(3) & ","
    mstrSql = mstrSql & Me.lvDetalleActiv.ListItems(i).SubItems(6) & ",'" & Me.lvDetalle.SelectedItem.SubItems(2) & "')"

    Conexion.SendHost mstrSql, , , , gcTiempoEspera
    
Next
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdGrabarOtrosServ_Click()
Dim i As Integer

Screen.MousePointer = vbHourglass

For i = 1 To Me.lvDetalleOtroS.ListItems.Count

    mstrSql = "Update Tllr_Otro_Ot set Mecanico_Asignado='" & Me.lvDetalleOtroS.ListItems(i).SubItems(5) & "',"
    mstrSql = mstrSql & "HorasReales=" & Me.lvDetalleOtroS.ListItems(i).SubItems(6)
    mstrSql = mstrSql & "WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_OT='" & Me.lvDetalle.SelectedItem & "' AND Seccion_OT ='" & Me.lvDetalle.SelectedItem.SubItems(7) & "' And Id_Otro_Servicio='" & Me.lvDetalleOtroS.ListItems(i) & "'"
    
    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apAbort Then
        MsgBox "Problemas para Grabar en Otros Servicios", vbExclamation, "ElisaTaller"
    End If
Next i
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()

If Not Atributos("Glbl", "Tllr_20_0090", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
    MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
    Unload Me
    Exit Sub
End If '/////////ojo

If SW Then
    pckFechaDesde = BOM(Date)
    pckFechaHasta = EOM(Date)
    FillMecanicos dtcSupervisor, datSupervisor
    SW = False
End If

End Sub

Private Sub Form_Load()
SW = True
Me.cckCriterios(1).Caption = gstrNombrePatente
End Sub

Private Sub lvDetalle_Click()
If Me.lvDetalle.ListItems.Count > 0 Then
    TraeServiciosMecanica Me.lvDetalle.SelectedItem.SubItems(7), Me.lvDetalle.SelectedItem
    TraeOtrosServicios Me.lvDetalle.SelectedItem, Me.lvDetalle.SelectedItem.SubItems(7)
End If
End Sub

Private Sub lvDetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'ReOrdenaLista lvDetalle, ColumnHeader
End Sub

Private Sub lvDetalleActiv_DblClick()
If Me.lvDetalleActiv.ListItems.Count > 0 Then
    If lvDetalleActiv.SelectedItem.ForeColor <> 192 Then
        If Me.lvDetalleActiv.ListItems.Count > 0 Then
            frmHorasActividadesMecanicoOT.Tag = "Mecanica"
            frmHorasActividadesMecanicoOT.Show vbModal
        End If
    End If
End If
End Sub

Private Sub lvDetalleActiv_KeyPress(KeyAscii As Integer)
If Me.lvDetalleActiv.ListItems.Count > 0 Then
    If KeyAscii = 13 Then
        frmHorasActividadesMecanicoOT.Tag = "Mecanica"
        frmHorasActividadesMecanicoOT.Show vbModal
    End If
End If
End Sub

Private Sub lvDetalleOtroS_DblClick()
If Me.lvDetalleOtroS.ListItems.Count > 0 Then
    frmHorasActividadesMecanicoOT.Tag = "Otros"
    frmHorasActividadesMecanicoOT.Show vbModal
End If
End Sub

Private Sub lvDetalleOtroS_KeyPress(KeyAscii As Integer)
If Me.lvDetalleOtroS.ListItems.Count > 0 Then
    If KeyAscii = 13 Then
        frmHorasActividadesMecanicoOT.Tag = "Otros"
        frmHorasActividadesMecanicoOT.Show vbModal
    End If
End If
End Sub

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Screen.MousePointer = vbHourglass
    Select Case Button.Key
        Case "Buscar"
            Buscar
        Case "Imprimir"
            ImprimirReporte
        Case "Cerrar"
            Unload Me
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
'KeyAscii = CheckIdCar(txtPatente.SelStart, mdLLNNNN, UpCaseLetter(KeyAscii))
'KeyAscii = UpCaseLetter(KeyAscii)
'kjcv 24-01-12 Valida Letras y numeros
If (KeyAscii <> 8) And Not (KeyAscii >= 48 And KeyAscii <= 57) And Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
    KeyAscii = 0: Beep
Else
    KeyAscii = UpCaseLetter(KeyAscii)
End If

End Sub


Sub TraeServiciosMecanica(strSeccion As String, strIdDocumento As String)

    Me.lvDetalleActiv.ListItems.Clear
    
    Screen.MousePointer = vbHourglass
    
    mstrSql = "Exec Tllr_CargaServicios_Mecanica " & "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "','" & strSeccion & "','" & strIdDocumento & "'"
    
    If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoTemp
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            Set itmAux = Me.lvDetalleActiv.ListItems.Add(, , ValorNulo(!ID))
            Set Me.lvDetalleActiv.SelectedItem = itmAux
            itmAux.SubItems(1) = ValorNulo(!Descripcion)
            itmAux.SubItems(2) = FormatoValor(!Horas, "", 1)
            itmAux.SubItems(3) = FormatoValor(!Total, "", gintDecimalesMoneda)
            itmAux.SubItems(4) = ValorNulo(!mec)
            itmAux.SubItems(5) = ValorNulo(!idmec)
            itmAux.SubItems(6) = ""
           
            lstrCodigoMecanico = !idmec
            lstrNombreMecanico = !mec
           
            'rojo es el servicio
            lvDetalleActiv.ListItems(Me.lvDetalleActiv.ListItems.Count).ForeColor = &HC0&
            Me.lvDetalleActiv.ListItems(Me.lvDetalleActiv.ListItems.Count).ListSubItems(1).ForeColor = &HC0&
            Me.lvDetalleActiv.ListItems(Me.lvDetalleActiv.ListItems.Count).ListSubItems(2).ForeColor = &HC0&
            Me.lvDetalleActiv.ListItems(Me.lvDetalleActiv.ListItems.Count).ListSubItems(3).ForeColor = &HC0&
            Me.lvDetalleActiv.ListItems(Me.lvDetalleActiv.ListItems.Count).ListSubItems(4).ForeColor = &HC0&
            Me.lvDetalleActiv.ListItems(Me.lvDetalleActiv.ListItems.Count).ListSubItems(5).ForeColor = &HC0&
            
            'busca si ya ingresaron actividades (edita los registros de Tllr_Actividades_mecanico)
            mstrSql = "Select * from Tllr_Actividades_Mecanico WHERE Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_OT='" & Me.lvDetalle.SelectedItem & "' AND Seccion_OT ='" & Me.lvDetalle.SelectedItem.SubItems(7) & "'"
            If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
                With AdoTemp
                    If Not .BOF And Not .EOF Then
                        .MoveFirst
                        While Not .EOF
                            Set itmAux = Me.lvDetalleActiv.ListItems.Add(, , ValorNulo(!Id_Actividad))
                            Set Me.lvDetalleActiv.SelectedItem = itmAux
                            itmAux.SubItems(1) = ValorNulo(!Descripcion)
                            itmAux.SubItems(2) = FormatoValor(ValorNulo(!HorasActividad), "", 1)
                            itmAux.SubItems(3) = FormatoValor(ValorNulo(!Valor) * gcurPrecioManoObra, "", gintDecimalesMoneda)
                            itmAux.SubItems(4) = TraeNombreMecanico(ValorNulo(!Id_Mecanico))
                            itmAux.SubItems(5) = ValorNulo(!Id_Mecanico)
                            itmAux.SubItems(6) = FormatoValor(ValorNulo(!HorasReales), "", 1)
                           .MoveNext
                        Wend
                    Else
                        'trae las actividades del servicio (primera vez que ingresa a la ot)
                        Actividades_del_Servicio Me.lvDetalle.SelectedItem.SubItems(8), Me.lvDetalle.SelectedItem.SubItems(9), Me.lvDetalleActiv.SelectedItem
                    End If
                End With
            End If
            .MoveNext
        Wend
    End If
    End With
End If
Conexion.CloseHost AdoTemp
Screen.MousePointer = vbDefault
End Sub


Sub Actividades_del_Servicio(strMarca As String, strModelo As String, strServicio As String)
Dim AdoAux As New ADODB.Recordset
Dim lstrSql As String
    
    lstrSql = " SELECT Tllr_Actividad_Servicio_Modelo.Id_Actividad AS CODIGO,"
    lstrSql = lstrSql & " Tllr_Actividad.Descripcion AS NOMBRE,"
    lstrSql = lstrSql & " Tllr_Actividad_Servicio_Modelo.Horas AS TIEMPO,"
    lstrSql = lstrSql & " Tllr_Actividad_Servicio_Modelo.Valor AS VALOR,"
    lstrSql = lstrSql & " Tllr_Actividad.Id_Especialidad AS IDESPE,"
    lstrSql = lstrSql & " Tllr_Especialidad.Descripcion AS ESPECIAL"
    lstrSql = lstrSql & " FROM Tllr_Actividad LEFT OUTER JOIN Tllr_Especialidad ON"
    lstrSql = lstrSql & " Tllr_Actividad.Id_Especialidad = Tllr_Especialidad.Id_Especialidad"
    lstrSql = lstrSql & " RIGHT OUTER JOIN Tllr_Actividad_Servicio_Modelo ON"
    lstrSql = lstrSql & " Tllr_Actividad.Id_Actividad = Tllr_Actividad_Servicio_Modelo.Id_Actividad"
    lstrSql = lstrSql & " WHERE Tllr_Actividad_Servicio_Modelo.Id_Marca = '" & strMarca & "' AND"
    lstrSql = lstrSql & " Tllr_Actividad_Servicio_Modelo.Id_Modelo = '" & strModelo & "' AND"
    lstrSql = lstrSql & " Tllr_Actividad_Servicio_Modelo.Id_Servicio = '" & strServicio & "' "

    If Conexion.SendHost(lstrSql, AdoAux, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
        With AdoAux
            If Not .BOF And Not .EOF Then
                .MoveFirst
                While Not .EOF
                    Set itmAux = Me.lvDetalleActiv.ListItems.Add(, , ValorNulo(!Codigo))
                    Set Me.lvDetalleActiv.SelectedItem = itmAux
                    itmAux.SubItems(1) = ValorNulo(!Nombre)
                    itmAux.SubItems(2) = FormatoValor(ValorNulo(!TIEMPO), "", 1)
                    itmAux.SubItems(3) = FormatoValor(!TIEMPO * gcurPrecioManoObra, "", gintDecimalesMoneda)
                    itmAux.SubItems(4) = lstrNombreMecanico
                    itmAux.SubItems(5) = lstrCodigoMecanico
                    itmAux.SubItems(6) = "0.0"
                    .MoveNext
                Wend
            End If
        End With
    End If
    
    Conexion.CloseHost AdoAux
End Sub

Sub TraeOtrosServicios(strIdDocumento As String, strSeccion As String)

lvDetalleOtroS.ListItems.Clear

mstrSql = "Exec Tllr_CargaServicios_Otro " & "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "','" & strSeccion & "','" & strIdDocumento & "'"

If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoTemp
    If Not .BOF And Not .EOF Then
        .MoveFirst
        While Not .EOF
            Set itmAux = Me.lvDetalleOtroS.ListItems.Add(, , !ID)
            itmAux.SubItems(1) = !Des
            itmAux.SubItems(2) = FormatoValor(!TIEMPO, "", 1)                                                 '///d_p
            itmAux.SubItems(3) = FormatoValor(!SubTotal, "", gintDecimalesMoneda)
            itmAux.SubItems(4) = MecanicoD(!idmec)
            itmAux.SubItems(5) = !idmec
            itmAux.SubItems(6) = IIf(IsNull(!HorasReales), "0.0", FormatoValor(ValorNulo(!HorasReales), "", 1))
            
            .MoveNext
        Wend
    End If
    End With
End If
Conexion.CloseHost AdoTemp

End Sub

