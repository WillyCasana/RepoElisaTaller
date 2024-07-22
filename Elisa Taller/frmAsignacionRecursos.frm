VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Begin VB.Form frmAsignacionRecursos 
   Caption         =   "Asignación de Recursos"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10815
   Icon            =   "frmAsignacionRecursos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   10815
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6600
      TabIndex        =   27
      Text            =   "Text6"
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6000
      TabIndex        =   26
      Text            =   "Text5"
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5520
      TabIndex        =   25
      Text            =   "Text4"
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5040
      TabIndex        =   24
      Text            =   "Text3"
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4560
      TabIndex        =   23
      Text            =   "Text2"
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   3960
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ImageList imgPunteros 
      Left            =   2520
      Top             =   8160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":179A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":1BEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":2042
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":2496
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":28EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":2D3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frTareas 
      Caption         =   "Tareas Pendientes"
      Height          =   2535
      Left            =   0
      TabIndex        =   19
      Top             =   5280
      Width           =   10575
      Begin MSComctlLib.ListView lvTareas 
         Height          =   2175
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   3836
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   16
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Id_Sucursal"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Sucursal"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "OT"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Seccion_Ot"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Seccion"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Cono"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Placa"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Marca"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Modelo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Cliente"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Id_Servicio"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Servicio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Text            =   "Horas"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Tabla"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgrDetalle 
      Height          =   3495
      Left            =   0
      TabIndex        =   15
      Top             =   1680
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   6165
      _Version        =   393216
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame frmFiltro 
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   10575
      Begin VB.TextBox txtHorizonte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5640
         TabIndex        =   18
         Text            =   "30"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdLimpiar 
         Height          =   315
         Index           =   1
         Left            =   3720
         Picture         =   "frmAsignacionRecursos.frx":3192
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   600
         Width           =   315
      End
      Begin VB.CommandButton cmdLimpiar 
         Height          =   315
         Index           =   0
         Left            =   3720
         Picture         =   "frmAsignacionRecursos.frx":3294
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   315
      End
      Begin MSComCtl2.UpDown udHorizonte 
         Height          =   315
         Left            =   6135
         TabIndex        =   14
         Top             =   600
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtHorizonte"
         BuddyDispid     =   196617
         OrigLeft        =   5760
         OrigTop         =   720
         OrigRight       =   6000
         OrigBottom      =   1035
         Max             =   15
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Frame Frame2 
         Caption         =   "Horas"
         Height          =   855
         Left            =   7440
         TabIndex        =   8
         Top             =   120
         Width           =   2655
         Begin VB.OptionButton optHoras 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Disponibles"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   1320
            TabIndex        =   12
            Top             =   480
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optHoras 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Ausencia"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optHoras 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Asignadas"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optHoras 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            Caption         =   "Compradas"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
      End
      Begin MSAdodcLib.Adodc datTurno 
         Height          =   330
         Left            =   1320
         Top             =   600
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
      Begin MSAdodcLib.Adodc datSucursal 
         Height          =   330
         Left            =   1440
         Top             =   240
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
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   5640
         TabIndex        =   7
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   178978817
         CurrentDate     =   37382
      End
      Begin MSDataListLib.DataCombo dbcSucursal 
         Bindings        =   "frmAsignacionRecursos.frx":3396
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "nombre"
         BoundColumn     =   "codigo"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo dbcTurno 
         Bindings        =   "frmAsignacionRecursos.frx":33B0
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "nombre"
         BoundColumn     =   "codigo"
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Horizonte:"
         Height          =   195
         Left            =   4440
         TabIndex        =   13
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial:"
         Height          =   195
         Left            =   4440
         TabIndex        =   6
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Turno:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   660
      End
   End
   Begin MSComctlLib.Toolbar BarraHerramientas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo Analisis"
            ImageKey        =   "Crear"
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar"
            ImageKey        =   "Buscar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Color"
            Object.ToolTipText     =   "Ajuste a los Colores de Fondo"
            ImageKey        =   "Dibujo"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (Ctrl+C)"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList dfs 
      Left            =   240
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":33C7
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":34D9
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":35EB
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":36FD
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":380F
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":3921
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":3A33
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":3B45
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":3C57
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":3D69
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":3E7B
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":3F8D
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":409F
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":41B1
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":42C3
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":43D5
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":44E7
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":4939
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":4D8B
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":4E9D
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":4FF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":5155
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":52B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":540D
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":5ED9
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":632D
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":6491
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":68ED
            Key             =   "Bmp"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":6A05
            Key             =   "Dibujo"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport rptRecursos 
      Left            =   8400
      Top             =   8040
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
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   5520
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   47
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":6B21
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":6C33
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":6D45
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":6E57
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":6F69
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":707B
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":718D
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":729F
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":73B1
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":74C3
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":75D5
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":76E7
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":77F9
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":790B
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":7A1D
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":7B2F
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":7C41
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":8093
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":84E5
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":85F7
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":8753
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":88AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":8A0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":8B67
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":9633
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":9A87
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":9BEB
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":A047
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":A1A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":B4AF
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":BA4B
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":BBA7
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":BD03
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":C057
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":C3AB
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":C6FF
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":CA53
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":CDA7
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":D2EB
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":D3FD
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":D751
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":DAA5
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":DBB9
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":DF0D
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":E01F
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":E373
            Key             =   "Cotizacion"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionRecursos.frx":E487
            Key             =   "Dibujo"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTarea 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   21
      Top             =   8160
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmAsignacionRecursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dblBackColorSabado As Double
Dim dblForeColorSabado As Double
Dim dblBackColorDomingo As Double
Dim dblForeColorDomingo As Double
Dim dblBackColorNormal As Double
Dim dblForeColorNormal As Double
Dim dblBackColorFeriado As Double
Dim dblForeColorFeriado As Double
Dim dblBackColorTotales As Double
Dim dblForeColorTotales As Double

Dim intColBase As Integer
Dim SW As Boolean
Dim strIdTarea As String
Dim Feriados() As Date
Dim arrHojaRecursos() As HojaRecurso
Private Sub BarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
  Screen.MousePointer = vbHourglass
    Select Case Button.Key
        Case "Nuevo"
            Limpiar
        Case "Buscar"
            Buscar
        Case "Cerrar"
            Unload Me
        Case "Color"
            AjustarColores
        Case "Imprimir"
            ImprimirReporte
    End Select
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdlimpiar_Click(Index As Integer)
    Select Case Index
        Case 0
            Me.dbcSucursal.BoundText = ""
        Case 1
            Me.dbcTurno.BoundText = ""
    End Select
End Sub

Private Sub fgrDetalle_DblClick()
Dim cuentalineas As Integer
    
    cuentalineas = Me.fgrDetalle.Rows - Me.fgrDetalle.Row
    If cuentalineas > 1 Then
    
    Text1.Text = arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Sucursal
    Text2.Text = arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Turno
    Text3.Text = arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Item
    Text4.Text = arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Mecanico
    Text5.Text = arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Fecha

    frmEditaAsignacionOt.Show vbModal
    
    If Me.Text1.Tag = "QUITO" Then
        Buscar
    End If
    End If
End Sub

Private Sub fgrDetalle_DragDrop(Source As Control, x As Single, Y As Single)
    Dim i As Integer, j As Integer
    Dim intColDesde As Integer
    Dim SW As Boolean

    SW = False
    With Me.fgrDetalle
        For i = 3 To .Rows - 2 'Solo llega hasta la fila antes del total...
            If Y >= .RowPos(i) And Y <= (.RowPos(i) + .RowHeight(i)) Then
                .Row = i
                For j = (intColBase - 2) To .Cols - 2
                    intColDesde = (.ColPos(j) + (.ColWidth(4) * (j - (intColBase - 1))))
                    If x >= intColDesde And x <= (intColDesde + .ColWidth(j)) Then
                        '//Aqui determina el destino de la tarea arrastrada...
                        .Col = IIf(j > 20, j - 1, j)
                        
                        'actualiza hoja de recursos
                        ActualizaHojaRecursos
                        
                        'Crear hoja de recursos de detalle
                        CreaHojaDetalle
                        
                        'Actualiza mano de obra servicio
                        ActualizaTareasPendientes
                        
                        'Actualiza numero de id_tarea
                        strIdTarea = Val(strIdTarea) + 1
                        Conexion.SendHost "Update Tllr_Parametro set Id_Tarea='" & strIdTarea & "'", , , , gcTiempoEspera
                        
                        'Actualiza horas en pantalla
                        arrHojaRecursos(.Col - (intColBase - 2), .Row - 2).Horas = arrHojaRecursos(.Col - (intColBase - 2), .Row - 2).Horas - CDbl(Me.lvTareas.SelectedItem.SubItems(14))
                        .Text = arrHojaRecursos(.Col - (intColBase - 2), .Row - 2).Horas
                        
                        'totaliza fila
                        .Col = .Cols - 1
                        .Text = CDbl(.Text) - CDbl(Me.lvTareas.SelectedItem.SubItems(14))
                        
                        'totaliza columna
                        .Col = IIf(j > 20, j - 1, j)
                        .Row = .Rows - 1
                        .Text = CDbl(.Text) - CDbl(Me.lvTareas.SelectedItem.SubItems(14))
                        
                        .Col = .Cols - 1
                        .Text = CDbl(.Text) - CDbl(Me.lvTareas.SelectedItem.SubItems(14))
                        
                        'vuelve al origen
                        .Col = IIf(j > 20, j - 1, j)
                        
                        'Elimina de la lista
                        If Me.lvTareas.ListItems.Count > 0 Then
                            If Not Me.lvTareas.SelectedItem Is Nothing Then
                                lvTareas.ListItems.Remove lvTareas.SelectedItem.Index
                            End If
                        End If

                        SW = True
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    End With
End Sub
Private Sub Form_Activate()
    If Not SW Then
        If Not Atributos("Glbl", "Tllr_20_0120", False, False, False, False) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
    
        FormateaGrilla
        SW = True
    End If
End Sub
Private Sub Form_Load()
    SW = False
    CargaSucursal
    CargaTurno
    CargaColores
    Limpiar
    CargaFeriados
End Sub
Private Sub CargaSucursal()
    Dim strSql As String
    Dim adoTemp As New ADODB.Recordset
    
    strSql = "select id_sucursal as codigo, descripcion as nombre from Glbl_Sucursal where TieneTaller='S' and id_empresa='" & gstrIdEmpresa & "' order by descripcion"
    If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        Set Me.datSucursal.Recordset = adoTemp
    End If
    Set adoTemp = New ADODB.Recordset
End Sub
Private Sub CargaTurno()
    Dim strSql As String
    Dim adoTemp As New ADODB.Recordset

    strSql = "select id_turno as codigo, descripcion as nombre from Tllr_Turnos where id_empresa='" & gstrIdEmpresa & "' order by descripcion"
    If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        Set Me.datTurno.Recordset = adoTemp
    End If
    Set adoTemp = New ADODB.Recordset
End Sub
Private Sub Buscar()
    Dim strSql As String
    Dim adoTemp As New ADODB.Recordset
    Dim strSucursal As String
    Dim strTurno As String
    Dim strMecanico As String
    Dim strHoras As String
    Dim i As Integer
    Dim intLinea As Integer
    Dim intFila As Integer
    Dim dblBackColor As Double
    Dim dblForeColor As Double
    Dim dblTotal As Double
    Dim intColMatriz As Integer
    Dim intFilMatriz As Integer

    '//Valida el horizonte...
    If txtHorizonte.Text = "" Then
        txtHorizonte.Text = 1
    ElseIf CDbl(txtHorizonte.Text) = 0 Then
        txtHorizonte.Text = 1
    ElseIf CDbl(txtHorizonte.Text) > 15 Then
        txtHorizonte.Text = 15
    End If


    For i = 0 To Me.optHoras.Count
        If Me.optHoras(i).Value Then
            strHoras = i + 1
            Exit For
        End If
    Next
   
    FormateaGrilla
    
    '//Crea matriz...
    ReDim arrHojaRecursos(0 To 0, 0 To 0) As HojaRecurso
    intColMatriz = Me.fgrDetalle.Cols - (intColBase - 1)
    intFilMatriz = 0
    ReDim arrHojaRecursos(0 To intColMatriz, 0 To intFilMatriz) As HojaRecurso
    strSql = "exec Tllr_Asignacion_Hoja_Recursos '" & gstrIdEmpresa & "','" & Me.dbcSucursal.BoundText & "','" & Me.dbcTurno.BoundText & "','" & Format(Me.dtpFecha.Value, "dd/mm/yyyy") & "','" & strHoras & "'," & Me.txtHorizonte
    If Conexion.SendHost(strSql, adoTemp, adOpenForwardOnly, adLockOptimistic, 10) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            adoTemp.MoveFirst
            strSucursal = ""
            strTurno = ""
            strMecanico = ""
            Me.fgrDetalle.Row = 2
            intLinea = 0
            dblTotal = 0
            While Not adoTemp.EOF
                If adoTemp!Id_Mecanico <> strMecanico Then
                    With Me.fgrDetalle
                        If .Row > 2 Then
                            .Col = .Cols - 1
                            .Text = dblTotal
                            dblTotal = 0
                            .AddItem ""
                        End If
                        .Row = .Row + 1
                        intFilMatriz = intFilMatriz + 1
                        ReDim Preserve arrHojaRecursos(0 To intColMatriz, 0 To intFilMatriz) As HojaRecurso
                        If .Row > 3 Then
                            ColoreaFilas .Row
                        End If
                        intLinea = intLinea + 1
                        .Col = 0
                        .Text = intLinea
                        .Col = 1
                        If Me.dbcSucursal.BoundText = "" Then
                            .Text = adoTemp!Nombre_Sucursal
                            .Col = .Col + 1
                        End If
                        If Me.dbcTurno.BoundText = "" Then
                            .Text = adoTemp!nombre_turno
                            .Col = .Col + 1
                        End If
                        .Text = adoTemp!nombre_mecanico
                    End With
                    strMecanico = adoTemp!Id_Mecanico
                    
                End If
                
                With Me.fgrDetalle
                    .Col = DateDiff("d", Me.dtpFecha.Value, adoTemp!Id_Fecha) + intColBase - 1
                    If .Row <> 3 Then
                        intFila = .Row
                        .Row = 3
                        dblBackColor = .CellBackColor
                        dblForeColor = .CellForeColor
                        .Row = intFila
                    Else
                        dblBackColor = .CellBackColor
                        dblForeColor = .CellForeColor
                    End If
                    .CellBackColor = dblBackColor
                    .CellForeColor = dblForeColor
'                    .Text = adoTemp!Horas
                    .Text = ValorNulo(adoTemp!Horas)
                    .CellAlignment = flexAlignRightCenter
                    
                    arrHojaRecursos(.Col - (intColBase - 2), .Row - 2).Id_Sucursal = adoTemp!Id_Sucursal
                    arrHojaRecursos(.Col - (intColBase - 2), .Row - 2).Id_Turno = adoTemp!Id_Turno
                    arrHojaRecursos(.Col - (intColBase - 2), .Row - 2).Id_Item = adoTemp!Id_Item
                    arrHojaRecursos(.Col - (intColBase - 2), .Row - 2).Id_Mecanico = adoTemp!Id_Mecanico
                    arrHojaRecursos(.Col - (intColBase - 2), .Row - 2).Id_Fecha = adoTemp!Id_Fecha
                    arrHojaRecursos(.Col - (intColBase - 2), .Row - 2).Horas = adoTemp!Horas
                    

'                    dblTotal = dblTotal + adoTemp!Horas
                    dblTotal = dblTotal + adoTemp!Horas
                    
                End With
                adoTemp.MoveNext
            Wend
            With Me.fgrDetalle
                .Col = .Cols - 1
                .Text = dblTotal
            End With
        End If
    End If
    Conexion.CloseHost adoTemp
    Totaliza
    With Me.fgrDetalle
        .Row = 3
        .Col = intColBase
    End With
    
    '//BuscarOt
    BuscarOt
End Sub
Private Sub Limpiar()
    Me.txtHorizonte = 15
    Me.dbcSucursal.BoundText = ""
    Me.dbcTurno.BoundText = ""
    Me.dtpFecha.Value = CDate("01/" & Month(Date) & "/" & Year(Date))
End Sub
Private Sub FormateaGrilla()
    Dim i As Date
    Dim j As Integer
    Dim intCol As Integer

    Dim Dia(7) As String
    Dim Mes(12) As String
    
    Dia(1) = "L"
    Dia(2) = "M"
    Dia(3) = "M"
    Dia(4) = "J"
    Dia(5) = "V"
    Dia(6) = "S"
    Dia(7) = "D"
    Mes(1) = "ENERO"
    Mes(2) = "FEBRERO"
    Mes(3) = "MARZO"
    Mes(4) = "ABRIL"
    Mes(5) = "MAYO"
    Mes(6) = "JUNIO"
    Mes(7) = "JULIO"
    Mes(8) = "AGOSTO"
    Mes(9) = "SEPTIEMBRE"
    Mes(10) = "OCTUBRE"
    Mes(11) = "NOVIEMBRE"
    Mes(12) = "DICIEMBRE"
    
    intColBase = 5
    If Me.dbcSucursal.BoundText <> "" Then
        intColBase = intColBase - 1
    End If
    If Me.dbcTurno.BoundText <> "" Then
        intColBase = intColBase - 1
    End If

    j = 2
    With Me.fgrDetalle
        .ClearStructure
        .BackColor = dblBackColorNormal
        .ForeColor = dblForeColorNormal
        .Rows = 4
        .Cols = intColBase + Int(Me.txtHorizonte)
        
        For j = 0 To .Cols - 1
            .MergeCol(j) = False
        Next
        .FixedRows = 3
        .FixedCols = intColBase - 1
        .Col = 0
        .ColWidth(.Col) = 400
        .Row = 0
        For j = 1 To 3
            .Text = "Nº"
            .CellAlignment = flexAlignCenterCenter
            .Row = .Row + 1
        Next
        .MergeCol(.Col) = True
        .Col = .Col + 1
        If Me.dbcSucursal.BoundText = "" Then
            .ColWidth(.Col) = 1500
            .Row = 0
            For j = 1 To 3
                .Text = "SUCURSAL"
                .CellAlignment = flexAlignCenterCenter
                .Row = .Row + 1
            Next
            .MergeCol(.Col) = True
            .Col = .Col + 1
        End If
        If Me.dbcTurno.BoundText = "" Then
            .ColWidth(.Col) = 1500
            .Row = 0
            For j = 1 To 3
                .Text = "TURNO"
                .CellAlignment = flexAlignCenterCenter
                .Row = .Row + 1
            Next
            .MergeCol(.Col) = True
            .Col = .Col + 1
        End If
        .ColWidth(.Col) = 1500
        .Row = 0
        For j = 1 To 3
            .Text = "MECANICO"
            .CellAlignment = flexAlignCenterCenter
            .Row = .Row + 1
        Next
        .MergeCol(.Col) = True
        .Col = .Col + 1
        .MergeRow(0) = False
        For i = Me.dtpFecha.Value To CDate(Me.dtpFecha.Value + (Int(Me.txtHorizonte) - 1))
            .ColWidth(.Col) = 450
            .Row = 0
            For j = 1 To 4
                Select Case j
                    Case 1
                        .Text = Mes(Month(i)) & "-" & Format(Year(i), "0000")
                        .CellAlignment = flexAlignCenterCenter
                    Case 2
                        .Text = Dia(Weekday(i, vbMonday))
                    Case 3
                        .Text = Format(i, "dd")
                    Case 4
                        '//Formato segun los colores establecidos...
                        If Weekday(i, vbMonday) = 6 Then
                            .CellBackColor = dblBackColorSabado
                            .CellForeColor = dblForeColorSabado
                        ElseIf Weekday(i, vbMonday) = 7 Then
                            .CellBackColor = dblBackColorDomingo
                            .CellForeColor = dblForeColorDomingo
                        Else
                            .CellBackColor = dblBackColorNormal
                            .CellForeColor = dblForeColorNormal
                        End If
                        If ColorFeriado(i) Then
                            .CellBackColor = dblBackColorFeriado
                            .CellForeColor = dblForeColorFeriado
                        End If
                End Select
                If .Row < 3 Then
                    .Row = .Row + 1
                    .CellAlignment = flexAlignCenterCenter
                End If
            Next
            '.MergeCol(.Col) = True
            .Col = .Col + 1
        Next
        .MergeRow(0) = True
        .ColWidth(.Col) = 1500
        .Row = 2
        For j = 3 To 3
            .Text = "TOTAL"
            .CellAlignment = flexAlignCenterCenter
            .Row = .Row + 1
        Next
        .CellBackColor = dblBackColorTotales
        .CellForeColor = dblForeColorTotales
        '.MergeCol(.Col) = True
        
        .MergeCells = flexMergeFree
        '.MergeCells = flexMergeNever
    End With

End Sub
Private Sub CargaFeriados()
    Dim strSql As String
    Dim adoTemp As New ADODB.Recordset
    Dim i As Integer
    '//Verifica Feriados...
    ReDim Feriados(0) As Date
    strSql = "select Id_Feriado from Tllr_Dia_No_Laboral where id_feriado between '" & Me.dtpFecha.Value & "' and '" & Me.dtpFecha.Value + Int(30) - 1 & "'"
    If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            While Not adoTemp.EOF
                i = UBound(Feriados) + 1
                ReDim Preserve Feriados(i) As Date
                Feriados(i) = adoTemp!Id_Feriado
                adoTemp.MoveNext
            Wend
        End If
    End If
    Conexion.CloseHost adoTemp
End Sub
Private Function ColorFeriado(Fecha As Date) As Boolean
    Dim i As Integer
    ColorFeriado = False
    For i = 1 To UBound(Feriados)
        If Feriados(i) = Fecha Then
            ColorFeriado = True
            Exit For
        End If
    Next
End Function
Private Sub CargaColores()
    Dim strSql As String
    Dim adoTemp As New ADODB.Recordset

    Do
        strSql = "select * from Tllr_Hoja_Recursos_Colores where Id_Empresa='" & gstrIdEmpresa & "' and id_usuario='" & gstrIdUsuario & "'"
        If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
            If Not adoTemp.BOF And Not adoTemp.EOF Then
                With adoTemp
                    dblBackColorSabado = !BackColorSabado
                    dblForeColorSabado = !ForeColorSabado
    
                    dblBackColorDomingo = !BackColorDomingo
                    dblForeColorDomingo = !ForeColorDomingo
    
                    dblBackColorNormal = !BackColorNormal
                    dblForeColorNormal = !ForeColorNormal
    
                    dblBackColorFeriado = !BackColorFestivos
                    dblForeColorFeriado = !ForeColorFestivos
                
                    dblBackColorTotales = !BackColorTotales
                    dblForeColorTotales = !ForeColorTotales
                    Exit Do
                End With
            Else
                strSql = "insert into Tllr_Hoja_Recursos_Colores (id_empresa, id_usuario, BackColorSabado, ForeColorSabado, BackColorDomingo, ForeColorDomingo, "
                strSql = strSql & "BackColorNormal, ForeColorNormal, BackColorFestivos, ForeColorFestivos, BackColorTotales, ForeColorTotales, Usr_Id, Usr_Fecha) "
                strSql = strSql & "values('" & gstrIdEmpresa & "', '" & gstrIdUsuario & "', 16777215, 0, 16777215, 0, 16777215, 0, 255, 0,  16777215 ,0, '" & gstrIdUsuario & "','" & Format(Date, "dd/mm/yyyy") & "')"
                Conexion.SendHost strSql, , , , 10
            End If
        Else
            dblBackColorSabado = &HE0E0E0
            dblForeColorSabado = vbBlack
        
            dblBackColorDomingo = &HC0C0C0
            dblForeColorDomingo = vbBlack
        
            dblBackColorNormal = &H80FFFF
            dblForeColorNormal = vbBlack
        
            dblBackColorFeriado = vbRed
            dblForeColorFeriado = vbBlack
        
            dblBackColorTotales = &HE0E0E0
            dblForeColorTotales = vbBlack
        End If
    Loop
    Conexion.CloseHost adoTemp
End Sub
Private Sub ColoreaFilas(intFila As Integer)
    Dim i As Integer, j As Integer
    Dim intFilaAnterior As Integer
    Dim dblBackColor As Double
    Dim dblForeColor As Double
    
    With Me.fgrDetalle
        For j = intColBase To .Cols
            .Col = j - 1
            intFilaAnterior = .Row
            .Row = 3
            dblBackColor = .CellBackColor
            dblForeColor = .CellForeColor
            .Row = intFilaAnterior
            .CellBackColor = dblBackColor
            .CellForeColor = dblForeColor
        Next
    End With
End Sub
Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        Me.frmFiltro.Width = Me.ScaleWidth
        Me.fgrDetalle.Left = 0
        Me.fgrDetalle.Width = Me.ScaleWidth
    End If
End Sub
Private Sub lvTareas_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim intCol As Integer
    Dim intFil As Integer
    If Button = vbRightButton Then
        If Me.lvTareas.ListItems.Count > 0 Then
            With Me.fgrDetalle
                intCol = .Col
                intFil = .Row
                .Row = 3
                .Col = intColBase
            End With
            With Me.lblTarea
                '.Move lvTareas.Left, frTareas.Top + lvTareas.SelectedItem.Height + y - TextHeight("A") / 2, lvTareas.Width, TextHeight("A")
                .Move x, frTareas.Top + Me.lvTareas.SelectedItem.Height + Y, Me.fgrDetalle.CellWidth, Me.fgrDetalle.CellHeight
                .Drag vbBeginDrag
            End With
            With Me.fgrDetalle
                .Col = intCol
                .Row = intFil
            End With
        End If
    End If
End Sub
Private Sub txtHorizonte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        Exit Sub
    End If
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Exit Sub
    End If
    KeyAscii = 0
End Sub
Private Sub Totaliza()
    Dim i As Integer, j As Integer
    Dim dblTotal As Double

    With Me.fgrDetalle
        .AddItem ""
        .Row = .Rows - 1
        .Col = intColBase - 2
        .Text = "TOTALES"
        .CellAlignment = flexAlignRightCenter
        
        '//Totaliza por columnas...
        For i = intColBase To .Cols
            .Col = i - 1
            dblTotal = 0
            For j = 3 To .Rows - 1
                .Row = j
                If j = (.Rows - 1) Then
                    .Text = dblTotal
                    .CellBackColor = dblBackColorTotales
                    .CellForeColor = dblForeColorTotales
                Else
                    If .Text <> "" Then
                        dblTotal = dblTotal + CDbl(.Text)
                    End If
                End If
            Next
        Next

'        '//Totaliza por filas...
'        For I = 3 To .Rows - 1
'            .Row = I
'            dblTotal = 0
'            For J = intColBase - 1 To .Cols - 1
'                .Col = J
'                If J = (.Cols - 1) Then
'                    .Text = dblTotal
'                    .CellBackColor = dblBackColorTotales
'                    .CellForeColor = dblForeColorTotales
'                Else
'                    If .Text <> "" Then
'                        dblTotal = dblTotal + CDbl(.Text)
'                    End If
'                End If
'            Next
'        Next
    End With
End Sub
Private Sub AjustarColores()
    frmAsignacionRecursos.Tag = "N"
    frmAsignacionRecursosColor.Show vbModal
    If frmAsignacionRecursos.Tag = "S" Then
        CargaColores
        FormateaGrilla
    End If
End Sub
Private Sub BuscarOt()

    Dim strSql As String
    Dim adoTemp As New ADODB.Recordset
    Dim Item As ListItem
    Dim i As Integer

    Me.lvTareas.ListItems.Clear
    i = 1
    strSql = "exec Tllr_Tareas_Pendientes '" & gstrIdEmpresa & "', '" & Me.dbcSucursal.BoundText & "'"
    If Conexion.SendHost(strSql, adoTemp, adOpenForwardOnly, adLockOptimistic, 10) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            While Not adoTemp.EOF
                With adoTemp
                    Set Item = Me.lvTareas.ListItems.Add(, , i)
                    
                    Item.SubItems(1) = !Id_Sucursal
                    Item.SubItems(2) = !Nombre_Sucursal
                    Item.SubItems(3) = !Id_OT
                    Item.SubItems(4) = !Seccion_OT
                    Item.SubItems(5) = IIf(!Seccion_OT = "M", "MECANICA", "CARROCERIA")
                    Item.SubItems(6) = Format(!Fecha_Emision, "dd/mm/yyyy")
                    Item.SubItems(7) = !Nro_Cono
                    Item.SubItems(8) = !Patente
                    Item.SubItems(9) = !Nombre_Marca
                    Item.SubItems(10) = !Nombre_Modelo
                    Item.SubItems(11) = !Nombre_Cliente
                    Item.SubItems(12) = !Id_servicio
                    Item.SubItems(13) = !Nombre_Servicio
                    Item.SubItems(14) = !Horas
                    Item.SubItems(15) = "M"
                End With
                
                adoTemp.MoveNext
                i = i + 1
            Wend
        End If
    End If
    Conexion.CloseHost adoTemp
    
    
    '/////// Otros Servicios
    strSql = "exec Tllr_Tareas_Pendientes_Otro '" & gstrIdEmpresa & "', '" & Me.dbcSucursal.BoundText & "'"
    If Conexion.SendHost(strSql, adoTemp, adOpenForwardOnly, adLockOptimistic, 10) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            While Not adoTemp.EOF
                With adoTemp
                    Set Item = Me.lvTareas.ListItems.Add(, , i)
                    
                    Item.SubItems(1) = !Id_Sucursal
                    Item.SubItems(2) = !Nombre_Sucursal
                    Item.SubItems(3) = !Id_OT
                    Item.SubItems(4) = !Seccion_OT
                    Item.SubItems(5) = IIf(!Seccion_OT = "M", "MECANICA", "CARROCERIA")
                    Item.SubItems(6) = Format(!Fecha_Emision, "dd/mm/yyyy")
                    Item.SubItems(7) = !Nro_Cono
                    Item.SubItems(8) = !Patente
                    Item.SubItems(9) = !Nombre_Marca
                    Item.SubItems(10) = !Nombre_Modelo
                    Item.SubItems(11) = !Nombre_Cliente
                    Item.SubItems(12) = !Id_Otro_Servicio
                    Item.SubItems(13) = !Descripcion_Otro
                    Item.SubItems(14) = !Horas
                    Item.SubItems(15) = "O"
                End With
                
                adoTemp.MoveNext
                i = i + 1
            Wend
        End If
    End If
    Conexion.CloseHost adoTemp
    
End Sub
Private Sub ActualizaTareasPendientes()
Dim strSql As String
Dim adoTemp As New ADODB.Recordset
    
    If Me.lvTareas.SelectedItem.SubItems(15) = "M" Then
        strSql = "Update Tllr_Mecanica_Ot Set Id_tarea='" & strIdTarea & "',"
        strSql = strSql & " Mecanico_Designado='" & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Mecanico & "'"
        strSql = strSql & " Where Id_Empresa='" & gstrIdEmpresa & "'"
        strSql = strSql & " And Id_Sucursal='" & Me.lvTareas.SelectedItem.SubItems(1) & "'"
        strSql = strSql & " And Id_Ot='" & Me.lvTareas.SelectedItem.SubItems(3) & "'"
        strSql = strSql & " And Seccion_Ot='" & Me.lvTareas.SelectedItem.SubItems(4) & "'"
        strSql = strSql & " And Id_Servicio='" & Me.lvTareas.SelectedItem.SubItems(12) & "'"
    Else
        strSql = "Update Tllr_Otro_Ot Set Id_tarea='" & strIdTarea & "',"
        strSql = strSql & " Mecanico_Asignado='" & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Mecanico & "'"
        strSql = strSql & " Where Id_Empresa='" & gstrIdEmpresa & "'"
        strSql = strSql & " And Id_Sucursal='" & Me.lvTareas.SelectedItem.SubItems(1) & "'"
        strSql = strSql & " And Id_Ot='" & Me.lvTareas.SelectedItem.SubItems(3) & "'"
        strSql = strSql & " And Seccion_Ot='" & Me.lvTareas.SelectedItem.SubItems(4) & "'"
        strSql = strSql & " And Id_Otro_Servicio='" & Me.lvTareas.SelectedItem.SubItems(12) & "'"
    End If
    Conexion.SendHost strSql, , , , 10
    Set adoTemp = New ADODB.Recordset
End Sub
Private Sub ActualizaHojaRecursos()
Dim strSql As String
Dim dblHorasAsignadas As Double
Dim adoTemp As New ADODB.Recordset
    
    dblHorasAsignadas = 0

    'rescato horas asignadas
    strSql = "Select Horas_Asignadas from Tllr_Hoja_Recursos where id_empresa='" & gstrIdEmpresa & "'"
    strSql = strSql & " And Id_sucursal='" & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Sucursal & "'"
    strSql = strSql & " And Id_Mecanico='" & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Mecanico & "'"
    strSql = strSql & " And Id_Turno='" & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Turno & "'"
    strSql = strSql & " And Id_Item=" & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Item
    strSql = strSql & " And id_Fecha='" & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Fecha & "'"
    
    If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            dblHorasAsignadas = adoTemp!Horas_Asignadas
        End If
    End If
    
    'actualiza horas asignadas y disponibles
    strSql = "Update Tllr_Hoja_Recursos Set"
    strSql = strSql & " Horas_Disponibles=" & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Horas - CDbl(Me.lvTareas.SelectedItem.SubItems(14))
    strSql = strSql & ", Horas_Asignadas=" & dblHorasAsignadas + CDbl(Me.lvTareas.SelectedItem.SubItems(14))
    strSql = strSql & " Where Id_Empresa='" & gstrIdEmpresa & "'"
    strSql = strSql & " And Id_sucursal='" & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Sucursal & "'"
    strSql = strSql & " And Id_Mecanico='" & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Mecanico & "'"
    strSql = strSql & " And Id_Turno='" & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Turno & "'"
    strSql = strSql & " And Id_Item=" & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Item
    strSql = strSql & " And id_Fecha='" & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Fecha & "'"
    
    Conexion.SendHost strSql, , , , 10
    Set adoTemp = New ADODB.Recordset
End Sub
Private Sub CreaHojaDetalle()
Dim strSql As String

Dim adoTemp As New ADODB.Recordset

    strSql = "Select Id_Tarea From Tllr_Parametro where Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Sucursal & "'"
    If Conexion.SendHost(strSql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If Not adoTemp.BOF And Not adoTemp.EOF Then
            strIdTarea = IIf(IsNull(adoTemp!Id_tarea), "1", adoTemp!Id_tarea)
        End If
    End If
    
    strSql = "Insert Into Tllr_Hoja_Recursos_Detalle (Id_Empresa,ID_Sucursal,Id_Mecanico,Id_Turno,Id_Item,Id_Fecha,Id_Tarea,Id_Ot,Id_Seccion,Horas_Asignadas,Id_Servicio)"
    strSql = strSql & " Values ('" & gstrIdEmpresa & "','"
    strSql = strSql & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Sucursal & "','"
    strSql = strSql & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Mecanico & "','"
    strSql = strSql & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Turno & "',"
    strSql = strSql & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Item & ",'"
    strSql = strSql & arrHojaRecursos(Me.fgrDetalle.Col - (intColBase - 2), Me.fgrDetalle.Row - 2).Id_Fecha & "','"
    strSql = strSql & strIdTarea & "','"
    strSql = strSql & Me.lvTareas.SelectedItem.SubItems(3) & "','"
    strSql = strSql & Me.lvTareas.SelectedItem.SubItems(4) & "',"
    strSql = strSql & Me.lvTareas.SelectedItem.SubItems(14) & ",'"
    strSql = strSql & Me.lvTareas.SelectedItem.SubItems(12) & "')"
    Conexion.SendHost strSql, , , , 10

    Set adoTemp = New ADODB.Recordset
End Sub
Private Sub ImprimirReporte()
Dim Dbsnueva As Database
Dim Tabla As DAO.Recordset
Dim i As Integer
Dim GcamBaseTem As String
Dim dblTotal As Double
Dim j As Integer
Dim lstrTipoHoras As String

    'Devuelve la ruta del directorio Windows
    Dim rc As Long
    Dim WinPath As String
    WinPath = Space$(300)
    rc = GetWindowsDirectory(WinPath, 300)
    GcamBaseTem = Trim$(WinPath)
    GcamBaseTem = Mid(GcamBaseTem, 1, Len(GcamBaseTem) - 1) & "\Temp"
    '---------------------------------------
    
'    If Me.fgrDetalle Then
'      MsgBox "No existen elementos en la lista", vbExclamation, "Imprimir"
'      Exit Sub
'    End If

    Screen.MousePointer = 11
    Dim wrkPredeterminado As Workspace
    Dim prpBucle As Property
    Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
'    If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    If Dir(gstrPathReporte & "\BDNueva.mdb") <> "" Then Kill gstrPathReporte & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
'    Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(gstrPathReporte & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    'Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (Sucursal text,Turno text,Mecanico text,FechaInicio text,Horizonte text,Horas text,Mes text,Dia1 text,Num1 text,Valor1 text,Dia2 text,Num2 text,Valor2 text,Dia3 text,Num3 text,Valor3 text,Dia4 text,Num4 text,Valor4 text,Dia5 text,Num5 text, Valor5 text,Dia6 text,Num6 text,Valor6 text,Dia7 text, Num7 text,Valor7 text," & _
    '                                              "Dia8 text,Num8 text,Valor8 text,Dia9 text,Num9 text,Valor9 text,Dia10 text,Num10 text,Valor10 text,Dia11 text,Num11 text,Valor11 text,Dia12 text,Num12 text,Valor12 text,Dia13 text,Num13 text,Valor13 text,Dia14 text,Num14 text,Valor14 text,Dia15 text,Num15 text,Valor15 text,Total1 text)"
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (Sucursal text,Turno text,FechaInicio text,Horizonte text,Horas text,col1 text,col2 text,col3 text,col4 text,col5 text,col6 text,col7 text,col8 text,col9 text,col10 text,col11 text,col12 text,col13 text,col14 text,col15 text, col16 text,col17 text,col18 text,col19 text)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")

    With Me.fgrDetalle
    
        For i = 1 To .Rows - 1
                .Row = i
                Tabla.AddNew
                Tabla!Sucursal = Me.dbcSucursal.Text
                Tabla!Turno = Me.dbcTurno.Text
                Tabla!FechaInicio = Me.dtpFecha.Value
                Tabla!Horizonte = Me.txtHorizonte
                Tabla!Horas = "DISPONIBLES"
                
                If i <> 1 Then
                    .Col = 1
                    Tabla!Col1 = .Text
                Else
                    .Col = 1
                    Tabla!Col1 = ""
                End If
                
                .Col = 2
                Tabla!Col2 = .Text
                
                .Col = 3
                Tabla!col3 = .Text
                
                .Col = 4
                Tabla!col4 = .Text
                
                .Col = 5
                Tabla!col5 = .Text
                
                .Col = 6
                Tabla!col6 = .Text
                
                .Col = 7
                Tabla!col7 = .Text
                
                .Col = 8
                Tabla!col8 = .Text
                
                .Col = 9
                Tabla!col9 = .Text
                
                .Col = 10
                Tabla!col10 = .Text
                
                .Col = 11
                Tabla!col11 = .Text
                
                .Col = 12
                Tabla!col12 = .Text
                
                .Col = 13
                Tabla!col13 = .Text
                
                .Col = 14
                Tabla!col14 = .Text
                
                .Col = 15
                Tabla!col15 = .Text
                
                .Col = 16
                Tabla!col16 = .Text
                
                .Col = 17
                Tabla!col17 = .Text
                
                If Me.dbcSucursal = "" Then
                    .Col = 18
                    Tabla!col18 = .Text
                End If
                
                If Me.dbcTurno = "" Then
                    .Col = 19
                    Tabla!col19 = .Text
                End If
                Tabla.Update
            Next
        
    End With
    
   Tabla.Close
   Dbsnueva.Close
   
   With rptRecursos
        If Me.optHoras(0).Value = True Then
            lstrTipoHoras = "COMPRADAS"
        ElseIf Me.optHoras(1).Value = True Then
            lstrTipoHoras = "ASIGNADAS"
        ElseIf Me.optHoras(2).Value = True Then
            lstrTipoHoras = "AUSENTES"
        ElseIf Me.optHoras(3).Value = True Then
            lstrTipoHoras = "DISPONIBLES"
        End If
   
        .ReportFileName = gstrPathReporte & "\AsigRecursosDetalle.rpt"
        .WindowTitle = "Reporte de Asignación de Recursos"
'        .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
        .DataFiles(0) = gstrPathReporte & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='REPORTE DE ASIGNACION DE RECURSOS'"
        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
        .Formulas(5) = "SucursalListado='" & Me.dbcSucursal.Text & "'"
        .Formulas(6) = "Turno='" & Me.dbcTurno.Text & "'"
        .Formulas(7) = "FechaInicio='" & Me.dtpFecha.Value & "'"
        .Formulas(8) = "Horizonte='" & Me.txtHorizonte & "'"
        .Formulas(9) = "TipoHoras='" & lstrTipoHoras & "'"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = True
   End With
   
'   Dbsnueva.Close
   Screen.MousePointer = 1

End Sub

