VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmAusenciaMecanicos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ausencia de Mecanicos"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   Icon            =   "frmAusenciaMecanicos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9255
      Begin VB.CommandButton cmdCancelaMecanico 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3840
         Picture         =   "frmAusenciaMecanicos.frx":179A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton cmdCancelaMotivo 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3840
         Picture         =   "frmAusenciaMecanicos.frx":189C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cmdCancelaSuc 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3840
         Picture         =   "frmAusenciaMecanicos.frx":199E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fechas"
         Height          =   1095
         Left            =   5280
         TabIndex        =   7
         Top             =   240
         Width           =   3015
         Begin MSComCtl2.DTPicker pckFechaHasta 
            Height          =   345
            Left            =   1320
            TabIndex        =   11
            Top             =   645
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Format          =   83427329
            CurrentDate     =   37382
         End
         Begin MSComCtl2.DTPicker pckFechaDesde 
            Height          =   345
            Left            =   1320
            TabIndex        =   10
            Top             =   225
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Format          =   83427329
            CurrentDate     =   37382
         End
         Begin VB.Label Label5 
            Caption         =   "Hasta"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Desde"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   615
         End
      End
      Begin MSDataListLib.DataCombo dtcSucursal 
         Bindings        =   "frmAusenciaMecanicos.frx":1AA0
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "NOMBRE"
         BoundColumn     =   "CODIGO"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc datSucursal 
         Height          =   330
         Left            =   1920
         Top             =   360
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
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
      Begin MSComctlLib.ListView lvwConceptos 
         Height          =   3315
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5847
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Codigo"
            Text            =   "Sucursal"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ausencia"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cod. Mecánico"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Mecánico"
            Object.Width           =   5468
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fecha"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Desde"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Hasta"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Horas"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "IdItem"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "IdSucursal"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "IdAusencia"
            Object.Width           =   0
         EndProperty
      End
      Begin Crystal.CrystalReport rptMantenedor 
         Left            =   6840
         Top             =   1200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
      End
      Begin MSDataListLib.DataCombo dtcMotivo 
         Bindings        =   "frmAusenciaMecanicos.frx":1ABA
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "NOMBRE"
         BoundColumn     =   "CODIGO"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc datMotivo 
         Height          =   330
         Left            =   1920
         Top             =   1320
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
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
      Begin MSDataListLib.DataCombo dtcSupervisor 
         Bindings        =   "frmAusenciaMecanicos.frx":1AD2
         Height          =   315
         Left            =   1200
         TabIndex        =   14
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
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
         Left            =   2520
         Top             =   840
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
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
      Begin VB.Label Label4 
         Caption         =   "Mecanico"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Motivo"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Crear"
            Object.ToolTipText     =   "Crear Registro (Ctrl+N)"
            ImageKey        =   "Crear"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar Registro (Ctrl+B)"
            ImageKey        =   "Buscar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Editar"
            Object.ToolTipText     =   "Editar(Ctrl+E)"
            ImageKey        =   "Editar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Borrar"
            Object.ToolTipText     =   "Borrar (Ctrl+D)"
            ImageKey        =   "Borrar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir (Ctrl+I)"
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Recursos"
            Object.ToolTipText     =   "Recursos(Ctrl+A)"
            ImageKey        =   "Seleccion"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (Ctrl+C)"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   3840
      Top             =   120
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
            Picture         =   "frmAusenciaMecanicos.frx":1AEE
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":1C00
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":1D12
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":1E24
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":1F36
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":2048
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":215A
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":226C
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":237E
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":2490
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":25A2
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":26B4
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":27C6
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":28D8
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":29EA
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":2AFC
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":2C0E
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":3060
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":34B2
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":35C4
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":3720
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":387C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":39D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":3B34
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":4600
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":4A54
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":4BB8
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":5014
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":5170
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":647C
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":6A18
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":6B74
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":6CD0
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":7024
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":7378
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":76CC
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":7A20
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":7D74
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":82B8
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":83CA
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":871E
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":8A72
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":8B86
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":8EDA
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":8FEC
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAusenciaMecanicos.frx":9340
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAusenciaMecanicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoPrincipal As New ADODB.Recordset
Dim AdoTemp As New ADODB.Recordset
Dim mstrSql As String
Dim mblnTablaVacia As Boolean
Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean
Dim mblnSW As Boolean
Dim Item As ListItem
Dim mstrWhere As String

Private Sub GenerarRecursos()
    'valida parametros para generar el procedimiento almacenado
    If Me.dtcSucursal.Text = "" Then
        MsgBox "La Sucursal debe contener un valor", vbInformation, "Genera Recursos"
        Me.dtcSucursal.SetFocus
        Exit Sub
    End If
    If Me.dtcMotivo.Text = "" Then
        MsgBox "El Turno debe contener un valor", vbInformation, "Genera Recursos"
        Me.dtcSucursal.SetFocus
        Exit Sub
    End If
    
    mstrSql = "Exec Tllr_Genera_Hora_Recursos " & "'" & gstrIdEmpresa & "','" & Me.dtcSucursal.BoundText & "','" & Me.dtcMotivo.BoundText & "','" & Me.pckFechaDesde & "','" & Me.pckFechaHasta & "','" & gstrUsuario & "','" & Format(Date, "DD/MM/YYYY") & "'"
    
    Screen.MousePointer = vbHourglass
    Conexion.SendHost mstrSql, , , , gcTiempoEspera
    Screen.MousePointer = vbDefault
    MsgBox "Proceso finalizado Exitosamente", vbInformation, "Genera Hora/Recursos"
End Sub

Sub CargaSucursal()
mstrSql = "SELECT Id_Sucursal AS CODIGO, Descripcion AS NOMBRE FROM Glbl_Sucursal where VIGENCIA = 'S' And Id_Empresa='" & gstrIdEmpresa & "'"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With datSucursal
        Set .Recordset = adoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcSucursal.ListField = "Nombre"
            dtcSucursal.BoundColumn = "Codigo"
        End If
    End With
End If
Set adoPrincipal = New ADODB.Recordset
Conexion.CloseHost adoPrincipal
End Sub

Sub CargaMotivo()
mstrSql = "SELECT Id_Ausencia AS CODIGO, Descripcion AS NOMBRE FROM Tllr_Motivo_Ausencia where VIGENCIA = 'S'"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With datMotivo
        Set .Recordset = adoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcMotivo.ListField = "Nombre"
            dtcMotivo.BoundColumn = "Codigo"
        End If
    End With
End If
Set adoPrincipal = New ADODB.Recordset
Conexion.CloseHost adoPrincipal
End Sub
Sub CargaMecanicos()
gstrSql = "SELECT Id_Mecanico AS Codigo, Nombre FROM Tllr_Mecanicos where Id_Empresa='" & gstrIdEmpresa & "' and Id_Sucursal ='" & gstrIdSucursal & "' and vigencia='S'"
If Conexion.SendHost(gstrSql, gadoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
With datSupervisor
    Set .Recordset = gadoPrincipal
    If Not .Recordset.BOF And Not .Recordset.EOF Then
        .Recordset.MoveFirst
        dtcSupervisor.ListField = "Nombre"
        dtcSupervisor.BoundColumn = "Codigo"
    End If
End With
End If
Set gadoPrincipal = New ADODB.Recordset
Conexion.CloseHost gadoPrincipal
End Sub

Private Sub cmdCancelaMecanico_Click()
Me.dtcSupervisor.Text = ""
End Sub

Private Sub cmdCancelaMotivo_Click()
Me.dtcMotivo.Text = ""
End Sub

Private Sub cmdCancelaSuc_Click()
Me.dtcSucursal.Text = ""
End Sub

Private Sub Form_Load()
    mblnSW = True
End Sub

Private Sub lvwConceptos_DblClick()
If Me.lvwConceptos.ListItems.Count > 0 Then
    With frmEditaMotivoAusencia
        .Tag = "Editar"
        .Show vbModal
    End With
End If
End Sub

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Screen.MousePointer = vbHourglass
    Select Case Button.Key
        Case "Crear"
            AgregarRegistro
        Case "Buscar"
            BuscarRegistro
        Case "Editar"
            EditarRegistro
        Case "Borrar"
            BorrarRegistro
        Case "Imprimir"
            ImprimirInforme
        Case "Cerrar"
            CerrarSalir
        Case "Recursos"
            GenerarRecursos
    End Select
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Activate()
    If mblnSW Then
        If Not Atributos("Glbl", "Tllr_20_0110", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
        
        CargaSucursal
        CargaMotivo
        CargaMecanicos
        
        pckFechaDesde = BOM(Date)
        pckFechaHasta = EOM(Date)

    End If
    mblnSW = False
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{tab}"
        Case vbKeyEscape
            KeyAscii = 0
            'CancelarAgregaRegistro
        Case 14 And tlbBarraHerramientas.Buttons.Item("Crear").Enabled
            KeyAscii = 0
            'AgregarRegistro
        Case 4 And tlbBarraHerramientas.Buttons.Item("Borrar").Enabled
            KeyAscii = 0
            'BorrarRegistro
        Case 2 And tlbBarraHerramientas.Buttons.Item("Buscar").Enabled
            KeyAscii = 0
            'BuscarRegistro
        Case 9 And tlbBarraHerramientas.Buttons.Item("Imprimir").Enabled
            KeyAscii = 0
            'ImprimirInforme
        Case 3 And tlbBarraHerramientas.Buttons.Item("Cerrar").Enabled
            KeyAscii = 0
            'CerrarSalir
    End Select
End Sub
Private Sub AgregarRegistro()
    Screen.MousePointer = vbDefault
    With frmEditaMotivoAusencia
        .Tag = "Crear"
        .Caption = "Nuevo"
        .Show vbModal
    End With
End Sub
Private Sub EditarRegistro()
    If Me.lvwConceptos.ListItems.Count > 0 Then
    Screen.MousePointer = vbDefault
    With frmEditaMotivoAusencia
        .Tag = "Editar"
        .Caption = "Editar"
        .Show vbModal
    End With
    End If
End Sub

Private Sub BorrarRegistro()
    If Me.lvwConceptos.ListItems.Count > 0 Then
        Screen.MousePointer = vbDefault
        If MsgBox("¿ Desea eliminar este registro ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
            'elimina
            mstrSql = "DELETE FROM TLLR_Mecanicos_Ausencias WHERE Id_Empresa ='" & gstrIdEmpresa & "' And  Id_sucursal='" & Me.lvwConceptos.SelectedItem.SubItems(9) & "' And Id_Ausencia='" & Me.lvwConceptos.SelectedItem.SubItems(10) & "' And Id_Mecanico='" & Me.lvwConceptos.SelectedItem.SubItems(2) & "' and Id_item=" & Me.lvwConceptos.SelectedItem.SubItems(8) & " And Id_Fecha='" & Me.lvwConceptos.SelectedItem.SubItems(4) & "'"
            Conexion.SendHost mstrSql, , , , gcTiempoEspera '//////////AQUI BORRA
            
            'lista
            If Not Me.lvwConceptos.SelectedItem Is Nothing Then
                Me.lvwConceptos.ListItems.Remove Me.lvwConceptos.SelectedItem.Index
            End If
        End If
    End If
    Conexion.CloseHost adoPrincipal
End Sub
Private Sub BuscarRegistro()
    
mstrWhere = ""
Me.lvwConceptos.ListItems.Clear

mstrWhere = " Where Tllr_Mecanicos_Ausencias.Id_Fecha Between '" & Me.pckFechaDesde & "' And '" & Me.pckFechaHasta & "'"

If Me.dtcSucursal.Text <> "" Then  '////////// sucursal
    mstrWhere = mstrWhere & " and Tllr_Mecanicos_Ausencias.Id_Sucursal ='" & Me.dtcSucursal.BoundText & "'"
End If

If Me.dtcMotivo.Text <> "" Then  '////////// sucursal
    mstrWhere = mstrWhere & " and Tllr_Mecanicos_Ausencias.Id_Ausencia ='" & Me.dtcMotivo.BoundText & "'"
End If

If Me.dtcSupervisor.Text <> "" Then  '////////// mecanico
    mstrWhere = mstrWhere & " and Tllr_Mecanicos_Ausencias.Id_Mecanico ='" & Me.dtcSupervisor.BoundText & "'"
End If

mstrSql = "SELECT Tllr_Mecanicos_Ausencias.Id_Sucursal, Glbl_Sucursal.Descripcion AS Sucursal, Tllr_Mecanicos_Ausencias.Id_Ausencia, "
mstrSql = mstrSql & "Tllr_Motivo_Ausencia.Descripcion AS Ausencia, Tllr_Mecanicos_Ausencias.Id_Mecanico, Tllr_Mecanicos.Nombre, "
mstrSql = mstrSql & "Tllr_Mecanicos_Ausencias.Id_Fecha, Tllr_Mecanicos_Ausencias.Id_Item, Tllr_Mecanicos_Ausencias.Hora_Desde, "
mstrSql = mstrSql & "Tllr_Mecanicos_Ausencias.Hora_Hasta , Tllr_Mecanicos_Ausencias.Total_Horas "
mstrSql = mstrSql & "FROM Tllr_Mecanicos_Ausencias LEFT OUTER JOIN "
mstrSql = mstrSql & "Tllr_Mecanicos ON Tllr_Mecanicos_Ausencias.Id_Empresa = Tllr_Mecanicos.Id_Empresa AND "
mstrSql = mstrSql & "Tllr_Mecanicos_Ausencias.Id_Sucursal = Tllr_Mecanicos.Id_Sucursal AND "
mstrSql = mstrSql & "Tllr_Mecanicos_Ausencias.Id_Mecanico = Tllr_Mecanicos.Id_Mecanico LEFT OUTER JOIN "
mstrSql = mstrSql & "Tllr_Motivo_Ausencia ON Tllr_Mecanicos_Ausencias.Id_Ausencia = Tllr_Motivo_Ausencia.Id_Ausencia LEFT OUTER JOIN "
mstrSql = mstrSql & "Glbl_Sucursal ON Tllr_Mecanicos_Ausencias.Id_Empresa = Glbl_Sucursal.Id_Empresa AND "
mstrSql = mstrSql & "Tllr_Mecanicos_Ausencias.Id_Sucursal = Glbl_Sucursal.Id_Sucursal "
mstrSql = mstrSql & mstrWhere

If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveLast: .MoveFirst
            While Not .EOF
                Set Item = lvwConceptos.ListItems.Add(, , !Sucursal)
                Item.SubItems(1) = !Ausencia
                Item.SubItems(2) = !Id_Mecanico
                Item.SubItems(3) = !Nombre
                Item.SubItems(4) = !Id_Fecha
                Item.SubItems(5) = Mid(!Hora_Desde, 12, 5)
                Item.SubItems(6) = Mid(!Hora_Hasta, 12, 5)
                Item.SubItems(7) = FormatoValor(!Total_Horas, "", 2)
                Item.SubItems(8) = !Id_Item
                Item.SubItems(9) = !Id_Sucursal
                Item.SubItems(10) = !Id_Ausencia
                .MoveNext
            Wend
        End If
    End With
End If ' por el otro
Set adoPrincipal = New ADODB.Recordset
Conexion.CloseHost adoPrincipal

End Sub
Private Sub ImprimirInforme()
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
    
    If Me.lvwConceptos.ListItems.Count = 0 Then
      MsgBox "No existen elementos en la lista", vbExclamation, "Imprimir"
      Exit Sub
    End If

    Screen.MousePointer = 11
    Dim wrkPredeterminado As Workspace
    Dim prpBucle As Property
    Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
    If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
    Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (Sucursal text,Ausencia text,CodMecanico text,Mecanico text,Fecha date,HoraDesde text,HoraHasta text,TotalHoras double)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To Me.lvwConceptos.ListItems.Count
        Set Me.lvwConceptos.SelectedItem = Me.lvwConceptos.ListItems(i)
        Tabla.AddNew
        Tabla!Sucursal = IIf(Me.lvwConceptos.SelectedItem = "", " ", Me.lvwConceptos.SelectedItem)
        Tabla!Ausencia = IIf(Me.lvwConceptos.SelectedItem.SubItems(1) = "", " ", Me.lvwConceptos.SelectedItem.SubItems(1))
        Tabla!CodMecanico = IIf(Me.lvwConceptos.SelectedItem.SubItems(2) = "", " ", Me.lvwConceptos.SelectedItem.SubItems(2))
        Tabla!Mecanico = IIf(Me.lvwConceptos.SelectedItem.SubItems(3) = "", " ", Me.lvwConceptos.SelectedItem.SubItems(3))
        Tabla!Fecha = IIf(Me.lvwConceptos.SelectedItem.SubItems(4) = "", " ", Me.lvwConceptos.SelectedItem.SubItems(4))
        Tabla!HoraDesde = IIf(Me.lvwConceptos.SelectedItem.SubItems(5) = "", " ", Me.lvwConceptos.SelectedItem.SubItems(5))
        Tabla!HoraHasta = IIf(Me.lvwConceptos.SelectedItem.SubItems(6) = "", " ", Me.lvwConceptos.SelectedItem.SubItems(6))
        Tabla!TotalHoras = IIf(Me.lvwConceptos.SelectedItem.SubItems(7) = "", " ", Me.lvwConceptos.SelectedItem.SubItems(7))
        Tabla.Update
    Next i
   Tabla.Close
   
   With rptMantenedor
        .ReportFileName = gstrPathReporte & "\MotivoAusencia.rpt"
        .WindowTitle = "Reporte de Ausencia de Mecanicos"
        .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='Listado de Ausencias de Mecanicos'"
        .Formulas(2) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(3) = "Sucursal='" & gstrSucursal & "'"
        .Formulas(4) = "Direccion='" & gstrDirSuc & "'"
        .Formulas(5) = "Desde='" & Me.pckFechaDesde & "'"
        .Formulas(6) = "Hasta='" & Me.pckFechaHasta & "'"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = True
   End With
   
   Dbsnueva.Close
   Screen.MousePointer = 1
End Sub
Private Sub CerrarSalir()
    Unload Me
End Sub
Private Sub Ayuda()
End Sub
Private Sub ActivaBotones()
    With tlbBarraHerramientas.Buttons
        .Item("Crear").Enabled = IIf(mblnAccesoCrear, True, False)
        .Item("Grabar").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoEditar, True, False))
        .Item("Cancelar").Enabled = False
        .Item("Borrar").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoBorrar, True, False))
        .Item("Buscar").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Imprimir").Enabled = IIf(mblnTablaVacia, False, IIf(mblnAccesoImprimir, True, False))
        .Item("Primero").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Anterior").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Siguiente").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Ultimo").Enabled = IIf(mblnTablaVacia, False, True)
        .Item("Renovar").Enabled = True
        .Item("Cerrar").Enabled = True
    End With
End Sub
Private Sub DesactivaBotones()
    With tlbBarraHerramientas.Buttons
        .Item("Crear").Enabled = False
        .Item("Grabar").Enabled = mblnAccesoEditar Or mblnAccesoCrear
        .Item("Cancelar").Enabled = True
        .Item("Borrar").Enabled = False
        .Item("Buscar").Enabled = False
        .Item("Imprimir").Enabled = False
        .Item("Primero").Enabled = False
        .Item("Anterior").Enabled = False
        .Item("Siguiente").Enabled = False
        .Item("Ultimo").Enabled = False
        .Item("Renovar").Enabled = False
        .Item("Cerrar").Enabled = True
    End With
End Sub
