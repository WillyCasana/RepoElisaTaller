VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Begin VB.Form frmAsignacionTurnos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignacion de Turnos"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   Icon            =   "frmAsignacionTurnos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9255
      Begin VB.CommandButton cmdCancelaTurno 
         Height          =   330
         Left            =   3840
         Picture         =   "frmAsignacionTurnos.frx":179A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdCancelaSuc 
         Height          =   330
         Left            =   3840
         Picture         =   "frmAsignacionTurnos.frx":189C
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
            Format          =   178913281
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
            Format          =   98304001
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
         Bindings        =   "frmAsignacionTurnos.frx":199E
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
         Top             =   1500
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Codigo"
            Text            =   "Sucursal"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Turno"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "Descripcion"
            Text            =   "Cod. Mecánico"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "DP"
            Text            =   "Mecánico"
            Object.Width           =   5468
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "Orden"
            Text            =   "Desde"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Hasta"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "IdItem"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "idSucursal"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "idturno"
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
      Begin MSDataListLib.DataCombo dtcTurnos 
         Bindings        =   "frmAsignacionTurnos.frx":19B8
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   960
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
      Begin MSAdodcLib.Adodc datTurnos 
         Height          =   330
         Left            =   1920
         Top             =   960
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
      Begin VB.Label Label2 
         Caption         =   "Turnos"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
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
         NumListImages   =   46
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":19D0
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":1AE2
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":1BF4
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":1D06
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":1E18
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":1F2A
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":203C
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":214E
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":2260
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":2372
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":2484
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":2596
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":26A8
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":27BA
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":28CC
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":29DE
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":2AF0
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":2F42
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":3394
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":34A6
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":3602
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":375E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":38BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":3A16
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":44E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":4936
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":4A9A
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":4EF6
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":5052
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":635E
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":68FA
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":6A56
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":6BB2
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":6F06
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":725A
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":75AE
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":7902
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":7C56
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":819A
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":82AC
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":8600
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":8954
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":8A68
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":8DBC
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":8ECE
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAsignacionTurnos.frx":9222
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAsignacionTurnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AdoPrincipal As New ADODB.Recordset
Dim adoTemp As New ADODB.Recordset
Dim mstrSQL As String
Dim mblnTablaVacia As Boolean
Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean
Dim mblnSW As Boolean
Dim Item As ListItem
Dim mstrWhere As String

Sub CargaSucursal()
mstrSQL = "SELECT Id_Sucursal AS CODIGO, Descripcion AS NOMBRE FROM Glbl_Sucursal where TieneTaller='S' and VIGENCIA = 'S' And Id_Empresa='" & gstrIdEmpresa & "'"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With datSucursal
        Set .Recordset = AdoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcSucursal.ListField = "Nombre"
            dtcSucursal.BoundColumn = "Codigo"
        End If
    End With
End If
Set AdoPrincipal = New ADODB.Recordset
Conexion.CloseHost AdoPrincipal
End Sub

Sub CargaTurnos()
mstrSQL = "SELECT Id_Turno AS CODIGO, Descripcion AS NOMBRE FROM Tllr_Turnos where VIGENCIA = 'S' And Id_Empresa='" & gstrIdEmpresa & "'"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With datTurnos
        Set .Recordset = AdoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcTurnos.ListField = "Nombre"
            dtcTurnos.BoundColumn = "Codigo"
        End If
    End With
End If
Set AdoPrincipal = New ADODB.Recordset
Conexion.CloseHost AdoPrincipal
End Sub

Private Sub cmdCancelaSuc_Click()
Me.dtcSucursal.Text = ""
End Sub

Private Sub cmdCancelaTurno_Click()
Me.dtcTurnos.Text = ""
End Sub

Private Sub Form_Load()
    mblnSW = True
End Sub

Private Sub lvwConceptos_DblClick()
If Me.lvwConceptos.ListItems.Count > 0 Then
    With frmEditaAsignacionRecursos
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
        If Not Atributos("Glbl", "Tllr_20_0100", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
        
        CargaSucursal
        CargaTurnos
        
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
    With frmEditaAsignacionRecursos
        .Tag = "Crear"
        .Caption = "Nuevo"
        .Show vbModal
    End With
End Sub
Private Sub EditarRegistro()
    If Me.lvwConceptos.ListItems.Count > 0 Then
    Screen.MousePointer = vbDefault
    With frmEditaAsignacionRecursos
        .Tag = "Editar"
        .Show vbModal
    End With
    End If
End Sub

Private Sub BorrarRegistro()
    If Me.lvwConceptos.ListItems.Count > 0 Then
        Screen.MousePointer = vbDefault
        If MsgBox("¿ Desea eliminar este registro ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
            'elimina
            mstrSQL = "DELETE FROM TLLR_Mecanicos_Turnos WHERE Id_Empresa ='" & gstrIdEmpresa & "' And  Id_sucursal='" & Me.lvwConceptos.SelectedItem.SubItems(7) & "' And Id_turno='" & Me.lvwConceptos.SelectedItem.SubItems(8) & "' And Id_Mecanico='" & Me.lvwConceptos.SelectedItem.SubItems(2) & "' and Id_item=" & Me.lvwConceptos.SelectedItem.SubItems(6)
            Conexion.SendHost mstrSQL, , , , gcTiempoEspera '//////////AQUI BORRA
            
            'lista
            If Not Me.lvwConceptos.SelectedItem Is Nothing Then
                Me.lvwConceptos.ListItems.Remove Me.lvwConceptos.SelectedItem.Index
            End If
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub BuscarRegistro()
    
mstrWhere = ""
Me.lvwConceptos.ListItems.Clear

mstrWhere = " Where Tllr_Mecanicos_Turnos.Fecha_Desde >='" & Me.pckFechaDesde & "' And Tllr_Mecanicos_Turnos.Fecha_Hasta <='" & Me.pckFechaHasta & "'"

If Me.dtcSucursal.Text <> "" Then  '////////// sucursal
    mstrWhere = mstrWhere & " and Tllr_Mecanicos_Turnos.Id_Sucursal ='" & Me.dtcSucursal.BoundText & "'"
End If

If Me.dtcTurnos.Text <> "" Then  '////////// turnos
    mstrWhere = mstrWhere & " and Tllr_Mecanicos_Turnos.Id_Turno ='" & Me.dtcTurnos.BoundText & "'"
End If

mstrSQL = "SELECT Tllr_Mecanicos_Turnos.Id_Sucursal, Glbl_Sucursal.Descripcion as Sucursal, Tllr_Mecanicos_Turnos.Id_Mecanico, Tllr_Mecanicos.Nombre, "
mstrSQL = mstrSQL & "Tllr_Mecanicos_Turnos.Id_Turno, Tllr_Turnos.Descripcion AS Turno, Tllr_Mecanicos_Turnos.Id_Item, "
mstrSQL = mstrSQL & "Tllr_Mecanicos_Turnos.Fecha_Desde , Tllr_Mecanicos_Turnos.Fecha_Hasta "
mstrSQL = mstrSQL & "FROM Tllr_Mecanicos_Turnos LEFT OUTER JOIN "
mstrSQL = mstrSQL & "Tllr_Mecanicos ON Tllr_Mecanicos_Turnos.Id_Mecanico = Tllr_Mecanicos.Id_Mecanico AND "
mstrSQL = mstrSQL & "Tllr_Mecanicos_Turnos.Id_Sucursal = Tllr_Mecanicos.Id_Sucursal AND "
mstrSQL = mstrSQL & "Tllr_Mecanicos_Turnos.Id_Empresa = Tllr_Mecanicos.Id_Empresa LEFT OUTER JOIN "
mstrSQL = mstrSQL & "Glbl_Sucursal ON Tllr_Mecanicos_Turnos.Id_Sucursal = Glbl_Sucursal.Id_Sucursal AND "
mstrSQL = mstrSQL & "Tllr_Mecanicos_Turnos.Id_Empresa = Glbl_Sucursal.Id_Empresa LEFT OUTER JOIN "
mstrSQL = mstrSQL & "Tllr_Turnos ON Tllr_Mecanicos_Turnos.Id_Turno = Tllr_Turnos.Id_Turno "
mstrSQL = mstrSQL & mstrWhere

If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With AdoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveLast: .MoveFirst
            While Not .EOF
                Set Item = lvwConceptos.ListItems.Add(, , !Sucursal)
                Item.SubItems(1) = !Turno
                Item.SubItems(2) = !Id_Mecanico
                Item.SubItems(3) = !Nombre
                Item.SubItems(4) = !Fecha_Desde
                Item.SubItems(5) = !Fecha_Hasta
                Item.SubItems(6) = !Id_Item
                Item.SubItems(7) = !Id_Sucursal
                Item.SubItems(8) = !Id_Turno
                .MoveNext
            Wend
        End If
    End With
End If ' por el otro
Set AdoPrincipal = New ADODB.Recordset
Conexion.CloseHost AdoPrincipal


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
    Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (Sucursal text,Turno text,CodMecanico text,Mecanico text,Fecdesde date,FecHasta date)"
    Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
    For i = 1 To Me.lvwConceptos.ListItems.Count
        Set Me.lvwConceptos.SelectedItem = Me.lvwConceptos.ListItems(i)
        Tabla.AddNew
        Tabla!Sucursal = IIf(Me.lvwConceptos.SelectedItem = "", " ", Me.lvwConceptos.SelectedItem)
        Tabla!Turno = IIf(Me.lvwConceptos.SelectedItem.SubItems(1) = "", " ", Me.lvwConceptos.SelectedItem.SubItems(1))
        Tabla!CodMecanico = IIf(Me.lvwConceptos.SelectedItem.SubItems(2) = "", " ", Me.lvwConceptos.SelectedItem.SubItems(2))
        Tabla!Mecanico = IIf(Me.lvwConceptos.SelectedItem.SubItems(3) = "", " ", Me.lvwConceptos.SelectedItem.SubItems(3))
        Tabla!FecDesde = IIf(Me.lvwConceptos.SelectedItem.SubItems(4) = "", " ", Me.lvwConceptos.SelectedItem.SubItems(4))
        Tabla!FecHasta = IIf(Me.lvwConceptos.SelectedItem.SubItems(5) = "", " ", Me.lvwConceptos.SelectedItem.SubItems(5))
        Tabla.Update
    Next i
   Tabla.Close
   
   With rptMantenedor
        .ReportFileName = gstrPathReporte & "\AsignacionRecursos.rpt"
        .WindowTitle = "Reporte de Asignacion de Turnos de Trabajo"
        .DataFiles(0) = GcamBaseTem & "\BDNueva.mdb"
        .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
        .Formulas(1) = "TITULO='Listado de Asignacion de Turnos'"
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
Private Sub GenerarRecursos()
'valida parametros para generar el procedimiento almacenado
If Me.dtcSucursal.Text = "" Then
    MsgBox "La Sucursal debe contener un valor", vbInformation, "Genera Recursos"
    Me.dtcSucursal.SetFocus
    Exit Sub
End If
If Me.dtcTurnos.Text = "" Then
    MsgBox "El Turno debe contener un valor", vbInformation, "Genera Recursos"
    Me.dtcSucursal.SetFocus
    Exit Sub
End If

mstrSQL = "Exec Tllr_Genera_Hora_Recursos " & "'" & gstrIdEmpresa & "','" & Me.dtcSucursal.BoundText & "','" & Me.dtcTurnos.BoundText & "','" & Me.pckFechaDesde & "','" & Me.pckFechaHasta & "','" & gstrUsuario & "','" & Format(Date, "DD/MM/YYYY") & "'"

Screen.MousePointer = vbHourglass
Conexion.SendHost mstrSQL, , , , gcTiempoEspera
Screen.MousePointer = vbDefault
MsgBox "Proceso finalizado Exitosamente", vbInformation, "Genera Hora/Recursos"
End Sub
