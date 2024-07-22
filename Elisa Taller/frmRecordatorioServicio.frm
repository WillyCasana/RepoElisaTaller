VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRecordatorioServicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recordatorio de Servicio"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13710
   Icon            =   "frmRecordatorioServicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   13710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   9930
      TabIndex        =   16
      Top             =   8400
      Width           =   1680
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "Seleccionar"
      Height          =   360
      Left            =   8160
      TabIndex        =   15
      Top             =   8400
      Width           =   1680
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   13335
      Begin VB.CommandButton cmdCancelaSuc 
         Height          =   330
         Left            =   6960
         Picture         =   "frmRecordatorioServicio.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdLimpiaEmpresa 
         Height          =   315
         Left            =   2880
         Picture         =   "frmRecordatorioServicio.frx":048C
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Limpia filtro por Empresa"
         Top             =   480
         Width           =   315
      End
      Begin VB.CommandButton cmdLimpiaFecha1 
         Height          =   315
         Left            =   9720
         Picture         =   "frmRecordatorioServicio.frx":09BE
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Limpia Fecha de Inicio"
         Top             =   480
         Width           =   315
      End
      Begin VB.CommandButton cmdLimpia2 
         Height          =   315
         Left            =   12240
         Picture         =   "frmRecordatorioServicio.frx":0EF0
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Limpia Fecha de Término"
         Top             =   480
         Width           =   315
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   315
         Left            =   10560
         TabIndex        =   4
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
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
         Format          =   50724865
         CurrentDate     =   36772
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   315
         HelpContextID   =   285
         Left            =   8040
         TabIndex        =   5
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
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
         Format          =   50724865
         CurrentDate     =   36772
      End
      Begin MSDataListLib.DataCombo dbcboEmpresa 
         Bindings        =   "frmRecordatorioServicio.frx":1422
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   -2147483640
         ListField       =   "Razon_Social"
         BoundColumn     =   "id_Empresa"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc datEmpresa 
         Height          =   375
         Left            =   1080
         Top             =   480
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   2
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
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
      Begin MSDataListLib.DataCombo dtcSucursal 
         Bindings        =   "frmRecordatorioServicio.frx":143B
         Height          =   315
         Left            =   4320
         TabIndex        =   13
         Top             =   480
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
         Left            =   4560
         Top             =   480
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
      Begin VB.Label Label1 
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   3360
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblEmpresa 
         Caption         =   "Empresa"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha Inicio"
         Height          =   255
         Left            =   8040
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Final"
         Height          =   255
         Left            =   10560
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lsvdetalle 
      Height          =   6420
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   11324
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ITEM"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "PLACA"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "MODELO"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CLIENTE"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "TELEFONO"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "CELULAR"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "E-MAIL"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "FECHA PROX. CITA"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ESTADO"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   600
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   36
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":1455
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":1567
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":1679
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":178B
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":189D
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":19AF
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":1AC1
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":1BD3
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":1CE5
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":1DF7
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":1F09
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":201B
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":212D
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":223F
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":2351
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":2463
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":2575
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":29C7
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":2E19
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":2F2B
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":3087
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":31E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":333F
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":349B
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":3F67
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":43BB
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":451F
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":497B
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":4AD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":5DE3
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":637F
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":64DB
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":6637
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":698B
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":6CDF
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecordatorioServicio.frx":7033
            Key             =   "outlook"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdExportar 
      Left            =   1920
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport rptKardex 
      Left            =   2640
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComctlLib.Toolbar BarraHerramientas 
      Height          =   330
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   15690
      _ExtentX        =   27675
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
            Object.ToolTipText     =   "Traer Datos"
            ImageKey        =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Imprimir"
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Excel"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc datDatos 
      Height          =   330
      Left            =   4320
      Top             =   7080
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
      Caption         =   "datDatos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRecordatorioServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'kjcv 31.10.14 Creacion de Formulario Recordatorio de Servicio
Dim AdoRecordSucursal As New ADODB.Recordset
Dim AdoRecordEmpresa As New ADODB.Recordset
Dim adoRecordset As New ADODB.Recordset
Dim mstrSQL As String
Dim adoPrincipal As New ADODB.Recordset
Dim Item As ListItem
Dim mblSW As Boolean
Dim mblSWRecordatorio As Boolean

Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean

Private Sub BarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
Screen.MousePointer = vbHourglass
Select Case Button.Key
    Case "Buscar"
        BuscarRegistro
    Case "Imprimir"
'        ImprimirConsulta
    Case "Excel"
        ExportarDatos Me.lsvdetalle, Me.cdExportar, Me.hwnd
    Case "Cerrar"
        Unload Me
End Select
Screen.MousePointer = vbDefault
End Sub
Sub CargaSucursal()
mstrSQL = "SELECT Id_Sucursal AS CODIGO, Descripcion AS NOMBRE FROM Glbl_Sucursal where VIGENCIA = 'S' And Id_Empresa='" & gstrIdEmpresa & "'"
If Conexion.SendHost(mstrSQL, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
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

Private Sub cmdSeleccionar_Click()
If Not lsvdetalle.SelectedItem Is Nothing Then
    gstrBusca = lsvdetalle.SelectedItem
End If
Unload Me
End Sub

Private Sub Form_Activate()
Dim blnBoolean As Boolean
    If mblSW Then
        mblSW = False
        
        Screen.MousePointer = vbDefault
        
        If Not Atributos("Glbl", "Tllr_20_0071", True, False, False, False) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
       
    End If
End Sub
Public Sub ListView_ColorearLinea(LaLista As ListView, linea As Long, Color As Long)
Dim x As Integer

'Verifico si la linea que quiere modificar existe
If linea > LaLista.ListItems.Count Then
    Exit Sub
End If

'modifico el color de la primer columna
LaLista.ListItems(linea).ForeColor = Color

'modifico el color de las demas columnas
For x = 1 To LaLista.ColumnHeaders.Count - 1
    'verifico que el subitem tenga algo escrito, por que si no tiene nada tira
    'error de "subindice fuera de intervalo"
    If Trim(LaLista.SelectedItem.SubItems(x) <> "") Then
        LaLista.ListItems(linea).ListSubItems(x).ForeColor = Color
    End If
Next x

'actualizo el list para que se vean los cambios
LaLista.Refresh
End Sub


Private Sub ColorearLista()
Dim i As Integer
Dim FechaInicio As Date
Dim fechaEvalua As Variant
Dim lstrFecha As Variant

FechaInicio = Date

fechaEmpieza = DateAdd("d", 7, FechaInicio)
fechafinaliza = DateAdd("d", 7, fechaEmpieza)


For i = 1 To Me.lsvdetalle.ListItems.Count

    fechaEvalua = Me.lsvdetalle.ListItems.Item(i).SubItems(7)
    If fechaEmpieza <= CDate(fechaEvalua) And fechafinaliza >= CDate(fechaEvalua) Then
     
        Me.lsvdetalle.ListItems.Item(i).ForeColor = vbBlue
        For x = 1 To lsvdetalle.ColumnHeaders.Count - 1

            If Trim(lsvdetalle.SelectedItem.SubItems(x) <> "") Then
                lsvdetalle.ListItems(i).ListSubItems(x).ForeColor = vbBlue
            End If
        Next x

    End If
    

Next i






End Sub

Private Sub Form_Load()
mblSW = True

mblSWRecordatorio = True

 dtpDesde = Format(Date, "dd/mm/yyyy")
 dtpHasta = EOM(Date)
 
  'Llena Empresa
mstrSQL = "SELECT Id_Empresa,Razon_Social FROM Glbl_Empresa WHERE Vigencia = 'S' ORDER BY Razon_Social"
 If Conexion.SendHost(mstrSQL, AdoRecordEmpresa, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    Set Me.datEmpresa.Recordset = AdoRecordEmpresa
 End If
  Me.dbcboEmpresa.BoundText = gstrIdEmpresa
CargaSucursal
Me.dtcSucursal.BoundText = gstrIdSucursal
End Sub

'Private Sub SeteaSpread()
'Dim lintHojas As Integer
'
'With Me.sprGrillaPrincipal
'    .Reset
'
'    ' crea las hojitas (sheets)
'    .SheetCount = 1
'    .Sheet = 1
'    .SheetName = "Consula Cotizaciones"
'
'    For lintHojas = 1 To .SheetCount
'        .Sheet = lintHojas
'        .ActiveSheet = lintHojas
'        SeteaSpreadSoloHoja Me.sprGrillaPrincipal
'    Next lintHojas
'End With
'
'End Sub

'Private Sub SeteaSpreadPost()
'Dim lintHojas As Integer
'Dim ldblCol As Double
'
'Screen.MousePointer = vbHourglass
'
'Me.sprGrillaPrincipal.Redraw = False
'
'' hace el seteo en cada hoja (sheet)
'For lintHojas = 1 To Me.sprGrillaPrincipal.SheetCount
'    Me.sprGrillaPrincipal.Sheet = lintHojas
'    Me.sprGrillaPrincipal.ActiveSheet = lintHojas
'    SeteaSpreadPostSoloHoja Me.sprGrillaPrincipal, 7
'Next lintHojas
'
'Me.sprGrillaPrincipal.Redraw = True
'
'Me.sprGrillaPrincipal.Row = -1
'
'ldblCol = TraeNumColSpread(Me.sprGrillaPrincipal, "FECHACITA")
'Me.sprGrillaPrincipal.Col = ldblCol
'
'Me.sprGrillaPrincipal.BackColor = &HC0FFC0
'
'Screen.MousePointer = vbDefault
'
'End Sub

Sub BuscarRegistro()
Dim strNumItem As Integer
Dim FechaInicio As Date
Dim fechaHasta As Date
Dim fechaEvalua As Date

FechaInicio = Date

fechaHasta = DateAdd("d", 7, FechaInicio)


mstrSQL = "EXEC Tllr_Reporte_RecordatorioServicio  '" & gstrIdEmpresa & "', '" & Format(Me.dtpDesde.Value, "dd/mm/yyyy") & "' , '" & Format(Me.dtpHasta.Value, "dd/mm/yyyy") & " 23:59:00" & "'"

Me.lsvdetalle.ListItems.Clear
If Conexion.SendHost(mstrSQL, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                    adoPrincipal.MoveFirst
    End If
            Do Until adoPrincipal.EOF
                strNumItem = strNumItem + 1
            Set Item = Me.lsvdetalle.ListItems.Add(, , strNumItem)
                Item.SubItems(1) = adoPrincipal!Placa
                Item.SubItems(2) = ValorNulo(adoPrincipal!Modelo)
                Item.SubItems(3) = ValorNulo(adoPrincipal!Cliente)
                Item.SubItems(4) = ValorNulo(adoPrincipal!Telefono)
                Item.SubItems(5) = ValorNulo(adoPrincipal!Celular)
                Item.SubItems(6) = ValorNulo(adoPrincipal!Email)
                Item.SubItems(7) = adoPrincipal!FechaCita
                               
                adoPrincipal.MoveNext
            Loop

End If

ColorearLista

End Sub

Private Sub lsvdetalle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ReOrdenaLista lsvdetalle, ColumnHeader
End Sub

Private Sub lsvdetalle_DblClick()
Dim lintFila As Integer
Dim lblModelo As String
Dim lblCliente As String
Dim lblFono As String
Dim lblFechaCita As String

    If Me.lsvdetalle.ListItems.Count <= 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

        lintFila = lsvdetalle.SelectedItem.Index
        Set lsvdetalle.SelectedItem = lsvdetalle.ListItems(lintFila)
                        
        
        Load frmReservadeHoras
        frmReservadeHoras.Visible = False
        frmReservadeHoras.Show
        
'        frmReservadeHoras.tlbBarraHerramientas.Buttons.Item("Crear").Value = tbrPressed
        swActivateRecorda = Atributos("Glbl", "Tllr_20_0070", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir)
        frmReservadeHoras.AgregarRegistro

        
        lblModelo = lsvdetalle.SelectedItem.SubItems(2)
        lblCliente = lsvdetalle.SelectedItem.SubItems(3)
        lblFono = lsvdetalle.SelectedItem.SubItems(4)
        lblFechaCita = lsvdetalle.SelectedItem.SubItems(7)
        
        frmReservadeHoras.lblModelo = lblModelo
        frmReservadeHoras.lblCliente = lblCliente
        frmReservadeHoras.lblFono = lblFono
        frmReservadeHoras.pckFechaEntrega = lblFechaCita
        frmReservadeHoras.optSinPatente.Value = True
        frmReservadeHoras.mblnSW = False
        
'        Unload Me
        
'        lstrNotaVenta = lvwListaVehiculos.SelectedItem.SubItems(6)
'        'kjcv 25.09.12
'        lstrIdSucu = lvwListaVehiculos.SelectedItem.SubItems(48)
'        TablaVenta.idsucursal = lstrIdSucu
''        TablaVenta.IdSucursal = lvwListaVehiculos.SelectedItem.SubItems(47)
'        TablaVenta.TipoDocto = "V"
'        TablaVenta.id_Vendedor = lvwListaVehiculos.SelectedItem.SubItems(85)
'        Load frmVentas
'        frmVentas.Visible = False
'        frmVentas.Show
'        'kjcv 25.09.12
'        frmVentas.Cancelar
'        frmVentas.TraeVenta CDbl(lstrNotaVenta), lstrIdSucu
'        frmVentas.mSwActivate = False

End Sub

Public Sub CargaDatos()

End Sub

