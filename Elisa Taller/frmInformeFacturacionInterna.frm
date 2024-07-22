VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmInformeFacturacionInterna 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Facturación Interna"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11970
   Icon            =   "frmInformeFacturacionInterna.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11655
      Begin VB.CommandButton cmdLimpiaEmpresa 
         Height          =   315
         Left            =   2880
         Picture         =   "frmInformeFacturacionInterna.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Limpia filtro por Empresa"
         Top             =   480
         Width           =   315
      End
      Begin VB.CommandButton cmdLimpiaFecha1 
         Height          =   315
         Left            =   8280
         Picture         =   "frmInformeFacturacionInterna.frx":08BC
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Limpia Fecha de Inicio"
         Top             =   480
         Width           =   315
      End
      Begin VB.CommandButton cmdLimpia2 
         Height          =   315
         Left            =   10800
         Picture         =   "frmInformeFacturacionInterna.frx":0DEE
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Limpia Fecha de Término"
         Top             =   480
         Width           =   315
      End
      Begin VB.CommandButton cmdlimpiar 
         Height          =   315
         Left            =   5880
         Picture         =   "frmInformeFacturacionInterna.frx":1320
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Limpiar"
         Top             =   480
         Width           =   315
      End
      Begin MSDataListLib.DataCombo cmbsucursal 
         Bindings        =   "frmInformeFacturacionInterna.frx":1422
         Height          =   315
         Left            =   3600
         TabIndex        =   5
         Top             =   480
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "Id_sucursal"
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
      Begin MSAdodcLib.Adodc adosucursal 
         Height          =   330
         Left            =   4080
         Top             =   480
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
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
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   315
         Left            =   9240
         TabIndex        =   6
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
         Format          =   32440321
         CurrentDate     =   36772
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   315
         HelpContextID   =   285
         Left            =   6720
         TabIndex        =   7
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
         Format          =   32440321
         CurrentDate     =   36772
      End
      Begin MSDataListLib.DataCombo dbcboEmpresa 
         Bindings        =   "frmInformeFacturacionInterna.frx":143C
         Height          =   315
         Left            =   240
         TabIndex        =   8
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
      Begin VB.Label lblEmpresa 
         Caption         =   "Empresa"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha Inicio"
         Height          =   255
         Left            =   6720
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Final"
         Height          =   255
         Left            =   9240
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Sucursal"
         Height          =   195
         Left            =   3600
         TabIndex        =   9
         Top             =   240
         Width           =   660
      End
   End
   Begin MSComctlLib.ListView lsvdetalle 
      Height          =   4980
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8784
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CENTRO COSTO"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "REPUESTOS"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "CARROCERIA"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "TERCEROS"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "MECANICA"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "OTROS"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   600
      Top             =   6840
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
            Picture         =   "frmInformeFacturacionInterna.frx":1455
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":1567
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":1679
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":178B
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":189D
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":19AF
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":1AC1
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":1BD3
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":1CE5
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":1DF7
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":1F09
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":201B
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":212D
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":223F
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":2351
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":2463
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":2575
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":29C7
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":2E19
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":2F2B
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":3087
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":31E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":333F
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":349B
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":3F67
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":43BB
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":451F
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":497B
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":4AD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":5DE3
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":637F
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":64DB
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":6637
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":698B
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":6CDF
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeFacturacionInterna.frx":7033
            Key             =   "outlook"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdExportar 
      Left            =   1920
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport rptKardex 
      Left            =   2640
      Top             =   6960
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
      TabIndex        =   14
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
End
Attribute VB_Name = "frmInformeFacturacionInterna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoRecordSucursal As New ADODB.Recordset
Dim AdoRecordEmpresa As New ADODB.Recordset
Dim adoRecordset As New ADODB.Recordset
Dim mstrSQL As String
Dim AdoPrincipal As New ADODB.Recordset
Dim Item As ListItem
Dim mblSW As Boolean

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

Private Sub cmdLimpia2_Click()
dtpHasta.Value = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub cmdLimpiaEmpresa_Click()
Me.dbcboEmpresa.Text = ""
End Sub

Private Sub cmdLimpiaFecha1_Click()
Me.dtpDesde.Value = FechaInicio()
End Sub

Private Sub cmdlimpiar_Click()
Me.cmbsucursal.Text = ""
End Sub

Private Sub Form_Activate()
Dim blnBoolean As Boolean
    If mblSW Then
        mblSW = False
        
        Screen.MousePointer = vbDefault
        'Solo habilitado para perfiles ADMIN,GG_01,Gg2
        If Not Atributos("Glbl", "Tllr_30_0085", True, True, True, True) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
       
    End If
End Sub

Function FechaInicio() As Date
    FechaInicio = "01/" & Format$(Date, "mm/yyyy")
End Function
Private Sub Form_Load()
 mblSW = True
 dtpHasta.Value = Format(Date, "dd/mm/yyyy")
 dtpDesde.Value = FechaInicio
 
  'Llena Empresa
mstrSQL = "SELECT Id_Empresa,Razon_Social FROM Glbl_Empresa WHERE Vigencia = 'S' ORDER BY Razon_Social"
 If Conexion.SendHost(mstrSQL, AdoRecordEmpresa, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    Set Me.datEmpresa.Recordset = AdoRecordEmpresa
 End If
  Me.dbcboEmpresa.BoundText = gstrIdEmpresa
 
 'Llena sucursal
mstrSQL = "Select Id_Sucursal, Descripcion From Glbl_Sucursal Where Id_Empresa ='" + gstrIdEmpresa + "' Order by Descripcion"
 If Conexion.SendHost(mstrSQL, AdoRecordSucursal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    Set Me.adosucursal.Recordset = AdoRecordSucursal
 End If
 Me.cmbsucursal.BoundText = gstrIdSucursal
  
End Sub

Sub BuscarRegistro()


'mstrSQL = " EXEC Tllr_OTs_Cargos_Internos @Id_Empresa='" & gstrIdEmpresa & "',@Id_Sucursal='" & Me.cmbsucursal.BoundText & "' , @FechaDesde='" & Format(Me.dtpDesde.Value, "dd/mm/yyyy") & "',@FechaHasta='" & Format(Me.dtpHasta.Value, "dd/mm/yyyy") & "'"
mstrSQL = "EXEC Tllr_Facturacion_Interna_CentroCosto  @Id_Empresa='" & gstrIdEmpresa & "',@Id_Sucursal='" & Me.cmbsucursal.BoundText & "' , @FechaDesde1='" & Format(Me.dtpDesde.Value, "dd/mm/yyyy") & "',@FechaHasta1='" & Format(Me.dtpHasta.Value, "dd/mm/yyyy") & "'"


Me.lsvdetalle.ListItems.Clear
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                    AdoPrincipal.MoveFirst
    End If
            Do Until AdoPrincipal.EOF
                strNumItem = strNumItem + 1
                
            Set Item = Me.lsvdetalle.ListItems.Add(, , strNumItem)
                Item.SubItems(1) = AdoPrincipal!Nombre
                Item.SubItems(2) = FormatoValor(ValorNulo(AdoPrincipal!Repuestos), "S/.", gintDecimalesMoneda)
                Item.SubItems(3) = FormatoValor(ValorNulo(AdoPrincipal!Carroceria), "S/.", gintDecimalesMoneda)
                Item.SubItems(4) = FormatoValor(ValorNulo(AdoPrincipal!Terceros), "S/.", gintDecimalesMoneda)
                Item.SubItems(5) = FormatoValor(ValorNulo(AdoPrincipal!Mecanica), "S/.", gintDecimalesMoneda)
                Item.SubItems(6) = FormatoValor(ValorNulo(AdoPrincipal!Otros), "S/.", gintDecimalesMoneda)
                
                AdoPrincipal.MoveNext
            Loop
End If
End Sub
