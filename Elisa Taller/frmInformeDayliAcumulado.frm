VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmInformeDayliAcumulado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Daily Acumulado"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   Icon            =   "frmInformeDayliAcumulado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   10815
      Begin VB.TextBox txtTipoCambio 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9000
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdlimpiar 
         Height          =   315
         Left            =   5880
         Picture         =   "frmInformeDayliAcumulado.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpiar"
         Top             =   480
         Width           =   315
      End
      Begin VB.CommandButton cmdLimpia2 
         Height          =   315
         Left            =   8160
         Picture         =   "frmInformeDayliAcumulado.frx":048C
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Limpia Fecha de Término"
         Top             =   480
         Width           =   315
      End
      Begin VB.CommandButton cmdLimpiaEmpresa 
         Height          =   315
         Left            =   2520
         Picture         =   "frmInformeDayliAcumulado.frx":09BE
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Limpia filtro por Empresa"
         Top             =   480
         Width           =   315
      End
      Begin MSDataListLib.DataCombo cmbsucursal 
         Bindings        =   "frmInformeDayliAcumulado.frx":0EF0
         Height          =   315
         Left            =   3600
         TabIndex        =   4
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
         Left            =   4440
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
         Left            =   6600
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
         Format          =   87359489
         CurrentDate     =   36772
      End
      Begin MSDataListLib.DataCombo dbcboEmpresa 
         Bindings        =   "frmInformeDayliAcumulado.frx":0F0A
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
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
      Begin VB.Label Label1 
         Caption         =   "Tipo de Cambio"
         Height          =   255
         Left            =   9000
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Sucursal"
         Height          =   195
         Left            =   3600
         TabIndex        =   9
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Final"
         Height          =   255
         Left            =   6600
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblEmpresa 
         Caption         =   "Empresa"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   600
      Top             =   4440
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
            Picture         =   "frmInformeDayliAcumulado.frx":0F23
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":1035
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":1147
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":1259
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":136B
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":147D
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":158F
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":16A1
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":17B3
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":18C5
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":19D7
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":1AE9
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":1BFB
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":1D0D
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":1E1F
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":1F31
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":2043
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":2495
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":28E7
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":29F9
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":2B55
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":2CB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":2E0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":2F69
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":3A35
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":3E89
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":3FED
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":4449
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":45A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":58B1
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":5E4D
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":5FA9
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":6105
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":6459
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":67AD
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeDayliAcumulado.frx":6B01
            Key             =   "outlook"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdExportar 
      Left            =   1920
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport rptKardex 
      Left            =   2640
      Top             =   4560
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
      TabIndex        =   10
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
   Begin MSComctlLib.ListView lsvdetalle 
      Height          =   2100
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   3704
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Fecha Facturacion"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "OT x Día"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Servicio"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Margen Servicio"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "D&P"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Margen D&P"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Repuestos"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Margen Repuestos"
         Object.Width           =   2999
      EndProperty
   End
End
Attribute VB_Name = "frmInformeDayliAcumulado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoRecordSucursal As New ADODB.Recordset
Dim AdoRecordEmpresa As New ADODB.Recordset
Dim adoRecordset As New ADODB.Recordset
Dim mstrSQL As String
Dim adoPrincipal As New ADODB.Recordset
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

Private Sub cmdlimpiar_Click()
Me.cmbsucursal.Text = ""
End Sub

Private Sub Form_Activate()
Dim blnBoolean As Boolean
    If mblSW Then
        mblSW = False
        
        Screen.MousePointer = vbDefault
        If Not Atributos("Glbl", "Tllr_30_0085", True, True, True, True) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
       
    End If
End Sub

Private Sub Form_Load()
 mblSW = True
 dtpHasta.Value = Format(Date, "dd/mm/yyyy")

 
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
Dim strNumItem As Integer

If txtTipoCambio = "" Then
    MsgBox "Debe ingresar el Tipo de Cambio", vbInformation, "Advertencia"
    txtTipoCambio.SetFocus
    Exit Sub
End If

mstrSQL = "EXEC Tllr_Informe_Dayli_Acum  '" & gstrIdEmpresa & "', '" & Me.cmbsucursal.BoundText & "', '" & Format(Me.dtpHasta.Value, "dd/mm/yyyy") & "','" & Me.txtTipoCambio & "'"

Me.lsvdetalle.ListItems.Clear
If Conexion.SendHost(mstrSQL, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                    adoPrincipal.MoveFirst
    End If
            Do Until adoPrincipal.EOF
                strNumItem = strNumItem + 1
                
            Set Item = Me.lsvdetalle.ListItems.Add(, , strNumItem)
                Item.SubItems(1) = adoPrincipal!Fecha
                Item.SubItems(2) = ValorNulo(adoPrincipal!NumOts)
                Item.SubItems(3) = FormatoValor(ValorNulo(adoPrincipal!servicio), "$", gintDecimalesMoneda)
                Item.SubItems(4) = FormatoValor(ValorNulo(adoPrincipal!mgservicio), "$", gintDecimalesMoneda)
                Item.SubItems(5) = FormatoValor(ValorNulo(adoPrincipal!dp), "$", gintDecimalesMoneda)
                Item.SubItems(6) = FormatoValor(ValorNulo(adoPrincipal!mgdp), "$", gintDecimalesMoneda)
                Item.SubItems(7) = FormatoValor(ValorNulo(adoPrincipal!Repuestos), "$", gintDecimalesMoneda)
                Item.SubItems(8) = FormatoValor(ValorNulo(adoPrincipal!MgRep), "$", gintDecimalesMoneda)
                adoPrincipal.MoveNext
            Loop
End If
End Sub



