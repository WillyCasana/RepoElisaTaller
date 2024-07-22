VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Begin VB.Form frmMantenedorHorasMecanicos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor Horas Mecanicos"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "frmMantenedorHorasMecanicos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   7575
      Begin MSDataListLib.DataCombo dtcMeses 
         Bindings        =   "frmMantenedorHorasMecanicos.frx":038A
         Height          =   315
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "NOMBRE"
         BoundColumn     =   "CODIGO"
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
      Begin MSAdodcLib.Adodc datMeses 
         Height          =   330
         Left            =   1080
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
         Height          =   3495
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Codigo"
            Text            =   "Codigo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Descripcion"
            Text            =   "Nombre Mecanico"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   "DP"
            Text            =   "Horas Compradas"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Key             =   "Orden"
            Text            =   "Horas Reales"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.TextBox txtAño 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox chkVigencia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Activo:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6120
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin Crystal.CrystalReport rptMantenedor 
         Left            =   5160
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "Lista de Mecanicos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Año:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Mes:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Crear"
            Object.ToolTipText     =   "Crear Registro (Ctrl+N)"
            ImageKey        =   "Crear"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Registro (Ctrl+G)"
            ImageKey        =   "Grabar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar (ESC)"
            ImageKey        =   "Cancelar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Borrar"
            Object.ToolTipText     =   "Borrar Registro (Ctrl+D)"
            ImageKey        =   "Borrar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar Registro (Ctrl+B)"
            ImageKey        =   "Buscar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir (Ctrl+I)"
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Primer Registro (Ctrl+P)"
            ImageKey        =   "Primero"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Registro Anterior (Ctrl+A)"
            ImageKey        =   "Anterior"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Registro Siguiente (Ctrl+S)"
            ImageKey        =   "Siguiente"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ultimo Registro (Ctrl+U)"
            ImageKey        =   "Ultimo"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Renovar"
            Object.ToolTipText     =   "Renovar Registros (Ctrl+R)"
            ImageKey        =   "Renovar"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (Ctrl+Q)"
            ImageKey        =   "Salir"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Actualizar"
            Object.ToolTipText     =   "Actualizar Horas"
            ImageKey        =   "Actualizar"
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
            Picture         =   "frmMantenedorHorasMecanicos.frx":03A1
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":04B3
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":05C5
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":06D7
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":07E9
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":08FB
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":0A0D
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":0B1F
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":0C31
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":0D43
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":0E55
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":0F67
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":1079
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":118B
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":129D
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":13AF
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":14C1
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":1913
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":1D65
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":1E77
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":1FD3
            Key             =   "Actualizar"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":212F
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":228B
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":23E7
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":2EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":3307
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":346B
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":38C7
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":3A23
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":4D2F
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":52CB
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":5427
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":5583
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":58D7
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":5C2B
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":5F7F
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":62D3
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":6627
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":6B6B
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":6C7D
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":6FD1
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":7325
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":7439
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":778D
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":789F
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorHorasMecanicos.frx":7BF3
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMantenedorHorasMecanicos"
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
Const mcNombreTabla = "Tllr_Mes_Año"
Const mcCampoCodigo = "Id_Mes"
Const mcCampoNombre = "Año"

Sub MecanicosMesAño(strMes As String, intAño As Integer)

lvwConceptos.ListItems.Clear
mstrSQL = "SELECT Tllr_Mes_Año_Mecanico.*, Tllr_Mecanicos.Nombre FROM Tllr_Mes_Año_Mecanico INNER JOIN Tllr_Mecanicos ON Tllr_Mes_Año_Mecanico.Id_Mecanico=Tllr_Mecanicos.Id_Mecanico WHERE Id_Mes = '" & strMes & "' And Año = " & intAño & " And tllr_mes_año_Mecanico.id_empresa='" & gstrIdEmpresa & "' And tllr_mes_año_mecanico.Id_Sucursal='" & gstrIdSucursal & "' and vigencia='S'  AND (Es_Recepcionista = 'N') AND (Es_Supervisor = 'N') AND (Es_Liquidador = 'N') AND Nombre not like '%definir%' order by Nombre "
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With AdoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveLast: .MoveFirst
            While Not .EOF
                Set Item = lvwConceptos.ListItems.Add(, , !Id_Mecanico)
                Item.SubItems(1) = !Nombre
                Item.SubItems(2) = FormatoValor(ValorNulo(!HorasCompradas), "", 1)
                Item.SubItems(3) = FormatoValor(ValorNulo(!HorasReales), "", 1)
                .MoveNext
            Wend
        End If
    End With
End If ' por el otro
Set AdoPrincipal = New ADODB.Recordset
Conexion.CloseHost AdoPrincipal

End Sub

Private Sub Check_Off()
Dim V As Integer

For V = 1 To lvwConceptos.ListItems.Count
    Set lvwConceptos.SelectedItem = lvwConceptos.ListItems(V)
    lvwConceptos.SelectedItem.Checked = False
Next
End Sub
Sub Fill_Meses()
mstrSQL = "SELECT Id_Mes AS CODIGO, Descripcion AS NOMBRE FROM Glbl_Mes where VIGENCIA = 'S'"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With datMeses
        Set .Recordset = AdoPrincipal
        If Not .Recordset.BOF And Not .Recordset.EOF Then
            .Recordset.MoveFirst
            dtcMeses.ListField = "Nombre"
            dtcMeses.BoundColumn = "Codigo"
        End If
    End With
End If
Set AdoPrincipal = New ADODB.Recordset
Conexion.CloseHost AdoPrincipal
End Sub

Private Sub Form_Load()
    mblnSW = True
End Sub

Private Sub lvwConceptos_DblClick()
If Me.lvwConceptos.ListItems.Count > 0 Then
    frmEditaHorasMecanico.Show vbModal
End If
End Sub

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Screen.MousePointer = vbHourglass
    Select Case Button.Key
        Case "Crear"
            AgregarRegistro
        Case "Grabar"
            GrabarRegistro
        Case "Cancelar"
            CancelarAgregaRegistro
        Case "Borrar"
            BorrarRegistro
        Case "Buscar"
            BuscarRegistro
        Case "Imprimir"
            ImprimirInforme
        Case "Primero"
            PrimerRegistro
        Case "Anterior"
            RegistroAnterior
        Case "Siguiente"
            RegistroSiguiente
        Case "Ultimo"
            UltimoRegistro
        Case "Renovar"
            Renovar
        Case "Cerrar"
            CerrarSalir
        Case "Actualizar"
            ActualizarHoras
    End Select
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Activate()
    If mblnSW Then
        If Not Atributos("Glbl", "Tllr_10_0051", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
        Fill_Meses
        If gapAccion = apcrear Then
           AgregarRegistro
           'txtCodigo = gstrBusca
        End If
        If gapAccion = apeditar Then
            If gstrBusca <> "" Then
                mstrSQL = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' order by " & mcCampoCodigo
                If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                        LeerCampos
                    End If
                End If
                Conexion.CloseHost AdoPrincipal
            End If
            'txtCodigo.Enabled = False
            Me.SetFocus
        End If
        If gapAccion = apninguno Then
           Renovar
        End If
    End If
    gapAccion = apninguno
    mblnSW = False
    'txtNombre.SetFocus
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{tab}"
        Case vbKeyEscape
            KeyAscii = 0
            CancelarAgregaRegistro
        Case 14 And tlbBarraHerramientas.Buttons.Item("Crear").Enabled
            KeyAscii = 0
            AgregarRegistro
        Case 7 And tlbBarraHerramientas.Buttons.Item("Grabar").Enabled
            KeyAscii = 0
            GrabarRegistro
        Case 4 And tlbBarraHerramientas.Buttons.Item("Borrar").Enabled
            KeyAscii = 0
            BorrarRegistro
        Case 2 And tlbBarraHerramientas.Buttons.Item("Buscar").Enabled
            KeyAscii = 0
            BuscarRegistro
        Case 9 And tlbBarraHerramientas.Buttons.Item("Imprimir").Enabled
            KeyAscii = 0
            ImprimirInforme
        Case 16 And tlbBarraHerramientas.Buttons.Item("Primero").Enabled
            KeyAscii = 0
            PrimerRegistro
        Case 1 And tlbBarraHerramientas.Buttons.Item("Anterior").Enabled
            KeyAscii = 0
            RegistroAnterior
        Case 19 And tlbBarraHerramientas.Buttons.Item("Siguiente").Enabled
            KeyAscii = 0
            RegistroSiguiente
        Case 21 And tlbBarraHerramientas.Buttons.Item("Ultimo").Enabled
            KeyAscii = 0
            UltimoRegistro
        Case 18 And tlbBarraHerramientas.Buttons.Item("Renovar").Enabled
            KeyAscii = 0
            Renovar
        Case 17 And tlbBarraHerramientas.Buttons.Item("Cerrar").Enabled
            KeyAscii = 0
            CerrarSalir
    End Select
End Sub
Private Sub AgregarRegistro()
    Me.Tag = "Crear"
    DesactivaBotones
    LimpiaCampos
    ValoresporDefecto
    CargaMecanicos
End Sub
Private Sub CancelarAgregaRegistro()
    Me.Tag = ""
    ActivaBotones
    
    'mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & dtcMeses.BoundText & "'order by " & mcCampoCodigo
    mstrSQL = "select TOP 1 * from " & mcNombreTabla & " Where id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by Año," & mcCampoCodigo
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            mstrSQL = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & dtcMeses.BoundText & "' order by " & mcCampoCodigo
            If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                    LeerCampos
                Else
                    mblnTablaVacia = True
                    LimpiaCampos
                End If
            End If
        End If
    End If
    Conexion.CloseHost AdoPrincipal
    'txtNombre.SetFocus
End Sub
Private Sub GrabarRegistro()
    If Not validacion() Then
        Exit Sub
    End If

    If Me.Tag = "Crear" Then
        mstrSQL = "INSERT INTO " & mcNombreTabla & " (" & mcCampoCodigo & ", " & mcCampoNombre & ", vigencia, "
        mstrSQL = mstrSQL & "usr_id, usr_fecha,Id_empresa,Id_Sucursal) "
        mstrSQL = mstrSQL & "values ('" & Me.dtcMeses.BoundText & "', " & Trim(txtAño) & ", '" & IIf(chkVigencia.Value = vbChecked, "S", "N") & "', "
        mstrSQL = mstrSQL & "'" & gstrUsuario & "', '" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "',"
        mstrSQL = mstrSQL & "'" & gstrIdEmpresa & "','" & gstrIdSucursal & "')"
    Else
        mstrSQL = "UPDATE " & mcNombreTabla & " SET " & mcCampoNombre & "=" & Trim(txtAño) & ", vigencia='" & IIf(chkVigencia.Value = vbChecked, "S", "N") & "', "
        mstrSQL = mstrSQL & "usr_id='" & gstrUsuario & "', usr_fecha='" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "'"
        mstrSQL = mstrSQL & " where " & mcCampoCodigo & "='" & Trim(dtcMeses.BoundText) & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
    End If
    If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
        mblnTablaVacia = False
        ActivaBotones
        Me.Tag = ""
    End If
    
    GuardaConceptos Trim(dtcMeses.BoundText), txtAño
    
End Sub

Private Sub GuardaConceptos(strMes As String, intAño As Integer)
Dim x As Integer

mstrSQL = "DELETE FROM TLLR_Mes_Año_Mecanico WHERE Id_Mes ='" & strMes & "' And  Año =" & txtAño & " And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
Conexion.SendHost mstrSQL, , , , gcTiempoEspera '//////////AQUI BORRA LAS QUE EXISTEN

For x = 1 To lvwConceptos.ListItems.Count
    Set lvwConceptos.SelectedItem = lvwConceptos.ListItems(x)
    mstrSQL = "INSERT INTO TLLR_Mes_Año_Mecanico ( ID_Mes, Año, Id_Mecanico, HorasCompradas, HorasReales, Id_empresa, Id_Sucursal )"
    mstrSQL = mstrSQL & " VALUES('" & strMes & "'," & intAño & ",'" & lvwConceptos.SelectedItem & "' , " & Me.lvwConceptos.SelectedItem.SubItems(2) & ", " & Me.lvwConceptos.SelectedItem.SubItems(3) & " ,'" & gstrIdEmpresa & "','" & gstrIdSucursal & "')"
    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
Next '///////////////AQUI GRABA LAS NUEVAS Y LAS QUE ESTABAN

End Sub

Private Sub BorrarRegistro()
    Screen.MousePointer = vbDefault
    If MsgBox("¿ Desea eliminar este registro ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
        'elimina primero los movimientos
        mstrSQL = "DELETE FROM TLLR_Mes_Año_Mecanico WHERE Id_Mes ='" & dtcMeses.BoundText & "' And  Año =" & txtAño & " And id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        Conexion.SendHost mstrSQL, , , , gcTiempoEspera '//////////AQUI BORRA LAS QUE EXISTEN

        mstrSQL = "DELETE FROM " & mcNombreTabla & " where " & mcCampoCodigo & "='" & dtcMeses.BoundText & "' And Año = " & txtAño & " And id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSQL, , , , gcTiempoEspera) = apOk Then
            If dtcMeses.BoundText = "12" Then
                mstrSQL = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & dtcMeses.BoundText & "' And Año >" & txtAño & " And id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by " & mcCampoCodigo
            Else
                mstrSQL = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & dtcMeses.BoundText & "' And Año =" & txtAño & " And id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by " & mcCampoCodigo
            End If
            If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                    LeerCampos
                Else
                    If dtcMeses.BoundText = "01" Then
                        mstrSQL = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & dtcMeses.BoundText & "' And Año < " & txtAño & " And id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by " & mcCampoCodigo
                    Else
                        mstrSQL = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & dtcMeses.BoundText & "' And Año = " & txtAño & " And id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by " & mcCampoCodigo
                    End If
                    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                            LeerCampos
                        Else
                            mblnTablaVacia = True
                            LimpiaCampos
                        End If
                    End If
                End If
            End If
        End If
        Conexion.CloseHost AdoPrincipal
    End If
End Sub
Private Sub BuscarRegistro()
    mstrSQL = "SELECT Tllr_Mes_Año.Id_mes,Glbl_Mes.Descripcion FROM Tllr_Mes_Año INNER JOIN Glbl_Mes ON Tllr_Mes_Año.Id_Mes=Glbl_Mes.Id_Mes"
    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "(SELECT Tllr_mes_Año.Id_mes + '/' + cast(Tllr_Mes_Año.Año as nvarchar) as Codigo,Glbl_Mes.Descripcion FROM Tllr_Mes_Año INNER JOIN Glbl_Mes ON Tllr_Mes_Año.Id_Mes=Glbl_Mes.Id_Mes where id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "') As MyTabla", "Codigo", "Descripcion", Me.Caption)
    If gstrBusca <> "" Then
        mstrSQL = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & Mid(gstrBusca, 1, 2) & "' And Año=" & Mid(gstrBusca, 4, 4)
        If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
                LeerCampos
            End If
        End If
        Conexion.CloseHost AdoPrincipal
    End If
    Me.SetFocus
End Sub
Private Sub ImprimirInforme()
   ' FormVol1.ImprimirRegistros Conexion, mcNombreTabla, mcCampoCodigo, mcCampoNombre, Me.Caption, gstrPathReporte, "APCARROC.RPT", gstrUSUARIO, gstrCodigoEmpresa
    With rptMantenedor
        .ReportFileName = gstrPathReporte & "\APHORASMECANICO.RPT"
        .Formulas(0) = "Titulo='Listado Horas Mecanico'"
        .Formulas(1) = "Razonsocial='" & gstrEmpresa & "'"
        .Formulas(2) = "Ruc='" & gstrIdEmpresa & "'"
        .Formulas(3) = "Direccion='" & gstrDirSuc & "'"
        .Formulas(4) = "Usuario='" & gstrUsuario & "'"
        .Formulas(5) = "Marcamodulo='ElisaTaller'"
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Connect = cnnAux.ConnectionString
        .Action = True
    End With

End Sub
Private Sub PrimerRegistro()
    
    mstrSQL = "select TOP 1 * from " & mcNombreTabla & " Where id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by Año," & mcCampoCodigo
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub RegistroAnterior()
    If dtcMeses.BoundText = "01" Then
        mstrSQL = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & dtcMeses.BoundText & "' And Año < " & txtAño & " And id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by " & mcCampoCodigo
    Else
        mstrSQL = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & dtcMeses.BoundText & "' And Año = " & txtAño & " And id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by " & mcCampoCodigo
    End If
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            AdoPrincipal.MoveLast
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub RegistroSiguiente()
    If dtcMeses.BoundText = "12" Then
        mstrSQL = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & dtcMeses.BoundText & "' And Año >" & txtAño & " And id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by " & mcCampoCodigo
    Else
        mstrSQL = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & dtcMeses.BoundText & "' And Año =" & txtAño & " And id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by " & mcCampoCodigo
    End If
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub UltimoRegistro()
    mstrSQL = "select * from " & mcNombreTabla & " Where id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by Año," & mcCampoCodigo
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            AdoPrincipal.MoveLast
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub Renovar()
    'Set adoPrincipal = New ADODB.Recordset
    mstrSQL = "select TOP 1 * from " & mcNombreTabla & " Where id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' order by Año," & mcCampoCodigo
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        VerificaTablaVacia
        ActivaBotones
        If Not mblnTablaVacia Then
            PrimerRegistro
        End If
    End If
    Conexion.CloseHost AdoPrincipal
End Sub
Private Sub CerrarSalir()
    Unload Me
End Sub
Private Sub Ayuda()
End Sub
Private Sub ActivaBotones()
    dtcMeses.Enabled = False
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
    dtcMeses.Enabled = True
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
Private Sub VerificaTablaVacia()
    If (Not AdoPrincipal.BOF And Not AdoPrincipal.EOF) And AdoPrincipal.RecordCount > 0 Then
        mblnTablaVacia = False
    Else
        mblnTablaVacia = True
        LimpiaCampos
        MsgBox "La tabla no contiene registros...", vbInformation, "Advertencia"
    End If
End Sub
Private Sub LeerCampos()

    If mblnTablaVacia Then
        LimpiaCampos
        Exit Sub
    End If

    With AdoPrincipal
        dtcMeses.BoundText = ValorNulo(.Fields(mcCampoCodigo))
        If IsNull(!vigencia) Then
            chkVigencia.Value = vbUnchecked
        Else
            If !vigencia = "S" Then
                chkVigencia.Value = vbChecked
            Else
                chkVigencia.Value = vbUnchecked
            End If
        End If
        Me.txtAño = .Fields("AÑO")
        'Check_Off
        MecanicosMesAño dtcMeses.BoundText, txtAño
        
    End With
End Sub
Private Sub LimpiaCampos()
    dtcMeses.Text = ""
    txtAño = ""
    chkVigencia.Value = vbUnchecked
    Me.lvwConceptos.ListItems.Clear
End Sub
Private Sub ValoresporDefecto()
Dim lstrMes As String

Me.chkVigencia.Value = vbChecked
Me.txtAño.Text = Year(Date)
lstrMes = Month(Date)
If CInt(lstrMes) < 10 Then lstrMes = "0" & lstrMes
Me.dtcMeses.BoundText = lstrMes

End Sub
Private Function validacion() As Boolean
    validacion = True
    If Me.dtcMeses.Text = "" Then
        MsgBox "El Valor Meses debe contener un valor...", vbInformation, "Advertencia"
        dtcMeses.SetFocus
        validacion = False
        Exit Function
    End If
    If txtAño = "" Then
        MsgBox "El Año debe contener un valor...", vbInformation, "Advertencia"
        txtAño.SetFocus
        validacion = False
        Exit Function
    End If
  
    If Len(txtAño) < 4 Or Len(txtAño) > 4 Then
        MsgBox "El Año debe contener un formato de 4 digitos...", vbInformation, "Advertencia"
        txtAño.SetFocus
        validacion = False
        Exit Function
    End If
    
    '//Verifica si existe un registro...
    If Me.Tag = "Crear" Then
        Dim adoTemp As New ADODB.Recordset
        mstrSQL = "select " & mcCampoCodigo & ", " & mcCampoNombre & " from " & mcNombreTabla & " where " & mcCampoCodigo & "='" & dtcMeses.BoundText & "' And Año=" & Me.txtAño & " And id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoTemp.BOF And Not adoTemp.EOF Then
                MsgBox "Este código ya esta registrado con la descripción " & Chr(13) & "[" & IIf(IsNull(adoTemp.Fields(mcCampoNombre)), "SIN DESCRIPCION", adoTemp.Fields(mcCampoNombre)) & "]", vbInformation, "Advertencia"
                validacion = False
                dtcMeses.SetFocus
            End If
        End If
        Conexion.CloseHost adoTemp
    End If
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmMantenedorHorasMecanicos = Nothing
    gstrBusca = dtcMeses.BoundText
End Sub
Private Sub RevizaAtributos()
    mblnAccesoCrear = True
    mblnAccesoEditar = True
    mblnAccesoBorrar = True
    mblnAccesoImprimir = True
End Sub
Private Sub CargaMecanicos()
mstrSQL = "SELECT Id_Mecanico,Nombre FROM Tllr_Mecanicos WHERE Vigencia='S' And Id_empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'  and vigencia='S'  AND (Es_Recepcionista = 'N') AND (Es_Supervisor = 'N') AND (Es_Liquidador = 'N') AND Nombre not like '%definir%' order by Nombre"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    With AdoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveLast: .MoveFirst
            While Not .EOF
                Set Item = lvwConceptos.ListItems.Add(, , !Id_Mecanico)
                Item.SubItems(1) = !Nombre
                Item.SubItems(2) = "0.00"
                Item.SubItems(3) = "0.00"
                .MoveNext
            Wend
        End If
    End With
End If ' por el otro
Set AdoPrincipal = New ADODB.Recordset
Conexion.CloseHost AdoPrincipal
End Sub

Private Sub txtAño_KeyPress(KeyAscii As Integer)
KeyAscii = CheckNumber(KeyAscii, txtAño, strDot)
End Sub
Private Sub ActualizarHoras()
Dim SumaHorasReales As Double
Dim i As Integer
Dim fecha1 As String
Dim fecha2 As String

'arma fecha
fecha1 = BOM("01/" & CDbl(Me.dtcMeses.BoundText) & "/" & Me.txtAño)
fecha2 = EOM("01/" & CDbl(Me.dtcMeses.BoundText) & "/" & Me.txtAño)


For i = 1 To Me.lvwConceptos.ListItems.Count
    'suma hora actividades
    SumaHorasReales = 0
    
    mstrSQL = "SELECT SUM(isnull(HorasReales,0)) AS SumaHoraActividades From TLLR_ACTIVIDADES_MECANICO "
    mstrSQL = mstrSQL & "Where Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "' And Id_Mecanico='" & Me.lvwConceptos.ListItems(i) & "' "
    mstrSQL = mstrSQL & "And FechaEmision between '" & fecha1 & "' And '" & fecha2 & "'"
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            SumaHorasReales = SumaHorasReales + IIf(IsNull(AdoPrincipal!SumaHoraActividades), 0, AdoPrincipal!SumaHoraActividades)
        End If
    End If
    
    Conexion.CloseHost AdoPrincipal
    
    mstrSQL = "SELECT sum(isnull(Tllr_Otro_OT.HorasReales,0)) AS SumaHoraOtro "
    mstrSQL = mstrSQL & "FROM Tllr_Otro_OT INNER JOIN "
    mstrSQL = mstrSQL & "Tllr_OT ON Tllr_Otro_OT.Id_OT = Tllr_OT.Id_OT AND Tllr_Otro_OT.Id_Empresa = Tllr_OT.Id_Empresa AND "
    mstrSQL = mstrSQL & "Tllr_Otro_OT.Id_Sucursal = Tllr_OT.Id_Sucursal And Tllr_Otro_OT.Seccion_OT = Tllr_OT.Seccion_OT "
    mstrSQL = mstrSQL & "Where Tllr_Otro_Ot.Id_Empresa='" & gstrIdEmpresa & "' And Tllr_Otro_Ot.Id_Sucursal='" & gstrIdSucursal & "' And Tllr_Otro_Ot.Mecanico_Asignado='" & Me.lvwConceptos.ListItems(i) & "' "
    mstrSQL = mstrSQL & "And Tllr_ot.Fecha_Emision between '" & fecha1 & "' And '" & fecha2 & "'"
    
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not AdoPrincipal.BOF And Not AdoPrincipal.EOF Then
            SumaHorasReales = SumaHorasReales + IIf(IsNull(AdoPrincipal!SumaHoraOtro), 0, AdoPrincipal!SumaHoraOtro)
        End If
    End If
    Conexion.CloseHost AdoPrincipal
    
    If SumaHorasReales > 0 Then
        Me.lvwConceptos.ListItems(i).SubItems(3) = FormatoValor(SumaHorasReales, "", 1)
    End If
Next
End Sub
