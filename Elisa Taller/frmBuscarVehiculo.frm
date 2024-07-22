VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "fpspr70.ocx"
Begin VB.Form frmBuscarVehiculo 
   Caption         =   "Búsqueda de Vehículos"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBuscarVehiculo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   ScaleHeight     =   6990
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fmeCriterios 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   10815
      Begin VB.CommandButton cmdLimpiarColor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         Picture         =   "frmBuscarVehiculo.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Limpia Color Seleccionado"
         Top             =   5400
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   315
      End
      Begin MSDataListLib.DataCombo dbcboColor 
         Bindings        =   "frmBuscarVehiculo.frx":08BC
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   5400
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "id_Color_Exterior"
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
      Begin VB.CommandButton cmdLimpiaFecha1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5160
         Picture         =   "frmBuscarVehiculo.frx":08D3
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Limpia Fecha de Inicio"
         Top             =   5400
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton cmdLimpiaFecha2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7200
         Picture         =   "frmBuscarVehiculo.frx":0E05
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Limpia Fecha de Término"
         Top             =   5400
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox txtPatente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8880
         TabIndex        =   12
         Top             =   5400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtNumeroCajon 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtPrecioVehiculo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   10
         Top             =   4800
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtAñoVehiculo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   7
         Top             =   4800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtFantasia 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         TabIndex        =   6
         Top             =   5400
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtPrecioVehiculo2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   11
         Top             =   4800
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtAñoVehiculo2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   8
         Top             =   4800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cboEstadoStock 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmBuscarVehiculo.frx":1337
         Left            =   6720
         List            =   "frmBuscarVehiculo.frx":1344
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   4800
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSAdodcLib.Adodc datColor 
         Height          =   375
         Left            =   1320
         Top             =   5280
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   315
         Left            =   5640
         TabIndex        =   25
         Top             =   5400
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   91095041
         CurrentDate     =   36772
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   315
         Left            =   3600
         TabIndex        =   26
         Top             =   5400
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   91095041
         CurrentDate     =   36772
      End
      Begin VB.Label Label11 
         Caption         =   "Desde"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3600
         TabIndex        =   28
         Top             =   5160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Hasta"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5640
         TabIndex        =   27
         Top             =   5160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblPatente 
         Caption         =   "Patente"
         Height          =   255
         Left            =   8880
         TabIndex        =   22
         Top             =   5160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblCajon 
         Caption         =   "VIN"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Desde Precio Venta"
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   4560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Desde Año"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   4560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Vendido/Stock/Reserva"
         Height          =   255
         Left            =   6720
         TabIndex        =   18
         Top             =   4560
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label14 
         Caption         =   "Fantasía"
         Height          =   255
         Left            =   7680
         TabIndex        =   17
         Top             =   5160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Hasta Año"
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   4560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Hasta Precio Venta"
         Height          =   255
         Left            =   4800
         TabIndex        =   15
         Top             =   4560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Color"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   5160
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin ComCtl3.CoolBar clbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   741
      BandCount       =   2
      BandBorders     =   0   'False
      VariantHeight   =   0   'False
      _CBWidth        =   11055
      _CBHeight       =   420
      _Version        =   "6.7.9782"
      Child1          =   "BarraHerramientas"
      MinHeight1      =   330
      Width1          =   3420
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      MinHeight2      =   360
      FixedBackground2=   0   'False
      NewRow2         =   0   'False
      BandStyle2      =   1
      AllowVertical2  =   0   'False
      Begin MSComctlLib.Toolbar BarraHerramientas 
         Height          =   330
         Left            =   165
         TabIndex        =   34
         Top             =   45
         Width           =   10800
         _ExtentX        =   19050
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Crear"
               Object.ToolTipText     =   "Nueva búsqueda"
               ImageKey        =   "Crear"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar "
               ImageKey        =   "Buscar"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Columnas"
               Object.ToolTipText     =   "Ver y Ocultar Columnas"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Imprimir"
               Object.ToolTipText     =   "Imprimir "
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cerrar"
               Object.ToolTipText     =   "Cerrar "
               ImageKey        =   "Salir"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ProgressBar pb1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   29
      Top             =   6585
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame fmeBotones 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   7920
      TabIndex        =   31
      Top             =   5760
      Width           =   3015
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "Seleccionar"
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   0
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList imgBtnSmall 
      Left            =   11280
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   8
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":1368
            Key             =   "Limpiar"
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread sprGrillaPrincipal 
      Bindings        =   "frmBuscarVehiculo.frx":147C
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   10815
      _Version        =   458752
      _ExtentX        =   19076
      _ExtentY        =   6588
      _StockProps     =   64
      ArrowsExitEditMode=   -1  'True
      DAutoSave       =   0   'False
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OperationMode   =   3
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "frmBuscarVehiculo.frx":1493
      TextTip         =   1
      TextTipDelay    =   200
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin MSAdodcLib.Adodc datDatos 
      Height          =   330
      Left            =   4560
      Top             =   7320
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
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   6735
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   450
      SimpleText      =   "0"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14552
         EndProperty
      EndProperty
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
   Begin MSComctlLib.ImageList imgBtnSmall_d 
      Left            =   11280
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   8
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":16CA
            Key             =   "Limpiar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTitulo 
      Left            =   840
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   37
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":17DE
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":18F0
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":1A02
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":1B14
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":1C26
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":1D38
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":1E4A
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":1F5C
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":206E
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":2180
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":2292
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":23A4
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":24B6
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":25C8
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":26DA
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":27EC
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":28FE
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":2D50
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":31A2
            Key             =   "CopiarX"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":32B4
            Key             =   "AgregarSucursal"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":3708
            Key             =   "VerSucursal"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":3E7C
            Key             =   "Abrir"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":3F9C
            Key             =   "Horizontal"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":4C90
            Key             =   "Resalte"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":50E4
            Key             =   "Cerrar2"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":5240
            Key             =   "Reset"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":539C
            Key             =   "Config_Col"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":5738
            Key             =   "otro"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":5B8C
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":5EA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":62FC
            Key             =   "Categorizar"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":6750
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":6870
            Key             =   "Pegar"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":6990
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":6ED4
            Key             =   "Categorias"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":6FE8
            Key             =   "UpDown"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":733C
            Key             =   "Excel2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbTitulo 
      Height          =   330
      Left            =   120
      TabIndex        =   32
      Top             =   1560
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imgTitulo"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Preview"
            Object.ToolTipText     =   "Vista Previa"
            ImageKey        =   "Preview"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Exportar a Microsoft Excel"
            ImageKey        =   "Excel2"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copiar"
            Object.ToolTipText     =   "Copiar"
            ImageKey        =   "Copiar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ordenar"
            Object.ToolTipText     =   "Ordenar"
            ImageKey        =   "SortAsc"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Vertical"
            Object.ToolTipText     =   "Vista Vertical de los Datos"
            ImageKey        =   "Horizontal"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   120
      Top             =   5760
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
            Picture         =   "frmBuscarVehiculo.frx":7690
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":77A2
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":78B4
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":79C6
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":7AD8
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":7BEA
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":7CFC
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":7E0E
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":7F20
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":8032
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":8144
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":8256
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":8368
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":847A
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":858C
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":869E
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":87B0
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":8C02
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":9054
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":9166
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":92C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":941E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":957A
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":96D6
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":A1A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":A5F6
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":A75A
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":ABB6
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":AD12
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":C01E
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":C5BA
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":C716
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":C872
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":CBC6
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":CF1A
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuscarVehiculo.frx":D26E
            Key             =   "outlook"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBuscarVehiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnTablaVacia As Boolean
Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean
Dim mstrSigla As String
Dim mintDecimales As Integer
Dim blnModeloCodigo As Boolean
Dim mSw As Boolean

Public Sub SetColoresSpread(ldblRows As Double)
'Dim lvarTemp As Variant
'Dim llngCol As Long
''
'If ldblRows = 0 Then Exit Sub
''
''Me.shpColor(0).BackColor = cstrColorVERDE ' vigente
''Me.lblColor(0).ForeColor = Me.shpColor(0).BackColor
''Me.lblColor(0).FontStrikethru = False
''Me.shpColor(1).BackColor = cstrColorROJO ' liquidada
''Me.lblColor(1).ForeColor = Me.shpColor(1).BackColor
''Me.lblColor(1).FontStrikethru = False
''Me.shpColor(2).BackColor = cstrColorNEGRO ' facturado
''Me.lblColor(2).ForeColor = Me.shpColor(2).BackColor
''Me.lblColor(2).FontStrikethru = False
''Me.shpColor(3).BackColor = cstrColorGRIS_CLARO ' nula
''Me.lblColor(3).ForeColor = Me.shpColor(3).BackColor
''Me.lblColor(3).FontStrikethru = True
''Me.shpColor(4).BackColor = cstrColorGRIS_CLARO ' presupuesto
''Me.lblColor(4).ForeColor = Me.shpColor(4).BackColor
''Me.lblColor(4).FontStrikethru = False
''Me.shpColor(5).BackColor = cstrColorGRIS_MEDIO_CLARO ' reserva
''Me.lblColor(5).ForeColor = Me.shpColor(5).BackColor
''Me.lblColor(5).FontStrikethru = False
''
'
'Me.sprGrillaPrincipal.Row = ldblRows
'Me.sprGrillaPrincipal.Col = -1
'Me.sprGrillaPrincipal.ForeColor = cstrColorNEGRO
'Me.sprGrillaPrincipal.FontStrikethru = False
'Me.sprGrillaPrincipal.FontBold = False
'
'If mblnAccesoAsignarVendedor = True Then ' si puede asignar vendedor
'    llngCol = TraeNumCol(Me.datDatos, "VENDEDOR ASIGNADO")
'    lvarTemp = ""
'    Me.sprGrillaPrincipal.GetText llngCol, ldblRows, lvarTemp
'    If lvarTemp <> "{sin asignar}" Then
'        Me.sprGrillaPrincipal.Row = ldblRows
'        Me.sprGrillaPrincipal.Col = -1
'        Me.sprGrillaPrincipal.ForeColor = cstrColorROJO
'        Me.sprGrillaPrincipal.FontStrikethru = False
'        Me.sprGrillaPrincipal.FontBold = False
'    End If
'End If
'
'llngCol = TraeNumCol(Me.datDatos, "GANADA/PERDIDA")
'lvarTemp = ""
'Me.sprGrillaPrincipal.GetText llngCol, ldblRows, lvarTemp
'If lvarTemp = "Ganada" Then
'    Me.sprGrillaPrincipal.Row = ldblRows
'    Me.sprGrillaPrincipal.Col = -1
'    Me.sprGrillaPrincipal.ForeColor = cstrColorVERDE
'    Me.sprGrillaPrincipal.FontStrikethru = False
'    Me.sprGrillaPrincipal.FontBold = False
'End If
'If lvarTemp = "Perdida" Then
'    Me.sprGrillaPrincipal.Row = ldblRows
'    Me.sprGrillaPrincipal.Col = -1
'    Me.sprGrillaPrincipal.ForeColor = cstrColorGRIS_CLARO
'    Me.sprGrillaPrincipal.FontStrikethru = True
'    Me.sprGrillaPrincipal.FontBold = False
'End If
    
End Sub

Public Sub SetNotaSpread(dblRow As Double)
'Dim lvarTemp As Variant
'Dim llngColTxtNota As Long
'Dim llngColConNota As Long
'
'llngColTxtNota = TraeNumCol(Me.datDatos, "NOTA_INT")
'llngColConNota = TraeNumCol(Me.datDatos, "OT")
'
'With Me.sprGrillaPrincipal
'    .Row = dblRow
'    .GetText llngColTxtNota, dblRow, lvarTemp
'    If lvarTemp <> "" Then
'        .Col = llngColConNota
'        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'        .CellNote = lvarTemp
'    End If
'End With

End Sub

Private Function ValidaDocumentosR(strIdSucursal As String, dblNumeroReserva As Double)
    Dim strTipo  As String
    Dim strParametros As String
    Dim hProcess As Long
    Dim retval As Long
    
    strTipo = "RESERVA"
    
    strParametros = "ESTADOS=S,STRTIME=" & Trim(gcTiempoEspera) & ",STRINI=" & gstrArchivoIni & ",STRIDEMPRESA=" & gstrIdEmpresa & ",STRIDSUCURSAL=" & strIdSucursal & ",DBLNUMERONOTA=" & dblNumeroReserva & ",STRTIPOMODULO=" & strTipo & ""
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, _
            Shell(App.Path & "\Auto_IT.exe " & strParametros, vbNormalFocus))
    Do
        'Get the status of the process
        GetExitCodeProcess hProcess, retval
        'Sleep command recommended as well
        'as DoEvents
        DoEvents
        Sleep 100
    'Loop while the process is active
    Loop While retval = STILL_ACTIVE
        
End Function

Private Sub TraeVehiculos()
Dim strEstadoVenta As String
Dim lstrSQL As String
Dim intAñoDesde As Integer
Dim intAñoHasta As Integer
Dim dblPrecioDesde As Double
Dim dblPrecioHasta As Double
Dim strEstado As String

With frmBuscarVehiculo
    If (.txtAñoVehiculo = "") Then
        intAñoDesde = 0
    Else
        intAñoDesde = CInt(.txtAñoVehiculo)
    End If
    If (.txtAñoVehiculo2 = "") Then
        intAñoHasta = 0
    Else
        intAñoHasta = CInt(.txtAñoVehiculo2)
    End If
    
    If (.txtPrecioVehiculo = "") Then
        dblPrecioDesde = 0
    Else
        dblPrecioDesde = CDbl(SacarFormatoValor(.txtPrecioVehiculo, mstrSigla))
    End If
    If (.txtPrecioVehiculo2 = "") Then
        dblPrecioHasta = 0
    Else
        dblPrecioHasta = CDbl(SacarFormatoValor(.txtPrecioVehiculo2, mstrSigla))
    End If

    strEstadoVenta = ""
    strEstado = "S"

    'lstrSQL = "EXEC Elisa_APVentaVehiculo2 'BUSCA_STOCK',@DesdeAño= " & intAñoDesde & ",@HastaAño= " & intAñoHasta & ",@id_Marca='" & Trim$(.dbcboMarca.BoundText) & "',@id_Modelo='" & Trim$(.dbcboModelo.BoundText) & "',@id_Color_Exterior='" & Trim$(.dbcboColor.BoundText) & "',@id_Tipo_Vehiculo='" & Trim$(.dbcboTipoVehiculo.BoundText) & "',@id_Estado_Vehiculo='" & .dbcEstado.BoundText & "',@Cajon='" & Trim$(.txtNumeroCajon.Text) & "',@id_Condicion_Vehiculo='" & .dbcboEstado.BoundText & "',@DesdePrecio=" & dblPrecioDesde & ",@HastaPrecio=" & dblPrecioHasta & ",@Patente='" & Trim$(.txtPatente.Text) & "',@Fantasia='" & .txtFantasia.Text & "',@Estado_Stock='" & strEstado & "',@id_Empresa='" & gstrIdEmpresa & "',@id_Sucursal='" & Me.dbcboSucursal.BoundText & "',@Estado_Venta='" & strEstadoVenta & "',@ID_LISTA_PRECIO='" & VGlob.gstrIdListaPrecio & "'"

lstrSQL = ""
lstrSQL = lstrSQL & "SELECT "
lstrSQL = lstrSQL & "Auto_Stock.VIN, "
lstrSQL = lstrSQL & "Glbl_Marca.Descripcion AS [MARCA], "
lstrSQL = lstrSQL & "Glbl_Modelo.Descripcion AS [MODELO], "
lstrSQL = lstrSQL & "Auto_Stock.ID_Cajon_Pedido AS [PEDIDO], "
lstrSQL = lstrSQL & "Glbl_Color_Exterior.Descripcion AS [COLOR], "
lstrSQL = lstrSQL & "Auto_Stock.NumeroChasis AS [CHASIS], "
lstrSQL = lstrSQL & "Glbl_Marca.Id_Marca, "
lstrSQL = lstrSQL & "Glbl_Modelo.Id_Modelo, "
lstrSQL = lstrSQL & "Auto_Stock.Id_Color_Exterior, "
lstrSQL = lstrSQL & "Auto_Stock.NumeroMotor, "
lstrSQL = lstrSQL & "Auto_Stock.Año "

lstrSQL = lstrSQL & "FROM Auto_Stock "
lstrSQL = lstrSQL & "LEFT OUTER JOIN Glbl_Color_Exterior "
lstrSQL = lstrSQL & "ON Auto_Stock.Id_Color_Exterior = Glbl_Color_Exterior.Id_Color_Exterior "
lstrSQL = lstrSQL & "LEFT OUTER JOIN Glbl_Marca "
lstrSQL = lstrSQL & "ON Auto_Stock.Id_Marca = Glbl_Marca.Id_Marca "
lstrSQL = lstrSQL & "RIGHT OUTER JOIN Glbl_Modelo "
lstrSQL = lstrSQL & "ON Auto_Stock.Id_Marca = Glbl_Modelo.Id_Marca And Auto_Stock.Id_Modelo = Glbl_Modelo.Id_Modelo "
lstrSQL = lstrSQL & "WHERE Auto_Stock.Vigencia='S' "
If Me.txtNumeroCajon.Text <> "" Then
    lstrSQL = lstrSQL & "AND Auto_Stock.VIN LIKE '%" & Me.txtNumeroCajon.Text & "%' "
End If
lstrSQL = lstrSQL & "Order by Auto_Stock.VIN ASC "
    
    
    
    
    
End With
    
' ///////////////////////////////////////////////////
SeteaSpread

GetData Me, lstrSQL, 1

SeteaSpreadPost
   
Me.statusBar.Panels(1).Text = "Registros: " & Me.sprGrillaPrincipal.MaxRows '    dblContador
End Sub

Private Sub SeteaSpread()
Dim lintHojas As Integer

With Me.sprGrillaPrincipal
    .Reset
    
    ' crea las hojitas (sheets)
    .SheetCount = 1
    .Sheet = 1
    .SheetName = "Rendicion Fondo Fijo"
    
    For lintHojas = 1 To .SheetCount
        .Sheet = lintHojas
        .ActiveSheet = lintHojas
        SeteaSpreadSoloHoja Me.sprGrillaPrincipal
    Next lintHojas
End With

End Sub

Private Sub SeteaSpreadPost()
Dim lintHojas As Integer

Screen.MousePointer = vbHourglass

Me.sprGrillaPrincipal.Redraw = False

' hace el seteo en cada hoja (sheet)
For lintHojas = 1 To Me.sprGrillaPrincipal.SheetCount
    Me.sprGrillaPrincipal.Sheet = lintHojas
    Me.sprGrillaPrincipal.ActiveSheet = lintHojas
    SeteaSpreadPostSoloHoja Me.sprGrillaPrincipal, 5
Next lintHojas

Me.sprGrillaPrincipal.Col = 1
Me.sprGrillaPrincipal.Col2 = 1
Me.sprGrillaPrincipal.Row = -1
Me.sprGrillaPrincipal.Row2 = -1
Me.sprGrillaPrincipal.Lock = False

Me.sprGrillaPrincipal.Col = 0
Me.sprGrillaPrincipal.ColHidden = True

Me.sprGrillaPrincipal.Col = 1
Me.sprGrillaPrincipal.Row = Me.sprGrillaPrincipal.DataRowCnt
Me.sprGrillaPrincipal.Lock = False

Me.sprGrillaPrincipal.Redraw = True

Screen.MousePointer = vbDefault

End Sub

Sub MuestraVehiculoParaCompra()
    Me.Visible = False
    Me.Refresh
'    frmComprasVehiculos.Refresh
'    frmComprasVehiculos.txtCodigo = lvwListaVehiculos.SelectedItem.SubItems(2)
'    frmComprasVehiculos.txtChasis = lvwListaVehiculos.SelectedItem.SubItems(5)
'    frmComprasVehiculos.dbcboMarca = lvwListaVehiculos.SelectedItem.SubItems(6)
'    frmComprasVehiculos.dbcModelo = lvwListaVehiculos.SelectedItem.SubItems(7)
End Sub
Sub MuestraVehiculoVendido(stridCajon As String, strIdSucursal As String)
'    Dim tbRegistros As New ADODB.Recordset
'    Dim Datos As TIPO_SIGLA_PARIDAD_DECIMALES
'    Dim dblIdNumero As Double
'    Dim strSql As String
'    Dim adoTemp As New ADODB.Recordset
'
'    VGlob.gblnMenuBuscarVenta = True
'    If MsgBox("El Vehículo esta Vendido..." & Chr(13) & "¿Desea Editar la Venta?", vbYesNo + vbQuestion + vbDefaultButton2, "Advertencia") = vbYes Then
'        SW = 1
'        dblIdNumero = 0
'        strSql = "SELECT isnull(ID_NUMERO,0) as Id_Numero FROM AUTO_VENTA WHERE ID_EMPRESA='" & gstrIdEmpresa & "' AND ID_SUCURSAL='" & strIdSucursal & "' and Id_Cajon_Pedido='" & stridCajon & "' AND ESTADO_VENTA <> 'N'"
'        If apConexion.SendHost(strSql, adoTemp, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
'            If Not adoTemp.EOF And Not adoTemp.BOF Then
'                dblIdNumero = adoTemp!Id_Numero
'            End If
'        End If
'        apConexion.CloseHost adoTemp
'        If ExisteVenta(gstrIdEmpresa, stridCajon, tbRegistros, TablaVenta.TipoDocto, strIdSucursal, dblIdNumero) Then
'            Load frmVentas
'            VGlob.gblnModificarVenta = True
'            mblnTablaVacia = False
'            Datos.Sigla = tbRegistros!Sigla
'            Datos.Descripcion = tbRegistros!Descripcion
'            Datos.Paridad = tbRegistros!Paridad
'            Datos.Decimales = tbRegistros!Decimales
'            Call LeerCampos(tbRegistros, tbRegistros!Id_Numero, Datos)
'            ActivaBotones
'            Unload Me
'        Else
'            MsgBox "Existen Problemas Para Encontrar la Nota de Venta!" & Chr(13) & "Consulte más Tarde...", vbOKCancel + vbCritical, "Información"
'            Screen.MousePointer = vbDefault
'            Exit Sub
'        End If
'        VGlob.gblnModificarVenta = True
'    End If
End Sub
Function MuestraVehiculoParaVenta() As Boolean
'    Dim tbRegistros As New ADODB.Recordset
'    Dim lstrQuery As String
'    Dim dblPrecioLista As Double
'    Dim dblBonoDcto As Double
'    Dim strSql As String
'    Dim adoTemp As New ADODB.Recordset
'    Dim x As Date
'    Dim lvarTemp As Variant
'    Dim llngFila As Long
'
'    llngFila = Me.sprGrillaPrincipal.ActiveRow
'
'    MuestraVehiculoParaVenta = True
'
'    Crear
'
'    lvarTemp = ""
'    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "marca"), llngFila, lvarTemp
'    frmVentas.lblMarca.Caption = lvarTemp
'
'    frmVentas.lblBonoDcto.Caption = "0"
'
'    lvarTemp = ""
'    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "modelo"), llngFila, lvarTemp
'    frmVentas.lblModelo.Caption = lvarTemp
'
'    lvarTemp = ""
'    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "cajon"), llngFila, lvarTemp
'    frmVentas.txtPedido.Text = lvarTemp
'
'    lvarTemp = ""
'    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "chasis"), llngFila, lvarTemp
'    frmVentas.lblChasis.Caption = lvarTemp
'
'    lvarTemp = ""
'    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "color"), llngFila, lvarTemp
'    frmVentas.lblColor.Caption = lvarTemp
'
'    lvarTemp = ""
'    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "id_color_exterior"), llngFila, lvarTemp
'    frmVentas.lblColor.Tag = lvarTemp
'
'    lvarTemp = ""
'    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "id_marca"), llngFila, lvarTemp
'    frmVentas.lblCodigoMarca.Caption = lvarTemp
'
'    lvarTemp = ""
'    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "id_modelo"), llngFila, lvarTemp
'    frmVentas.lblCodigoModelo.Caption = lvarTemp
'
'
'    'kjcv 15-02-12
'    'Retorna la moneda de venta
'    lvarTemp = ""
'    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "Id_MonedaLista"), llngFila, lvarTemp
'
'    ParamGlob.strIdMonedaLocal = lvarTemp
'    frmVentas.dbcboMoneda1.BoundText = lvarTemp
'    Call frmVentas.dbcboMoneda_Click(1)
'    frmVentas.txtTipoCambio.Text = "1"
'
'    lvarTemp = ""
'    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "id_sucursal"), llngFila, lvarTemp
'    frmVentas.dbcboSucursal.BoundText = lvarTemp
'
'    frmVentas.dbcboSucursal.Enabled = False
'
'    '// Bodega
'    Set tbRegistros = New ADODB.Recordset
'    If frmVentas.dbcboSucursal.Text = "" Then
'        lstrQuery = "SELECT * FROM Glbl_Bodega WHERE Id_Sucursal =  '" & ParamGlob.strIdSucursal & "' AND Vigencia = 'S'  and tipo_BODEGA='V' ORDER BY Descripcion"
'    Else
'        lstrQuery = "SELECT * FROM Glbl_Bodega WHERE Id_Sucursal =  '" & frmVentas.dbcboSucursal.BoundText & "' AND Vigencia = 'S' AND TIPO_BODEGA='V' ORDER BY Descripcion"
'    End If
'
'    If apConexion.SendHost(lstrQuery, tbRegistros, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
'        Set frmVentas.datBodega.Recordset = tbRegistros
'    End If
'
'    lvarTemp = ""
'    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "id_bodega"), llngFila, lvarTemp
'    frmVentas.dbcboBodega.BoundText = lvarTemp
'
'    VGlob.gblnChange = True
'
'    lvarTemp = ""
'    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "precio de lista"), llngFila, lvarTemp
'    frmVentas.lblListaPrecio = lvarTemp
'
'    dblPrecioLista = lvarTemp
'
'    dblBonoDcto = 0
'
''    frmVentas.lblPrecioLista = FormatoValor(dblPrecioLista, "$", 0)
''    frmVentas.lblBonoDcto = FormatoValor(0, "$", 0)
''    frmVentas.lblListaPrecio = FormatoValor(dblPrecioLista, "$", 0)
''kjcv 15.02.12
''    frmVentas.lblPrecioLista = FormatoValor(dblPrecioLista, "$", 0)
'    frmVentas.lblPrecioLista = FormatoValor(dblPrecioLista, Trim$(ParamGlob.strSiglaMonedaLocal), 0)
''    frmVentas.lblBonoDcto = FormatoValor(0, "$", 0)
'     frmVentas.lblBonoDcto = FormatoValor(0, Trim$(ParamGlob.strSiglaMonedaLocal), 0)
''    frmVentas.lblListaPrecio = FormatoValor(dblPrecioLista, "$", 0)
'    frmVentas.lblListaPrecio = FormatoValor(dblPrecioLista, Trim$(ParamGlob.strSiglaMonedaLocal), 0)
'
'    VGlob.gblnChange = False
'
'    strSql = "SELECT ISNULL(ID_GRUPO_CENTRO_COSTO,'') AS ID_GRUPO_CENTRO_COSTO,ISNULL(ID_CENTRO_COSTO,'') AS ID_CENTRO_COSTO FROM REMU_EMPLEADO WHERE ID_EMPLEADO = '" & frmVentas.dbcboVendedor.BoundText & "'"
'
'    If apConexion.SendHost(strSql, adoTemp, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
'        If Not adoTemp.EOF And Not adoTemp.BOF Then
'            frmVentas.dbcboGrupoCosto.BoundText = adoTemp!Id_Grupo_Centro_Costo
'             '// Centro Costo
'            If Trim$(frmVentas.dbcboGrupoCosto.BoundText) <> "" Then
'               Set tbRegistros = New ADODB.Recordset
'
'               lstrQuery = "SELECT * FROM Cont_Centro_Costo WHERE id_Empresa = '" & gstrIdEmpresa & "' and id_grupo_centro_costo = '" & Trim$(frmVentas.dbcboGrupoCosto.BoundText) & "' and Vigencia = 'S' ORDER BY Nombre"
'
'               If apConexion.SendHost(lstrQuery, tbRegistros, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
'                   Set frmVentas.datResponsabilidad.Recordset = tbRegistros
'               End If
'            End If
'            frmVentas.dbcboCentroResponsabilidad.BoundText = adoTemp!Id_Centro_Costo
'        End If
'    End If
'    apConexion.CloseHost adoTemp
'
'
'
'
'' actualiza carrocería en auto_stock para joda camiones
'Dim lstrIdCarroceria As String
'Dim tbCarroceria As New ADODB.Recordset
'Dim lstrSQL As String
'lstrSQL = ""
'lstrSQL = lstrSQL & "SELECT Id_Carroceria FROM Auto_Stock WHERE Id_Cajon_Pedido = '" & frmVentas.txtPedido.Text & "' "
'If apConexion.SendHost(lstrSQL, tbCarroceria, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
'    If Not tbCarroceria.BOF And Not tbCarroceria.EOF Then
'        If IsNull(tbCarroceria.Fields("Id_Carroceria")) Or Trim$(ValorNulo(tbCarroceria.Fields("Id_Carroceria"), "")) = "" Then
'            lstrIdCarroceria = Retorna_Valor_General("select id_carroceria from glbl_modelo where id_modelo='" & frmVentas.lblCodigoModelo.Caption & "'")
'            lstrSQL = ""
'            lstrSQL = lstrSQL & "UPDATE Auto_Stock SET Id_Carroceria = '" & lstrIdCarroceria & "' WHERE Id_Cajon_Pedido = '" & frmVentas.txtPedido.Text & "' "
'            apConexion.SendHost lstrSQL, , , , gcTiempoEspera
'        End If
'    End If
'End If
'apConexion.CloseHost tbCarroceria
    
End Function
Private Function MostrarProductosCompra(stridCajon As String, strIdSucursal As String, strIdEmpresa As String)
'    Dim item As ListItem
'    Dim tbRegistros As New ADODB.Recordset
'    Dim strSql As String
'    Dim dblSubTotal As Double
'    Set tbRegistros = New ADODB.Recordset
'    Dim strQuery As String
'    Dim adoTemp As New ADODB.Recordset
'    Dim strModeloCambiado As String
'
'    strSql = "EXEC Elisa_APVentaVehiculo 'PRODUCTOSDESTOCK',@ID_EMPRESA='" & strIdEmpresa & "',@ID_SUCURSAL='" & strIdSucursal & "',@CAJON='" & stridCajon & "'"
'
'    If apConexion.SendHost(strSql, tbRegistros, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
'        If Not tbRegistros.BOF And Not tbRegistros.EOF Then
'            'If MsgBox("El Vehículo Seleccionado, Tiene Accesorios Asociados..." & Chr(13) & "Desea Agregarlos a la Venta?", vbYesNo + vbQuestion + vbDefaultButton2, "Accesorios") = vbYes Then
'                tbRegistros.MoveFirst
'                Do While Not tbRegistros.EOF
'                    Set item = frmVentas.lvwProductos.ListItems.Add(, , ValorNulo(tbRegistros!Id_Item))
'                    item.SubItems(1) = ValorNulo(tbRegistros!DescripcionProducto)
'                    item.SubItems(2) = ValorNulo(Numero, tbRegistros!cantidad)
'
'                    strQuery = "Select isnull(ModeloCambiado,'N') as ModeloCambiado "
'                    strQuery = strQuery & " FROM AUTO_STOCK "
'                    strQuery = strQuery & " WHERE ID_EMPRESA = '" & strIdEmpresa & "' and"
'                    strQuery = strQuery & " ID_SUCURSAL = '" & strIdSucursal & "' and"
'                    strQuery = strQuery & " Id_CAJON_PEDIDO = '" & stridCajon & "'"
'
'                    Set adoTemp = New ADODB.Recordset
'
'                    If apConexion.SendHost(strQuery, adoTemp, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
'                        If Not adoTemp.EOF And Not adoTemp.BOF Then
'                            strModeloCambiado = adoTemp!ModeloCambiado
'                        End If
'                    End If
'
'                    apConexion.CloseHost adoTemp
'
'
''                    If ValorNulo(Numero, tbRegistros!Precio_Venta) = 0 Then
''
''                        If strModeloCambiado = "S" And VGlob.gblnCambioModelo Then
''                            item.SubItems(3) = FormatoValor(0, tbRegistros!siglamonedaventa, tbRegistros!decimalesMonedaventa)
''                        Else
''                            Do
''                                item.SubItems(3) = InputBox("Ingrese el Precio de Venta del Accesorio: " & Chr(13) & Chr(10) & "Accesorio: " & ValorNulo(tbRegistros!DescripcionProducto) & Chr(13) & Chr(10) & "Costo: " & SacarFormatoValor(tbRegistros!Precio_Compra, tbRegistros!siglaMonedaCompra, tbRegistros!DecimalesMonedaCompra), "Venta de Accesorio", 0)
''
''                                If IsNumeric(item.SubItems(3)) Then
''                                    If Val(item.SubItems(3)) > 0 Then
''                                        item.SubItems(3) = FormatoValor(Item.SubItems(3), tbRegistros!siglamonedaventa, tbRegistros!decimalesMonedaventa)
''                                        Exit Do
''                                    End If
''                                End If
''                            Loop
''                        End If
''                    Else
'                        item.SubItems(3) = FormatoValor(ValorNulo(Numero, tbRegistros!Precio_Venta), tbRegistros!siglamonedaventa, tbRegistros!decimalesMonedaventa)
''                    End If
'                    item.SubItems(4) = FormatoValor(ValorNulo(Numero, tbRegistros!Descto_Recgo), "%", 2)
'                    If VGlob.gblnUsaInterface Then
'                        If tbRegistros!Genera_OT = "S" Then
'                            item.SubItems(5) = "Factura Dpto. Ventas"
'                        Else
'                            item.SubItems(5) = "Automotriz"
'                        End If
'                    Else
'                        item.SubItems(5) = "Automotriz"
'                    End If
'                    dblSubTotal = SacarFormatoValor(item.SubItems(3), tbRegistros!siglamonedaventa) * ValorNulo(Numero, tbRegistros!cantidad) 'ValorNulo(Numero, tbRegistros!Precio_Venta) * ValorNulo(Numero, tbRegistros!Cantidad)
'                    item.SubItems(6) = FormatoValor(dblSubTotal, tbRegistros!siglamonedaventa, tbRegistros!decimalesMonedaventa)
'                    item.SubItems(7) = ValorNulo(tbRegistros!DescripcionMonedaVenta)
'                    item.SubItems(8) = tbRegistros!siglamonedaventa
'                    item.SubItems(9) = ValorNulo(Numero, tbRegistros!ParidadMonedaVenta)
'                    item.SubItems(10) = tbRegistros!idMonedaventa
'                    item.SubItems(11) = tbRegistros!decimalesMonedaventa
'
'                    If ValorNulo(Numero, tbRegistros!Precio_Compra) = 0 Then
'                        Do
'                            item.SubItems(12) = InputBox("Ingrese el Costo del Accesorio: " & Chr(13) & Chr(10) & "Accesorio: " & ValorNulo(tbRegistros!DescripcionProducto) & Chr(13) & Chr(10) & "Precio de Venta: " & SacarFormatoValor(tbRegistros!Precio_Venta, tbRegistros!siglamonedaventa, tbRegistros!decimalesMonedaventa), "Costo de Accesorio", 0)
'
'                            If IsNumeric(item.SubItems(12)) Then
'                                If Val(item.SubItems(12)) > 0 Then
'                                    item.SubItems(12) = FormatoValor(item.SubItems(12), tbRegistros!siglaMonedaCompra, tbRegistros!DecimalesMonedaCompra)
'                                    Exit Do
'                                End If
'                            End If
'                        Loop
'                    Else
'                        item.SubItems(12) = FormatoValor(ValorNulo(Numero, tbRegistros!Precio_Compra), tbRegistros!siglaMonedaCompra, tbRegistros!DecimalesMonedaCompra)
'                    End If
'
'                    item.SubItems(13) = tbRegistros!ID_OT
'                    If tbRegistros!Genera_OT = "S" Then
'                        item.SubItems(14) = "Si"
'                    Else
'                        item.SubItems(14) = "No"
'                    End If
'                    item.SubItems(15) = "No"
'                    item.SubItems(16) = ""
'                    item.SubItems(17) = ""
'                    item.SubItems(18) = "0"
'                    item.SubItems(19) = " "
'                    item.SubItems(20) = " "
'                    tbRegistros.MoveNext
'                Loop
'            End If
'        'End If
'    End If
'    VGlob.gblnChange = True
'    Call CalculaProductos
'    VGlob.gblnChange = False
'    apConexion.CloseHost tbRegistros
'
'Dim lstrCodColor As String
'Dim R, G, b As Byte
'
'If frmVentas.lblColor.Tag <> "" Then
'    lstrCodColor = CStr(Retorna_Valor_General("select CodigoRGB from Glbl_Color_Exterior where id_color_exterior = '" & frmVentas.lblColor.Tag & "'"))
'
'    R = CByte("&H" & Mid(lstrCodColor, 1, 2))
'    G = CByte("&H" & Mid(lstrCodColor, 3, 2))
'    b = CByte("&H" & Mid(lstrCodColor, 5, 2))
'
'    frmVentas.lblElColor.BackColor = RGB(R, G, b)
'End If
End Function
Sub DesactivaBotones()
'    frmComprasVehiculos.tlbBarraHerramientas.Buttons("Crear").Enabled = False
'    frmComprasVehiculos.tlbBarraHerramientas.Buttons("Cancelar").Enabled = False
'    frmComprasVehiculos.tlbBarraHerramientas.Buttons("Cancelar").Enabled = False
'    frmComprasVehiculos.tlbBarraHerramientas.Buttons("Borrar").Enabled = False
'    frmComprasVehiculos.tlbBarraHerramientas.Buttons("Buscar").Enabled = False
'    frmComprasVehiculos.tlbBarraHerramientas.Buttons("Primero").Enabled = False
'    frmComprasVehiculos.tlbBarraHerramientas.Buttons("Anterior").Enabled = False
'    frmComprasVehiculos.tlbBarraHerramientas.Buttons("Siguiente").Enabled = False
'    frmComprasVehiculos.tlbBarraHerramientas.Buttons("Ultimo").Enabled = False
'    frmComprasVehiculos.tlbBarraHerramientas.Buttons("Renovar").Enabled = False
'    frmComprasVehiculos.tlbBarraHerramientas.Buttons("CambiarCajon").Enabled = False
'    frmComprasVehiculos.FraUbicacion.Enabled = False
End Sub
Sub InhabilitaVenta()
'    frmComprasVehiculos.tlbBarraHerramientas.Visible = False
'    frmComprasVehiculos.cmdEliminar.Enabled = False
'    frmComprasVehiculos.cmdEliminarFormaPago.Enabled = False
'    frmComprasVehiculos.cmdEliminarGasto.Enabled = False
'    frmComprasVehiculos.cmdEliminarFormaPago.Enabled = False
'    frmComprasVehiculos.cmdAgregarFormaPago.Enabled = False
'    frmComprasVehiculos.cmdAgregarGasto.Enabled = False
'    frmComprasVehiculos.cmdSeleccionar.Enabled = False
'    frmComprasVehiculos.fraComentario.Enabled = False
'    frmComprasVehiculos.FraCompra.Enabled = False
'    frmComprasVehiculos.FraDatosTecnicos.Enabled = False
'    frmComprasVehiculos.FraDatosVehiculo.Enabled = False
'    frmComprasVehiculos.FraUbicacion.Enabled = False
'    frmComprasVehiculos.fraLista.Enabled = False
'    frmComprasVehiculos.lvwDetalleFormaPago.Enabled = False
'    frmComprasVehiculos.lvwGastos.Enabled = False
'    frmComprasVehiculos.lvwProductos.Enabled = False
End Sub
Sub LimpiarBuscarVehiculo()
    Me.dbcboColor.Text = ""
    Me.txtAñoVehiculo = ""
    Me.txtAñoVehiculo2 = ""
    Me.txtNumeroCajon = ""
    Me.txtPatente = ""
    Me.txtPrecioVehiculo = FormatoValor(0, mstrSigla, 2)
    Me.txtPrecioVehiculo2 = FormatoValor(0, mstrSigla, 2)
    Me.cboEstadoStock.ListIndex = 0

End Sub
Private Sub BarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Screen.MousePointer = vbHourglass
    Me.SetFocus
    Select Case Button.Key
        Case "Crear"
            LimpiarBuscarVehiculo
        Case "Buscar"
            TraeVehiculos
        Case "Columnas"
            'ObjHideColumnHeader Me.lvwListaVehiculos
        Case "Cerrar"
            Unload Me
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub cboEstadoStock_Click()
    If Me.cboEstadoStock.Text = "En Stock" Then
        Label11.Enabled = False
        Label12.Enabled = False
        Me.dtpDesde.Enabled = False
        Me.dtpHasta.Enabled = False
        Me.cmdLimpiaFecha1.Enabled = False
        Me.cmdLimpiaFecha2.Enabled = False
    Else
        Label11.Enabled = True
        Label12.Enabled = True
        Me.dtpDesde.Enabled = True
        Me.dtpHasta.Enabled = True
        Me.cmdLimpiaFecha1.Enabled = True
        Me.cmdLimpiaFecha2.Enabled = True
    End If
End Sub

Private Sub cmdLimpiarColor_Click()
    dbcboColor.Text = ""
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSeleccionar_Click()
If Me.datDatos.Recordset.RecordCount > 0 Then
    Selecciona Me.sprGrillaPrincipal.ActiveRow
End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If mSw Then
        mSw = False
        
'        If Atributos("Auto_10_0016_02", True, True, True, True) = True Then
'            Me.dbcboSucursal.Enabled = False
'            Me.tlbBtnSmall(0).Buttons("Limpiar").Enabled = False
'        End If
        
        TraeVehiculos
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{tab}"
        End Select
End Sub

Private Sub Form_Load()
    Dim adoRecordset As New ADODB.Recordset
    Dim strSql As String
    Dim item As ListItem
    Dim lblnBoolean As Boolean
    
    mSw = True
    
End Sub

Private Sub Form_Resize()
'Dim ldblAncho As Double
'Dim ldblAnchoCol As Double
'Dim ldblAnchoBtnSmall As Double
'
'Screen.MousePointer = vbHourglass
'
'ldblAncho = 120
'ldblAnchoBtnSmall = 240
'
'Me.fmeCriterios.Left = ldblAncho
'Me.fmeCriterios.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0)
'
'ldblAnchoCol = IIf(((Me.fmeCriterios.Width - ldblAncho) / 3) - ldblAncho >= 0, ((Me.fmeCriterios.Width - ldblAncho) / 3) - ldblAncho, 0)
'Me.dbcboSucursal.Left = ldblAncho
'Me.lblSucursal.Left = Me.dbcboSucursal.Left
'Me.dbcboSucursal.Width = ldblAnchoCol
'Me.tlbBtnSmall(0).Left = Me.dbcboSucursal.Left + Me.dbcboSucursal.Width - ldblAnchoBtnSmall
''
'Me.dbcboMarca.Left = Me.dbcboSucursal.Left + Me.dbcboSucursal.Width + ldblAncho
'Me.lblMarca.Left = Me.dbcboMarca.Left
'Me.dbcboMarca.Width = ldblAnchoCol
'Me.tlbBtnSmall(1).Left = Me.dbcboMarca.Left + Me.dbcboMarca.Width - ldblAnchoBtnSmall
''
'Me.dbcboModelo.Left = Me.dbcboMarca.Left + Me.dbcboMarca.Width + ldblAncho
'Me.lblModelo.Left = Me.dbcboModelo.Left
'Me.dbcboModelo.Width = ldblAnchoCol
'Me.tlbBtnSmall(2).Left = Me.dbcboModelo.Left + Me.dbcboModelo.Width - ldblAnchoBtnSmall
'
'
'ldblAnchoCol = IIf(((Me.fmeCriterios.Width - ldblAncho) / 4) - ldblAncho >= 0, ((Me.fmeCriterios.Width - ldblAncho) / 4) - ldblAncho, 0)
'Me.txtNumeroCajon.Left = ldblAncho
'Me.lblCajon.Left = Me.txtNumeroCajon.Left
'Me.txtNumeroCajon.Width = ldblAnchoCol
''
'Me.dbcboEstado.Left = Me.txtNumeroCajon.Left + Me.txtNumeroCajon.Width + ldblAncho
'Me.lblCondicion.Left = Me.dbcboEstado.Left
'Me.dbcboEstado.Width = ldblAnchoCol
'Me.tlbBtnSmall(3).Left = Me.dbcboEstado.Left + Me.dbcboEstado.Width - ldblAnchoBtnSmall
''
'Me.dbcboTipoVehiculo.Left = Me.dbcboEstado.Left + Me.dbcboEstado.Width + ldblAncho
'Me.lblTipoVehiculo.Left = Me.dbcboTipoVehiculo.Left
'Me.dbcboTipoVehiculo.Width = ldblAnchoCol
'Me.tlbBtnSmall(4).Left = Me.dbcboTipoVehiculo.Left + Me.dbcboTipoVehiculo.Width - ldblAnchoBtnSmall
''
'Me.dbcEstado.Left = Me.dbcboTipoVehiculo.Left + Me.dbcboTipoVehiculo.Width + ldblAncho
'Me.lblestado.Left = Me.dbcEstado.Left
'Me.dbcEstado.Width = ldblAnchoCol
'Me.tlbBtnSmall(5).Left = Me.dbcEstado.Left + Me.dbcEstado.Width - ldblAnchoBtnSmall
'
'
'Me.sprGrillaPrincipal.Left = ldblAncho
'Me.sprGrillaPrincipal.Width = IIf(Me.ScaleWidth - (ldblAncho * 2) >= 0, Me.ScaleWidth - (ldblAncho * 2), 0)
'Me.sprGrillaPrincipal.Height = IIf(Me.ScaleHeight - Me.sprGrillaPrincipal.Top - Me.fmeBotones.Height - Me.pb1.Height - Me.statusBar.Height - ldblAncho >= 0, Me.ScaleHeight - Me.sprGrillaPrincipal.Top - Me.fmeBotones.Height - Me.pb1.Height - Me.statusBar.Height - ldblAncho, 0)
'
'Me.fmeBotones.Top = Me.sprGrillaPrincipal.Top + Me.sprGrillaPrincipal.Height + ldblAncho
'Me.fmeBotones.Left = Me.ScaleWidth - Me.fmeBotones.Width - ldblAncho
'Me.fmeBotones.ZOrder 0
'
'Screen.MousePointer = vbDefault
End Sub

Private Sub Selecciona(lngFila As Long)
Dim lvarTemp As Variant
Dim lstrIdCajon As String
Dim lstrIdSucursal As String

Screen.MousePointer = vbHourglass
    
If Me.sprGrillaPrincipal.MaxRows > 0 Then
    
    lvarTemp = ""
    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "VIN"), Me.sprGrillaPrincipal.ActiveRow, lvarTemp
    frmMantenedorVehiculoCliente.txtPatente.Text = lvarTemp
    
    lvarTemp = ""
    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "id_marca"), Me.sprGrillaPrincipal.ActiveRow, lvarTemp
    frmMantenedorVehiculoCliente.lblMarca.Tag = lvarTemp
   
    lvarTemp = ""
    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "marca"), Me.sprGrillaPrincipal.ActiveRow, lvarTemp
    frmMantenedorVehiculoCliente.lblMarca.Caption = lvarTemp
    
    lvarTemp = ""
    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "id_modelo"), Me.sprGrillaPrincipal.ActiveRow, lvarTemp
    frmMantenedorVehiculoCliente.lblModelo.Tag = lvarTemp
    
    lvarTemp = ""
    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "modelo"), Me.sprGrillaPrincipal.ActiveRow, lvarTemp
    frmMantenedorVehiculoCliente.lblModelo.Caption = lvarTemp
    
    lvarTemp = ""
    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "id_color_exterior"), Me.sprGrillaPrincipal.ActiveRow, lvarTemp
    frmMantenedorVehiculoCliente.lblColorExt.Tag = lvarTemp
    
    lvarTemp = ""
    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "color"), Me.sprGrillaPrincipal.ActiveRow, lvarTemp
    frmMantenedorVehiculoCliente.lblColorExt.Caption = lvarTemp

    lvarTemp = ""
    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "chasis"), Me.sprGrillaPrincipal.ActiveRow, lvarTemp
    frmMantenedorVehiculoCliente.txtNroChasis.Text = lvarTemp
    
    lvarTemp = ""
    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "NumeroMotor"), Me.sprGrillaPrincipal.ActiveRow, lvarTemp
    frmMantenedorVehiculoCliente.txtNroMotor.Text = lvarTemp

    lvarTemp = ""
    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "vin"), Me.sprGrillaPrincipal.ActiveRow, lvarTemp
    frmMantenedorVehiculoCliente.txtNroVin.Text = lvarTemp
    
    lvarTemp = ""
    Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "año"), Me.sprGrillaPrincipal.ActiveRow, lvarTemp
    frmMantenedorVehiculoCliente.txtAño.Text = lvarTemp
    
'    TablaVenta.IdSucursal = lvarTemp
'    TablaVenta.id_Vendedor = ParamGlob.strRutEmpleado   ' ParamGlob.strRutEmpleado
'    TablaVenta.TipoDocto = "V"
'
'    VGlob.gblnVehiculoVendido = False
'
'    If MuestraVehiculoParaVenta Then
'        lvarTemp = ""
''        Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "cajon"), lngFila, lvarTemp
''        'kjcv 06/06/12
'        Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "pedido"), lngFila, lvarTemp
'        lstrIdCajon = lvarTemp
'        'kjcv 07.06.12
'        frmVentas.txtPedido = lstrIdCajon
'
'        lvarTemp = ""
'        Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "id_sucursal"), lngFila, lvarTemp
'        lstrIdSucursal = lvarTemp
'
'        MostrarProductosCompra lstrIdCajon, lstrIdSucursal, gstrIdEmpresa
'
'        lvarTemp = ""
'        Me.sprGrillaPrincipal.GetText TraeNumCol(Me.datDatos, "id_color_exterior"), Me.sprGrillaPrincipal.ActiveRow, lvarTemp
'        PoneColorAuto CStr(lvarTemp)

        Unload Me
'    End If
End If

Screen.MousePointer = vbDefault

End Sub
Private Function MostrarAccesoriosVenta(strSucursal As String, dblNumeroCotiza As Double)
'    Dim item As ListItem
'    Dim lstrSQL As String
'    Dim adoTemp As New ADODB.Recordset
'
'
'    Set adoTemp = New ADODB.Recordset
'    lstrSQL = ""
'    lstrSQL = "SELECT Glbl_Moneda.*, Auto_Cotizacion_Accesorios.*,"
'    lstrSQL = lstrSQL & " Stck_Item.Descripcion AS DescripcionProducto,"
'    lstrSQL = lstrSQL & " isnull(Stck_Item.Genera_Ot,'N') AS Genera_Ot"
'    lstrSQL = lstrSQL & " FROM Auto_Cotizacion_Accesorios LEFT OUTER JOIN"
'    lstrSQL = lstrSQL & " Glbl_Moneda ON"
'    lstrSQL = lstrSQL & " Auto_Cotizacion_Accesorios.Id_Moneda = Glbl_Moneda.Id_Moneda LEFT"
'    lstrSQL = lstrSQL & " Outer Join"
'    lstrSQL = lstrSQL & " Stck_Item ON"
'    lstrSQL = lstrSQL & " Auto_Cotizacion_Accesorios.Id_Item = Stck_Item.Id_Item"
'    lstrSQL = lstrSQL & " WHERE (Auto_Cotizacion_Accesorios.Id_Empresa = '" & gstrIdEmpresa & "') AND"
'    lstrSQL = lstrSQL & " (Auto_Cotizacion_Accesorios.Id_Sucursal = '" & strSucursal & "') AND"
'    lstrSQL = lstrSQL & " (Auto_Cotizacion_Accesorios.Id_Numero = " & dblNumeroCotiza & ")"
'
'    If apConexion.SendHost(lstrSQL, adoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
'        If Not adoTemp.BOF And Not adoTemp.EOF Then
'            adoTemp.MoveFirst
'            Do While Not adoTemp.EOF
'                Set item = frmVentas.lvwProductos.ListItems.Add(, , adoTemp!Id_Item)
'                item.SubItems(1) = adoTemp!DescripcionProducto
'                item.SubItems(2) = adoTemp!cantidad
'                item.SubItems(3) = FormatoValor(adoTemp!Precio_Venta, adoTemp!Sigla, adoTemp!Decimales)
'                item.SubItems(4) = FormatoValor(adoTemp!Descto_Recargo, "%", 2)
'                If VGlob.gblnUsaInterface Then
'                    If adoTemp!Genera_OT = "S" Then
'                        item.SubItems(5) = IIf(adoTemp!Cancela = "C", "Factura Servicios", "Factura Dpto. Ventas")
'                    Else
'                        item.SubItems(5) = IIf(adoTemp!Cancela = "C", "Cliente", "Automotriz")
'                    End If
'                Else
'                    item.SubItems(5) = IIf(adoTemp!Cancela = "C", "Cliente", "Automotriz")
'                End If
'                item.SubItems(6) = FormatoValor(adoTemp!Subtotal, adoTemp!Sigla, adoTemp!Decimales)
'                item.SubItems(7) = adoTemp!Descripcion
'                item.SubItems(8) = adoTemp!Sigla
'                item.SubItems(9) = adoTemp!Paridad
'                item.SubItems(10) = adoTemp!Id_Moneda
'                item.SubItems(11) = adoTemp!Decimales
'                item.SubItems(12) = " "
'                item.SubItems(13) = " "
'                If adoTemp!Genera_OT = "S" Then
'                    item.SubItems(14) = "Si"
'                Else
'                    item.SubItems(14) = "No"
'                End If
'                item.SubItems(15) = "No"
'                item.SubItems(16) = " "
'                item.SubItems(17) = " "
'                item.SubItems(18) = "0"
'                item.SubItems(19) = " "
'                item.SubItems(20) = " "
'                adoTemp.MoveNext
'            Loop
'        End If
'    End If
'
'    VGlob.gblnChange = True
'    Call CalculaProductos
'    VGlob.gblnChange = False
'    apConexion.CloseHost adoTemp
End Function


Private Sub sprGrillaPrincipal_DblClick(ByVal Col As Long, ByVal Row As Long)
If Me.datDatos.Recordset.RecordCount > 0 Then
    Selecciona Row
End If
End Sub


Private Sub tlbTitulo_ButtonClick(ByVal Button As MSComctlLib.Button)
'Dim lstrSQL As String
'Dim lvarPaso As Variant
'
'Select Case Button.Key
'    Case "Imprimir"
'        sprLib.Spread_Imprimir Me.sprGrillaPrincipal, ParamGlob.strNombreEmpresa, "Búsqueda de Vehículos", True, ParamGlob.strApp
'    Case "Preview"
'        sprLib.Spread_Preview Me.sprGrillaPrincipal, ParamGlob.strNombreEmpresa, "Búsqueda de Vehículos", True, ParamGlob.strApp
'    Case "Excel"
'        sprLib.Spread_ExportarExcel Me.sprGrillaPrincipal, ""
'    Case "Copiar"
'        sprLib.Spread_Copiar Me.sprGrillaPrincipal, VGlob.spr_Block_Col, VGlob.spr_Block_Col2, VGlob.spr_Block_Row, VGlob.spr_Block_Row2
'    Case "Ordenar"
'        sprLib.Spread_Ordenar Me.sprGrillaPrincipal, 1, False, False
'    Case "Vertical"
'        sprLib.Spread_Detalle Me.sprGrillaPrincipal, ParamGlob.strNombreEmpresa, "Búsqueda de Vehículos", ParamGlob.strApp, ""
'End Select
End Sub

Private Sub txtAñoVehiculo_GotFocus()
     MarcaTexto txtAñoVehiculo
End Sub

Private Sub txtAñoVehiculo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        Exit Sub
    End If
    If (KeyAscii >= 48 And KeyAscii <= 57) Then
        Exit Sub
    Else
        If KeyAscii = 46 Then   '//Punto...
            Exit Sub
        End If
    End If
    KeyAscii = 0
End Sub

Private Sub txtAñoVehiculo2_GotFocus()
     MarcaTexto txtAñoVehiculo2
End Sub

Private Sub txtañovehiculo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        Exit Sub
    End If
    If (KeyAscii >= 48 And KeyAscii <= 57) Then
        Exit Sub
    Else
        If KeyAscii = 46 Then   '//Punto...
            Exit Sub
        End If
    End If
    KeyAscii = 0
End Sub

Private Sub txtNumeroCajon_Click()
MarcaTexto Me.txtNumeroCajon
End Sub

Private Sub txtNumeroCajon_GotFocus()
MarcaTexto Me.txtNumeroCajon
End Sub

Private Sub txtNumeroCajon_KeyPress(KeyAscii As Integer)

KeyAscii = Asc(UCase(Chr(KeyAscii)))


If KeyAscii < 65 Or KeyAscii > 90 Then
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 And KeyAscii <> 209 And KeyAscii <> 13 Then ' 8=borrar, 209=Ñ, 13=enter
            KeyAscii = 0
        End If
    End If
End If

End Sub

Private Sub txtPrecioVehiculo_GotFocus()
    txtPrecioVehiculo = SacarFormatoValor(txtPrecioVehiculo, mstrSigla)
    MarcaTexto txtPrecioVehiculo
End Sub

Private Sub txtPrecioVehiculo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        Exit Sub
    End If
    If (KeyAscii >= 48 And KeyAscii <= 57) Then
        Exit Sub
    Else
        If KeyAscii = 46 Then   '//Punto...
            Exit Sub
        End If
    End If
    KeyAscii = 0
End Sub

Private Sub txtPrecioVehiculo_LostFocus()
     txtPrecioVehiculo = FormatoValor(txtPrecioVehiculo, mstrSigla, 2)
End Sub


Private Sub txtPrecioVehiculo2_GotFocus()
    txtPrecioVehiculo2 = SacarFormatoValor(txtPrecioVehiculo2, mstrSigla)
    MarcaTexto txtPrecioVehiculo2
End Sub

Private Sub txtPrecioVehiculo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        Exit Sub
    End If
    If (KeyAscii >= 48 And KeyAscii <= 57) Then
        Exit Sub
    Else
        If KeyAscii = 46 Then   '//Punto...
            Exit Sub
        End If
    End If
    KeyAscii = 0
End Sub

Private Sub txtPrecioVehiculo2_LostFocus()
     txtPrecioVehiculo2 = FormatoValor(txtPrecioVehiculo2, mstrSigla, 2)
End Sub
Function CodigoCondicion(strCajon As String) As String
    Dim strSql As String
    Dim adoTemp As New ADODB.Recordset
    
    strSql = "SELECT    isnull(Id_Condicion_Vehiculo,'') as Id_Condicion_Vehiculo"
    strSql = strSql & " From dbo.Auto_Stock"
    strSql = strSql & " WHERE Id_Cajon_Pedido='" & strCajon & "'"
    
    If apConexion.SendHost(strSql, adoTemp, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
        If Not adoTemp.EOF And Not adoTemp.BOF Then
            CodigoCondicion = adoTemp!Id_Condicion_Vehiculo
        End If
    End If
    apConexion.CloseHost adoTemp
    
End Function
