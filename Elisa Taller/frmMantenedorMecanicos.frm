VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmMantenedorMecanicos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mecánicos"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frmMantenedorMecanicos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5175
      Begin VB.TextBox txtRut 
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
         Left            =   1320
         MaxLength       =   11
         TabIndex        =   16
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox cckLiquidador 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Liquidador"
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
         Left            =   3840
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dtcSupervisor 
         Bindings        =   "frmMantenedorMecanicos.frx":038A
         Height          =   315
         Left            =   1320
         TabIndex        =   14
         Top             =   1920
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
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
      Begin VB.TextBox TxtValorHora 
         Alignment       =   1  'Right Justify
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
         Left            =   1320
         TabIndex        =   11
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox cckRecepcionista 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Recepcionista"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CheckBox cckSupervisor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Supervisor"
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
         Left            =   2040
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Especialidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   4935
         Begin MSComctlLib.ListView lvwEspecialidad 
            Height          =   2895
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   5106
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "Codigo"
               Text            =   "Código"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "Des"
               Text            =   "Descripción"
               Object.Width           =   6174
            EndProperty
         End
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   2
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   0
         Top             =   240
         Width           =   1695
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
         Left            =   3960
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc datSupervisor 
         Height          =   330
         Left            =   3720
         Top             =   2040
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
      Begin Crystal.CrystalReport rptMantenedor 
         Left            =   3600
         Top             =   480
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
      Begin VB.Label label3 
         Caption         =   "D.N.I"
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
         TabIndex        =   17
         Top             =   640
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Valor Hora"
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
         TabIndex        =   12
         Top             =   2440
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Supervisor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   1960
         Width           =   930
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre"
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
         TabIndex        =   5
         Top             =   1000
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
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
         TabIndex        =   4
         Top             =   280
         Width           =   975
      End
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
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
            Key             =   "Password"
            Object.ToolTipText     =   "Cambiar Password"
            ImageKey        =   "Seleccion1"
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
            Picture         =   "frmMantenedorMecanicos.frx":03A6
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":04B8
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":05CA
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":06DC
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":07EE
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":0900
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":0A12
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":0B24
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":0C36
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":0D48
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":0E5A
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":0F6C
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":107E
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":1190
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":12A2
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":13B4
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":14C6
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":1918
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":1D6A
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":1E7C
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":1FD8
            Key             =   "Actualizar"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":2134
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":2290
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":23EC
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":2EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":330C
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":3470
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":38CC
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":3A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":4D34
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":52D0
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":542C
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":5588
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":58DC
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":5C30
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":5F84
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":62D8
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":662C
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":6B70
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":6C82
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":6FD6
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":732A
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":743E
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":7792
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":78A4
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorMecanicos.frx":7BF8
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMantenedorMecanicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoPrincipal As New ADODB.Recordset

Dim mstrSql As String
Dim mblnTablaVacia As Boolean

Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean

Dim mblnSW As Boolean

Const mcNombreTabla = "Tllr_Mecanicos"
Const mcCampoCodigo = "Id_Mecanico"
Const mcCampoNombre = "Nombre"


Private Sub Llena_Especialidades()
Dim Item As ListItem

mstrSql = "SELECT Id_Especialidad, Descripcion FROM Tllr_Especialidad WHERE Vigencia = 'S' "
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            While Not .EOF
                Set Item = lvwEspecialidad.ListItems.Add(, , !ID_ESPECIALIDAD)
                Item.SubItems(1) = !Descripcion
                .MoveNext
            Wend
        End If
    End With
End If

End Sub
Private Sub Especialidades_Mecanico(strCodigoMecanico As String)

mstrSql = "SELECT Id_Especialidad FROM Tllr_Especialidad_Mecanico WHERE Id_Mecanico = '" & strCodigoMecanico & "'"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera) = apOk Then
    With adoPrincipal
        If Not .BOF And Not .EOF Then
            While Not .EOF
                Set lvwEspecialidad.SelectedItem = lvwEspecialidad.FindItem(CStr(!ID_ESPECIALIDAD), , , 1)
                lvwEspecialidad.SelectedItem.Checked = True
                .MoveNext
            Wend
        End If
    End With
End If

End Sub
Private Sub GuardaEspecialidad(strMecanico As String)
Dim x As Integer

mstrSql = "DELETE FROM TLLR_ESPECIALIDAD_MECANICO WHERE ID_MECANICO ='" & strMecanico & "' "
Conexion.SendHost mstrSql, , , , gcTiempoEspera '//////////AQUI BORRA LAS QUE EXISTEN

For x = 1 To Me.lvwEspecialidad.ListItems.Count
    Set lvwEspecialidad.SelectedItem = lvwEspecialidad.ListItems(x)
    If lvwEspecialidad.SelectedItem.Checked = True Then
        mstrSql = "INSERT INTO TLLR_ESPECIALIDAD_MECANICO (Id_Empresa,Id_sucursal, ID_ESPECIALIDAD, ID_MECANICO)"
        mstrSql = mstrSql & " VALUES('" & gstrIdEmpresa & "','" & gstrIdSucursal & "','" & lvwEspecialidad.SelectedItem & "' , '" & strMecanico & "' ) "
        Conexion.SendHost mstrSql, , , , gcTiempoEspera
    End If
Next '///////////////AQUI GRABA LAS NUEVAS Y LAS QUE ESTABAN

End Sub

Private Sub Supervisores()
    mstrSql = "SELECT Id_Mecanico AS Codigo, Nombre FROM Tllr_Mecanicos Where Es_supervisor= 'S' "
    dtcSupervisor.Enabled = True
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        With datSupervisor
            Set .Recordset = adoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcSupervisor.ListField = "Nombre"
                dtcSupervisor.BoundColumn = "Codigo"
'                dtcSupervisor.BoundText = .Recordset!Codigo
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
End Sub


Private Sub Form_Load()
mblnSW = True
Me.label3.Caption = UCase(gstrNombreRut)
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
        Case "Password"
            CambiaPassword
    End Select
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Activate()
    If mblnSW Then
        If Not Atributos("Glbl", "Tllr_10_0030", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If
        Llena_Especialidades
        Supervisores
        If gapAccion = apcrear Then
           AgregarRegistro
           txtCodigo = gstrBusca
        End If
        If gapAccion = apeditar Then
            If gstrBusca <> "" Then
                mstrSql = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' order by " & mcCampoCodigo
                If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                        LeerCampos
                    End If
                End If
                Conexion.CloseHost adoPrincipal
            End If
            txtCodigo.Enabled = False
            Me.SetFocus
        End If
        If gapAccion = apninguno Then
           Renovar
        End If
    End If
    gapAccion = apninguno
    mblnSW = False
    txtNombre.SetFocus
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
    'txtCodigo.SetFocus
    Me.txtNombre.SetFocus
End Sub
Private Sub CancelarAgregaRegistro()
    Me.Tag = ""
    ActivaBotones
    
    mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & txtCodigo & "' order by " & mcCampoCodigo
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & txtCodigo & "' order by " & mcCampoCodigo
            If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                    LeerCampos
                Else
                    mblnTablaVacia = True
                    LimpiaCampos
                End If
            End If
        End If
    End If
    Conexion.CloseHost adoPrincipal
    txtNombre.SetFocus
End Sub
Private Sub GrabarRegistro()
If Not Validacion() Then
    Exit Sub
End If

If Me.Tag = "Crear" Then
    mstrSql = "INSERT INTO " & mcNombreTabla & " (" & mcCampoCodigo & ", " & mcCampoNombre & ", Es_Recepcionista, Es_Supervisor,Quien_Supervisa, Valor_Hora, Vigencia, usr_id, usr_fecha,Id_Empresa,Id_Sucursal,Es_Liquidador,Rut_Mecanico ) "
    mstrSql = mstrSql & "values ('" & Trim(txtCodigo) & "', '" & Trim(txtNombre) & "', '" & IIf(cckRecepcionista.Value = vbChecked, "S", "N") & "', '" & IIf(cckSupervisor.Value = vbChecked, "S", "N") & "','" & dtcSupervisor.BoundText & "', " & Trim(TxtValorHora.Text) & ", '" & IIf(chkVigencia.Value = vbChecked, "S", "N") & "', "
    mstrSql = mstrSql & " '" & gstrUsuario & "', '" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "','" & gstrIdEmpresa & "','" & gstrIdSucursal & "','" & IIf(cckLiquidador.Value = vbChecked, "S", "N") & "','" & txtRut & "')"
Else
    mstrSql = "UPDATE " & mcNombreTabla & " SET " & mcCampoNombre & "='" & Trim(txtNombre) & "',"
    mstrSql = mstrSql & " Es_Recepcionista='" & IIf(cckRecepcionista.Value = vbChecked, "S", "N") & "',"
    mstrSql = mstrSql & " Es_Supervisor='" & IIf(cckSupervisor.Value = vbChecked, "S", "N") & "',"
    mstrSql = mstrSql & " Es_Liquidador='" & IIf(cckLiquidador.Value = vbChecked, "S", "N") & "',"
    mstrSql = mstrSql & " Quien_Supervisa='" & dtcSupervisor.BoundText & "',"
    mstrSql = mstrSql & " vigencia='" & IIf(chkVigencia.Value = vbChecked, "S", "N") & "',"
    mstrSql = mstrSql & " Valor_Hora = " & Trim(TxtValorHora.Text) & ","
    mstrSql = mstrSql & " usr_id='" & gstrUsuario & "', usr_fecha='" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "',"
    mstrSql = mstrSql & " Rut_Mecanico='" & txtRut & "'"
    mstrSql = mstrSql & " WHERE " & mcCampoCodigo & "='" & Trim(txtCodigo) & "'"
End If
If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
    mblnTablaVacia = False
    ActivaBotones
    Me.Tag = ""
End If

GuardaEspecialidad Trim(txtCodigo)
'datSupervisor.Recordset.Requery
'datSupervisor.Refresh

Me.tlbBarraHerramientas.Buttons.Item("Password").Enabled = IIf(Me.cckLiquidador.Value = vbChecked, True, False)
End Sub
Private Sub BorrarRegistro()
    Screen.MousePointer = vbDefault
    If MsgBox("¿ Desea eliminar este registro ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
        mstrSql = "DELETE FROM TLLR_ESPECIALIDAD_MECANICO where id_mecanico = '" & txtCodigo & "'"
        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
            mstrSql = "DELETE FROM " & mcNombreTabla & " where " & mcCampoCodigo & "='" & txtCodigo & "'"
            If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
                mstrSql = "select TOP 1 * from " & mcNombreTabla & " where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "' and " & mcCampoCodigo & ">'" & txtCodigo & "' order by " & mcCampoCodigo
                If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                        LeerCampos
                    Else
                        mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "' and " & mcCampoCodigo & "<'" & txtCodigo & "' order by " & mcCampoCodigo
                        If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                            If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                                LeerCampos
                            Else
                                mblnTablaVacia = True
                                LimpiaCampos
                            End If
                        End If
                    End If
                End If
            End If
            Conexion.CloseHost adoPrincipal
        End If
    End If
End Sub

Private Sub BuscarRegistro()
    'gstrBusca = BuscarRegistros(mcNombreTabla, mcCampoCodigo, mcCampoNombre, Me.Caption)
    gstrBusca = apfFormulario.BuscarRegistros(Conexion, "(Select Id_mecanico,Nombre from Tllr_Mecanicos Where id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "') as MyTabla", "Id_Mecanico", "Nombre", Me.Caption)
    If gstrBusca <> "" Then
        mstrSql = "select * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "='" & gstrBusca & "' order by " & mcCampoCodigo
        If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                LeerCampos
            End If
        End If
        Conexion.CloseHost adoPrincipal
    End If
    Me.SetFocus
End Sub
Private Sub ImprimirInforme()
    'ImprimirRegistros mcNombreTabla, mcCampoCodigo, mcCampoNombre, Me.Caption, gstrPathReporte, "APCARROC.RPT", gstrUSUARIO
    With rptMantenedor
        .ReportFileName = gstrPathReporte & "\APMECANICOS.RPT"
        .Formulas(0) = "Titulo='Listado De Mecanicos'"
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
mstrSql = "select TOP 1 * from " & mcNombreTabla & " where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "' order by " & mcCampoCodigo
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos
    Else
        Beep
    End If
End If
Conexion.CloseHost adoPrincipal
End Sub
Private Sub RegistroAnterior()
mstrSql = "select TOP 1 * from " & mcNombreTabla & " where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "'  and  " & mcCampoCodigo & "<'" & txtCodigo & "' order by " & mcCampoCodigo & " DESC"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos
    Else
        Beep
    End If
End If
Conexion.CloseHost adoPrincipal
End Sub
Private Sub RegistroSiguiente()
mstrSql = "select TOP 1 * from " & mcNombreTabla & " where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "'  and " & mcCampoCodigo & ">'" & txtCodigo & "' order by " & mcCampoCodigo
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos
    Else
        Beep
    End If
End If
Conexion.CloseHost adoPrincipal
End Sub
Private Sub UltimoRegistro()
mstrSql = "select TOP 1 * from " & mcNombreTabla & "  where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "'  order by " & mcCampoCodigo & " DESC"
If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
        LeerCampos
    Else
        Beep
    End If
End If
Conexion.CloseHost adoPrincipal
End Sub
Private Sub Renovar()
Set adoPrincipal = New ADODB.Recordset
mstrSql = "select TOP 1 * from " & mcNombreTabla & " where Id_Empresa='" & gstrIdEmpresa & "'  and Id_Sucursal ='" & gstrIdSucursal & "'  order by " & mcCampoCodigo

If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
    VerificaTablaVacia
    ActivaBotones
    If Not mblnTablaVacia Then
        PrimerRegistro
    End If
End If
Conexion.CloseHost adoPrincipal
End Sub
Private Sub CerrarSalir()
    Unload Me
End Sub
Private Sub Ayuda()
End Sub
Private Sub ActivaBotones()
    txtCodigo.Enabled = False
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
        .Item("Password").Enabled = IIf(mblnTablaVacia, False, True)
    End With
End Sub
Private Sub DesactivaBotones()
    'txtCodigo.Enabled = True
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
    If (Not adoPrincipal.BOF And Not adoPrincipal.EOF) And adoPrincipal.RecordCount > 0 Then
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

    With adoPrincipal
        txtCodigo.Text = ValorNulo(.Fields(mcCampoCodigo))
        If IsNull(!vigencia) Then
            chkVigencia.Value = vbUnchecked
        Else
            If !vigencia = "S" Then
                chkVigencia.Value = vbChecked
            Else
                chkVigencia.Value = vbUnchecked
            End If
        End If
        txtNombre.Text = ValorNulo(.Fields(mcCampoNombre))
        If !es_recepcionista = "S" Then
            cckRecepcionista.Value = vbChecked
        Else
            cckRecepcionista.Value = vbUnchecked
        End If
        If !es_supervisor = "S" Then
            cckSupervisor.Value = vbChecked
        Else
            cckSupervisor.Value = vbUnchecked
        End If
        If !es_Liquidador = "S" Then
            cckLiquidador.Value = vbChecked
        Else
            cckLiquidador.Value = vbUnchecked
        End If
        TxtValorHora = .Fields("Valor_Hora")
        txtRut = ValorNulo(!Rut_Mecanico)
        dtcSupervisor.BoundText = !Quien_Supervisa
        SetCheckOff lvwEspecialidad
        Especialidades_Mecanico ValorNulo(.Fields(mcCampoCodigo))
        Me.tlbBarraHerramientas.Buttons.Item("Password").Enabled = IIf(Me.cckLiquidador.Value = vbChecked, True, False)
    End With
End Sub
Private Sub LimpiaCampos()
    txtCodigo.Text = "": txtNombre.Text = "": dtcSupervisor.BoundText = "": TxtValorHora = ""
    chkVigencia.Value = vbUnchecked: cckRecepcionista.Value = 0: cckSupervisor.Value = 0
    cckLiquidador.Value = 0
    txtRut = ""
    SetCheckOff lvwEspecialidad
End Sub

Private Sub ValoresporDefecto()
Dim lvarCODIGO As Variant
Dim lvarDESCRIP As Variant
Dim lstrSql As String
Dim AdoTemp As New ADODB.Recordset

' obtiene siguiente codigo
lvarCODIGO = 0
' hace la query para extraer el siguiente codigo
lstrSql = ""
lstrSql = lstrSql & "SELECT TOP 1 " & mcCampoCodigo & " "
lstrSql = lstrSql & "From " & mcNombreTabla & " "
lstrSql = lstrSql & "ORDER BY CAST(" & mcCampoCodigo & " AS FLOAT) DESC "
Set AdoTemp = New ADODB.Recordset
If Conexion.SendHost(lstrSql, AdoTemp, adOpenStatic, adLockReadOnly, gcTiempoEspera) <> apAbort Then
    If AdoTemp.EOF = False And AdoTemp.BOF = False Then
        lvarCODIGO = ValorNulo(AdoTemp.Fields(mcCampoCodigo))
    End If
End If
AdoTemp.Close
If IsNumeric(lvarCODIGO) = True Then
    lvarCODIGO = CDbl(lvarCODIGO) + 1
    lvarCODIGO = Format(lvarCODIGO, "0#")
Else
    lvarCODIGO = Asc(lvarCODIGO) + 1
    lvarCODIGO = Chr(lvarCODIGO)
End If

Me.txtCodigo.Text = lvarCODIGO
Me.chkVigencia.Value = vbChecked
Me.TxtValorHora = "0"
End Sub
Private Function Validacion() As Boolean
    Validacion = True
    If txtCodigo = "" Then
        MsgBox "El código debe contener un valor...", vbInformation, "Advertencia"
        txtCodigo.SetFocus
        Validacion = False
        Exit Function
    End If
    If txtNombre = "" Then
        MsgBox "La descripción debe contener un valor...", vbInformation, "Advertencia"
        txtNombre.SetFocus
        Validacion = False
        Exit Function
    End If
  
    
    '//Verifica si existe un registro...
    If Me.Tag = "Crear" Then
        Dim AdoTemp As New ADODB.Recordset
        mstrSql = "select " & mcCampoCodigo & ", " & mcCampoNombre & " from " & mcNombreTabla & " where " & mcCampoCodigo & "='" & txtCodigo & "' And Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
        If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not AdoTemp.BOF And Not AdoTemp.EOF Then
                MsgBox "Este código ya esta registrado con la descripción " & Chr(13) & "[" & IIf(IsNull(AdoTemp.Fields(mcCampoNombre)), "SIN DESCRIPCION", AdoTemp.Fields(mcCampoNombre)) & "]", vbInformation, "Advertencia"
                Validacion = False
                txtCodigo.SetFocus
            End If
        End If
        Conexion.CloseHost AdoTemp
    End If
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmMantenedorMecanicos = Nothing
    gstrBusca = txtCodigo.Text
End Sub
Private Sub RevizaAtributos()
'    mblnAccesoCrear = rsUsuarios!OPC_AUXILIAR_CREAR
'    mblnAccesoEditar = rsUsuarios!OPC_AUXILIAR_EDITAR
'    mblnAccesoBorrar = rsUsuarios!OPC_AUXILIAR_BORRAR
'    mblnAccesoImprimir = rsUsuarios!OPC_AUXILIAR_IMPRIMIR

    mblnAccesoCrear = True
    mblnAccesoEditar = True
    mblnAccesoBorrar = True
    mblnAccesoImprimir = True

End Sub

Private Sub CambiaPassword()
    Screen.MousePointer = vbDefault
    frmCambiaPasswordLiquidador.Show 1
End Sub

Function VerificaRut(ByVal rut As String) As String
Dim taum As Integer
Dim sp, xru, xid As String
Dim p, N, i As Integer
Dim digito As String
Dim LARGO As Long
Dim re As Double

If gstrValidaRut = "S" Then
 taum = 0

 LARGO = Len(rut)
 xid = Mid$(rut, LARGO, LARGO)
 LARGO = LARGO - 1
 xru = Mid$(rut, 1, LARGO)
 N = Len(Trim$(xru))
 i = 2
 While N > 0
  sp = Mid$(Trim$(xru), N, 1)
  p = Val(sp)
  taum = taum + (i * p)
  If i = 7 Then
   i = 2
  Else
   i = i + 1
  End If
  N = N - 1
 Wend
 re = Int(taum / 11)
 re = taum - (re * 11)
 re = Int(11 - re)

Select Case re
 Case 10
  digito = "K"
 Case 11
  digito = "0"
 Case Else
  digito = Trim$(Str$(re))
 End Select

 If xid = digito Then
  VerificaRut = "1"
 Else
  VerificaRut = "0"
 End If
End If
End Function
'kjcv 27-01-12
Private Sub txtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = UpCaseLetter(KeyAscii)
End Sub

Private Sub txtRut_LostFocus()
If txtRut <> "" Then
'    If VerificaRut(Trim(txtRut.Text)) = "0" Then
    If Trim(txtRut.Text) = "" Then

        MsgBox gstrNombreRut & " Incorrecto, Favor de Corregir", vbCritical + vbOKOnly, "Incorrecto " & gstrNombreRut
        txtRut.SetFocus
    End If
End If
End Sub
