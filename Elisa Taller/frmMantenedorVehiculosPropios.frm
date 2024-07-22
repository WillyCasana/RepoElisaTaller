VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmMantenedorVehiculosPropios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor Vehiculos Propios"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmMantenedorVehiculosPropios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   6615
      Begin VB.TextBox txtCuentaAsociada 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1350
         Width           =   1725
      End
      Begin MSDataListLib.DataCombo dtcCentroCosto 
         Bindings        =   "frmMantenedorVehiculosPropios.frx":179A
         Height          =   315
         Left            =   1905
         TabIndex        =   9
         Top             =   960
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "nombre"
         BoundColumn     =   "codigo"
         Text            =   ""
      End
      Begin VB.TextBox txtPatente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   1905
         MaxLength       =   10
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   210
         Width           =   1200
      End
      Begin MSDataListLib.DataCombo dtcGrupoCosto 
         Bindings        =   "frmMantenedorVehiculosPropios.frx":17B7
         Height          =   315
         Left            =   1905
         TabIndex        =   5
         Top             =   600
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "nombre"
         BoundColumn     =   "codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc datGrupoCosto 
         Height          =   330
         Left            =   1905
         Top             =   585
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
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
      Begin VB.CheckBox chkVigencia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         Caption         =   "Activo:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5520
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin MSComctlLib.Toolbar tlbPatente 
         Height          =   330
         Left            =   3180
         TabIndex        =   8
         Top             =   180
         Visible         =   0   'False
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImgBarraHerramienta"
         DisabledImageList=   "ImgBarraHerramienta"
         HotImageList    =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Nuevo"
               Object.ToolTipText     =   "Nuevo Patente"
               ImageKey        =   "Crear"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar Patente"
               ImageKey        =   "Buscar"
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc datCentroCosto 
         Height          =   330
         Left            =   1920
         Top             =   945
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
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
      Begin MSComctlLib.ImageList ImgBuscaCuenta 
         Left            =   5085
         Top             =   1485
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   19
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":17D3
               Key             =   "Crear"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":18E5
               Key             =   "Editar"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":19F7
               Key             =   "Grabar"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":1B09
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":1C1B
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":1D2D
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":1E3F
               Key             =   "Imprimir"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":1F51
               Key             =   "Cerrar"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":2063
               Key             =   "Ayuda"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":2175
               Key             =   "Primero"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":2287
               Key             =   "Anterior"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":2399
               Key             =   "Siguiente"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":24AB
               Key             =   "Ultimo"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":25BD
               Key             =   "Renovar"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":26CF
               Key             =   "SortAsc"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":27E1
               Key             =   "SortDesc"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":28F3
               Key             =   "Seleccion"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":2D45
               Key             =   "Seleccion1"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenedorVehiculosPropios.frx":3197
               Key             =   "Copiar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbCuenta 
         Height          =   330
         Left            =   3705
         TabIndex        =   12
         Top             =   1350
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImgBuscaCuenta"
         DisabledImageList=   "ImgBuscaCuenta"
         HotImageList    =   "ImgBuscaCuenta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar Patente"
               ImageKey        =   "Buscar"
            EndProperty
         EndProperty
      End
      Begin Crystal.CrystalReport rptMantenedor 
         Left            =   5040
         Top             =   600
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
         Caption         =   "Cuenta Asociada :"
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
         Left            =   105
         TabIndex        =   10
         Top             =   1350
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Centro Costo :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   105
         TabIndex        =   6
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Grupo Centro Costo :"
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
         Left            =   105
         TabIndex        =   4
         Top             =   600
         Width           =   1860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Placa:"
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
         Left            =   105
         TabIndex        =   2
         Top             =   240
         Width           =   525
      End
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
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
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar Registro (Ctrl+B)"
            ImageKey        =   "Buscar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir (Ctrl+I)"
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Key             =   "Renovar"
            Object.ToolTipText     =   "Renovar Registros (Ctrl+R)"
            ImageKey        =   "Renovar"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (Ctrl+Q)"
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
            Picture         =   "frmMantenedorVehiculosPropios.frx":32A9
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":33BB
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":34CD
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":35DF
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":36F1
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":3803
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":3915
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":3A27
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":3B39
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":3C4B
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":3D5D
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":3E6F
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":3F81
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":4093
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":41A5
            Key             =   "Ascendente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":42B7
            Key             =   "Descendente"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":43C9
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":481B
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":4C6D
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":4D7F
            Key             =   "Archivar"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":4EDB
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":5037
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":5193
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":52EF
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":5DBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":620F
            Key             =   "sii"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":6373
            Key             =   "siid"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":67CF
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":692B
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":7C37
            Key             =   "Ins"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":81D3
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":832F
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":848B
            Key             =   "Ir"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":87DF
            Key             =   "IrAold"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":8B33
            Key             =   "IrA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":8E87
            Key             =   "outlook"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":91DB
            Key             =   "Porcent"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":952F
            Key             =   "Copiar2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":9A73
            Key             =   "Tambor"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":9B85
            Key             =   "Cajon_mal"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":9ED9
            Key             =   "Cajon"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":A22D
            Key             =   "Bono"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":A341
            Key             =   "Bono2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":A695
            Key             =   "Picking"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":A7A7
            Key             =   "Pago"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorVehiculosPropios.frx":AAFB
            Key             =   "Cotizacion"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMantenedorVehiculosPropios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoPrincipal As New ADODB.Recordset
Dim mstrSql As String
Dim strEntrada As String
Dim mblnTablaVacia As Boolean
Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean
Dim mblnSW As Boolean
Dim mstrD_P As String
Const mcNombreTabla = "Tllr_Vehiculo_Propio"
Const mcCampoCodigo = "Id_Vehiculo"
Const mcCampoNombre = "Id_Centro_Costo"

Private Sub dtcGrupoCosto_Change()
If Me.dtcGrupoCosto <> "" Then
    LlenaCentroCosto
End If
End Sub

Private Sub Form_Load()
    mblnSW = True
    Label1.Caption = gstrNombrePatente
End Sub

Private Sub Form_Unload(Cancel As Integer)
gintProcedencia = 0
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
    End Select
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Activate()
    If mblnSW Then
        
        LlenaGrupoCosto
        'LlenaCentroCosto
        
        If Not Atributos("Glbl", "Tllr_10_0140", mblnAccesoCrear, mblnAccesoEditar, mblnAccesoBorrar, mblnAccesoImprimir) Then
            MsgBox "Acceso no permitido...", vbInformation, "Advertencia"
            Unload Me
            Exit Sub
        End If

        If gapAccion = apcrear Then
           AgregarRegistro
           txtPatente = gstrBusca
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
            txtPatente.Enabled = False
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
    txtPatente.SetFocus
End Sub
Private Sub CancelarAgregaRegistro()
    Me.Tag = ""
    ActivaBotones
    
    mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & txtPatente & "' order by " & mcCampoCodigo
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & txtPatente & "' order by " & mcCampoCodigo
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
    'txtNombre.SetFocus
End Sub
Private Sub GrabarRegistro()
    If Not Validacion() Then
        Exit Sub
    End If

    If Me.Tag = "Crear" Then
        mstrSql = "INSERT INTO " & mcNombreTabla & " (" & mcCampoCodigo & ", "
        mstrSql = mstrSql & "id_grupo_centro_Costo,id_centro_costo,id_cuenta_asociada,Id_empresa, Id_Sucursal) "
        mstrSql = mstrSql & " values ('" & UCase(Trim(txtPatente)) & "', '" & dtcGrupoCosto.BoundText & "', '" & dtcCentroCosto.BoundText & "', '" & txtCuentaAsociada & "',"
        mstrSql = mstrSql & "'" & gstrIdEmpresa & "', '" & gstrIdSucursal & "')"
    Else
        mstrSql = "UPDATE " & mcNombreTabla & " SET "
        mstrSql = mstrSql & "id_grupo_centro_costo = '" & dtcGrupoCosto.BoundText & "',"
        mstrSql = mstrSql & " id_centro_costo = '" & dtcCentroCosto.BoundText & "',"
        mstrSql = mstrSql & " id_cuenta_asociada = '" & txtCuentaAsociada & "',"
        mstrSql = mstrSql & " Id_Empresa = '" & gstrIdEmpresa & "',"
        mstrSql = mstrSql & " Id_Sucursal = '" & gstrIdSucursal & "'"
        mstrSql = mstrSql & " where " & mcCampoCodigo & "='" & Trim(txtPatente) & "'"
    End If
    If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
        mblnTablaVacia = False
        ActivaBotones
        Me.Tag = ""
    End If
    
    If gintProcedencia = 1 Then
        Unload Me
    End If
End Sub
Private Sub BorrarRegistro()
    Screen.MousePointer = vbDefault
    If MsgBox("¿ Desea eliminar este registro ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
        mstrSql = "DELETE FROM " & mcNombreTabla & " where " & mcCampoCodigo & "='" & txtPatente & "'"
        If Conexion.SendHost(mstrSql, , , , gcTiempoEspera) = apOk Then
            mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & txtPatente & "' order by " & mcCampoCodigo
            If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                    LeerCampos
                Else
                    mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & txtPatente & "' order by " & mcCampoCodigo
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
End Sub
Private Sub BuscarRegistro()
'    Set FormVol1 = New APFORM1.APFORM
    gstrBusca = apfFormulario.BuscarRegistros(Conexion, mcNombreTabla, mcCampoCodigo, mcCampoNombre, Me.Caption)
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
   ' FormVol1.ImprimirRegistros Conexion, mcNombreTabla, mcCampoCodigo, mcCampoNombre, Me.Caption, gstrPathReporte, "APCARROC.RPT", gstrUSUARIO, gstrCodigoEmpresa
    With rptMantenedor
        .ReportFileName = gstrPathReporte & "\APVEHICULOPROPIO.RPT"
        .Formulas(0) = "Titulo='Listado Vehículos Propios'"
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
    
    mstrSql = "select TOP 1 * from " & mcNombreTabla & " order by " & mcCampoCodigo
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
    
    mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & "<'" & txtPatente & "' order by " & mcCampoCodigo & " DESC"
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

    mstrSql = "select TOP 1 * from " & mcNombreTabla & " WHERE " & mcCampoCodigo & ">'" & txtPatente & "' order by " & mcCampoCodigo
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
    mstrSql = "select TOP 1 * from " & mcNombreTabla & " order by " & mcCampoCodigo & " DESC"
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
    mstrSql = "select TOP 1 * from " & mcNombreTabla & " order by " & mcCampoCodigo
    
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
    txtPatente.Enabled = False
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
    txtPatente.Enabled = True
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
        txtPatente.Text = ValorNulo(.Fields(mcCampoCodigo))
        dtcGrupoCosto.BoundText = !id_grupo_centro_costo
        dtcCentroCosto.BoundText = !id_centro_costo
        txtCuentaAsociada = ValorNulo(!id_cuenta_asociada)
    End With
End Sub
Private Sub LimpiaCampos()
    txtPatente.Text = ""
    chkVigencia.Value = vbUnchecked
    txtCuentaAsociada = ""
End Sub
Private Sub ValoresporDefecto()
    With adoPrincipal
        chkVigencia.Value = vbChecked
    End With
End Sub
Private Function Validacion() As Boolean
    Validacion = True
    If txtPatente = "" Then
        MsgBox "El código debe contener un valor...", vbInformation, "Advertencia"
        txtPatente.SetFocus
        Validacion = False
        Exit Function
    End If
    If txtCuentaAsociada = "" Then
        MsgBox "La Cuenta Asociada debe contener un valor...", vbInformation, "Advertencia"
        txtCuentaAsociada.SetFocus
        Validacion = False
        Exit Function
    End If

'    If txtValor = "" Then
'        MsgBox "Valor del servicio Incorrecto...", vbInformation, "Advertencia"
'        txtValor.SetFocus
'        Validacion = False
'        Exit Function
'    End If
'
'    If txtComision = "" Then
'        MsgBox "El valor de la Comisión es Incorrecto...", vbInformation, "Advertencia"
'        txtComision.SetFocus
'        Validacion = False
'        Exit Function
'    End If
    
    '//Verifica si existe un registro...
    If Me.Tag = "Crear" Then
        Dim AdoTemp As New ADODB.Recordset
        mstrSql = "select " & mcCampoCodigo & ", " & mcCampoNombre & " from " & mcNombreTabla & " where " & mcCampoCodigo & "='" & txtPatente & "'"
        If Conexion.SendHost(mstrSql, AdoTemp, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
            If Not AdoTemp.BOF And Not AdoTemp.EOF Then
                MsgBox "Este código ya esta registrado con la descripción " & Chr(13) & "[" & IIf(IsNull(AdoTemp.Fields(mcCampoNombre)), "SIN DESCRIPCION", AdoTemp.Fields(mcCampoNombre)) & "]", vbInformation, "Advertencia"
                Validacion = False
                txtPatente.SetFocus
            End If
        End If
        Conexion.CloseHost AdoTemp
    End If
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmMantenedorVehiculosPropios = Nothing
    gstrBusca = txtPatente.Text
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

Sub LlenaGrupoCosto()
    dtcGrupoCosto.Enabled = True
    mstrSql = "Select Id_Grupo_Centro_Costo as CODIGO, Nombre from Cont_Grupo_Centro_Costo where Id_Empresa='" & gstrIdEmpresa & "' And VIGENCIA = 'S' order by nombre"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With datGrupoCosto
            Set .Recordset = adoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                '.Recordset.MoveFirst
                dtcGrupoCosto.ListField = "Nombre"
                dtcGrupoCosto.BoundColumn = "Codigo"
                dtcGrupoCosto.BoundText = .Recordset!Codigo
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
End Sub

Sub LlenaCentroCosto()
    dtcCentroCosto.Enabled = True
    mstrSql = "Select Id_centro_Costo as CODIGO, Nombre from Cont_Centro_Costo where VIGENCIA = 'S' And Id_empresa='" & gstrIdEmpresa & "' AND id_grupo_centro_costo='" & dtcGrupoCosto.BoundText & "' order by nombre"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With datCentroCosto
            Set .Recordset = adoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                dtcCentroCosto.ListField = "Nombre"
                dtcCentroCosto.BoundColumn = "Codigo"
                dtcCentroCosto.BoundText = .Recordset!Codigo
            End If
        End With
    End If ' por el otro
    Set adoPrincipal = New ADODB.Recordset
    Conexion.CloseHost adoPrincipal
End Sub

Private Sub tlbCuenta_ButtonClick(ByVal Button As MSComctlLib.Button)
'    Select Case Button.Key
'    Case "Buscar"
'        strEntrada = APFORM1.BuscarRegistros(Conexion, "(Select id_Cuenta,Nombre from Cont_Plan_Cuentas where Id_Empresa='" & gstrIdEmpresa & "') as tabla", "id_Cuenta", "Nombre", "Búsqueda de Sub-Cuentas")
'        If strEntrada <> "" Then
'            txtCuentaAsociada = strEntrada
'        End If
'    End Select
End Sub

Private Sub tlbPatente_ButtonClick(ByVal Button As MSComctlLib.Button)
If Me.Tag = "Crear" Then
    Select Case Button.Key
    Case "Nuevo"
        txtPatente = Vehiculos(Conexion, gstrIdUsuario, "TLLR", "", gstrIdEmpresa, gstrPathReporte, txtPatente, apcrear)
        DatosVehiculo txtPatente
    Case "Buscar"
        gstrProcedencia = "Movimientos"
        frmBuscaVehiculo.Show vbModal
    End Select
Else
    Select Case Button.Key
    Case "Nuevo"
        gstrProcedencia = "MantenedorPropio"
        txtPatente = Vehiculos(Conexion, gstrIdUsuario, "TLLR", "", gstrIdEmpresa, gstrPathReporte, txtPatente, apeditar)
        DatosVehiculo txtPatente
    End Select
End If

End Sub

Sub DatosVehiculo(strPatente As String)
If strPatente <> "" Then
    mstrSql = "SELECT Tllr_Vehiculo_Cliente.Patente,"
    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Id_Marca AS IDMARCA,"
    mstrSql = mstrSql & " Glbl_Marca.Descripcion AS MARCA,"
    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Id_Modelo AS IDMODELO,"
    mstrSql = mstrSql & " Glbl_Modelo.Descripcion AS MODELO,"
    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Año,"
    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Id_Color_Exterior AS IDCOLOR,"
    mstrSql = mstrSql & " Glbl_Color_Exterior.Descripcion AS COLOR,"
    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Kilometros_Actuales AS KILACT,"
    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Nro_Motor AS MOTOR,"
    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Nro_Chasis AS CHASIS,"
    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.VIN AS VIN,"
    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor AS IDCLI,"
    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Fecha_Venta AS FECVTA,"
    mstrSql = mstrSql & " Tllr_Vehiculo_Cliente.Concesionario AS CONCES"
    mstrSql = mstrSql & " FROM Glbl_Cliente_Proveedor RIGHT OUTER JOIN Glbl_Color_Exterior RIGHT OUTER JOIN Tllr_Vehiculo_Cliente ON Glbl_Color_Exterior.Id_Color_Exterior = Tllr_Vehiculo_Cliente.Id_Color_Exterior LEFT OUTER JOIN Glbl_Modelo LEFT OUTER JOIN Glbl_Marca ON Glbl_Modelo.Id_Marca = Glbl_Marca.Id_Marca ON Tllr_Vehiculo_Cliente.Id_Modelo = Glbl_Modelo.Id_Modelo AND Tllr_Vehiculo_Cliente.Id_Marca = Glbl_Modelo.Id_Marca ON Glbl_Cliente_Proveedor.Id_Cliente_Proveedor = Tllr_Vehiculo_Cliente.Id_Cliente_Proveedor"
    mstrSql = mstrSql & " WHERE Tllr_Vehiculo_Cliente.Patente='" & txtPatente & "'"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            With adoPrincipal
                'lblMarca = ValorNulo(!Marca)
                'lblIdMarca = ValorNulo(!IdMarca)
                'lblModelo = ValorNulo(!Modelo)
                'lblIdModelo = ValorNulo(!IdModelo)
                'lblChasis = ValorNulo(!chasis)
                'lblMotor = ValorNulo(!motor)
                'lblVIN = ValorNulo(!VIN)
                'txtAño = !Año
                'lblColorE = ValorNulo(!Color)
                'lblCliente = ValorNulo(!idCLI)
                'txtConcesionario = ValorNulo(!CONCES)
                'pckFecVta.Value = IIf(Not IsNull(!FECVTA), !FECVTA, Now)
                'txtKilAct = IIf(Not IsNull(!kilact), !kilact, "0")
                'lblIdCliente = ValorNulo(!idCLI)
            End With
        End If
    End If
    Conexion.CloseHost adoPrincipal
End If
End Sub

Private Sub txtPatente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If txtPatente <> "" Then
            If Len(txtPatente) = 6 Then
                If ConsultaVehiculo(txtPatente) = False Then
                    If MsgBox("La " & gstrNombrePatente & " " & txtPatente & " Debe Ingresarla como " & gstrNombrePatente & "Cliente Primero", 32, gstrNombrePatente & " Existente") Then
                    End If
                    gstrProcedencia = "MantenedorPropio"
                    gapAccion = apcrear
                    frmMantenedorVehiculoCliente.Show vbModal
                End If
            Else
                MsgBox LoadResString(326)
            End If
        Else
            MsgBox LoadResString(327)
        End If
    End If

End Sub

