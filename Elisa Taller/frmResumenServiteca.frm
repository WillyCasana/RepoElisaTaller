VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmResumenServiteca 
   Caption         =   "Resumen Servicios de Serviteca"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   Icon            =   "frmResumenServiteca.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8055
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport rptOT 
      Left            =   4860
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11655
      Begin VB.TextBox txtOtHasta 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9960
         TabIndex        =   20
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtOtDesde 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9960
         TabIndex        =   18
         Top             =   480
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Height          =   1575
         Left            =   9720
         TabIndex        =   16
         Top             =   120
         Width           =   30
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   315
         Left            =   7680
         TabIndex        =   15
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   162267137
         CurrentDate     =   36880
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   315
         Left            =   7680
         TabIndex        =   13
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   162267137
         CurrentDate     =   36880
      End
      Begin MSComctlLib.Toolbar tlbBotones 
         Height          =   330
         Index           =   0
         Left            =   3480
         TabIndex        =   9
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cancelar"
               Object.ToolTipText     =   "Limpiar"
               ImageKey        =   "Cancelar"
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   7440
         TabIndex        =   8
         Top             =   120
         Width           =   30
      End
      Begin MSDataListLib.DataCombo dbcboMecanico 
         Bindings        =   "frmResumenServiteca.frx":179A
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Id_Mecanico"
         Text            =   "dbcboMecanico"
      End
      Begin MSAdodcLib.Adodc datMecanico 
         Height          =   270
         Left            =   1920
         Top             =   720
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   476
         ConnectMode     =   0
         CursorLocation  =   2
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
         LockType        =   1
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   0
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
         Caption         =   "datMecanico"
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
      Begin MSAdodcLib.Adodc datConceptoServicio 
         Height          =   270
         Left            =   240
         Top             =   1440
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   476
         ConnectMode     =   0
         CursorLocation  =   2
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
         LockType        =   1
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   0
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
         Caption         =   "datConceptoServicio"
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
      Begin MSDataListLib.DataCombo dbCboConceptoServicio 
         Bindings        =   "frmResumenServiteca.frx":17B4
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "Id_Concepto_Servicio"
         Text            =   "dbCboConceptoServicio"
      End
      Begin MSAdodcLib.Adodc datServicio 
         Height          =   270
         Left            =   3960
         Top             =   1440
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   476
         ConnectMode     =   0
         CursorLocation  =   2
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
         LockType        =   1
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   0
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
         Caption         =   "datServicio"
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
      Begin MSDataListLib.DataCombo dbCboServicio 
         Bindings        =   "frmResumenServiteca.frx":17D6
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         Top             =   1200
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "Id_Servicio"
         Text            =   "dbCboServicio"
      End
      Begin MSComctlLib.Toolbar tlbBotones 
         Height          =   330
         Index           =   1
         Left            =   3000
         TabIndex        =   10
         Top             =   1200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cancelar"
               Object.ToolTipText     =   "Limpiar"
               ImageKey        =   "Cancelar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbBotones 
         Height          =   330
         Index           =   2
         Left            =   6840
         TabIndex        =   11
         Top             =   1200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cancelar"
               Object.ToolTipText     =   "Limpiar"
               ImageKey        =   "Cancelar"
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo dbcboSucursal 
         Bindings        =   "frmResumenServiteca.frx":17F0
         Height          =   315
         Left            =   3960
         TabIndex        =   29
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "Id_Sucursal"
         Text            =   "dbcboSucursal"
      End
      Begin MSAdodcLib.Adodc datSucursal 
         Height          =   270
         Left            =   4080
         Top             =   720
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   476
         ConnectMode     =   0
         CursorLocation  =   2
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
         LockType        =   1
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   0
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
         Caption         =   "datSucursal"
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
      Begin MSComctlLib.Toolbar tlbOtrosCriterios 
         Height          =   330
         Left            =   5880
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         ButtonWidth     =   2566
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Otros Criterios"
               Key             =   "Otros"
               Object.ToolTipText     =   "Selección de otros criterios"
               ImageKey        =   "Editar"
               Style           =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbBotones 
         Height          =   330
         Index           =   3
         Left            =   6840
         TabIndex        =   30
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImgBarraHerramienta"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cancelar"
               Object.ToolTipText     =   "Limpiar"
               ImageKey        =   "Cancelar"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label8 
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   3960
         TabIndex        =   31
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "O.T. Hasta"
         Height          =   255
         Left            =   9960
         TabIndex        =   19
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "O.T. Desde"
         Height          =   255
         Left            =   9960
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   7680
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   255
         Left            =   7680
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Servicio"
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto Servicio"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Mecánico"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":180A
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":191C
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":1A2E
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":1B40
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":1C52
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":1D64
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":1E76
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":1F88
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":209A
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":21AC
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":22BE
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":23D0
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":24E2
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":25F4
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":2706
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":2818
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":292A
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":2D7C
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":31CE
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":32E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":3824
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResumenServiteca.frx":3B44
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Crear"
            Object.ToolTipText     =   "Crear Registro (Ctrl+N)"
            ImageKey        =   "Crear"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar Registro (Ctrl+G)"
            ImageKey        =   "Grabar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar (ESC)"
            ImageKey        =   "Cancelar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Borrar"
            Object.ToolTipText     =   "Borrar Registro (Ctrl+D)"
            ImageKey        =   "Borrar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Primero"
            Object.ToolTipText     =   "Primer Registro (Ctrl+P)"
            ImageKey        =   "Primero"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Anterior"
            Object.ToolTipText     =   "Registro Anterior (Ctrl+A)"
            ImageKey        =   "Anterior"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Registro Siguiente (Ctrl+S)"
            ImageKey        =   "Siguiente"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.ToolTipText     =   "Cerrar (Ctrl+C)"
            ImageKey        =   "Cerrar"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Liquidar"
            Object.ToolTipText     =   "Estados Orden de Trabajo"
            ImageIndex      =   17
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Liquidar"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Anular"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsvComisiones 
      Height          =   4815
      Left            =   120
      TabIndex        =   21
      Top             =   2400
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "linea"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CODIGO CONCEPTO SERVICIO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CONCEPTO SERVICIO"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CODIGO SERVICIO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "SERVICIO"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "TOTAL MONTO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "CANTIDAD"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox OtrosCriterios 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   5280
      ScaleHeight     =   1665
      ScaleWidth      =   2145
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
      Begin VB.CheckBox chkFacturadas 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ordenes Facturadas"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkLiquidadas 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ordenes Liquidadas"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkVigentes 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ordenes Vigentes"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.OptionButton opcAgrupaOT 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Agrupar por O.T."
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton opcAgrupaServicio 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Agrupar por Servicio"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2040
         Y1              =   720
         Y2              =   720
      End
   End
End
Attribute VB_Name = "frmResumenServiteca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Item As ListItem

Public gstrPrefijoSistema As String
Public gstrCodigoAcceso As String

Private Sub ImprimirConsulta()
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

If lsvComisiones.ListItems.Count = 0 Then
  MsgBox "No existen elementos en la lista", vbExclamation, "Imprimir"
  Exit Sub
End If

Screen.MousePointer = 11
Dim wrkPredeterminado As Workspace
Dim prpBucle As Property
Set wrkPredeterminado = DBEngine.Workspaces(0)  ' Obtiene el Workspace predeterminado.
If Dir(GcamBaseTem & "\BDNueva.mdb") <> "" Then Kill GcamBaseTem & "\BDNueva.mdb" ' Asegúrese de que no existe un archivo con el nombre de la base de datos nueva.
Set Dbsnueva = wrkPredeterminado.CreateDatabase(GcamBaseTem & "\BDNueva.mdb", dbLangGeneral) ' Crea a una base de datos nueva
Dbsnueva.Execute "CREATE TABLE T_PARAREPORTE (ID_CONCEPTO_SERV TEXT, CONCEPTO_SERV TEXT, ID_SERVICIO TEXT, SERVICIO TEXT, VALOR DOUBLE, CANTIDAD DOUBLE)"
Set Tabla = Dbsnueva.OpenRecordset("SELECT * FROM T_PARAREPORTE")
For i = 1 To lsvComisiones.ListItems.Count - 1
    Tabla.AddNew
    Set lsvComisiones.SelectedItem = lsvComisiones.ListItems(i)
    Tabla!ID_CONCEPTO_SERV = IIf(lsvComisiones.SelectedItem.SubItems(1) = "", " ", Me.lsvComisiones.SelectedItem.SubItems(1))
    Tabla!CONCEPTO_SERV = IIf(lsvComisiones.SelectedItem.SubItems(2) = "", " ", Me.lsvComisiones.SelectedItem.SubItems(2))
    Tabla!Id_servicio = IIf(lsvComisiones.SelectedItem.SubItems(3) = "", " ", Me.lsvComisiones.SelectedItem.SubItems(3))
    Tabla!servicio = IIf(lsvComisiones.SelectedItem.SubItems(4) = "", " ", Me.lsvComisiones.SelectedItem.SubItems(4))
    Tabla!Valor = IIf(lsvComisiones.SelectedItem.SubItems(5) = "", " ", SacarFormatoValor(Me.lsvComisiones.SelectedItem.SubItems(5), gstrMonedaLocal))
    Tabla!cantidad = IIf(lsvComisiones.SelectedItem.SubItems(6) = "", " ", Me.lsvComisiones.SelectedItem.SubItems(6))
    Tabla.Update
Next i
Tabla.Close

With rptOT
    .ReportFileName = gstrPathReporte & "\RESUMEN.RPT"
    .Formulas(0) = "USUARIO='" & gstrUsuario & "'"
    .Formulas(1) = "TITULO='RESUMEN SERVICIOS DE SERVITECA'"
    .Formulas(2) = "RazonSocial='" & gstrEmpresa & "'"
    .Formulas(3) = "SUCURSAL='" & gstrSucursal & "'"
    .Formulas(4) = "DIRECCION='" & gstrDirSuc & "'"
    .Formulas(5) = "MECANICO='" & IIf(Me.dbcboMecanico.Text <> "", Me.dbcboMecanico.Text, "[TODOS]") & "'"
    .Formulas(6) = "CONCEPTO='" & IIf(Me.dbCboConceptoServicio.Text <> "", Me.dbCboConceptoServicio.Text, "[TODOS]") & "'"
    .Formulas(7) = "SERVICIO='" & IIf(Me.dbCboServicio.Text <> "", Me.dbCboServicio.Text, "[TODOS]") & "'"
    If Me.dtpDesde.Value <> Me.dtpHasta.Value Then
        .Formulas(8) = "RANGO_FECHAS='" & "DESDE EL " & Me.dtpDesde.Value & " HASTA EL " & Me.dtpHasta.Value & "'"
    Else
        .Formulas(8) = "RANGO_FECHAS='" & Me.dtpDesde.Value & "'"
    End If
    .Destination = crptToWindow
    .Action = True
End With

End Sub
Private Sub dbCboConceptoServicio_Click(Area As Integer)
If Area = 2 Then
    LLena_Servicio
    Me.lsvComisiones.ListItems.Clear
End If
End Sub


Private Function TraeMecanico(CodMecanico As String) As String
Dim tablaMec As New ADODB.Recordset
Dim lsql As String

lsql = ""
lsql = "SELECT Nombre FROM Tllr_Mecanicos WHERE Id_Mecanico = '" & CodMecanico & "'"
If Conexion.SendHost(lsql, tablaMec, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    If tablaMec.EOF = False And tablaMec.BOF = False Then
        TraeMecanico = tablaMec!Nombre
    Else
        TraeMecanico = "."
    End If
End If
Conexion.CloseHost tablaMec


End Function

Private Sub Buscar()
Dim Tabla As New ADODB.Recordset
Dim tablapaso As New ADODB.Recordset
Dim sql As String
Dim ldblTotalNetoOT As Double
Dim ldblTotalCantidad As Double
Dim ldblAcumValorServicio As Double
Dim ldblAcumCantidadServicio As Double

ldblTotalNetoOT = 0
ldblTotalCantidad = 0
ldblAcumValorServicio = 0
ldblAcumCantidadServicio = 0

Me.lsvComisiones.ListItems.Clear

ProcesoRegistros gcInicioProceso
Me.Refresh
ProcesoRegistros gcAvanceProceso, 20
sql = ""
sql = sql & "SELECT SUM(Srvt_OT.Valor_OT) AS Total_Ot, "
sql = sql & "Srvt_Servicios_OT.Id_Concepto_Servicio AS Id_Concepto_Servicio, "
sql = sql & "Srvt_Concepto_Servicio.Descripcion AS Concepto, "
sql = sql & "Srvt_Servicios_OT.Id_Servicio AS Id_Servicio, "
sql = sql & "Srvt_Servicios.Descripcion AS Servicio, "
sql = sql & "SUM(Srvt_Servicios_OT.Valor) AS Valor_Servicio_En_OT, "
sql = sql & "SUM(Srvt_Servicios_OT.Cantidad) AS Cantidad, "
sql = sql & "SUM (Srvt_Servicios_OT.Descuento) "
sql = sql & "AS Desc_Servicio_En_OT, SUM(Srvt_Servicios_OT.Total) "
sql = sql & "AS TOTAL "
sql = sql & "FROM Srvt_Mecanico_Factor RIGHT OUTER JOIN "
sql = sql & "Factor_Recepcionista RIGHT OUTER JOIN "
sql = sql & "Srvt_Concepto_Servicio RIGHT OUTER JOIN "
sql = sql & "Srvt_Servicios ON "
sql = sql & "Srvt_Concepto_Servicio.Id_Concepto_Servicio = Srvt_Servicios.Id_Concepto_Servicio "
sql = sql & "RIGHT OUTER JOIN "
sql = sql & "Srvt_Servicios_OT ON "
sql = sql & "Srvt_Servicios.Id_Concepto_Servicio = Srvt_Servicios_OT.Id_Concepto_Servicio AND "
sql = sql & "Srvt_Servicios.Id_Servicio = Srvt_Servicios_OT.Id_Servicio LEFT "
sql = sql & "Outer Join "
sql = sql & "Srvt_OT LEFT OUTER JOIN "
sql = sql & "Copia_Mecanicos ON "
sql = sql & "Srvt_OT.Id_Mecanico = Copia_Mecanicos.Id_Mecanico ON "
sql = sql & "Srvt_Servicios_OT.Id_OT = Srvt_OT.Id_OT AND "
sql = sql & "Srvt_Servicios_OT.Id_Sucursal = Srvt_OT.Id_Sucursal AND "
sql = sql & "Srvt_Servicios_OT.Id_Empresa = Srvt_OT.Id_Empresa ON "
sql = sql & "Factor_Recepcionista.Id_Mecanico = Srvt_OT.Id_Mecanico ON "
sql = sql & "Srvt_Mecanico_Factor.Id_Mecanico = Srvt_Servicios_OT.Id_Mecanico "
sql = sql & "LEFT OUTER JOIN "
sql = sql & "Tllr_Mecanicos ON "
sql = sql & "Srvt_Servicios_OT.Id_Mecanico = Tllr_Mecanicos.Id_Mecanico "
sql = sql & "WHERE "
sql = sql & "(Srvt_OT.Fecha_Apertura BETWEEN "
sql = sql & "'" & Me.dtpDesde.Value & " 00:00:01' AND "
sql = sql & "'" & Me.dtpHasta.Value & " 23:59:00') "
sql = sql & "AND (Srvt_OT.Estado = '.' OR "
sql = sql & "Srvt_OT.Estado = 'L' OR "
'sql = sql & "Srvt_OT.Estado = 'N' OR "
sql = sql & "Srvt_OT.Estado = 'V' OR "
sql = sql & "Srvt_OT.Estado = 'F' OR "
sql = sql & "Srvt_OT.Estado = 'B') "
If Me.dbcboSucursal.BoundText <> "" Then
    sql = sql & "AND (Srvt_Servicios_OT.Id_Empresa = '" & gstrIdEmpresa & "' AND Srvt_Servicios_OT.Id_Sucursal = '" & Me.dbcboSucursal.BoundText & "') "
End If
If Me.dbcboMecanico.Text <> "" Then
    sql = sql & "AND Tllr_Mecanicos.Vigencia = 'S' "
    sql = sql & "AND (Tllr_Mecanicos.Id_Mecanico = '" & Me.dbcboMecanico.BoundText & "' "
    sql = sql & "OR Srvt_Servicios_OT.Id_Mecanico = '" & Me.dbcboMecanico.BoundText & "') "
End If
If Me.dbCboConceptoServicio.Text <> "" Then
    sql = sql & "AND Srvt_Servicios_OT.Id_Concepto_Servicio = '" & Me.dbCboConceptoServicio.BoundText & "' "
End If
If Me.dbCboServicio.Text <> "" Then
    sql = sql & "AND Srvt_Servicios_OT.Id_Servicio = '" & Me.dbCboServicio.BoundText & "' "
End If
If Me.txtOtDesde.Text <> "" And Me.txtOtHasta.Text <> "" Then
    sql = sql & "AND Srvt_OT.Id_OT BETWEEN '" & Me.txtOtDesde.Text & "' AND '" & Me.txtOtHasta.Text & "' "
End If
sql = sql & "GROUP BY Srvt_Servicios_OT.Id_Concepto_Servicio,"
sql = sql & "Srvt_Concepto_Servicio.Descripcion,"
sql = sql & "Srvt_Servicios.Descripcion , Srvt_Servicios_OT.Id_Servicio"

If Conexion.SendHost(sql, Tabla, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera * 2) = apOk Then
    If Tabla.EOF = False And Tabla.BOF = False Then
        Tabla.MoveFirst
        While Tabla.EOF = False
            Set Item = Me.lsvComisiones.ListItems.Add(, , Me.lsvComisiones.ListItems.Count + 1)
            Item.SubItems(1) = Tabla!Id_Concepto_Servicio
            Item.SubItems(2) = Tabla!Concepto
            Item.SubItems(3) = Tabla!Id_servicio
            Item.SubItems(4) = Tabla!servicio
            Item.SubItems(5) = FormatoValor(ValorNuloNum(Tabla!Total), gstrMonedaLocal, gintDecimalesMoneda)
            ldblTotalNetoOT = ldblTotalNetoOT + ValorNuloNum(Tabla!Total)
            Item.SubItems(6) = ValorNuloNum(Tabla!cantidad)
            ldblTotalCantidad = ldblTotalCantidad + ValorNuloNum(Tabla!cantidad)
            Tabla.MoveNext
            ProcesoRegistros gcAvanceProceso, 50
        Wend
    End If
End If
Conexion.CloseHost Tabla

Set Item = Me.lsvComisiones.ListItems.Add(, , Me.lsvComisiones.ListItems.Count + 1)
Me.lsvComisiones.ListItems.Item(1).Bold = True
Item.SubItems(1) = "TOTALES"
Item.ListSubItems(1).Bold = True
Item.ListSubItems(1).ForeColor = &HC00000
Item.SubItems(2) = "---"
Item.ListSubItems(2).Bold = False
Item.SubItems(3) = "---"
Item.ListSubItems(3).Bold = False
Item.SubItems(4) = "---"
Item.ListSubItems(4).Bold = False
Item.SubItems(5) = FormatoValor(ldblTotalNetoOT, gstrMonedaLocal, gintDecimalesMoneda)
Item.ListSubItems(5).Bold = True
Item.ListSubItems(5).ForeColor = &HC00000
Item.SubItems(6) = ldblTotalCantidad
Item.ListSubItems(6).Bold = True
Item.ListSubItems(6).ForeColor = &HC00000

ProcesoRegistros gcFinProceso
End Sub

Private Function TraeValoresResumen(idServicio As String) As Double
Dim tablaValores As New ADODB.Recordset

sql = ""
sql = sql & "SELECT Srvt_Servicios_OT.Id_Servicio, "
sql = sql & "SUM(Srvt_Servicios_OT.Valor) AS Valor_OT, "
sql = sql & "SUM(Srvt_Servicios_OT.Cantidad) AS Cantidad_OT, "
sql = sql & "SUM(Srvt_Servicios_OT.Descuento) AS Desc_OT, "
sql = sql & "SUM(Srvt_Servicios_OT.Total) AS Total_OT, "
sql = sql & "SUM(Fact_Con_Detalle.Precio_Venta) AS Valor_Fac, "
sql = sql & "SUM(Fact_Con_Detalle.Despachado) AS Cantidad_Fac, "
sql = sql & "SUM(Fact_Con_Detalle.Valor_Descto) AS Desc_Fac, "
sql = sql & "SUM(Fact_Con_Detalle.TOTAL) As Total_Fac "
sql = sql & "FROM Srvt_Servicios_OT INNER JOIN "
sql = sql & "Srvt_OT ON "
sql = sql & "Srvt_Servicios_OT.Id_Empresa = Srvt_OT.Id_Empresa AND "
sql = sql & "Srvt_Servicios_OT.Id_Sucursal = Srvt_OT.Id_Sucursal AND "
sql = sql & "Srvt_Servicios_OT.Id_OT = Srvt_OT.Id_OT LEFT OUTER Join "
sql = sql & "Fact_Con_Detalle ON CONVERT(nvarchar(25), "
sql = sql & "Srvt_Servicios_OT.Id_OT) "
sql = sql & "= Fact_Con_Detalle.Numero_Rescate AND "
sql = sql & "Srvt_Servicios_OT.Id_Empresa = Fact_Con_Detalle.Id_Empresa AND "
sql = sql & "Srvt_Servicios_OT.Id_Sucursal = Fact_Con_Detalle.Id_Sucursal AND "
sql = sql & "Srvt_Servicios_OT.Id_Servicio = SUBSTRING(CONVERT(nvarchar(30), "
sql = sql & "Fact_Con_Detalle.Id_Item), 5, "
sql = sql & "len(Fact_Con_Detalle.Id_Item)) "
sql = sql & "WHERE "
sql = sql & "(Srvt_OT.Fecha_Apertura BETWEEN "
sql = sql & "'" & Me.dtpDesde.Value & " 00:00:01' AND "
sql = sql & "'" & Me.dtpHasta.Value & " 23:59:00') "
sql = sql & "AND (Srvt_OT.Estado = '.' OR "
sql = sql & "Srvt_OT.Estado = 'L' OR "
'sql = sql & "Srvt_OT.Estado = 'N' OR "
sql = sql & "Srvt_OT.Estado = 'V' OR "
sql = sql & "Srvt_OT.Estado = 'F' OR "
sql = sql & "Srvt_OT.Estado = 'B') "
If Me.dbcboSucursal.BoundText <> "" Then
    sql = sql & "AND (Srvt_Servicios_OT.Id_Empresa = '" & gstrIdEmpresa & "' AND Srvt_Servicios_OT.Id_Sucursal = '" & Me.dbcboSucursal.BoundText & "') "
End If
If Me.dbcboMecanico.Text <> "" Then
    sql = sql & "AND Tllr_Mecanicos.Vigencia = 'S' "
    sql = sql & "AND (Tllr_Mecanicos.Id_Mecanico = '" & Me.dbcboMecanico.BoundText & "' "
    sql = sql & "OR Srvt_Servicios_OT.Id_Mecanico = '" & Me.dbcboMecanico.BoundText & "') "
End If
If Me.dbCboConceptoServicio.Text <> "" Then
    sql = sql & "AND Srvt_Servicios_OT.Id_Concepto_Servicio = '" & Me.dbCboConceptoServicio.BoundText & "' "
End If
If Me.dbCboServicio.Text <> "" Then
    sql = sql & "AND Srvt_Servicios_OT.Id_Servicio = '" & Me.dbCboServicio.BoundText & "' "
End If
If Me.txtOtDesde.Text <> "" And Me.txtOtHasta.Text <> "" Then
    sql = sql & "AND Srvt_OT.Id_OT BETWEEN '" & Me.txtOtDesde.Text & "' AND '" & Me.txtOtHasta.Text & "' "
End If
sql = sql & "GROUP BY Srvt_Servicios_OT.Id_Servicio "
sql = sql & "HAVING (Srvt_Servicios_OT.Id_Servicio = '" & idServicio & "')"
If Conexion.SendHost(sql, tablaValores, adOpenForwardOnly, adLockReadOnly, gcTiempoEspera * 2) = apOk Then
    If tablaValores.EOF = False And tablaValores.BOF = False Then
        If Not IsNull(tablaValores!Valor_Fac) Then
            TraeValoresResumen = ValorNuloNum(tablaValores!Valor_Fac)
        Else
            TraeValoresResumen = ValorNuloNum(tablaValores!Valor_OT)
        End If
    End If
End If
Conexion.CloseHost tablaValores
End Function

Private Sub dbcboMecanico_Click(Area As Integer)
If Area = 2 Then
    Me.lsvComisiones.ListItems.Clear
End If
End Sub

Private Sub dbCboServicio_Click(Area As Integer)
If Area = 2 Then
    Me.lsvComisiones.ListItems.Clear
End If
End Sub

Private Sub dtpDesde_Change()
Me.lsvComisiones.ListItems.Clear
End Sub

Private Sub dtpHasta_Change()
Me.lsvComisiones.ListItems.Clear
End Sub

Private Sub Form_Load()
Me.Label8.Caption = gstrNombreSucursal
LLena_Mecanico
LLena_TipoServicio
LLena_Sucursales
Me.dtpDesde.Value = "01/" & Format$(Date, "mm/yyyy")
Me.dtpHasta.Value = Date
Me.dbcboSucursal.BoundText = gstrIdSucursal
End Sub

Public Sub LLena_Sucursales()
Dim Tabla As New ADODB.Recordset
Dim sql As String

sql = ""
sql = "SELECT Id_Sucursal, Descripcion FROM Glbl_Sucursal WHERE Id_Empresa = '" & gstrIdEmpresa & "' AND Vigencia = 'S' ORDER BY Descripcion"
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    Set Me.datSucursal.Recordset = Tabla
    Set Tabla = New ADODB.Recordset
End If

End Sub

Public Sub LLena_Mecanico()
Dim Tabla As New ADODB.Recordset
Dim sql As String

sql = ""
sql = "SELECT Id_Mecanico, Nombre FROM Tllr_Mecanicos WHERE Vigencia='S' ORDER BY Nombre"
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    Set Me.datMecanico.Recordset = Tabla
    Set Tabla = New ADODB.Recordset
End If

End Sub

Public Sub LLena_TipoServicio()
Dim Tabla As New ADODB.Recordset
Dim sql As String

sql = ""
sql = "SELECT Id_Concepto_Servicio, Descripcion FROM Srvt_Concepto_Servicio WHERE Vigencia='S' ORDER BY Descripcion"
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    Set Me.datConceptoServicio.Recordset = Tabla
    Set Tabla = New ADODB.Recordset
End If

End Sub

Public Sub LLena_Servicio()
Dim Tabla As New ADODB.Recordset
Dim sql As String

Me.dbCboServicio.Text = ""
sql = ""
sql = "SELECT Id_Servicio, Descripcion FROM Srvt_Servicios WHERE Id_Concepto_Servicio='" & Me.dbCboConceptoServicio.BoundText & "' AND Vigencia='S' ORDER BY Descripcion"
If Conexion.SendHost(sql, Tabla, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    Set Me.datServicio.Recordset = Tabla
    Set Tabla = New ADODB.Recordset
End If

End Sub

Private Sub opcAgrupaOT_Click()
If Me.opcAgrupaOT.Value = True Then
    Me.dbCboConceptoServicio.Text = ""
    Me.dbCboServicio.Text = ""
    Me.dbCboConceptoServicio.Enabled = False
    Me.dbCboServicio.Enabled = False
    Me.tlbBotones(1).Buttons(1).Enabled = False
    Me.tlbBotones(2).Buttons(1).Enabled = False
Else
    Me.dbCboConceptoServicio.Enabled = True
    Me.dbCboServicio.Enabled = True
    Me.tlbBotones(1).Buttons(1).Enabled = True
    Me.tlbBotones(2).Buttons(1).Enabled = True
End If
End Sub

Private Sub opcAgrupaServicio_Click()
If Me.opcAgrupaOT.Value = True Then
    Me.dbCboConceptoServicio.Text = ""
    Me.dbCboServicio.Text = ""
    Me.dbCboConceptoServicio.Enabled = False
    Me.dbCboServicio.Enabled = False
    Me.tlbBotones(1).Buttons(1).Enabled = False
    Me.tlbBotones(2).Buttons(1).Enabled = False
Else
    Me.dbCboConceptoServicio.Enabled = True
    Me.dbCboServicio.Enabled = True
    Me.tlbBotones(1).Buttons(1).Enabled = True
    Me.tlbBotones(2).Buttons(1).Enabled = True
End If
End Sub

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
Screen.MousePointer = vbHourglass
Select Case Button.Key
    Case "Buscar"
        Buscar
    Case "Imprimir"
        ImprimirConsulta
    Case "Cerrar"
        Unload Me
End Select
Screen.MousePointer = vbDefault

End Sub

Private Sub tlbBotones_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

Screen.MousePointer = vbHourglass
Select Case Button.Key
    Case "Cancelar"
        Select Case Index
            Case 0
                Me.dbcboMecanico.Text = ""
            Case 1
                Me.dbCboConceptoServicio.Text = ""
            Case 2
                Me.dbCboServicio.Text = ""
            Case 3
                Me.dbcboSucursal.Text = ""
        End Select
End Select
Screen.MousePointer = vbDefault
Me.lsvComisiones.ListItems.Clear
End Sub

Private Sub tlbOtrosCriterios_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Value = tbrPressed Then
    Me.OtrosCriterios.Visible = True
Else
    Me.OtrosCriterios.Visible = False
End If
End Sub

Private Sub txtOtDesde_Change()
Me.lsvComisiones.ListItems.Clear
End Sub

Private Sub txtOtDesde_GotFocus()
Me.txtOtDesde.SelStart = 0
Me.txtOtDesde.SelLength = Len(Me.txtOtDesde.Text)
End Sub

Private Sub txtOtDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtOtHasta.Text = Me.txtOtDesde.Text
    Me.txtOtHasta.SetFocus
End If
End Sub

Private Sub txtOtHasta_Change()
Me.lsvComisiones.ListItems.Clear
End Sub

Private Sub txtOtHasta_GotFocus()
Me.txtOtHasta.SelStart = 0
Me.txtOtHasta.SelLength = Len(Me.txtOtHasta.Text)
End Sub
