VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSelTempRepuestos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Repuestos"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   Icon            =   "frmSelTempRepuestos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   8295
      Top             =   2565
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":179A
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":18AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":1D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":215C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":25B4
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":26C6
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":27D8
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":28EA
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":29FC
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":2B0E
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":2C20
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":2D32
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":2E44
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":2F56
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":3068
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":317A
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":328C
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":339E
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":34B0
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":35C2
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":3A14
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":3E66
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelTempRepuestos.frx":3F78
            Key             =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwRepuestos 
      Height          =   3450
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   6085
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción Repuesto"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Familia"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Procedencia"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Prefijo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Basico"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Sufijo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "IdFam"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "IdPro"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Aplicación"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Precio_Venta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Precio_Flota"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbOpciones 
      Height          =   330
      Index           =   1
      Left            =   5280
      TabIndex        =   9
      Top             =   6285
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      ButtonWidth     =   1746
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            Key             =   "Buscar"
            Object.ToolTipText     =   "Ejecutar Busqueda"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Agregar"
            Key             =   "Agregar"
            Object.ToolTipText     =   "Agregar la Selección de Repuestos"
            ImageKey        =   "Editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar"
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar Formulario de Busqueda"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbOpciones 
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   21
      Top             =   6240
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ButtonWidth     =   1746
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Todos"
            Key             =   "SelectAll"
            Object.ToolTipText     =   "Seleccionar toda la Lista"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ninguno"
            Key             =   "UnSelectAll"
            Object.ToolTipText     =   "Quitar Seleccion a toda la Lista"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Criterios de Busqueda"
      Height          =   4680
      Left            =   30
      TabIndex        =   10
      Top             =   0
      Width           =   8415
      Begin VB.CheckBox optCriterios 
         Caption         =   "Aplicación"
         Height          =   195
         Index           =   2
         Left            =   6045
         TabIndex        =   4
         Tag             =   "APLICACION"
         Top             =   3510
         Width           =   1200
      End
      Begin VB.TextBox txtAplicacion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1365
         TabIndex        =   5
         Top             =   960
         Width           =   4650
      End
      Begin MSDataListLib.DataCombo dtcSubFamilia 
         Bindings        =   "frmSelTempRepuestos.frx":42CA
         Height          =   315
         Left            =   3600
         TabIndex        =   7
         Top             =   1635
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSComctlLib.Toolbar tlbLimpiaMarca 
         Height          =   330
         Left            =   3120
         TabIndex        =   28
         Top             =   2355
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
               Key             =   "Limpiar"
               Object.ToolTipText     =   "Limpiar"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.UpDown updNroRecord 
         Height          =   315
         Left            =   7980
         TabIndex        =   23
         Top             =   270
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   5
         BuddyControl    =   "txtNroRecord"
         BuddyDispid     =   196612
         OrigLeft        =   8445
         OrigTop         =   300
         OrigRight       =   8685
         OrigBottom      =   615
         Max             =   100
         Min             =   5
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtNroRecord 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7470
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "10"
         Top             =   270
         Width           =   540
      End
      Begin VB.CheckBox optCriterios 
         Caption         =   "Sufijo"
         Height          =   195
         Index           =   7
         Left            =   6975
         TabIndex        =   20
         Tag             =   "SUFIJO"
         Top             =   3675
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox optCriterios 
         Caption         =   "Básico"
         Height          =   195
         Index           =   6
         Left            =   5490
         TabIndex        =   19
         Tag             =   "BASICO"
         Top             =   3675
         Visible         =   0   'False
         Width           =   795
      End
      Begin MSDataListLib.DataCombo dtcProcedencia 
         Bindings        =   "frmSelTempRepuestos.frx":42E6
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   2355
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin VB.TextBox txtSufijo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4560
         TabIndex        =   16
         Top             =   270
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox txtBasico 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2940
         TabIndex        =   15
         Top             =   270
         Width           =   1470
      End
      Begin VB.TextBox txtPrefijo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1365
         TabIndex        =   14
         Top             =   270
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.ComboBox cboCoincidir 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSelTempRepuestos.frx":4303
         Left            =   6120
         List            =   "frmSelTempRepuestos.frx":4313
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   960
         Width           =   1980
      End
      Begin VB.TextBox txtDes 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1365
         TabIndex        =   3
         Top             =   600
         Width           =   4650
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   3600
         Width           =   2235
      End
      Begin VB.CheckBox optCriterios 
         Caption         =   "Código"
         Height          =   195
         Index           =   0
         Left            =   6045
         TabIndex        =   0
         Tag             =   "CODIGO"
         Top             =   2805
         Width           =   840
      End
      Begin MSDataListLib.DataCombo dtcFamilia 
         Bindings        =   "frmSelTempRepuestos.frx":4366
         Height          =   315
         Left            =   150
         TabIndex        =   6
         Top             =   1635
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc datFamilia 
         Height          =   330
         Left            =   840
         Top             =   1680
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
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
      Begin VB.CheckBox optCriterios 
         Caption         =   "Descripción"
         Height          =   195
         Index           =   1
         Left            =   6045
         TabIndex        =   2
         Tag             =   "DESCRIPCION"
         Top             =   3120
         Width           =   1200
      End
      Begin MSAdodcLib.Adodc datProcedencia 
         Height          =   330
         Left            =   720
         Top             =   2400
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
         Caption         =   "Adodc4"
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
      Begin MSDataListLib.DataCombo dtcUnidadMedida 
         Bindings        =   "frmSelTempRepuestos.frx":437F
         Height          =   315
         Left            =   3600
         TabIndex        =   13
         Top             =   2355
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Nombre"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc datUnidadMedida 
         Height          =   330
         Left            =   4200
         Top             =   2400
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
      Begin VB.CheckBox optCriterios 
         Caption         =   "Prefijo"
         Height          =   195
         Index           =   5
         Left            =   4020
         TabIndex        =   18
         Tag             =   "PREFIJO"
         Top             =   3675
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.Toolbar tlbLimpiaModelo 
         Height          =   330
         Left            =   7080
         TabIndex        =   29
         Top             =   2355
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
               Key             =   "Limpiar"
               Object.ToolTipText     =   "Limpiar"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc datSubFamilia 
         Height          =   330
         Left            =   4080
         Top             =   1680
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
      Begin MSComctlLib.Toolbar tlbLimpiaSubFamilia 
         Height          =   330
         Left            =   7080
         TabIndex        =   31
         Top             =   1635
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
               Key             =   "Limpiar"
               Object.ToolTipText     =   "Limpiar"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   3120
         TabIndex        =   35
         Top             =   1635
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
               Key             =   "Limpiar"
               Object.ToolTipText     =   "Limpiar"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Aplicación:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label5 
         Caption         =   "SubFamilia"
         Height          =   195
         Left            =   3600
         TabIndex        =   30
         Top             =   1440
         Width           =   2205
      End
      Begin VB.Label Label4 
         Caption         =   "Modelo Vehículo"
         Height          =   255
         Left            =   3600
         TabIndex        =   27
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Marca Vehículo"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Familia"
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   1440
         Width           =   2280
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro. de Registros :"
         Height          =   195
         Index           =   1
         Left            =   6090
         TabIndex        =   24
         Top             =   270
         Width           =   1320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   2
         X1              =   150
         X2              =   8235
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   3
         X1              =   150
         X2              =   8235
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   150
         X2              =   8235
         Y1              =   2070
         Y2              =   2070
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   165
         X2              =   8250
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Coincidir en :"
         Height          =   195
         Index           =   0
         Left            =   6120
         TabIndex        =   12
         Top             =   720
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmSelTempRepuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrSQL As String
Dim strWhere As String, strOrder As String
Dim lsiItem As ListItem
Dim mblnSW As Boolean
Dim AdoPrincipal As New ADODB.Recordset
Dim intContador As Integer
Dim itmFound As ListItem
Dim itmLista As ListItem
Public VerPreSufijo As String
Dim mstrSQLCli As String
Dim AdorecordCliente As New ADODB.Recordset
Public Flotista As String

Sub FillFamilias()
    Set AdoPrincipal = New ADODB.Recordset
    mstrSQL = "SELECT Id_Familia as codigo, Descripcion as Nombre FROM Glbl_Familia WHERE Id_Empresa='" & gstrIdEmpresa & "' and Vigencia = 'S' ORDER BY DESCRIPCION"
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        With datFamilia
            Set .Recordset = AdoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcFamilia.ListField = "Nombre"
                dtcFamilia.BoundColumn = "Codigo"
            End If
        End With
    End If ' por el otro
    Set AdoPrincipal = New ADODB.Recordset
    Conexion.CloseHost AdoPrincipal
End Sub
Sub FillUnidades()
    Set AdoPrincipal = New ADODB.Recordset
    mstrSQL = "SELECT Id_Unidad_Medida as codigo, Descripcion as nombre FROM Glbl_Unidad_Medida WHERE (Vigencia = N'S')"
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        With datUnidadMedida
            Set .Recordset = AdoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcUnidadMedida.ListField = "Nombre"
                dtcUnidadMedida.BoundColumn = "Codigo"
            End If
        End With
    End If ' por el otro
    Set AdoPrincipal = New ADODB.Recordset
    Conexion.CloseHost AdoPrincipal
End Sub
Sub FillProcedencias()
    Set AdoPrincipal = New ADODB.Recordset
    mstrSQL = "SELECT Id_Pais AS Codigo, Descripcion AS Nombre FROM Glbl_Pais WHERE Vigencia = 'S' "
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
        With datProcedencia
            Set .Recordset = AdoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                dtcProcedencia.ListField = "Nombre"
                dtcProcedencia.BoundColumn = "Codigo"
            End If
        End With
    End If ' por el otro
    Set AdoPrincipal = New ADODB.Recordset
    Conexion.CloseHost AdoPrincipal
End Sub

Private Sub dtcFamilia_Change()
If dtcFamilia.BoundText <> "" Then
    dtcSubFamilia.Text = ""
    FillSubFamilia dtcFamilia.BoundText
End If
End Sub

Private Sub dtcProcedencia_Change()
If dtcProcedencia.BoundText <> "" Then
    dtcUnidadMedida.Text = ""
    FillModelos dtcProcedencia.BoundText
End If
End Sub

Private Sub Form_Activate()
If mblnSW Then
    FillFamilias
    FillMarcas
    If gstrProcedencia = "Movimientos" Or gstrProcedencia = "Presupuestos" Then
        Me.dtcProcedencia.BoundText = frmRecepcion.lblIdMarca
        Me.dtcUnidadMedida.BoundText = frmRecepcion.lblIdModelo
    Else
        Me.dtcProcedencia.BoundText = frmTempServiciosMarMod.dtcMarca.BoundText
        Me.dtcUnidadMedida.BoundText = frmTempServiciosMarMod.dtcModelo.BoundText
    End If
    
    'FillUnidades
    'FillProcedencias
    
'    If gstrProcedencia = "Movimientos" Then
'    With Me
'        .optCriterios.Item(2).Value = 1
'        .optCriterios.Item(2).Enabled = False
'        '.dtcFamilia.BoundText = FamMateriales("Materiales")
''        .dtcFamilia.Enabled = False
'        .dtcProcedencia.Enabled = False
'        .dtcUnidadMedida.Enabled = False
'    End With
'    End If

    mblnSW = False
    'kjcv 09.06.14
    
    VerPreSufijo = Retorna_Valor_General("Select Verprefijosufijo from stck_parametro where id_empresa='" & gstrIdEmpresa & "'and id_sucursal='" & gstrIdSucursal & "'")
    If VerPreSufijo = "1" Then
    '    Me.txtprefijo.Enabled = True
        Me.txtPrefijo.Visible = True
    '    Me.txtsufijo.Enabled = True
        Me.txtSufijo.Visible = True
    Else
        Me.txtPrefijo.Enabled = False
        Me.txtPrefijo.Visible = False
        Me.txtSufijo.Enabled = False
        Me.txtSufijo.Visible = False
    End If
    
    If VerPreSufijo = "1" Then
        Me.txtPrefijo.SetFocus
    Else
        Me.txtBasico.SetFocus
    End If
End If
End Sub
Private Sub Form_Load()
mblnSW = True
updNroRecord.Value = gintNroRecDefectoQry

'parametriza familia y subfamilia como grupo y subgrupo
mstrSQL = "Select Nombre_Grupo,Nombre_Subgrupo from Stck_Parametro Where Id_Empresa='" & gstrIdEmpresa & "' And Id_Sucursal='" & gstrIdSucursal & "'"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
        If Not .BOF And Not .EOF Then
            Label2.Caption = !Nombre_Grupo
            Label5.Caption = !Nombre_SubGrupo
            
            Me.lvwRepuestos.ColumnHeaders(5).Text = !Nombre_Grupo
        End If
    End With
End If
AdoPrincipal.Close
Me.cboCoincidir.ListIndex = 0
End Sub
Sub FillRepuestos(strCondicion As String, strOrdenamiento As String)
    
    If txtCodigo = "" Then
        lvwRepuestos.ListItems.Clear
    End If
'    mstrSql = "SELECT TOP " & CStr(updNroRecord.Value) & " Stck_Item.Id_Item AS CODIGO,"
'    mstrSql = mstrSql & " Stck_Item.Descripcion AS NOMBRE,"
'    mstrSql = mstrSql & " Stck_Item.Precio_Venta AS PRECIO,"
'    mstrSql = mstrSql & " Glbl_Familia.Descripcion AS FAMILIA,"
'    mstrSql = mstrSql & " Glbl_Pais.Descripcion AS PROCEDENCIA,"
'    mstrSql = mstrSql & " Stck_Item.Prefijo AS PREFIJO, Stck_Item.Basico AS BASICO,"
'    mstrSql = mstrSql & " Stck_Item.Sufijo AS SUFIJO, Stck_Item.Id_Familia AS IDFAM,"
'    mstrSql = mstrSql & " Stck_Item.Procedencia AS IDPRO"
'    mstrSql = mstrSql & " FROM Stck_Item LEFT OUTER JOIN"
'    mstrSql = mstrSql & " Stck_Item_Modelo ON"
'    mstrSql = mstrSql & " Stck_Item.Id_Item = Stck_Item_Modelo.Id_Item LEFT OUTER JOIN"
'    mstrSql = mstrSql & " Glbl_Pais ON"
'    mstrSql = mstrSql & " Stck_Item.Procedencia = Glbl_Pais.Id_Pais LEFT OUTER JOIN"
'    mstrSql = mstrSql & " Glbl_Familia ON"
'    mstrSql = mstrSql & " Stck_Item.Id_Familia = Glbl_Familia.Id_Familia"
'    mstrSql = mstrSql & strCondicion & strOrdenamiento

    Dim strCoincidir As String
    strCoincidir = Me.cboCoincidir.ListIndex + 1
    mstrSQL = "exec dbo.Stck_Buscar_Item @Id_Empresa='" & gstrIdEmpresa & "', @Id_Familia='" & Me.dtcFamilia.BoundText & "', @Id_SubFamilia='" & Me.dtcSubFamilia.BoundText & "', @Prefijo='" & Me.txtPrefijo & "', @Basico='" & Me.txtBasico & "', @Sufijo='" & Me.txtSufijo & "', @Descripcion='" & Me.txtDes & "', @Coincidir='" & strCoincidir & "', @Id_Marca='" & Me.dtcProcedencia.BoundText & "', @Id_Modelo='" & Me.dtcUnidadMedida.BoundText & "', @CodModelo='" & Me.txtAplicacion & "'"
If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
    With AdoPrincipal
        If Not .BOF And Not .EOF Then
            .MoveFirst
            While Not .EOF
                'Set lsiItem = lvwRepuestos.ListItems.Add(, , !Codigo)
                Set lsiItem = lvwRepuestos.ListItems.Add(, , !Id_Item)
                'lsiItem.SubItems(1) = !Nombre
                lsiItem.SubItems(1) = !Descripcion
                'lsiItem.SubItems(2) = Format(!Precio, "###,##0.0")
'                lsiItem.SubItems(2) = Format(!Precio_Venta * traeParidadMonedaMes("02", frmRecepcion.pckFechaAtencion), "###,##0.00")
                lsiItem.SubItems(2) = Format(!Precio_Venta * traeParidadMonedaMes("02", Now()), "###,##0.00")
                'Se extrae Precio Flota
                lsiItem.SubItems(12) = Format(!Precio_Flota * traeParidadMonedaMes("02", Now()), "###,##0.00")
                
'                'kjcv 01.03.13 Cambio de Taller a dolares
''                lsiItem.SubItems(2) = Format(!Precio_Venta, "###,##0.00")
                'lsiItem.SubItems(3) = IIf(Not IsNull(!Familia), !Familia, "(Ninguna)")
                lsiItem.SubItems(3) = !Nombre_familia
                'lsiItem.SubItems(4) = IIf(Not IsNull(!Procedencia), !Procedencia, "(Ninguna)")
                lsiItem.SubItems(4) = !Procedencia
                lsiItem.SubItems(5) = !prefijo
                lsiItem.SubItems(6) = !basico
                lsiItem.SubItems(7) = !sufijo
                'lsiItem.SubItems(8) = IIf(Not IsNull(!IDFAM), !IDFAM, "")
                lsiItem.SubItems(8) = !Id_Familia
                'lsiItem.SubItems(9) = IIf(Not IsNull(!IDPRO), !IDPRO, "")
                lsiItem.SubItems(10) = !Cod_Modelo
                lsiItem.SubItems(11) = !Precio_Venta
                .MoveNext
            Wend
        End If
    End With
End If
End Sub

Private Sub lvwRepuestos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'MsgBox item.Index
End Sub

'Private Sub optCriterios_Click(Index As Integer)
'    Select Case Index
'    Case 0 '//////////////CODIGO
'        If optCriterios(Index).Value = 1 Then
'            optCriterios(Index + 1).Value = 0: txtDes.Text = "": txtDes.Enabled = False
'            cboCoincidir.ListIndex = 0
'            txtCodigo.Enabled = True
'            txtCodigo.SetFocus
'        Else
'            optCriterios(Index).Value = 0
'            txtCodigo.Enabled = False
'            txtCodigo.Text = ""
'        End If
'    Case 1 '//////////////DESCRIPCION
'        If optCriterios(Index).Value = 1 Then
'            optCriterios(Index - 1).Value = 0: txtCodigo.Text = "": txtCodigo.Enabled = False
'            cboCoincidir.ListIndex = 0
'            txtDes.Enabled = True
'            txtDes.SetFocus
'        Else
'            optCriterios(Index).Value = 0
'            txtDes.Enabled = False
'            txtDes.Text = ""
'        End If
'    End Select
'End Sub

Private Sub tlbLimpiaMarca_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Limpiar"
    Me.dtcProcedencia.Text = ""
    Me.dtcUnidadMedida.Text = ""
End Select

End Sub

Private Sub tlbLimpiaModelo_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Limpiar"
    'Me.dtcProcedencia.Text = ""
    Me.dtcUnidadMedida.Text = ""
End Select
End Sub

Private Sub tlbLimpiaSubFamilia_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Limpiar"
    Me.dtcSubFamilia.Text = ""
End Select

End Sub

Sub tlbOpciones_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Select Case Index
Case 0
    Select Case Button.Key
    Case "SelectAll"
        SelectingItem lvwRepuestos, gcSelectAll
    Case "UnSelectAll"
        SelectingItem lvwRepuestos, gcUnSelectAll
    End Select
Case 1
    Select Case Button.Key
    Case "Agregar"
    If gstrProcedencia = "Temparios" Then
        For intContador = 1 To lvwRepuestos.ListItems.Count - 1
            Set lvwRepuestos.SelectedItem = lvwRepuestos.ListItems(intContador)
            If lvwRepuestos.ListItems(intContador).Checked = True Then
                Set itmFound = frmTempServiciosMarMod.lvwRepuestos.FindItem(lvwRepuestos.SelectedItem, lvwText, , 0)
                If itmFound Is Nothing Then   ' Si no hay coincidencia                                    ' usuario y sale.
                    Set itmFound = frmTempServiciosMarMod.lvwRepuestos.ListItems.Add(, , lvwRepuestos.ListItems(intContador))
                    itmFound.SubItems(1) = lvwRepuestos.ListItems(intContador).SubItems(1) ' descripcion
                    itmFound.SubItems(2) = FormatoValor(1, "", 1) 'LvwRepuestos.ListItems(intContador).SubItems(2) ' cantidad
                    itmFound.SubItems(3) = lvwRepuestos.ListItems(intContador).SubItems(2) ' valor
                    itmFound.SubItems(4) = lvwRepuestos.ListItems(intContador).SubItems(3) ' familia
                    itmFound.SubItems(5) = lvwRepuestos.ListItems(intContador).SubItems(8) ' idfam
                    '/*//////////////////////////
                    
                    '// pregunto si las lista de actividades viene vacia, si es asi agrega la actividad generica
                    If frmTempServiciosMarMod.lvwActividades.SelectedItem Is Nothing Then
                        mstrSQL = "INSERT INTO Tllr_Actividad_Servicio_Modelo ( Id_Marca, Id_Modelo, Id_Servicio, ID_ACTIVIDAD, Horas, Valor ) "
                        mstrSQL = mstrSQL & " VALUES( '" & frmTempServiciosMarMod.dtcMarca.BoundText & "' , "
                        mstrSQL = mstrSQL & " '" & frmTempServiciosMarMod.dtcModelo.BoundText & "' , '" & frmTempServiciosMarMod.lvwServicios.SelectedItem & "' , "
                        mstrSQL = mstrSQL & " 'CAREP' ,0 ,0 " & ") "
                        Conexion.SendHost mstrSQL, , , , gcTiempoEspera
                        '/// lleno la lista con la actividad generica
                        Set itmFound = frmTempServiciosMarMod.lvwActividades.ListItems.Add(, , "CAREP")
                        
                    End If
                    
                    mstrSQL = "INSERT INTO Tllr_Actividad_Repuesto ( Id_Marca, Id_Modelo, Id_Servicio, Id_Actividad, Id_Item, Cantidad , Valor) "
                    mstrSQL = mstrSQL & " VALUES( '" & frmTempServiciosMarMod.dtcMarca.BoundText & "' , "
                    mstrSQL = mstrSQL & " '" & frmTempServiciosMarMod.dtcModelo.BoundText & "' , "
                    mstrSQL = mstrSQL & " '" & frmTempServiciosMarMod.lvwServicios.SelectedItem & "' , "
                    mstrSQL = mstrSQL & " '" & IIf(frmTempServiciosMarMod.lvwActividades.SelectedItem Is Nothing, "CAREP", frmTempServiciosMarMod.lvwActividades.SelectedItem) & "' , "
                    mstrSQL = mstrSQL & " '" & lvwRepuestos.ListItems(intContador) & "' , "
                    mstrSQL = mstrSQL & " " & CCur(1) & " , "
                    mstrSQL = mstrSQL & " " & CCur(Format(lvwRepuestos.ListItems(intContador).SubItems(2), "####0")) & ") "
                    
                    
                    Conexion.SendHost mstrSQL, , , , gcTiempoEspera
                End If
            End If
        Next
    ElseIf gstrProcedencia = "Movimientos" Or gstrProcedencia = "Presupuestos" Then
        If gstrProcedenciaRptos <> "Mantencion" Then
            'gstrIdCargo = gstrIdCargoDefecto
            For intContador = 1 To lvwRepuestos.ListItems.Count
                Set lvwRepuestos.SelectedItem = lvwRepuestos.ListItems(intContador)
                If lvwRepuestos.ListItems(intContador).Checked = True Then
                    Set itmFound = frmRecepcion.lvwRepuestos.FindItem(lvwRepuestos.SelectedItem, lvwText, , 0)
                    If itmFound Is Nothing Then   ' Si no hay coincidencia                                    ' usuario y sale.
                        Set itmFound = frmRecepcion.lvwRepuestos.ListItems.Add(, , lvwRepuestos.ListItems(intContador))
                        Set frmRecepcion.lvwRepuestos.SelectedItem = itmFound
                        itmFound.SubItems(9) = lvwRepuestos.ListItems(intContador).SubItems(8)
                        itmFound.SubItems(1) = lvwRepuestos.ListItems(intContador).SubItems(1)
                        itmFound.SubItems(2) = Format(1, "#0.00")
                        ' Precio Flota
                        If Flotista <> "" Then
                            itmFound.SubItems(3) = lvwRepuestos.ListItems(intContador).SubItems(12)
                        Else
                            itmFound.SubItems(3) = lvwRepuestos.ListItems(intContador).SubItems(2)
                        End If
                        itmFound.SubItems(4) = Format(0, "#0.00")
                        itmFound.SubItems(5) = Format(0, "###,##0.00")
                        itmFound.SubItems(6) = TraeCargoDes(gstrIdCargo)
                        itmFound.SubItems(7) = gstrIdCargo
                        itmFound.SubItems(8) = Format(frmRecepcion.CalculoSubTotal(mcFichaRepuestos), "###,##0.00")
                        itmFound.SubItems(10) = "N"
                        itmFound.SubItems(11) = "PRESUPUESTO"
                        'kjcv 23.08.16
                        itmFound.SubItems(15) = lvwRepuestos.ListItems(intContador).SubItems(11)
                    End If
                End If
            Next
        Else  '//// agrega repuestos de servicio de mantenciones
            For intContador = 1 To lvwRepuestos.ListItems.Count
                Set lvwRepuestos.SelectedItem = lvwRepuestos.ListItems(intContador)
                If lvwRepuestos.ListItems(intContador).Checked = True Then
                    Set itmFound = frmRecepcion.lvwRepuestosMantencion.FindItem(lvwRepuestos.SelectedItem, lvwText, , 0)
                    If itmFound Is Nothing Then   ' Si no hay coincidencia                                    ' usuario y sale.
                        Set itmFound = frmRecepcion.lvwRepuestosMantencion.ListItems.Add(, , lvwRepuestos.ListItems(intContador))
                        Set frmRecepcion.lvwRepuestosMantencion.SelectedItem = itmFound
                        itmFound.SubItems(1) = lvwRepuestos.ListItems(intContador).SubItems(1)
                        itmFound.SubItems(2) = FormatoValor(1, "", 1)
                        itmFound.SubItems(3) = lvwRepuestos.ListItems(intContador).SubItems(2)
                        itmFound.SubItems(4) = lvwRepuestos.ListItems(intContador).SubItems(3)
                        itmFound.SubItems(5) = lvwRepuestos.ListItems(intContador).SubItems(8)
                        itmFound.SubItems(6) = "N"
                    End If
                End If
            Next
        End If
        'Unload Me
    ElseIf gstrProcedencia = "Presupuesto" Then
        gstrIdCargo = gstrIdCargoDefecto
        'kjcv 21.05.21
            Flotista = ""
            gCliente = Retorna_Valor_General("Select Id_Cliente_Proveedor from Tllr_OT where Estado='P' and Id_OT='" & frmRecepcion.lblNroRecepcion & "'")
            
            mstrSQLCli = "SELECT isnull(id_lista_precio,'') as Lista, isnull(Glbl_Lista_Precios.Descuento_Maximo,0) as Dscto "
            mstrSQLCli = mstrSQLCli & " from Glbl_Cliente_Proveedor inner join Glbl_Lista_Precios on "
            mstrSQLCli = mstrSQLCli & "Glbl_Lista_Precios.Id_Lista = Glbl_Cliente_Proveedor.id_lista_precio "
            mstrSQLCli = mstrSQLCli & "where Glbl_Cliente_Proveedor.Id_Cliente_Proveedor='" & gCliente & "' and  Glbl_Cliente_Proveedor.id_lista_precio='02'"
            If Conexion.SendHost(mstrSQLCli, AdorecordCliente, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
                If Not AdorecordCliente.BOF And Not AdorecordCliente.EOF Then
                'se obtiene el descuento
                    'lDscto = AdorecordCliente!Dscto
                    Flotista = AdorecordCliente!Lista
                End If
            End If
            Conexion.CloseHost AdorecordCliente
                    
        
        For intContador = 1 To lvwRepuestos.ListItems.Count
            Set lvwRepuestos.SelectedItem = lvwRepuestos.ListItems(intContador)
            If lvwRepuestos.ListItems(intContador).Checked = True Then
                Set itmFound = frmPresupuesto.lvwRepuestos.FindItem(lvwRepuestos.SelectedItem, lvwText, , 0)
                If itmFound Is Nothing Then   ' Si no hay coincidencia                                    ' usuario y sale.
                    Set itmFound = frmPresupuesto.lvwRepuestos.ListItems.Add(, , lvwRepuestos.ListItems(intContador))
                    Set frmPresupuesto.lvwRepuestos.SelectedItem = itmFound
                    itmFound.SubItems(9) = lvwRepuestos.ListItems(intContador).SubItems(8)
                    itmFound.SubItems(1) = lvwRepuestos.ListItems(intContador).SubItems(1)
                    itmFound.SubItems(2) = Format(1, "#0.00")
                    itmFound.SubItems(3) = lvwRepuestos.ListItems(intContador).SubItems(2)
                    itmFound.SubItems(4) = Format(0, "#0.00")
                    itmFound.SubItems(5) = Format(0, "###,##0.00")
                    itmFound.SubItems(6) = TraeCargoDes(gstrIdCargo)
                    itmFound.SubItems(7) = gstrIdCargo
                    itmFound.SubItems(8) = Format(frmPresupuesto.CalculoSubTotal(mcFichaRepuestos), "###,##0.00")
                End If
            End If
        Next
        Unload Me
    Else
        DoEvents
    End If
    Case "Buscar" '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''    Flotista = ""
''    gCliente = Retorna_Valor_General("Select Id_Cliente_Proveedor from Tllr_OT where Estado='P' and Id_OT='" & frmRecepcion.lblNroRecepcion & "'")
''
''                    mstrSQLCli = "SELECT isnull(id_lista_precio,'') as Lista, isnull(Glbl_Lista_Precios.Descuento_Maximo,0) as Dscto "
''                    mstrSQLCli = mstrSQLCli & " from Glbl_Cliente_Proveedor inner join Glbl_Lista_Precios on "
''                    mstrSQLCli = mstrSQLCli & "Glbl_Lista_Precios.Id_Lista = Glbl_Cliente_Proveedor.id_lista_precio "
''                    mstrSQLCli = mstrSQLCli & "where Glbl_Cliente_Proveedor.Id_Cliente_Proveedor='" & gCliente & "' and  Glbl_Cliente_Proveedor.id_lista_precio='02'"
''                    If Conexion.SendHost(mstrSQLCli, AdorecordCliente, adOpenKeyset, adLockOptimistic, gcTiempoEspera) = apOk Then
''                        If Not AdorecordCliente.BOF And Not AdorecordCliente.EOF Then
''                        'se obtiene el descuento
''                            'lDscto = AdorecordCliente!Dscto
''                            Flotista = AdorecordCliente!Lista
''                        End If
''                    End If
''                    Conexion.CloseHost AdorecordCliente
                    
                    
                    
    
        FillRepuestos "", ""
'        If txtCodigo <> "" Then
'            If strWhere <> "" Then
'                strWhere = strWhere & " AND Stck_Item.ID_ITEM LIKE '" & MatchMode(txtCodigo, cboCoincidir.Text, apSqlServer) & "' "
'                strOrder = strOrder & " , Stck_Item.Id_Item"
'            Else
'                strWhere = " WHERE Stck_Item.ID_ITEM LIKE '" & MatchMode(txtCodigo, cboCoincidir.Text, apSqlServer) & "' "
'                strOrder = " Order By Stck_Item.Id_Item"
'            End If
'        End If
'        If txtDes <> "" Then
'            If strWhere <> "" Then
'                strWhere = strWhere & " AND Stck_Item.DESCRIPCION LIKE '" & MatchMode(txtDes, cboCoincidir.Text, apSqlServer) & "' "
'                strOrder = strOrder & " , Stck_Item.Descripcion"
'            Else
'                strWhere = " WHERE Stck_Item.DESCRIPCION LIKE '" & MatchMode(txtDes, cboCoincidir.Text, apSqlServer) & "' "
'                strOrder = " Order By Stck_Item.Descripcion"
'            End If
'        End If
'        If dtcFamilia.BoundText <> "" Then
'            If strWhere <> "" Then
'                strWhere = strWhere & " AND Stck_Item.ID_FAMILIA = '" & dtcFamilia.BoundText & "'"
'                strOrder = strOrder & " , Stck_Item.Id_familia"
'            Else
'                strWhere = " WHERE Stck_Item.ID_FAMILIA = '" & dtcFamilia.BoundText & "'"
'                strOrder = " Order By Stck_Item.Id_familia"
'            End If
'        End If
'        If Me.dtcSubFamilia.BoundText <> "" Then
'            If strWhere <> "" Then
'                strWhere = strWhere & " AND Stck_Item.ID_SUBFAMILIA = '" & dtcSubFamilia.BoundText & "'"
'            Else
'                strWhere = " WHERE Stck_Item.ID_SUBFAMILIA = '" & dtcSubFamilia.BoundText & "'"
'            End If
'        End If
'        If dtcProcedencia.BoundText <> "" Then
'            If strWhere <> "" Then
'                'strWhere = strWhere & " AND Stck_Item.PROCEDENCIA ='" & dtcProcedencia.BoundText & "'"
'                'strOrder = strOrder & " , Stck_Item.Procedencia"
'                strWhere = strWhere & " AND Stck_Item_Modelo.Id_Marca='" & dtcProcedencia.BoundText & "'"
'            Else
'                'strWhere = " WHERE Stck_Item.PROCEDENCIA ='" & dtcProcedencia.BoundText & "'"
'                'strOrder = " Order By Stck_Item.Procedencia"
'                strWhere = " WHERE Stck_Item_Modelo.Id_Marca='" & dtcProcedencia.BoundText & "'"
'            End If
'        End If
'        If dtcUnidadMedida.BoundText <> "" Then
'            If strWhere <> "" Then
'                'strWhere = strWhere & " AND Stck_Item.ID_UNIDAD_MEDIDA= '" & dtcUnidadMedida.BoundText & "'"
'                'strOrder = strOrder & " Stck_Item.Id_Unidad_Medida"
'                strWhere = strWhere & " AND Stck_Item_Modelo.ID_Modelo= '" & dtcUnidadMedida.BoundText & "'"
'                'strOrder = strOrder & " Stck_Item_Modelo.Id_Modelo"
'            Else
'                'strWhere = " WHERE Stck_Item.ID_UNIDAD_MEDIDA= '" & dtcUnidadMedida.BoundText & "'"
'                'strOrder = " Order By Stck_Item.Id_Unidad_Medida"
'                strWhere = " WHERE Stck_Item_Modelo.ID_Modelo= '" & dtcUnidadMedida.BoundText & "'"
'                'strOrder = " Order By Stck_Item_Modelo.Id_Modelo"
'            End If
'        End If
'        If txtprefijo <> "" Then
'            If strWhere <> "" Then
'                strWhere = strWhere & " AND Stck_Item.PREFIJO LIKE '" & MatchMode(txtprefijo, cboCoincidir.Text, apSqlServer) & "'"
'                strOrder = strOrder & " Stck_Item.PREFIJO"
'            Else
'                strWhere = " WHERE Stck_Item.PREFIJO LIKE '" & MatchMode(txtprefijo, cboCoincidir.Text, apSqlServer) & "'"
'                strOrder = " Order By Stck_Item.PREFIJO"
'            End If
'        End If
'        If txtbasico <> "" Then
'            If strWhere <> "" Then
'                strWhere = strWhere & " AND Stck_Item.BASICO LIKE '" & MatchMode(txtbasico, cboCoincidir.Text, apSqlServer) & "' "
'                strOrder = strOrder & " Stck_Item.BASICO"
'            Else
'                strWhere = " WHERE Stck_Item.BASICO LIKE '" & MatchMode(txtbasico, cboCoincidir.Text, apSqlServer) & "' "
'                strOrder = " Order By Stck_Item.BASICO"
'            End If
'        End If
'        If txtsufijo <> "" Then
'            If strWhere <> "" Then
'                strWhere = strWhere & " AND Stck_Item.SUFIJO LIKE '" & MatchMode(txtsufijo, cboCoincidir.Text, apSqlServer) & "' "
'                strOrder = strOrder & " Stck_Item.SUFIJO"
'            Else
'                strWhere = " WHERE Stck_Item.SUFIJO LIKE '" & MatchMode(txtsufijo, cboCoincidir.Text, apSqlServer) & "' "
'                strOrder = " Order By Stck_Item.SUFIJO"
'            End If
'        End If
'
'
'        If strWhere <> "" Then
'            FillRepuestos strWhere, strOrder
'            strWhere = ""
'            strOrder = ""
'        Else
'            MsgBox "No Hay Criterio Seleccionado, Por Favor Verifique"
'            strWhere = ""
'            strOrder = ""
'            lvwRepuestos.ListItems.Clear
'        End If
    Case "Cerrar"
        Unload Me
    End Select
End Select
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Me.dtcFamilia.BoundText = ""
    Me.dtcSubFamilia.BoundText = ""
    FillSubFamilia dtcFamilia.BoundText
End Sub


Private Sub txtAplicacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtBasico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCodigo_GotFocus()
    MarcaTexto txtCodigo
End Sub
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    Dim Paso As MSComctlLib.Button
    If KeyAscii = 13 Then
        Set Paso = Me.tlbOpciones.Item(1).Buttons("Buscar")
        tlbOpciones_ButtonClick 1, Paso
        MarcaTexto txtCodigo
    End If
End Sub
Private Sub txtDes_KeyPress(KeyAscii As Integer)
Dim Paso As MSComctlLib.Button
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
If KeyAscii = 13 Then
    Set Paso = Me.tlbOpciones.Item(1).Buttons("Buscar")
    tlbOpciones_ButtonClick 1, Paso
End If
End Sub
Sub FillMarcas()
    Me.dtcProcedencia.Enabled = True
    mstrSQL = "Select Id_marca as CODIGO, Descripcion as Nombre from Glbl_Marca where VIGENCIA = 'S' order by Descripcion"
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With Me.datProcedencia
            Set .Recordset = AdoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                Me.dtcProcedencia.ListField = "Nombre"
                Me.dtcProcedencia.BoundColumn = "Codigo"
                'Me.dtcProcedencia.BoundText = .Recordset!codigo
            End If
        End With
    End If ' por el otro
    Set AdoPrincipal = New ADODB.Recordset
    Conexion.CloseHost AdoPrincipal
End Sub

Sub FillModelos(strMarca As String)
    Me.dtcUnidadMedida.Enabled = True
    mstrSQL = "Select Id_modelo as CODIGO, Descripcion as Nombre from Glbl_Modelo where VIGENCIA = 'S' and Id_marca = '" & strMarca & "'  order by Descripcion"
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With Me.datUnidadMedida
            Set .Recordset = AdoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                Me.dtcUnidadMedida.ListField = "Nombre"
                Me.dtcUnidadMedida.BoundColumn = "Codigo"
                Me.dtcUnidadMedida.BoundText = .Recordset!Codigo
            End If
        End With
    End If ' por el otro
    Set AdoPrincipal = New ADODB.Recordset
    Conexion.CloseHost AdoPrincipal
End Sub

Sub FillSubFamilia(strSubFamilia As String)
    'Me.dtcUnidadMedida.Enabled = True
    mstrSQL = "Select Id_SubFamilia as CODIGO, Descripcion as Nombre from Glbl_Subfamilia where VIGENCIA = 'S' and Id_Familia = '" & strSubFamilia & "'  order by Descripcion"
    If Conexion.SendHost(mstrSQL, AdoPrincipal, adOpenKeyset, adLockReadOnly, gcTiempoEspera) = apOk Then
        With Me.datSubFamilia
            Set .Recordset = AdoPrincipal
            If Not .Recordset.BOF And Not .Recordset.EOF Then
                .Recordset.MoveFirst
                Me.dtcSubFamilia.ListField = "Nombre"
                Me.dtcSubFamilia.BoundColumn = "Codigo"
                Me.dtcSubFamilia.BoundText = .Recordset!Codigo
            End If
        End With
    End If ' por el otro
    Set AdoPrincipal = New ADODB.Recordset
    Conexion.CloseHost AdoPrincipal
End Sub

Private Sub txtPrefijo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSufijo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub
