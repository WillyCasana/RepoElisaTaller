VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMantenedorClientes 
   Caption         =   "Mantenedor de Clientes"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   Icon            =   "frmMantenedorClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc datCondicionVenta 
      Height          =   330
      Left            =   2580
      Top             =   3390
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
   Begin MSComCtl2.DTPicker dtpFechaIncorporacion 
      Height          =   315
      Left            =   4740
      TabIndex        =   9
      Top             =   2760
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24707073
      CurrentDate     =   36715
   End
   Begin MSComctlLib.ImageList ImgBarraHerramienta 
      Left            =   6420
      Top             =   -30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":0442
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":0554
            Key             =   "Editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":0666
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":0778
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":088A
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":099C
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":0AAE
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":0BC0
            Key             =   "Cerrar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":0CD2
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":0DE4
            Key             =   "Primero"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":0EF6
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":1008
            Key             =   "Siguiente"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":111A
            Key             =   "Ultimo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":122C
            Key             =   "Renovar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":133E
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":1450
            Key             =   "SortDesc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":1562
            Key             =   "Seleccion"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":19B4
            Key             =   "Seleccion1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":1E06
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":1F18
            Key             =   "AgregarSucursal"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenedorClientes.frx":236C
            Key             =   "VerSucursal"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBarraHerramientas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImgBarraHerramienta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
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
            Object.ToolTipText     =   "Cerrar"
            ImageKey        =   "Cerrar"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AgregarSucursal"
            Object.ToolTipText     =   "Agregar Sucursal"
            ImageKey        =   "AgregarSucursal"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "VerSucursal"
            Object.ToolTipText     =   "Ver Sucursales"
            ImageKey        =   "VerSucursal"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraDatosCliente 
      Height          =   7245
      Left            =   30
      TabIndex        =   28
      Top             =   450
      Width           =   11145
      Begin VB.ComboBox cboClasificacion 
         Height          =   315
         ItemData        =   "frmMantenedorClientes.frx":2AE0
         Left            =   210
         List            =   "frmMantenedorClientes.frx":2AED
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1065
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker dtpFechaNacimiento 
         Height          =   315
         Left            =   2580
         TabIndex        =   8
         Top             =   2310
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   36715
      End
      Begin VB.TextBox txtPaginaWeb 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   225
         MaxLength       =   100
         TabIndex        =   18
         Top             =   4800
         Width           =   10755
      End
      Begin MSAdodcLib.Adodc datTipoCliente 
         Height          =   330
         Left            =   2940
         Top             =   1080
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
      Begin MSAdodcLib.Adodc datComuna 
         Height          =   330
         Left            =   255
         Top             =   2310
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
      Begin MSAdodcLib.Adodc datCiudad 
         Height          =   330
         Left            =   8295
         Top             =   1650
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Caption         =   "Vigente:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9930
         TabIndex        =   48
         Top             =   540
         Width           =   1005
      End
      Begin VB.TextBox txtNombreRazonSocial 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5265
         MaxLength       =   60
         TabIndex        =   3
         Top             =   1050
         Width           =   5715
      End
      Begin VB.TextBox txtDireccion 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   210
         MaxLength       =   60
         TabIndex        =   4
         Top             =   1650
         Width           =   5175
      End
      Begin VB.TextBox txtCodigoPostal 
         Height          =   315
         Left            =   9015
         MaxLength       =   30
         TabIndex        =   26
         Text            =   "12345678901234567890"
         Top             =   2280
         Width           =   1965
      End
      Begin VB.TextBox txtTelefono 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6780
         MaxLength       =   30
         TabIndex        =   10
         Top             =   2280
         Width           =   1950
      End
      Begin VB.TextBox txtFax 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5220
         MaxLength       =   20
         TabIndex        =   15
         Top             =   3540
         Width           =   2805
      End
      Begin VB.TextBox txtCasilla 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8175
         MaxLength       =   20
         TabIndex        =   16
         Top             =   3540
         Width           =   2805
      End
      Begin VB.TextBox txtEMail 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   225
         MaxLength       =   100
         TabIndex        =   17
         Top             =   4170
         Width           =   10755
      End
      Begin VB.TextBox txtActividadEconomica 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   210
         MaxLength       =   50
         TabIndex        =   14
         Top             =   3540
         Width           =   4860
      End
      Begin VB.TextBox txtNombreContacto 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5595
         MaxLength       =   50
         TabIndex        =   13
         Top             =   2940
         Width           =   5385
      End
      Begin VB.TextBox txtComentario 
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   195
         MaxLength       =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   6270
         Width           =   10785
      End
      Begin MSAdodcLib.Adodc datPais 
         Height          =   330
         Left            =   5520
         Top             =   1635
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
      Begin MSDataListLib.DataCombo dbcboComuna 
         Bindings        =   "frmMantenedorClientes.frx":2B18
         DataSource      =   "datComuna"
         Height          =   315
         Left            =   210
         TabIndex        =   7
         Top             =   2310
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "id_Comuna"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcboPais 
         Bindings        =   "frmMantenedorClientes.frx":2B30
         DataSource      =   "datPais"
         Height          =   315
         Left            =   5490
         TabIndex        =   5
         Top             =   1650
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "id_Pais"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcboCiudad 
         Bindings        =   "frmMantenedorClientes.frx":2B46
         DataSource      =   "datCiudad"
         Height          =   315
         Left            =   8265
         TabIndex        =   6
         Top             =   1650
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "id_Ciudad"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcboCondicionVenta 
         Bindings        =   "frmMantenedorClientes.frx":2B5E
         DataSource      =   "datCondicionVenta"
         Height          =   315
         Left            =   2550
         TabIndex        =   12
         Top             =   2940
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "id_Condicion_Venta"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcboTipoCliente 
         Bindings        =   "frmMantenedorClientes.frx":2B7E
         DataSource      =   "datTipoCliente"
         Height          =   315
         Left            =   2190
         TabIndex        =   2
         Top             =   1050
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Descripcion"
         BoundColumn     =   "id_Tipo_Cliente"
         Text            =   ""
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2220
         TabIndex        =   58
         Top             =   270
         Width           =   495
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   210
         TabIndex        =   57
         Top             =   480
         Width           =   1845
         VariousPropertyBits=   746604569
         MaxLength       =   12
         Size            =   "3254;556"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.TextBox txtCuposindocumento 
         Height          =   315
         Left            =   9240
         TabIndex        =   24
         Top             =   5490
         Width           =   1740
         VariousPropertyBits=   746604571
         MaxLength       =   13
         Size            =   "3069;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.TextBox txtSaldosindocumento 
         Height          =   315
         Left            =   7470
         TabIndex        =   23
         Top             =   5490
         Width           =   1740
         VariousPropertyBits=   746604571
         MaxLength       =   13
         Size            =   "3069;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.TextBox txtCreditosindocumento 
         Height          =   315
         Left            =   5670
         TabIndex        =   22
         Top             =   5490
         Width           =   1740
         VariousPropertyBits=   746604571
         MaxLength       =   13
         Size            =   "3069;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.TextBox txtCupocondocumento 
         Height          =   315
         Left            =   3870
         TabIndex        =   21
         Top             =   5490
         Width           =   1740
         VariousPropertyBits=   746604571
         MaxLength       =   13
         Size            =   "3069;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.TextBox txtSaldocondocumento 
         Height          =   315
         Left            =   2070
         TabIndex        =   20
         Top             =   5490
         Width           =   1740
         VariousPropertyBits=   746604571
         MaxLength       =   13
         Size            =   "3069;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.TextBox txtCreditoconDocumento 
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Top             =   5490
         Width           =   1740
         VariousPropertyBits=   746604571
         MaxLength       =   13
         Size            =   "3069;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.TextBox txtDescuento 
         Height          =   315
         Left            =   210
         TabIndex        =   11
         Top             =   2940
         Width           =   2175
         VariousPropertyBits=   746604571
         MaxLength       =   5
         Size            =   "3836;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.TextBox txtRut 
         Height          =   315
         Left            =   2190
         TabIndex        =   0
         Top             =   480
         Width           =   1845
         VariousPropertyBits=   746604571
         MaxLength       =   12
         Size            =   "3254;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.Label lblCuposinDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Cupo sin Documento"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9240
         TabIndex        =   56
         Top             =   5280
         Width           =   1485
      End
      Begin VB.Label lblSaldosinDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Saldo sin Documento"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7470
         TabIndex        =   55
         Top             =   5280
         Width           =   1515
      End
      Begin VB.Label lblCreditosinDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Crédito sin Documento"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5700
         TabIndex        =   54
         Top             =   5280
         Width           =   1605
      End
      Begin VB.Label lblCupoconDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Cupo con Documento"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3900
         TabIndex        =   53
         Top             =   5280
         Width           =   1560
      End
      Begin VB.Label lblSaldoconDocumento 
         AutoSize        =   -1  'True
         Caption         =   "Saldo con Documento"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2100
         TabIndex        =   52
         Top             =   5280
         Width           =   1590
      End
      Begin VB.Label lblCredConDoc 
         AutoSize        =   -1  'True
         Caption         =   "Crédito con Documento"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   51
         Top             =   5280
         Width           =   1680
      End
      Begin VB.Label lblClasificacion 
         AutoSize        =   -1  'True
         Caption         =   "Clasificación"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   50
         Top             =   840
         Width           =   885
      End
      Begin VB.Label lblPaginaWeb 
         AutoSize        =   -1  'True
         Caption         =   "Página Web"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   210
         TabIndex        =   49
         Top             =   4590
         Width           =   885
      End
      Begin VB.Label lblRut 
         AutoSize        =   -1  'True
         Caption         =   "Rut del Cliente"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   47
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label lblNombreRazonSocial 
         AutoSize        =   -1  'True
         Caption         =   "Nombre o Razón Social"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5280
         TabIndex        =   46
         Top             =   840
         Width           =   1680
      End
      Begin VB.Label lblDireccion 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   210
         TabIndex        =   45
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label lblPais 
         AutoSize        =   -1  'True
         Caption         =   "País"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5520
         TabIndex        =   44
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label lblCiudad 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8280
         TabIndex        =   43
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblComuna 
         AutoSize        =   -1  'True
         Caption         =   "Comuna"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   210
         TabIndex        =   42
         Top             =   2070
         Width           =   585
      End
      Begin VB.Label lblTipoCliente 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cliente"
         Height          =   195
         Left            =   2220
         TabIndex        =   41
         Top             =   840
         Width           =   840
      End
      Begin VB.Label lblCodigoPostal 
         AutoSize        =   -1  'True
         Caption         =   "Código Postal"
         Height          =   195
         Left            =   9030
         TabIndex        =   40
         Top             =   2070
         Width           =   975
      End
      Begin VB.Label lblTelefono 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono"
         Height          =   195
         Left            =   6780
         TabIndex        =   39
         Top             =   2070
         Width           =   630
      End
      Begin VB.Label lblFechaNacimiento 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Nacimiento "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2610
         TabIndex        =   38
         Top             =   2070
         Width           =   1560
      End
      Begin VB.Label lblFechaIncorporacion 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Incorporación"
         Height          =   195
         Left            =   4710
         TabIndex        =   37
         Top             =   2070
         Width           =   1695
      End
      Begin VB.Label lblDescuento 
         AutoSize        =   -1  'True
         Caption         =   "% Descuento Asociado"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   2730
         Width           =   1650
      End
      Begin VB.Label lblCondicionVenta 
         AutoSize        =   -1  'True
         Caption         =   "Condición de Venta"
         Height          =   195
         Left            =   2550
         TabIndex        =   35
         Top             =   2730
         Width           =   1395
      End
      Begin VB.Label lblFax 
         AutoSize        =   -1  'True
         Caption         =   "Fax / Telex"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5250
         TabIndex        =   34
         Top             =   3330
         Width           =   810
      End
      Begin VB.Label lblCasilla 
         AutoSize        =   -1  'True
         Caption         =   "Casilla"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8190
         TabIndex        =   33
         Top             =   3330
         Width           =   450
      End
      Begin VB.Label lblEMail 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail (Internet)"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   210
         TabIndex        =   32
         Top             =   3960
         Width           =   1110
      End
      Begin VB.Label lblGiro 
         AutoSize        =   -1  'True
         Caption         =   "Giro de  la Empresa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   3330
         Width           =   1380
      End
      Begin VB.Label lblNombreContacto 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de Contacto "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5610
         TabIndex        =   30
         Top             =   2730
         Width           =   1515
      End
      Begin VB.Label lblComentario 
         AutoSize        =   -1  'True
         Caption         =   "Comentarios"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   6030
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmMantenedorClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoPrincipal As New ADODB.Recordset

Dim mstrSql As String
Dim mblnTablaVacia As Boolean
Dim mblnSw As Boolean
Dim mblnAccesoCrear As Boolean
Dim mblnAccesoEditar As Boolean
Dim mblnAccesoBorrar As Boolean
Dim mblnAccesoImprimir As Boolean

Sub LimpiaHabilita()
    txtRut.Text = ""
    dbcboTipoCliente.Text = ""
    txtNombreRazonSocial.Text = ""
    txtDireccion.Text = ""
    dbcboPais.Text = ""
    dbcboCiudad.Text = ""
    dbcboComuna.Text = ""
    txtTelefono.Text = ""
    txtCodigoPostal.Text = ""
    txtDescuento.Text = "0"
    txtDescuento.Text = FormatoValor(txtDescuento, "%", 1)
    dbcboCondicionVenta.Text = ""
    txtNombreContacto.Text = ""
    txtActividadEconomica.Text = ""
    txtFax.Text = ""
    txtCasilla.Text = ""
    txtEMail.Text = ""
    txtPaginaWeb.Text = ""
    txtComentario.Text = ""
    txtCreditoconDocumento = "0"
    txtCreditoconDocumento.Text = FormatoValor(txtCreditoconDocumento, "$", 0)
    txtSaldocondocumento.Text = "0"
    txtSaldocondocumento.Text = FormatoValor(txtSaldocondocumento, "$", 0)
    txtCupocondocumento.Text = "0"
    txtCupocondocumento.Text = FormatoValor(txtCupocondocumento, "$", 0)
    txtCreditosindocumento.Text = "0"
    txtCreditosindocumento.Text = FormatoValor(txtCreditosindocumento, "$", 0)
    txtSaldosindocumento.Text = "0"
    txtSaldosindocumento.Text = FormatoValor(txtSaldosindocumento, "$", 0)
    txtCuposindocumento.Text = "0"
    txtCuposindocumento.Text = FormatoValor(txtCuposindocumento, "$", 0)
    txtCodigo.Text = ""
    cboClasificacion.Text = "Cliente"
End Sub
Sub Deshabilita()
    txtRut.Enabled = False
    dbcboTipoCliente.Enabled = False
    txtNombreRazonSocial.Enabled = False
    txtDireccion.Enabled = False
    dbcboPais.Enabled = False
    dbcboCiudad.Enabled = False
    dbcboComuna.Enabled = False
    mskFechaNacimiento.Enabled = False
    mskFechaIncorporacion.Enabled = False
    txtTelefono.Enabled = False
    txtCodigoPostal.Enabled = False
    txtDescuento.Enabled = False
    dbcboCondicionVenta.Enabled = False
    txtNombreContacto.Enabled = False
    txtActividadEconomica.Enabled = False
    txtFax.Enabled = False
    txtCasilla.Enabled = False
    txtEMail.Enabled = False
    txtPaginaWeb.Enabled = False
    txtComentario.Enabled = False
End Sub
Sub FillComuna(strPais As String, strCiudad As String)
    mstrSql = "SELECT * FROM glbl_Comuna WHERE Id_Ciudad = '" & strCiudad & "' and id_Pais = '" & strPais & "' and Vigencia = 'S' ORDER BY Descripcion "
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, 10) = apOk Then
        Set datComuna.Recordset = adoPrincipal
    End If
End Sub

Sub FillCiudad(strPais As String)
    mstrSql = "SELECT * FROM glbl_Ciudad WHERE id_Pais = '" & strPais & "' and Vigencia = 'S' ORDER BY Descripcion "
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockReadOnly, 10) = apOk Then
        Set datCiudad.Recordset = adoPrincipal
    End If
End Sub


Private Sub dbcboCiudad_Change()
If dbcboCiudad.BoundText <> "" Then
    dbcboComuna.Text = ""
    FillComuna dbcboPais.BoundText, dbcboCiudad.BoundText
End If
End Sub

Private Sub dbcboPais_Change()

If dbcboPais.BoundText <> "" Then
    dbcboCiudad.Text = ""
    FillCiudad dbcboPais.BoundText
End If

End Sub

Private Sub Form_Load()
    mblnSw = True
'    Dim tbRegistros As ADODB.Recordset
'    Dim strConnect As String
'    Dim lstrQuery As String
'    Dim llngRetorno As Long
'
'    pstrRut = Empty
'    strConnect = "DRIVER={SQL Server};SERVER=SERVIDOR_NT;UID=sa;PWD=;DATABASE=AUTOPRO"
'
'    Set Conexion = New APCONADO.ConnectionAdo
'
'    llngRetorno = Conexion.ConnectHost(Cn2, adUseServer, strConnect, 10)
'
'    '// Pais
'    Set tbRegistros = New ADODB.Recordset
'
'    lstrQuery = "SELECT * FROM glbl_Pais WHERE Vigencia = 'S' ORDER BY Descripcion"
'
'    If Conexion.SendHost(lstrQuery, tbRegistros, adOpenKeyset, adLockReadOnly, 10) = apOk Then
'        Set datPais.Recordset = tbRegistros
'    End If
'
'    '//CondicionVenta
'    Set tbRegistros = New ADODB.Recordset
'
'    lstrQuery = "SELECT * FROM glbl_Condicion_Venta WHERE Vigencia = 'S' ORDER BY Descripcion"
'
'    If Conexion.SendHost(lstrQuery, tbRegistros, adOpenKeyset, adLockReadOnly, 10) = apOk Then
'        Set datCondicionVenta.Recordset = tbRegistros
'    End If
'
'
'    '// TipoCliente
'    Set tbRegistros = New ADODB.Recordset
'
'    lstrQuery = "SELECT * FROM glbl_Tipo_Cliente WHERE Vigencia = 'S' ORDER BY Descripcion"
'
'    If Conexion.SendHost(lstrQuery, tbRegistros, adOpenKeyset, adLockReadOnly, 10) = apOk Then
'        Set datTipoCliente.Recordset = tbRegistros
'    End If
    
    LimpiaHabilita
    RevizaAtributos
    Renovar
    Screen.MousePointer = vbDefault
End Sub

Private Sub tlbBarraHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lintRespuesta As Integer
    
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
            Screen.MousePointer = vbDefault
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
        Case "AgregarSucursal"
            AgregarSucursal
        Case "VerSucursal"
            Screen.MousePointer = vbDefault
            frmSucursales.Show 1
        Case "Cerrar"
            CerrarSalir
    End Select
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Activate()
    If mblnSw Then
    
        Renovar
        
    End If
    
'    If TablaDatosExport.Rut <> "" Then
'        MuestraDatos (TablaDatosExport.Rut)
'        cboClasificacion.SetFocus
'        Screen.MousePointer = vbDefault
'    End If
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
        Case 3 And tlbBarraHerramientas.Buttons.Item("Cerrar").Enabled
            KeyAscii = 0
            CerrarSalir
    End Select
End Sub
Private Sub AgregarRegistro()
    Me.Tag = "Crear"
    DesactivaBotones
    LimpiaHabilita
    ValoresporDefecto
    txtCodigo.SetFocus
End Sub
Private Sub CancelarAgregaRegistro()
    Me.Tag = ""
    ActivaBotones
    
    mstrSql = "SELECT TOP 1 * FROM Glbl_Cliente_Proveedor WHERE Id_Cliente_Proveedor >'" & Trim$(txtCodigo.Text) & "' ORDER BY Id_Cliente_Proveedor"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            mstrSql = "SELECT TOP 1 * FROM Glbl_Cliente_Proveedor WHERE Glbl_Cliente_Proveedor < '" & Trim$(txtCodigo.Text) & "' ORDER BY Id_Cliente_Proveedor"
            If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, 10) = apOk Then
                If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                    LeerCampos
                Else
                    mblnTablaVacia = True
                    LimpiaHabilita
                End If
            End If
        End If
    End If
    Conexion.CloseHost adoPrincipal
    cboClasificacion.SetFocus
End Sub
Private Sub GrabarRegistro()
    Dim lstrClasificacion As String * 1
    Dim lstrCodigoSucursal As String
    If Not Validacion() Then
        Exit Sub
    End If
        
    If cboClasificacion.Text = "Cliente" Then
        lstrClasificacion = "C"
    Else
        If cboClasificacion.Text = "Proveedor" Then
            lstrClasificacion = "P"
        Else
            lstrClasificacion = "A"
        End If
    End If
    If txtDescuento.Text = "" Then
        txtDescuento.Text = "0"
    End If
    
    If txtCreditoconDocumento.Text = "" Then
        txtCreditoconDocumento.Text = "0"
    End If
    If txtSaldocondocumento.Text = "" Then
        txtSaldocondocumento.Text = "0"
    End If
    If txtCupocondocumento.Text = "" Then
        txtCupocondocumento.Text = "0"
    End If
    If txtCreditosindocumento.Text = "" Then
        txtCreditosindocumento.Text = "0"
    End If
    If txtSaldosindocumento.Text = "" Then
        txtSaldosindocumento.Text = "0"
    End If
    If txtCuposindocumento.Text = "" Then
        txtCuposindocumento.Text = "0"
    End If
    If Me.Tag = "Crear" Then
        mstrSql = "INSERT INTO Glbl_Cliente_Proveedor " & " (id_Cliente_Proveedor, id_Condicion_Venta, id_Tipo_Cliente,id_Comuna,id_Ciudad,id_Pais,Razon_Social,Direccion,Fecha_Nacimiento,Fecha_Incorporacion,Telefono,Fax,NombreContacto,Casilla,PaginaWeb,Descuento_Asociado,Comentario,CodigoPostal,E_Mail,Cliente_Proveedor,Usr_Id,Usr_Fecha,Vigencia,Giro_Comercial,Credito_con_documento,Saldo_con_documento,Cupo_con_documento,Credito_sin_documento,Saldo_sin_documento,Cupo_sin_documento,rut )" _
                & "VALUES ('" & Trim(UCase(txtCodigo)) & "', '" & Trim(dbcboCondicionVenta.BoundText) & "','" & Trim(dbcboTipoCliente.BoundText) & "', '" & Trim(dbcboComuna.BoundText) & "','" & Trim(dbcboCiudad.BoundText) & "','" & Trim(dbcboPais.BoundText) & "','" & Trim$(UCase(txtNombreRazonSocial.Text)) & "','" & Trim(UCase(txtDireccion.Text)) & "','" & Trim(dtpFechaNacimiento.Value) & "','" & Trim(dtpFechaIncorporacion.Value) & "','" & Trim(txtTelefono.Text) & "','" & Trim(txtFax.Text) & "','" & Trim$(UCase(txtNombreContacto.Text)) & "','" & Trim(txtCasilla.Text) & "','" & Trim(txtPaginaWeb.Text) & "'," & CDbl(SacarFormatoValor(txtDescuento.Text, "%")) & ",'" & Trim$(UCase(txtComentario.Text)) & "','" & Trim(txtCodigoPostal.Text) & "','" & Trim(txtEMail.Text) & "','" & lstrClasificacion & "','" & "ARIEL" & "', '" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "','" & IIf(chkVigencia.Value = vbChecked, "S", "N") & "', '" & Trim$(UCase(txtActividadEconomica.Text)) & "'" _
                & "," & CDbl(SacarFormatoValor(txtCreditoconDocumento.Text, "$")) & "," & CDbl(SacarFormatoValor(txtSaldocondocumento.Text, "$")) & "," & CDbl(SacarFormatoValor(txtCupocondocumento.Text, "$")) & "," & CDbl(SacarFormatoValor(txtCreditosindocumento.Text, "$")) & "," & CDbl(SacarFormatoValor(txtSaldosindocumento.Text, "$")) & "," & CDbl(SacarFormatoValor(txtCuposindocumento.Text, "$")) & ",'" & Trim$(UCase(txtCodigo.Text)) & "')"
    Else
        If Me.Tag = "AgregarSucursal" Then
            lstrCodigoSucursal = TraeCodigo(txtCodigo.Text)
            mstrSql = "INSERT INTO Glbl_Cliente_Proveedor " & " (id_Cliente_Proveedor, id_Condicion_Venta, id_Tipo_Cliente,id_Comuna,id_Ciudad,id_Pais,Razon_Social,Direccion,Fecha_Nacimiento,Fecha_Incorporacion,Telefono,Fax,NombreContacto,Casilla,PaginaWeb,Descuento_Asociado,Comentario,CodigoPostal,E_Mail,Cliente_Proveedor,Usr_Id,Usr_Fecha,Vigencia,Giro_Comercial,Credito_con_documento,Saldo_con_documento,Cupo_con_documento,Credito_sin_documento,Saldo_sin_documento,Cupo_sin_documento,rut)" _
                & "VALUES ('" & Trim(UCase(lstrCodigoSucursal)) & "', '" & Trim(dbcboCondicionVenta.BoundText) & "','" & Trim(dbcboTipoCliente.BoundText) & "', '" & Trim(dbcboComuna.BoundText) & "','" & Trim(dbcboCiudad.BoundText) & "','" & Trim(dbcboPais.BoundText) & "','" & Trim$(UCase(txtNombreRazonSocial.Text)) & "','" & Trim(UCase(txtDireccion.Text)) & "','" & Trim(dtpFechaNacimiento.Value) & "','" & Trim(dtpFechaIncorporacion.Value) & "','" & Trim(txtTelefono.Text) & "','" & Trim(txtFax.Text) & "','" & Trim$(UCase(txtNombreContacto.Text)) & "','" & Trim(txtCasilla.Text) & "','" & Trim(txtPaginaWeb.Text) & "'," & CDbl(SacarFormatoValor(txtDescuento.Text, "%")) & ",'" & Trim$(UCase(txtComentario.Text)) & "','" & Trim(txtCodigoPostal.Text) & "','" & Trim(txtEMail.Text) & "','" & lstrClasificacion & "','" & "ARIEL" & "', '" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "','" & IIf(chkVigencia.Value = vbChecked, "S", "N") & "', '" & Trim$(UCase(txtActividadEconomica.Text)) & "'" _
                & "," & CDbl(SacarFormatoValor(txtCreditoconDocumento.Text, "$")) & "," & CDbl(SacarFormatoValor(txtSaldocondocumento.Text, "$")) & "," & CDbl(SacarFormatoValor(txtCupocondocumento.Text, "$")) & "," & CDbl(SacarFormatoValor(txtCreditosindocumento.Text, "$")) & "," & CDbl(SacarFormatoValor(txtSaldosindocumento.Text, "$")) & "," & CDbl(SacarFormatoValor(txtCuposindocumento.Text, "$")) & ",'" & Trim(UCase(txtCodigo)) & "')"
        
        Else
            mstrSql = "UPDATE Glbl_Cliente_Proveedor SET Id_Condicion_Venta  ='" & Trim(dbcboCondicionVenta.BoundText) & "', vigencia='" & IIf(chkVigencia.Value = vbChecked, "S", "N") & "',Id_Tipo_Cliente ='" & Trim$(dbcboTipoCliente.BoundText) & "',Id_Comuna ='" & Trim$(dbcboComuna.BoundText) & "',Id_Ciudad ='" & Trim$(dbcboCiudad.BoundText) & "',Id_Pais ='" & Trim$(dbcboPais.BoundText) & "',Razon_Social ='" & Trim$(UCase(txtNombreRazonSocial.Text)) & "',Direccion ='" & Trim$(UCase(txtDireccion.Text)) & "',Fecha_Nacimiento ='" & Trim$(dtpFechaNacimiento.Value) & "',Fecha_Incorporacion ='" & Trim$(dtpFechaIncorporacion.Value) & "',Telefono ='" & Trim$(txtTelefono.Text) & "',Fax ='" & Trim$(txtFax.Text) & "',NombreContacto ='" & Trim$(UCase(txtNombreContacto.Text)) & "',Casilla ='" & Trim$(txtCasilla.Text) & "',PaginaWeb ='" & Trim$(txtPaginaWeb.Text) & "',Descuento_Asociado =" & CDbl(SacarFormatoValor(txtDescuento.Text, "%")) & ",Comentario ='" & Trim$(UCase(txtComentario.Text)) & "'," _
                    & " E_Mail ='" & Trim$(txtEMail.Text) & "',Cliente_Proveedor ='" & lstrClasificacion & "',usr_id='ariel', usr_fecha='" & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "HH:MM:SS") & "',Credito_con_documento =" & CDbl(SacarFormatoValor(txtCreditoconDocumento.Text, "$")) & ",Saldo_con_documento =" & CDbl(SacarFormatoValor(txtSaldocondocumento.Text, "$")) & ",Cupo_con_documento =" & CDbl(SacarFormatoValor(txtCupocondocumento.Text, "$")) & ",Credito_sin_documento =" & CDbl(SacarFormatoValor(txtCreditosindocumento.Text, "$")) & ",Saldo_sin_documento =" & CDbl(SacarFormatoValor(txtSaldosindocumento.Text, "$")) & ",Cupo_sin_documento =" & CDbl(SacarFormatoValor(txtCuposindocumento.Text, "$")) & ",rut =" & Trim$(txtCodigo.Text) & "" _
                    & " WHERE Id_Cliente_Proveedor ='" & Trim(UCase(txtCodigo.Text)) & "' and rut ='" & Trim$(txtCodigo.Text) & "'"
        End If
    End If
    If Conexion.SendHost(mstrSql, , , , 10) = apOk Then
        mblnTablaVacia = False
        ActivaBotones
    End If
    Me.Tag = ""
End Sub
Private Sub BorrarRegistro()
    
    Screen.MousePointer = vbDefault
    If MsgBox("¿ Desea eliminar este registro ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar") = vbYes Then
        mstrSql = "DELETE FROM Glbl_Cliente_Proveedor WHERE  Id_Cliente_Proveedor = '" & Trim$(txtCodigo.Text) & "'"
        If Conexion.SendHost(mstrSql, , , , 10) = apOk Then
            mstrSql = "SELECT TOP 1 * Glbl_Cliente_Proveedor WHERE Id_Cliente_Proveedor > '" & Trim$(txtCodigo.Text) & "' ORDER BY Id_Cliente_Proveedor"
            If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, 10) = apOk Then
                If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                    LeerCampos
                Else
                    mstrSql = "SELECT TOP 1 * FROM Glbl_Cliente_Proveedor WHERE Id_Cliente_Proveedor < '" & Trim$(txtCodigo.Text) & "' ORDER BY Id_Cliente_Proveedor"
                    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, 10) = apOk Then
                        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
                            LeerCampos
                        Else
                            mblnTablaVacia = True
                            LimpiaHabilita
                        End If
                    End If
                End If
            End If
        End If
        Conexion.CloseHost adoPrincipal
    End If
End Sub

Private Sub BuscarRegistro()
    Load frmBuscarCliente
    frmBuscarCliente.Show 1
End Sub
Private Sub ImprimirInforme()
    frmInforme.Show 1
End Sub
Private Sub PrimerRegistro()
    mstrSql = "SELECT TOP 1 * FROM  Glbl_Cliente_Proveedor ORDER BY Id_Cliente_Proveedor ASC"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost adoPrincipal
End Sub
Private Sub RegistroAnterior()
    
    mstrSql = "SELECT TOP 1 * FROM Glbl_Cliente_Proveedor WHERE Id_Cliente_Proveedor < '" & Trim$(txtRut.Text) & "' ORDER BY Id_Cliente_Proveedor DESC"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost adoPrincipal
End Sub
Private Sub RegistroSiguiente()

    mstrSql = "SELECT TOP 1 * FROM Glbl_Cliente_Proveedor WHERE Id_Cliente_Proveedor  > '" & Trim$(txtRut.Text) & "' ORDER BY Id_Cliente_Proveedor"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost adoPrincipal
End Sub
Private Sub UltimoRegistro()
    mstrSql = "SELECT TOP 1 * FROM Glbl_Cliente_Proveedor ORDER BY Id_Cliente_Proveedor DESC"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If Not adoPrincipal.BOF And Not adoPrincipal.EOF Then
            LeerCampos
        Else
            Beep
        End If
    End If
    Conexion.CloseHost adoPrincipal
End Sub
Private Sub Renovar()
'    Set adoPrincipal = New ADODB.Recordset
    mstrSql = "SELECT TOP 1 * FROM Glbl_Cliente_Proveedor ORDER BY Id_Cliente_Proveedor ASC"
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, 10) = apOk Then
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
    txtRut.Enabled = False
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
        .Item("AgregarSucursal").Enabled = True
        .Item("VerSucursal").Enabled = True
    End With
End Sub
Private Sub DesactivaBotones()
    txtCodigo.Enabled = True
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
        .Item("AgregarSucursal").Enabled = False
        .Item("VerSucursal").Enabled = False
    End With
End Sub
Private Sub VerificaTablaVacia()
    If (Not adoPrincipal.BOF And Not adoPrincipal.EOF) And adoPrincipal.RecordCount > 0 Then
        mblnTablaVacia = False
    Else
        mblnTablaVacia = True
        LimpiaHabilita
        MsgBox "La tabla no contiene registros...", vbInformation, "Advertencia"
    End If
End Sub
Private Sub LeerCampos()
    
    If mblnTablaVacia Then
        LimpiaHabilita
        Exit Sub
    End If

    With adoPrincipal
        txtRut.Text = ValorNulo(.Fields("id_Cliente_Proveedor"))
        If IsNull(!vigencia) Then
            chkVigencia.Value = vbUnchecked
        Else
            If !vigencia = "S" Then
                chkVigencia.Value = vbChecked
            Else
                chkVigencia.Value = vbUnchecked
            End If
        End If
        
        
        dbcboTipoCliente.BoundText = ValorNulo(.Fields("id_Tipo_Cliente"))
        txtNombreRazonSocial.Text = ValorNulo(.Fields("Razon_Social"))
        txtDireccion.Text = ValorNulo(.Fields("Direccion"))
        dbcboPais.BoundText = ValorNulo(.Fields("id_Pais"))
        dbcboCiudad.BoundText = ValorNulo(.Fields("id_Ciudad"))
        dbcboComuna.BoundText = ValorNulo(.Fields("id_Comuna"))
        dtpFechaNacimiento.Value = ValorNulo(.Fields("Fecha_Nacimiento"))
        dtpFechaIncorporacion.Value = ValorNulo(.Fields("Fecha_Incorporacion"))
        txtTelefono.Text = ValorNulo(.Fields("Telefono"))
        txtCodigoPostal.Text = ValorNulo(.Fields("CodigoPostal"))
        txtDescuento.Text = FormatoValor(ValorNulo(.Fields("Descuento_Asociado")), "%", 1)
        dbcboCondicionVenta.BoundText = ValorNulo(.Fields("id_Condicion_Venta"))
        txtNombreContacto.Text = ValorNulo(.Fields("NombreContacto"))
        txtActividadEconomica.Text = ValorNulo(.Fields("Giro_Comercial"))
        txtFax.Text = ValorNulo(.Fields("Fax"))
        txtCasilla.Text = ValorNulo(.Fields("Casilla"))
        txtEMail.Text = ValorNulo(.Fields("E_Mail"))
        txtPaginaWeb.Text = ValorNulo(.Fields("PaginaWeb"))
        txtComentario.Text = ValorNulo(.Fields("Comentario"))
        If ValorNulo(.Fields("Cliente_Proveedor")) = "C" Then
            cboClasificacion.Text = "Cliente"
        Else
            If ValorNulo(.Fields("Cliente_Proveedor")) = "P" Then
                cboClasificacion.Text = "Proveedor"
            Else
                cboClasificacion.Text = "Cliente-Proveedor"
            End If
        End If
        txtCreditoconDocumento.Text = FormatoValor(ValorNulo(.Fields("Credito_con_documento")), "$", 0)
        txtSaldocondocumento.Text = FormatoValor(ValorNulo(.Fields("Saldo_con_documento")), "$", 0)
        txtCupocondocumento.Text = FormatoValor(ValorNulo(.Fields("Cupo_con_documento")), "$", 0)
        txtCreditosindocumento.Text = FormatoValor(ValorNulo(.Fields("Credito_sin_documento")), "$", 0)
        txtSaldosindocumento.Text = FormatoValor(ValorNulo(.Fields("Saldo_sin_documento")), "$", 0)
        txtCuposindocumento.Text = FormatoValor(ValorNulo(.Fields("Cupo_sin_documento")), "$", 0)
        txtCodigo = ValorNulo(.Fields("rut"))
    End With
End Sub

Private Sub ValoresporDefecto()
    With adoPrincipal
        chkVigencia.Value = vbChecked
    End With
End Sub
Private Function Validacion() As Boolean
    Validacion = True
    If txtCodigo.Text = "" Then
        MsgBox "Debe Ingresar Rut...", vbInformation, "Advertencia"
        txtCodigo.SetFocus
        Validacion = False
        Exit Function
    End If
    If cboClasificacion.Text = "" Then
        MsgBox "La Clasificación debe contener un valor...", vbInformation, "Advertencia"
        cboClasificacion.SetFocus
        Validacion = False
        Exit Function
    End If
    If dbcboTipoCliente.Text = "" Then
        MsgBox "Debe Seleccionar Tipo de Cliente...", vbInformation, "Advertencia"
        dbcboTipoCliente.SetFocus
        Validacion = False
        Exit Function
    End If
    
    If txtNombreRazonSocial.Text = "" Then
        MsgBox "Debe Ingresar Nombre de Cliente...", vbInformation, "Advertencia"
        txtNombreRazonSocial.SetFocus
        Validacion = False
        Exit Function
    End If
    If txtDireccion.Text = "" Then
        MsgBox "Debe Ingresar Dirección...", vbInformation, "Advertencia"
        txtDireccion.SetFocus
        Validacion = False
        Exit Function
    End If
    If dbcboPais.Text = "" Then
        MsgBox "Debe Seleccionar Pais...", vbInformation, "Advertencia"
        dbcboPais.SetFocus
        Validacion = False
        Exit Function
    End If
    If dbcboCiudad.Text = "" Then
        MsgBox "Debe Seleccionar Ciudad...", vbInformation, "Advertencia"
        dbcboCiudad.SetFocus
        Validacion = False
        Exit Function
    End If
    If dbcboComuna.Text = "" Then
        MsgBox "Debe Seleccionar Comuna...", vbInformation, "Advertencia"
        dbcboComuna.SetFocus
        Validacion = False
        Exit Function
    End If
    If dbcboCondicionVenta.Text = "" Then
        MsgBox "Debe Seleccionar Condición de Venta...", vbInformation, "Advertencia"
        dbcboCondicionVenta.SetFocus
        Validacion = False
        Exit Function
    End If
    
    
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Set frmMarcaVehiculo = Nothing
'    gstrBusca = txtCodigo.Text
    
    Conexion.DisconnectHost
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


Private Sub txtActividadEconomica_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCasilla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCodigo_LostFocus()
    Dim adoTemp As ADODB.Recordset
    Dim lintValor As Integer
    Dim i As Integer
    Dim lstrCaracter As String
    
    If ValidaRut(txtCodigo.Text) Then
        '//Verifica si existe un registro...
        If Me.Tag = "Crear" Then
            mstrSql = "SELECT  Id_Cliente_Proveedor, Razon_Social FROM Glbl_Cliente_Proveedor WHERE  Id_Cliente_Proveedor ='" & Trim$(txtCodigo.Text) & "'"
            If Conexion.SendHost(mstrSql, adoTemp, adOpenKeyset, adLockOptimistic, 10) = apOk Then
                If Not adoTemp.BOF And Not adoTemp.EOF Then
                    MsgBox "Este Rut ya esta registrado con la descripción " & Chr(13) & "[" & IIf(IsNull(adoTemp.Fields("Razon_Social")), "SIN DESCRIPCION", adoTemp.Fields("Razon_Social")) & "]", vbInformation, "Advertencia"
                    txtCodigo.Text = ""
                    txtCodigo.SetFocus
                Else
                    lintValor = InStr(1, txtCodigo, ".")
                    If lintValor > 0 Then
                        For i = 1 To Len(txtCodigo.Text)
                            If Not (Mid$(txtCodigo, i, 1) = ".") Then
                                lstrCaracter = lstrCaracter & Mid$(txtCodigo, i, 1)
                            End If
                        Next i
                        txtCodigo.Text = lstrCaracter
                    End If
                    
                    lstrCaracter = ""
                    lintValor = InStr(1, txtCodigo.Text, "-")
                    If lintValor > 0 Then
                        For i = 1 To Len(txtCodigo.Text)
                            If Not (Mid$(txtCodigo, i, 1) = "-") Then
                                lstrCaracter = lstrCaracter & Mid$(txtCodigo, i, 1)
                            End If
                        Next i
                        txtCodigo.Text = lstrCaracter
                        txtRut.Text = lstrCaracter
                    End If
                    txtRut = txtCodigo
                End If
            End If
            Conexion.CloseHost adoTemp
        End If
    Else
        MsgBox "Rut Inválido", vbOKOnly + vbInformation, "Valida Rut"
        txtCodigo.SetFocus
    End If
End Sub

Private Sub txtCodigoPostal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCreditoconDocumento_GotFocus()
    txtCreditoconDocumento = SacarFormatoValor(txtCreditoconDocumento, "$")
End Sub

Private Sub txtCreditoconDocumento_KeyPress(KeyAscii As VB.ReturnInteger)
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCreditoconDocumento_LostFocus()
    If Not IsNumeric(txtCreditoconDocumento.Text) Then
        MsgBox "Debe Ingresar un Valor Numérico...", vbOKOnly + vbInformation, "Información"
        txtCreditoconDocumento.Text = "0"
        txtCreditoconDocumento.SetFocus
    Else
        txtCreditoconDocumento = FormatoValor(txtCreditoconDocumento, "$", 0)
    End If
End Sub

Private Sub txtCreditosindocumento_GotFocus()
    txtCreditosindocumento = SacarFormatoValor(txtCreditosindocumento, "$")
End Sub

Private Sub txtCreditosindocumento_KeyPress(KeyAscii As VB.ReturnInteger)
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCreditosindocumento_LostFocus()
    If Not IsNumeric(txtCreditosindocumento.Text) Then
        MsgBox "Debe Ingresar un Valor Numérico...", vbOKOnly + vbInformation, "Información"
        txtCreditosindocumento.Text = "0"
        txtCreditosindocumento.SetFocus
    Else
        txtCreditosindocumento = FormatoValor(txtCreditosindocumento, "$", 0)
    End If
End Sub

Private Sub txtCupocondocumento_GotFocus()
    txtCupocondocumento = SacarFormatoValor(txtCupocondocumento, "$")
End Sub

Private Sub txtCupocondocumento_KeyPress(KeyAscii As VB.ReturnInteger)
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCupocondocumento_LostFocus()
    If Not IsNumeric(txtCupocondocumento.Text) Then
        MsgBox "Debe Ingresar un Valor Numérico...", vbOKOnly + vbInformation, "Información"
        txtCupocondocumento.Text = "0"
        txtCupocondocumento.SetFocus
    Else
        txtCupocondocumento = FormatoValor(txtCupocondocumento, "$", 0)
    End If
End Sub


Private Sub txtCuposindocumento_GotFocus()
    txtCuposindocumento = SacarFormatoValor(txtCuposindocumento, "$")
End Sub

Private Sub txtCuposindocumento_KeyPress(KeyAscii As VB.ReturnInteger)
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCuposindocumento_LostFocus()
    If Not IsNumeric(txtCuposindocumento.Text) Then
        MsgBox "Debe Ingresar un Valor Numérico...", vbOKOnly + vbInformation, "Información"
        txtCuposindocumento.Text = "0"
        txtCuposindocumento.SetFocus
    Else
        txtCuposindocumento = FormatoValor(txtCuposindocumento, "$", 0)
    End If
End Sub

Private Sub txtDescuento_GotFocus()
    txtDescuento = SacarFormatoValor(txtDescuento, "%")
End Sub

Private Sub txtDescuento_KeyPress(KeyAscii As VB.ReturnInteger)
    
    If Not IsNumeric(Chr(KeyAscii)) Then
        If KeyAscii <> 46 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtDescuento_LostFocus()
    If Not IsNumeric(txtDescuento.Text) Then
        MsgBox "Debe Ingresar un Valor Numérico...", vbOKOnly + vbInformation, "Información"
        txtDescuento.Text = "0"
        txtDescuento.SetFocus
    Else
        txtDescuento = FormatoValor(txtDescuento, "%", 1)
    End If
End Sub

Private Sub MuestraDatos(strRut As String)
    Set adoPrincipal = New ADODB.Recordset
    mstrSql = "SELECT * FROM Glbl_Cliente_Proveedor WHERE Id_Cliente_Proveedor = '" & strRut & "'"
    
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        VerificaTablaVacia
        ActivaBotones
        If Not mblnTablaVacia Then
            LeerCampos
        End If
    End If
    Conexion.CloseHost adoPrincipal
End Sub

Private Sub AgregarSucursal()
    Dim lintRespuesta As Integer
    lintRespuesta = MsgBox("Esta Seguro de Agregar Sucursal ?", vbYesNo + vbQuestion, "Confirmación")
    If lintRespuesta = vbYes Then
        Me.Tag = "AgregarSucursal"
        txtDireccion.Text = ""
        txtDireccion.SetFocus
        dbcboPais.Text = ""
        dbcboCiudad.Text = ""
        dbcboComuna.Text = ""
    End If
End Sub

Private Function TraeCodigo(strRut As String) As String
    Dim lintindice As Integer
    
    Set adoPrincipal = New ADODB.Recordset
    mstrSql = "SELECT * FROM Glbl_Cliente_Proveedor WHERE Rut = '" & strRut & "' ORDER BY Id_Cliente_Proveedor DESC"
    
    If Conexion.SendHost(mstrSql, adoPrincipal, adOpenKeyset, adLockOptimistic, 10) = apOk Then
        If adoPrincipal.RecordCount > 0 Then
            If Mid$(Right(adoPrincipal.Fields("Id_Cliente_Proveedor"), 2), 1, 1) = "-" Then
                lintindice = Val(Right(adoPrincipal.Fields("Id_Cliente_Proveedor"), 1))
                lintindice = lintindice + 1
                TraeCodigo = adoPrincipal.Fields("Rut") & "-" & Trim$(lintindice)
            Else
                TraeCodigo = adoPrincipal.Fields("Rut") & "-1"
            End If
        End If
        
        
    End If
    Conexion.CloseHost adoPrincipal
End Function

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtEMail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNombreContacto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNombreRazonSocial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPaginaWeb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSaldocondocumento_GotFocus()
    txtSaldocondocumento = SacarFormatoValor(txtSaldocondocumento, "$")
End Sub

Private Sub txtSaldocondocumento_KeyPress(KeyAscii As VB.ReturnInteger)
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSaldocondocumento_LostFocus()
    If Not IsNumeric(txtSaldocondocumento.Text) Then
        MsgBox "Debe Ingresar un Valor Numérico...", vbOKOnly + vbInformation, "Información"
        txtSaldocondocumento.Text = "0"
        txtSaldocondocumento.SetFocus
    Else
        txtSaldocondocumento = FormatoValor(txtSaldocondocumento, "$", 0)
    End If
End Sub

Private Sub txtSaldosindocumento_GotFocus()
    txtSaldosindocumento = SacarFormatoValor(txtSaldosindocumento, "$")
End Sub

Private Sub txtSaldosindocumento_KeyPress(KeyAscii As VB.ReturnInteger)
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSaldosindocumento_LostFocus()
    If Not IsNumeric(txtSaldosindocumento.Text) Then
        MsgBox "Debe Ingresar un Valor Numérico...", vbOKOnly + vbInformation, "Información"
        txtSaldosindocumento.Text = "0"
        txtSaldosindocumento.SetFocus
    Else
        txtSaldosindocumento = FormatoValor(txtSaldosindocumento, "$", 0)
    End If
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub
